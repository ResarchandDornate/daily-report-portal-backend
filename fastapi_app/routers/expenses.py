"""Employee expense claims with approval workflow.

Endpoints:
  POST   /api/expenses              — employee submits a new expense (optionally
                                       with a bill image as multipart upload).
  GET    /api/expenses              — list expenses.  Employees see only their
                                       own; HR + named approvers (TARINI, SMITA)
                                       see ALL.
  GET    /api/expenses/{id}/bill    — stream the bill image bytes.  Same
                                       visibility rules.
  POST   /api/expenses/{id}/decide  — HR / approver approves or rejects.
  DELETE /api/expenses/{id}         — owner can withdraw a PENDING expense;
                                       HR can delete any.

Bills live in MinIO under `expenses/{uuid}.{ext}` in the existing bucket.
"""
from __future__ import annotations

import io
import uuid
from datetime import date as date_type, datetime, timedelta, timezone
from pathlib import Path

from fastapi import (
    APIRouter,
    Body,
    Depends,
    File,
    Form,
    HTTPException,
    Query,
    Response,
    UploadFile,
    status,
)
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from sqlalchemy.orm import Session

import storage
from auth import get_current_user
from database import get_db
from models import Expense, User
from schemas import ExpenseBillOut, ExpenseDecideIn, ExpenseOut, ExpensePatchIn


router = APIRouter(prefix="/api/expenses", tags=["expenses"])

# Hard-coded approver emails — TARINI + SMITA can approve / reject in addition
# to anyone with role=hr.  Stored uppercase for case-insensitive comparison.
APPROVER_EMAILS = {
    "tarini@ornatesolar.com",
    "smita@ornatesolar.com",
}

# Allowed expense types — server-side enum so the frontend can't smuggle in
# arbitrary values.  "travel" requires `travel_type` to also be one of the
# travel sub-types below.
ALLOWED_EXPENSE_TYPES = {
    # Legacy values kept for backward-compat with existing rows ("material",
    # "fuel") even though they're no longer in the dropdown.
    "material", "food", "travel", "hotel", "fuel", "others",
    "officereimburse", "sitematerial", "officematerial",
}
ALLOWED_TRAVEL_TYPES = {
    "bus", "cab", "bike", "rapido", "car", "auto", "metro", "other",
}
ALLOWED_MODES = {"cash", "upi", "card", "bank", "other", ""}

MAX_BILL_BYTES = 5 * 1024 * 1024   # per-file size cap
MAX_BILLS_PER_EXPENSE = 10          # upper bound on attachments per expense
ALLOWED_BILL_EXTENSIONS = {
    ".jpg", ".jpeg", ".png", ".webp", ".heic", ".pdf",
}
# Mime types accepted as the bill's Content-Type header from the browser.
ALLOWED_BILL_MIMETYPES = {
    "image/jpeg", "image/png", "image/webp", "image/heic", "image/heif",
    "application/pdf",
}


def _is_hr(user: User) -> bool:
    return user.role == "hr"


def _is_approver(user: User) -> bool:
    """HR users + the named approvers (TARINI / SMITA) can decide expenses."""
    if _is_hr(user):
        return True
    email = (user.email or "").strip().lower()
    return email in APPROVER_EMAILS


def _full_name(u: User | None) -> str:
    if not u:
        return ""
    first = (u.first_name or "").strip()
    last = (u.last_name or "").strip()
    return (first + " " + last).strip() or (u.username or "")


def _dept_name(u: User | None) -> str:
    if not u or not getattr(u, "department", None):
        return ""
    return (u.department.name or "")


def _to_out(exp: Expense) -> ExpenseOut:
    raw_bills = exp.bills or []
    bills_out = [
        ExpenseBillOut(index=i, filename=(b.get("filename") or f"bill-{i + 1}"))
        for i, b in enumerate(raw_bills)
    ]
    return ExpenseOut(
        id=exp.id,
        user_id=exp.user_id,
        user_name=_full_name(exp.user),
        user_department=_dept_name(exp.user),
        date=exp.date,
        mode=exp.mode or "",
        expense_type=exp.expense_type,
        travel_type=exp.travel_type or "",
        amount=exp.amount or 0,
        remarks=exp.remarks or "",
        bills=bills_out,
        status=exp.status,
        decided_by_id=exp.decided_by_id,
        decided_by_name=_full_name(exp.decided_by),
        decided_at=exp.decided_at,
        decision_note=exp.decision_note or "",
        created_at=exp.created_at,
    )


@router.post("", response_model=ExpenseOut, status_code=status.HTTP_201_CREATED)
async def create_expense(
    date: date_type = Form(...),
    mode: str = Form(""),
    expense_type: str = Form(...),
    travel_type: str = Form(""),
    amount: int = Form(...),
    remarks: str = Form(""),
    # Multi-file: the browser sends one `bills` form field per file.  Single
    # files still work (the list arrives with one element).  We also accept
    # the legacy `bill` field name for backward compat if any client still
    # uses it; that file is appended after `bills`.
    bills: list[UploadFile] = File(default=[]),
    bill: UploadFile | None = File(None),
    db: Session = Depends(get_db),
    user: User = Depends(get_current_user),
):
    """Employee creates a new expense.  Multiple bill files are supported."""
    expense_type = (expense_type or "").strip().lower()
    travel_type = (travel_type or "").strip().lower()
    mode = (mode or "").strip().lower()

    if expense_type not in ALLOWED_EXPENSE_TYPES:
        raise HTTPException(
            status.HTTP_400_BAD_REQUEST,
            f"Invalid expense_type. Allowed: {sorted(ALLOWED_EXPENSE_TYPES)}",
        )
    if expense_type == "travel":
        if not travel_type:
            raise HTTPException(
                status.HTTP_400_BAD_REQUEST,
                "travel_type is required when expense_type is 'travel'",
            )
        if travel_type not in ALLOWED_TRAVEL_TYPES:
            raise HTTPException(
                status.HTTP_400_BAD_REQUEST,
                f"Invalid travel_type. Allowed: {sorted(ALLOWED_TRAVEL_TYPES)}",
            )
    else:
        travel_type = ""
    if mode and mode not in ALLOWED_MODES:
        raise HTTPException(
            status.HTTP_400_BAD_REQUEST,
            f"Invalid mode. Allowed: {sorted(ALLOWED_MODES - {''})}",
        )
    if amount is None or amount < 0:
        raise HTTPException(
            status.HTTP_400_BAD_REQUEST, "amount must be a non-negative integer"
        )

    # Combine new + legacy bill fields, drop any "empty" UploadFile entries
    # (browsers sometimes send an empty file with no filename when the input
    # was left untouched).
    incoming: list[UploadFile] = []
    for f in (bills or []):
        if f and f.filename:
            incoming.append(f)
    if bill and bill.filename:
        incoming.append(bill)
    if len(incoming) > MAX_BILLS_PER_EXPENSE:
        raise HTTPException(
            status.HTTP_400_BAD_REQUEST,
            f"At most {MAX_BILLS_PER_EXPENSE} bills can be attached to one expense.",
        )

    stored_bills: list[dict] = []
    for f in incoming:
        suffix = Path(f.filename).suffix.lower()
        if suffix not in ALLOWED_BILL_EXTENSIONS:
            raise HTTPException(
                status.HTTP_400_BAD_REQUEST,
                f"Bill {f.filename!r}: must be one of {sorted(ALLOWED_BILL_EXTENSIONS)}",
            )
        content_type = (f.content_type or "").lower()
        if content_type and content_type not in ALLOWED_BILL_MIMETYPES:
            raise HTTPException(
                status.HTTP_400_BAD_REQUEST,
                f"Bill {f.filename!r}: content-type {content_type!r} not allowed.",
            )
        data = await f.read()
        if len(data) > MAX_BILL_BYTES:
            raise HTTPException(
                status.HTTP_400_BAD_REQUEST,
                f"Bill {f.filename!r}: exceeds {MAX_BILL_BYTES // (1024 * 1024)} MB limit.",
            )
        object_key = f"expenses/{uuid.uuid4().hex}{suffix}"
        try:
            storage.put_object(
                object_key,
                data,
                content_type=content_type or "application/octet-stream",
            )
        except Exception as e:
            # Best-effort cleanup of any earlier bills we already stored so
            # we don't leave orphans on failure mid-upload.
            for prior in stored_bills:
                try:
                    storage.delete_object(prior["object_key"])
                except Exception:
                    pass
            raise HTTPException(
                status.HTTP_500_INTERNAL_SERVER_ERROR,
                f"Could not store bill {f.filename!r}: {e}",
            )
        stored_bills.append({"filename": f.filename, "object_key": object_key})

    exp = Expense(
        user_id=user.id,
        date=date,
        mode=mode,
        expense_type=expense_type,
        travel_type=travel_type,
        amount=amount,
        remarks=(remarks or "").strip(),
        bills=stored_bills,
        status="pending",
    )
    db.add(exp)
    db.commit()
    db.refresh(exp)
    return _to_out(exp)


@router.get("", response_model=list[ExpenseOut])
def list_expenses(
    db: Session = Depends(get_db),
    user: User = Depends(get_current_user),
):
    """List expenses visible to the caller.

    Employees see ONLY their own expenses.  HR + named approvers see ALL
    expenses across the org (newest first).
    """
    q = db.query(Expense).join(User, Expense.user_id == User.id)
    if not _is_approver(user):
        q = q.filter(Expense.user_id == user.id)
    rows = q.order_by(Expense.created_at.desc()).limit(2000).all()
    return [_to_out(r) for r in rows]


def _serve_bill(exp: Expense, bill: dict) -> Response:
    """Stream a single bill from MinIO with the right Content-Type."""
    object_key = bill.get("object_key") or ""
    filename = bill.get("filename") or ""
    if not object_key:
        raise HTTPException(status.HTTP_404_NOT_FOUND, "Bill is missing its storage key")
    try:
        data = storage.get_object_bytes(object_key)
    except Exception as e:
        raise HTTPException(
            status.HTTP_500_INTERNAL_SERVER_ERROR, f"Could not fetch bill: {e}"
        )
    suffix = Path(filename or object_key).suffix.lower()
    mime = {
        ".jpg":  "image/jpeg",
        ".jpeg": "image/jpeg",
        ".png":  "image/png",
        ".webp": "image/webp",
        ".heic": "image/heic",
        ".pdf":  "application/pdf",
    }.get(suffix, "application/octet-stream")
    headers = {
        "Content-Disposition": (
            f'inline; filename="{filename or ("bill" + suffix)}"'
        ),
    }
    return Response(content=data, media_type=mime, headers=headers)


@router.get("/{expense_id}/bill/{index}")
def download_bill_at(
    expense_id: int,
    index: int,
    db: Session = Depends(get_db),
    user: User = Depends(get_current_user),
):
    """Stream a specific bill (by position in the `bills` list)."""
    exp = db.query(Expense).filter(Expense.id == expense_id).first()
    if not exp:
        raise HTTPException(status.HTTP_404_NOT_FOUND, "Expense not found")
    if exp.user_id != user.id and not _is_approver(user):
        raise HTTPException(
            status.HTTP_403_FORBIDDEN, "You can't view this expense's bills"
        )
    raw_bills = exp.bills or []
    if index < 0 or index >= len(raw_bills):
        raise HTTPException(status.HTTP_404_NOT_FOUND, "Bill index out of range")
    return _serve_bill(exp, raw_bills[index])


@router.get("/{expense_id}/bill")
def download_bill_legacy(
    expense_id: int,
    db: Session = Depends(get_db),
    user: User = Depends(get_current_user),
):
    """Backward-compat: serve the FIRST bill when no index is given."""
    exp = db.query(Expense).filter(Expense.id == expense_id).first()
    if not exp:
        raise HTTPException(status.HTTP_404_NOT_FOUND, "Expense not found")
    if exp.user_id != user.id and not _is_approver(user):
        raise HTTPException(
            status.HTTP_403_FORBIDDEN, "You can't view this expense's bills"
        )
    raw_bills = exp.bills or []
    if not raw_bills:
        raise HTTPException(status.HTTP_404_NOT_FOUND, "No bill attached")
    return _serve_bill(exp, raw_bills[0])


@router.post("/{expense_id}/decide", response_model=ExpenseOut)
def decide_expense(
    expense_id: int,
    payload: ExpenseDecideIn,
    db: Session = Depends(get_db),
    user: User = Depends(get_current_user),
):
    """Approve or reject an expense.  HR + named approvers only."""
    if not _is_approver(user):
        raise HTTPException(
            status.HTTP_403_FORBIDDEN,
            "Only HR / approvers (TARINI, SMITA) can decide expenses.",
        )
    decision = (payload.decision or "").strip().lower()
    if decision not in ("approved", "rejected", "onhold"):
        raise HTTPException(
            status.HTTP_400_BAD_REQUEST,
            "decision must be 'approved', 'rejected', or 'onhold'",
        )
    exp = db.query(Expense).filter(Expense.id == expense_id).first()
    if not exp:
        raise HTTPException(status.HTTP_404_NOT_FOUND, "Expense not found")
    exp.status = decision
    exp.decided_by_id = user.id
    exp.decided_at = datetime.now(timezone.utc)
    exp.decision_note = (payload.note or "").strip()
    db.commit()
    db.refresh(exp)
    return _to_out(exp)


@router.patch("/{expense_id}", response_model=ExpenseOut)
def update_expense(
    expense_id: int,
    payload: ExpensePatchIn,
    db: Session = Depends(get_db),
    user: User = Depends(get_current_user),
):
    """Owner edits their own expense.  Only allowed while the expense is
    still pending or on hold; once approved / rejected the row is locked.

    Bills attached to the expense are NOT touched here — bill add/remove
    flows go through the upload + delete endpoints.
    """
    exp = db.query(Expense).filter(Expense.id == expense_id).first()
    if not exp:
        raise HTTPException(status.HTTP_404_NOT_FOUND, "Expense not found")
    if exp.user_id != user.id and not _is_hr(user):
        raise HTTPException(
            status.HTTP_403_FORBIDDEN,
            "You can only edit your own expenses (HR can edit any).",
        )
    if exp.status not in ("pending", "onhold"):
        raise HTTPException(
            status.HTTP_409_CONFLICT,
            "This expense has already been decided and can no longer be edited.",
        )

    # Apply each provided field with the same validation rules as create.
    if payload.date is not None:
        exp.date = payload.date
    if payload.mode is not None:
        m = (payload.mode or "").strip().lower()
        if m and m not in ALLOWED_MODES:
            raise HTTPException(
                status.HTTP_400_BAD_REQUEST,
                f"Invalid mode. Allowed: {sorted(ALLOWED_MODES - {''})}",
            )
        exp.mode = m
    if payload.expense_type is not None:
        et = (payload.expense_type or "").strip().lower()
        if et not in ALLOWED_EXPENSE_TYPES:
            raise HTTPException(
                status.HTTP_400_BAD_REQUEST,
                f"Invalid expense_type. Allowed: {sorted(ALLOWED_EXPENSE_TYPES)}",
            )
        exp.expense_type = et
    if payload.travel_type is not None:
        tt = (payload.travel_type or "").strip().lower()
        if tt and tt not in ALLOWED_TRAVEL_TYPES:
            raise HTTPException(
                status.HTTP_400_BAD_REQUEST,
                f"Invalid travel_type. Allowed: {sorted(ALLOWED_TRAVEL_TYPES)}",
            )
        exp.travel_type = tt
    if payload.amount is not None:
        if payload.amount < 0:
            raise HTTPException(
                status.HTTP_400_BAD_REQUEST,
                "amount must be a non-negative integer",
            )
        exp.amount = payload.amount
    if payload.remarks is not None:
        exp.remarks = (payload.remarks or "").strip()

    # Editing an on-hold expense bumps it back to pending so the approver
    # sees a fresh request to review.  Reset the prior decision metadata too.
    if exp.status == "onhold":
        exp.status = "pending"
        exp.decided_by_id = None
        exp.decided_at = None
        exp.decision_note = ""

    db.commit()
    db.refresh(exp)
    return _to_out(exp)


@router.delete("/{expense_id}", status_code=status.HTTP_204_NO_CONTENT)
def delete_expense(
    expense_id: int,
    db: Session = Depends(get_db),
    user: User = Depends(get_current_user),
):
    """Owner can withdraw a PENDING expense; HR can delete any."""
    exp = db.query(Expense).filter(Expense.id == expense_id).first()
    if not exp:
        raise HTTPException(status.HTTP_404_NOT_FOUND, "Expense not found")
    is_owner = exp.user_id == user.id
    if not (_is_hr(user) or (is_owner and exp.status == "pending")):
        raise HTTPException(
            status.HTTP_403_FORBIDDEN,
            "Only the owner (while pending) or HR can delete this expense.",
        )
    # Best-effort cleanup of every attached bill — swallow errors so the
    # row still deletes even if MinIO is briefly unreachable.
    for bill in (exp.bills or []):
        key = bill.get("object_key") or ""
        if key:
            try:
                storage.delete_object(key)
            except Exception:
                pass
    db.delete(exp)
    db.commit()
    return Response(status_code=status.HTTP_204_NO_CONTENT)


# ============================================================
# Monthly summary Excel export
# ============================================================

def _month_bounds(year: int, month: int) -> tuple[date_type, date_type]:
    if month < 1 or month > 12 or year < 2000 or year > 2100:
        raise HTTPException(
            status.HTTP_400_BAD_REQUEST,
            f"Invalid year/month: {year}/{month}",
        )
    first = date_type(year, month, 1)
    last = (
        date_type(year + 1, 1, 1) if month == 12 else date_type(year, month + 1, 1)
    ) - timedelta(days=1)
    return first, last


@router.post("/monthly-summary.xlsx")
def monthly_summary_xlsx(
    payload: dict = Body(...),
    db: Session = Depends(get_db),
    user: User = Depends(get_current_user),
):
    """One-tab Excel of monthly expense summary per employee.

    Body shape:
        {
          "year": 2026,
          "month": 6,                  # 1..12
          "advances": { "<user_id>": 500, ... }   # optional advance per emp
        }

    Columns:
        Employee | Total Amount | Status | Advance | Subtotal | Submit Date

    Footer row: Total Amount  ·  Total Advance  ·  Subtotal.

    Restricted to HR + named approvers (Tarini / Smita).
    """
    if not _is_approver(user):
        raise HTTPException(
            status.HTTP_403_FORBIDDEN,
            "Only HR / approvers can download the monthly summary.",
        )
    year = int(payload.get("year") or 0)
    month = int(payload.get("month") or 0)
    advances_in = payload.get("advances") or {}
    if not isinstance(advances_in, dict):
        raise HTTPException(
            status.HTTP_400_BAD_REQUEST, "`advances` must be an object"
        )
    advances: dict[int, int] = {}
    for k, v in advances_in.items():
        try:
            advances[int(k)] = int(float(v or 0))
        except (TypeError, ValueError):
            continue

    first, last = _month_bounds(year, month)

    # All expenses in the month (any status).
    rows = (
        db.query(Expense)
        .join(User, Expense.user_id == User.id)
        .filter(Expense.date >= first, Expense.date <= last)
        .order_by(User.first_name, Expense.created_at)
        .all()
    )

    # Group by user_id.
    by_user: dict[int, dict] = {}
    for r in rows:
        u = r.user
        full_name = _full_name(u) or "—"
        g = by_user.setdefault(r.user_id, {
            "name": full_name,
            "total": 0,
            "pending": 0,
            "approved": 0,
            "rejected": 0,
            "onhold": 0,
            "latest_submit": r.created_at,
        })
        g["total"] += (r.amount or 0)
        if r.status in g:
            g[r.status] += 1
        if r.created_at and (g["latest_submit"] is None or r.created_at > g["latest_submit"]):
            g["latest_submit"] = r.created_at

    # ---- Build the workbook ----
    wb = Workbook()
    ws = wb.active
    month_label = first.strftime("%B %Y")  # e.g. "June 2026"
    ws.title = f"Expenses {first.strftime('%b %Y')}"[:31]

    title_font = Font(bold=True, size=14, color="1B5E8B")
    subtitle_font = Font(size=10, color="666666", italic=True)
    header_font = Font(bold=True, size=10, color="FFFFFF")
    header_fill = PatternFill("solid", fgColor="1A2A3A")
    total_font = Font(bold=True, size=10, color="1B5E8B")
    total_fill = PatternFill("solid", fgColor="EAF1F8")
    body_font = Font(size=10, color="1A1A1A")
    border = Border(*[Side(style="thin", color="C9D2DC")] * 4)
    center = Alignment(horizontal="center", vertical="center")
    right = Alignment(horizontal="right", vertical="center")
    left = Alignment(horizontal="left", vertical="center")

    ws["B1"] = f"Monthly Expense Summary — {month_label}"
    ws["B1"].font = title_font
    ws.merge_cells("B1:G1")
    ws["B2"] = (
        f"{len(by_user)} employee{'s' if len(by_user) != 1 else ''}"
        f"  ·  {len(rows)} expense{'s' if len(rows) != 1 else ''}"
        f"  ·  Generated: {date_type.today().strftime('%d %b %Y')}"
    )
    ws["B2"].font = subtitle_font
    ws.merge_cells("B2:G2")

    headers = ["Employee", "Total Amount", "Status", "Advance", "Subtotal", "Submit Date"]
    widths = [28, 18, 22, 14, 16, 18]
    ws.column_dimensions["A"].width = 3
    for i, (h, w) in enumerate(zip(headers, widths), start=2):
        c = ws.cell(row=4, column=i, value=h)
        c.font = header_font
        c.fill = header_fill
        c.alignment = center
        c.border = border
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.row_dimensions[4].height = 24

    sorted_groups = sorted(by_user.items(), key=lambda kv: kv[1]["name"].lower())

    total_amount_sum = 0
    total_advance_sum = 0
    total_subtotal_sum = 0
    row_idx = 5
    for uid, g in sorted_groups:
        amt = g["total"]
        adv = advances.get(uid, 0)
        sub = amt - adv
        total_amount_sum += amt
        total_advance_sum += adv
        total_subtotal_sum += sub
        status_bits = []
        if g["pending"]:  status_bits.append(f"Pending: {g['pending']}")
        if g["onhold"]:   status_bits.append(f"On Hold: {g['onhold']}")
        if g["approved"]: status_bits.append(f"Approved: {g['approved']}")
        if g["rejected"]: status_bits.append(f"Rejected: {g['rejected']}")
        status_text = ", ".join(status_bits) or "—"
        submit_str = (
            g["latest_submit"].strftime("%d %b %Y")
            if g["latest_submit"] else "—"
        )

        cells = [
            (g["name"], left),
            (f"₹{amt:,.0f}", right),
            (status_text, left),
            (f"₹{adv:,.0f}", right),
            (f"₹{sub:,.0f}", right),
            (submit_str, center),
        ]
        for col_offset, (val, align) in enumerate(cells):
            c = ws.cell(row=row_idx, column=2 + col_offset, value=val)
            c.font = body_font
            c.alignment = align
            c.border = border
        row_idx += 1

    # Footer totals row.
    if sorted_groups:
        ws.cell(row=row_idx, column=2, value="Total")
        ws.cell(row=row_idx, column=3, value=f"₹{total_amount_sum:,.0f}")
        ws.cell(row=row_idx, column=4, value="")
        ws.cell(row=row_idx, column=5, value=f"₹{total_advance_sum:,.0f}")
        ws.cell(row=row_idx, column=6, value=f"₹{total_subtotal_sum:,.0f}")
        ws.cell(row=row_idx, column=7, value="")
        for col in range(2, 8):
            c = ws.cell(row=row_idx, column=col)
            c.font = total_font
            c.fill = total_fill
            c.border = border
            c.alignment = right if col in (3, 5, 6) else left

    # Stream the file out.
    buf = io.BytesIO()
    wb.save(buf)
    payload_bytes = buf.getvalue()
    filename = f"monthly-expenses-{first.strftime('%Y-%m')}.xlsx"
    return Response(
        content=payload_bytes,
        media_type=(
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        ),
        headers={
            "Content-Disposition": f'attachment; filename="{filename}"',
            "Content-Length": str(len(payload_bytes)),
        },
    )
