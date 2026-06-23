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

import uuid
from datetime import date as date_type, datetime, timezone
from pathlib import Path

from fastapi import (
    APIRouter,
    Depends,
    File,
    Form,
    HTTPException,
    Response,
    UploadFile,
    status,
)
from sqlalchemy.orm import Session

import storage
from auth import get_current_user
from database import get_db
from models import Expense, User
from schemas import ExpenseBillOut, ExpenseDecideIn, ExpenseOut


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
    if decision not in ("approved", "rejected"):
        raise HTTPException(
            status.HTTP_400_BAD_REQUEST,
            "decision must be 'approved' or 'rejected'",
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
