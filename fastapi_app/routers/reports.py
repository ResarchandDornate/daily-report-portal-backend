import io
from datetime import date as date_type, datetime, timedelta, timezone

from fastapi import APIRouter, Depends, HTTPException, Query, Response, status
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from sqlalchemy import and_
from sqlalchemy.orm import Session

from auth import get_current_user
from database import get_db
from models import DailyReport, Department, User
from schemas import LeaveIn, ReportIn, ReportListOut, ReportOut

router = APIRouter(prefix="/api/reports", tags=["reports"])

# Excel sheet-name rules: max 31 chars, no  : \ / ? * [ ]
_INVALID_SHEET_CHARS = set(r":\/?*[]")


def _sanitize_sheet_name(name: str) -> str:
    cleaned = "".join("_" if c in _INVALID_SHEET_CHARS else c for c in (name or "Sheet"))
    cleaned = cleaned.strip() or "Sheet"
    return cleaned[:31]


def _unique_sheet_name(wb: Workbook, base: str) -> str:
    name = _sanitize_sheet_name(base)
    existing = set(wb.sheetnames)
    if name not in existing:
        return name
    # Append (2), (3), ... while keeping total <= 31 chars.
    n = 2
    while True:
        suffix = f" ({n})"
        candidate = name[: 31 - len(suffix)] + suffix
        if candidate not in existing:
            return candidate
        n += 1


def _summarise_field(values_with_dates):
    """Combine multiple per-day cells into ONE summary cell.

    `values_with_dates` is a list of `(date, raw_value)` tuples for the same
    employee + field across the date range.

    Heuristic:
      - Drop empty values and leave markers.
      - If every remaining value parses as a number (allowing ₹, commas,
        whitespace), return the numeric sum — int if it lands on a whole
        number, otherwise rounded to 2 dp.
      - Otherwise, concatenate `date: value` lines so HR can still see what
        was filled when.
    """
    pairs: list[tuple] = []
    for d, v in values_with_dates:
        s = ("" if v is None else str(v)).strip()
        if not s or s.lower() == "on leave":
            continue
        pairs.append((d, s))
    if not pairs:
        return ""

    nums: list[float] = []
    is_numeric = True
    for _, s in pairs:
        cleaned = s.replace("₹", "").replace("$", "").replace(",", "").replace(" ", "")
        try:
            nums.append(float(cleaned))
        except ValueError:
            is_numeric = False
            break

    if is_numeric:
        total = sum(nums)
        return int(total) if total == int(total) else round(total, 2)

    return "\n".join(
        f"{(d.isoformat() if d else '-')}: {s}" for d, s in pairs
    )


@router.get("", response_model=ReportListOut)
def list_reports(
    employee: int | None = Query(None, description="Filter by employee id"),
    department: str | None = Query(None, description="Filter by department slug"),
    start: date_type | None = Query(None, description="Inclusive start date"),
    end: date_type | None = Query(None, description="Inclusive end date"),
    limit: int = Query(1000, ge=1, le=5000, description="Page size"),
    offset: int = Query(0, ge=0, description="Rows to skip"),
    db: Session = Depends(get_db),
    _: User = Depends(get_current_user),
):
    q = db.query(DailyReport)
    if employee is not None:
        q = q.filter(DailyReport.user_id == employee)
    if department:
        dept = db.query(Department).filter(Department.slug == department).first()
        if not dept:
            return {"items": [], "total": 0, "limit": limit, "offset": offset}
        q = q.join(User, DailyReport.user_id == User.id).filter(User.department_id == dept.id)
    if start:
        q = q.filter(DailyReport.date >= start)
    if end:
        q = q.filter(DailyReport.date <= end)

    total = q.count()
    items = (
        q.order_by(DailyReport.date.desc(), DailyReport.user_id)
        .offset(offset)
        .limit(limit)
        .all()
    )
    return {"items": items, "total": total, "limit": limit, "offset": offset}


@router.get("/export.xlsx")
def export_xlsx(
    employee: int | None = Query(None, description="Filter by employee id"),
    department: str | None = Query(None, description="Filter by department slug"),
    start: date_type | None = Query(None, description="Inclusive start date"),
    end: date_type | None = Query(None, description="Inclusive end date"),
    db: Session = Depends(get_db),
    _: User = Depends(get_current_user),
):
    """Multi-sheet Excel export — one sheet per department.

    The same filters as `list_reports` apply, but pagination is replaced with
    a hard cap (50k rows) so HR can pull a full window in one click.  Each
    department gets its own sheet whose columns are Date / Employee / Title
    / Email + the department's own `report_fields`.  Leave days are written
    as-is (every field reads "On Leave"), matching what the UI shows.
    """
    q = db.query(DailyReport).join(User, DailyReport.user_id == User.id)
    if employee is not None:
        q = q.filter(DailyReport.user_id == employee)
    if department:
        dept = db.query(Department).filter(Department.slug == department).first()
        if not dept:
            return _empty_xlsx("daily-reports.xlsx")
        q = q.filter(User.department_id == dept.id)
    if start:
        q = q.filter(DailyReport.date >= start)
    if end:
        q = q.filter(DailyReport.date <= end)

    rows = (
        q.order_by(DailyReport.date.desc(), DailyReport.user_id)
        .limit(50000)
        .all()
    )

    # Group by department.  Reports for users with no department fall under
    # a synthetic "No Department" sheet so they aren't silently dropped.
    by_dept: dict[int | None, dict] = {}
    for r in rows:
        u = r.user
        d = u.department if u else None
        key = d.id if d else None
        if key not in by_dept:
            by_dept[key] = {
                "dept": d,
                "fields": list(d.report_fields) if d and d.report_fields else [],
                "reports": [],
            }
        by_dept[key]["reports"].append(r)

    wb = Workbook()
    # Drop the default empty sheet; we'll re-add a placeholder if there's nothing.
    wb.remove(wb.active)

    header_font = Font(bold=True, color="FFFFFF", size=11)
    header_fill = PatternFill("solid", fgColor="EA580C")  # orange-600
    header_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    # Data cells wrap long text and pin to the top — keeps columns at a sane
    # width and grows the row height instead of stretching off the screen.
    body_align = Alignment(horizontal="left", vertical="top", wrap_text=True)
    date_align = Alignment(horizontal="left", vertical="top")

    sorted_groups = sorted(
        by_dept.values(),
        key=lambda g: ((g["dept"].name if g["dept"] else "No Department")).lower(),
    )

    for group in sorted_groups:
        d = group["dept"]
        title = d.name if d else "No Department"
        ws = wb.create_sheet(title=_unique_sheet_name(wb, title))

        fields = group["fields"]
        headers = (
            ["Date range", "Employee"]
            + [(f.get("label") or f.get("key") or "") for f in fields]
            + ["Reports"]
        )
        ws.append(headers)
        for col_idx in range(1, len(headers) + 1):
            cell = ws.cell(row=1, column=col_idx)
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = header_align
        ws.row_dimensions[1].height = 28

        # Bucket all reports by employee so we can emit one summary row per
        # person rather than one row per (person, date).
        by_user: dict[int, list] = {}
        for r in group["reports"]:
            by_user.setdefault(r.user_id, []).append(r)

        def _user_sort_key(uid):
            u = next((r.user for r in group["reports"] if r.user_id == uid), None)
            if not u:
                return ("", "")
            return ((u.last_name or "").lower(), (u.first_name or "").lower())

        for uid in sorted(by_user.keys(), key=_user_sort_key):
            user_reports = sorted(
                by_user[uid], key=lambda r: r.date or date_type.min
            )
            u = user_reports[0].user
            full_name = ""
            if u:
                full_name = (
                    f"{(u.first_name or '').strip()} {(u.last_name or '').strip()}".strip()
                    or u.username
                )

            dates = [r.date for r in user_reports if r.date]
            if not dates:
                date_range = ""
            elif min(dates) == max(dates):
                date_range = min(dates).isoformat()
            else:
                date_range = f"{min(dates).isoformat()} → {max(dates).isoformat()}"

            row_vals = [date_range, full_name]
            for f in fields:
                key = f.get("key", "")
                pairs = [((r.date), (r.data or {}).get(key, "")) for r in user_reports]
                row_vals.append(_summarise_field(pairs))
            row_vals.append(len(user_reports))
            ws.append(row_vals)

            row_idx = ws.max_row
            for col_idx in range(1, len(row_vals) + 1):
                cell = ws.cell(row=row_idx, column=col_idx)
                cell.alignment = date_align if col_idx in (1, len(row_vals)) else body_align

        # Date range needs more room ("YYYY-MM-DD → YYYY-MM-DD"); per-field
        # cells are wider so summed/concatenated values aren't cramped; the
        # tail Reports column is tiny.
        widths = [24, 24] + [34] * len(fields) + [10]
        for i, w in enumerate(widths, start=1):
            ws.column_dimensions[get_column_letter(i)].width = w
        ws.freeze_panes = "C2"

    if not wb.sheetnames:
        ws = wb.create_sheet(title="No reports")
        ws["A1"] = "No reports match the current filters."

    buf = io.BytesIO()
    wb.save(buf)
    payload = buf.getvalue()

    # Filename — include the date range when supplied so HR can keep multiple
    # exports side-by-side without overwriting.
    parts = ["daily-reports"]
    if start:
        parts.append(start.isoformat())
    if end:
        parts.append(end.isoformat())
    if not (start or end):
        parts.append(date_type.today().isoformat())
    filename = "_".join(parts) + ".xlsx"

    return Response(
        content=payload,
        media_type=(
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        ),
        headers={
            "Content-Disposition": f'attachment; filename="{filename}"',
            "Content-Length": str(len(payload)),
        },
    )


def _empty_xlsx(filename: str) -> Response:
    wb = Workbook()
    ws = wb.active
    ws.title = "No reports"
    ws["A1"] = "No reports match the current filters."
    buf = io.BytesIO()
    wb.save(buf)
    data = buf.getvalue()
    return Response(
        content=data,
        media_type=(
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        ),
        headers={
            "Content-Disposition": f'attachment; filename="{filename}"',
            "Content-Length": str(len(data)),
        },
    )


@router.post("", response_model=ReportOut, status_code=status.HTTP_201_CREATED)
def upsert_report(
    payload: ReportIn,
    db: Session = Depends(get_db),
    user: User = Depends(get_current_user),
):
    """Create or update a daily report for (user, date).

    Regular employees can only submit/edit their own reports.  HR users may
    submit/edit on behalf of any employee by passing `user_id` in the body —
    this is how the dashboard's "Edit" feature works.
    """
    target_user_id = user.id
    if payload.user_id is not None and payload.user_id != user.id:
        # Confirm the target exists before we evaluate permission — it's the
        # same 404 either way.
        target = db.query(User).filter(User.id == payload.user_id).first()
        if not target:
            raise HTTPException(
                status.HTTP_404_NOT_FOUND,
                f"No user with id={payload.user_id}",
            )
        # Who can submit on behalf of whom:
        #   - HR can submit for anyone.
        #   - A "team head" can submit for any employee in their MANAGED
        #     department.  Managed dept is `team_head_dept_id` when set,
        #     otherwise falls back to the team head's own department.
        #   - Everyone else: forbidden.
        is_team_head = bool(getattr(user, "is_team_head", False))
        managed_dept_id = (
            getattr(user, "team_head_dept_id", None) or user.department_id
        )
        target_in_managed = (
            managed_dept_id is not None
            and target.department_id == managed_dept_id
        )
        if user.role != "hr" and not (is_team_head and target_in_managed):
            raise HTTPException(
                status.HTTP_403_FORBIDDEN,
                "You don't have permission to submit reports for this employee.",
            )
        target_user_id = target.id

    existing = db.query(DailyReport).filter(
        and_(DailyReport.user_id == target_user_id, DailyReport.date == payload.date)
    ).first()

    now = datetime.now(timezone.utc)
    if existing:
        # Employees can now edit their OWN past reports.  HR can edit anyone's
        # (the HR-on-behalf-of override earlier in this function is already
        # gated to HR-only, so non-HR callers can only ever touch their own
        # row).
        existing.data = payload.data
        existing.submitted_at = now
        report = existing
    else:
        report = DailyReport(
            user_id=target_user_id,
            date=payload.date,
            data=payload.data,
            submitted_at=now,
            created_at=now,
        )
        db.add(report)
    db.commit()
    db.refresh(report)
    return report


@router.post("/leave", response_model=list[ReportOut], status_code=status.HTTP_201_CREATED)
def apply_leave(
    payload: LeaveIn,
    db: Session = Depends(get_db),
    user: User = Depends(get_current_user),
):
    """Mark N consecutive days as leave for the current user (or any user, if HR).

    Each day becomes a DailyReport row where every department field is filled
    with "On Leave" — so all tables, summaries, and detail views render the
    leave state automatically.  Two hidden marker keys (`__leave__`,
    `__leave_reason__`) are also stored so callers can detect leave rows.
    """
    if payload.days < 1 or payload.days > 60:
        raise HTTPException(status.HTTP_400_BAD_REQUEST, "days must be between 1 and 60")

    target = user
    if payload.user_id is not None and payload.user_id != user.id:
        if user.role != "hr":
            raise HTTPException(status.HTTP_403_FORBIDDEN, "Only HR can apply leave for others")
        target = db.query(User).filter(User.id == payload.user_id).first()
        if not target:
            raise HTTPException(status.HTTP_404_NOT_FOUND, f"No user with id={payload.user_id}")

    fields = (target.department.report_fields or []) if target.department else []
    leave_payload: dict[str, str] = {f["key"]: "On Leave" for f in fields}
    leave_payload["__leave__"] = "1"
    leave_payload["__leave_reason__"] = payload.reason or ""

    now = datetime.now(timezone.utc)
    created: list[DailyReport] = []
    for i in range(payload.days):
        d = payload.start_date + timedelta(days=i)
        existing = (
            db.query(DailyReport)
            .filter(and_(DailyReport.user_id == target.id, DailyReport.date == d))
            .first()
        )
        if existing:
            existing.data = leave_payload
            existing.submitted_at = now
            created.append(existing)
        else:
            row = DailyReport(
                user_id=target.id,
                date=d,
                data=leave_payload,
                submitted_at=now,
                created_at=now,
            )
            db.add(row)
            created.append(row)
    db.commit()
    for r in created:
        db.refresh(r)
    return created


@router.get("/missing-today", response_model=list[int])
def missing_today(
    db: Session = Depends(get_db),
    _: User = Depends(get_current_user),
):
    """Returns the list of user ids that have NOT submitted a report today.

    Weekends (Sat / Sun) return an empty list — employees aren't expected to
    submit reports on non-working days, so nobody is "missing".
    """
    today = date_type.today()
    if today.weekday() >= 5:  # Mon=0, Fri=4, Sat=5, Sun=6
        return []
    submitted_ids = {r.user_id for r in db.query(DailyReport).filter(DailyReport.date == today).all()}
    all_active = db.query(User.id).filter(User.is_active.is_(True), User.role != "hr").all()
    return [uid for (uid,) in all_active if uid not in submitted_ids]


@router.delete("/{report_id}", status_code=status.HTTP_204_NO_CONTENT)
def delete_report(
    report_id: int,
    db: Session = Depends(get_db),
    user: User = Depends(get_current_user),
):
    report = db.query(DailyReport).filter(DailyReport.id == report_id).first()
    if not report:
        raise HTTPException(status.HTTP_404_NOT_FOUND, "Report not found")
    if report.user_id != user.id and user.role != "hr":
        raise HTTPException(status.HTTP_403_FORBIDDEN, "You can only delete your own reports")
    db.delete(report)
    db.commit()
    return None
