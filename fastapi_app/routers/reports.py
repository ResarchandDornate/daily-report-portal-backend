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


# ---------- Styled XLSX export ----------
#
# Match the look of the manual "All Department Employee Summary Report" file:
# 16-pt title row, grey subtitle, navy-blue header band with white text, light
# zebra fill on data, blue badge for the department column, bold employee
# names, em-dashes in place of empty values.  See _build_styled_sheet for the
# per-sheet layout.

_DASH = "—"

_TITLE_FONT = Font(bold=True, size=16, color="1A1A1A")
_SUBTITLE_FONT = Font(size=10, color="666666", italic=True)
_HEADER_FONT = Font(bold=True, size=10, color="FFFFFF")
_HEADER_FILL = PatternFill("solid", fgColor="1A2A3A")
_NAME_FONT = Font(bold=True, size=10, color="1A1A1A")
_BADGE_FONT = Font(bold=True, size=9, color="1B5E8B")
_BADGE_FILL = PatternFill("solid", fgColor="D6E8F5")
_BODY_FONT = Font(size=9, color="333333")
_MUTED_FONT = Font(size=9, color="888888")
_ROW_FILL = PatternFill("solid", fgColor="FAFAFA")
_TOTAL_FONT = Font(bold=True, size=10, color="1A1A1A")
_TOTAL_FILL = PatternFill("solid", fgColor="E8EEF5")
_LEFT_TOP_WRAP = Alignment(horizontal="left", vertical="top", wrap_text=True)
_CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)


def _summarise_user_reports(reports):
    """Build a (full_name, role, date_range, field_values, n_reports) tuple
    for one employee's set of daily reports.  `field_values` is a dict keyed
    by report-field key whose value is the per-cell summary (sum if numeric,
    date-prefixed concatenation otherwise).
    """
    user = reports[0].user
    full_name = ""
    role = ""
    if user:
        full_name = (
            f"{(user.first_name or '').strip()} {(user.last_name or '').strip()}".strip()
            or user.username
        )
        role = (user.title or "").strip()

    dates = sorted([r.date for r in reports if r.date])
    if not dates:
        date_range = _DASH
    elif dates[0] == dates[-1]:
        date_range = dates[0].isoformat()
    else:
        date_range = f"{dates[0].isoformat()} → {dates[-1].isoformat()}"

    dept = user.department if user else None
    fields = list(dept.report_fields) if dept and dept.report_fields else []
    field_values: dict[str, object] = {}
    for f in fields:
        key = f.get("key") or ""
        pairs = [(r.date, (r.data or {}).get(key, "")) for r in reports]
        field_values[key] = _summarise_field(pairs)

    return full_name, role, date_range, field_values, len(reports), fields


@router.get("/export.xlsx")
def export_xlsx(
    employee: int | None = Query(None, description="Filter by employee id"),
    department: str | None = Query(None, description="Filter by department slug"),
    start: date_type | None = Query(None, description="Inclusive start date"),
    end: date_type | None = Query(None, description="Inclusive end date"),
    db: Session = Depends(get_db),
    _: User = Depends(get_current_user),
):
    """Styled Excel export — Summary sheet first, one sheet per department.

    Matches HR's manual report layout: title row, subtitle, navy header band,
    bold employee names, department badge, light row fill, em-dashes for
    empty cells.  Numeric report fields are summed across the date range;
    text fields are concatenated as `date: value` lines.  Each employee
    appears on exactly one row per sheet.
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

    # Group by (department, user).  Skip leave-day rows from the rollups so
    # totals don't get contaminated by "On Leave" placeholders.
    by_dept: dict[int | None, dict] = {}
    for r in rows:
        if (r.data or {}).get("__leave__") == "1":
            continue
        u = r.user
        d = u.department if u else None
        key = d.id if d else None
        if key not in by_dept:
            by_dept[key] = {"dept": d, "users": {}}
        by_dept[key]["users"].setdefault(u.id if u else 0, []).append(r)

    wb = Workbook()
    wb.remove(wb.active)

    range_label = _format_range_label(start, end)
    today_label = date_type.today().strftime("%d %b %Y")
    total_reports = sum(
        len(rs) for grp in by_dept.values() for rs in grp["users"].values()
    )
    total_employees = sum(len(grp["users"]) for grp in by_dept.values())
    dept_count = len([g for g in by_dept.values() if g["dept"] is not None])

    # ----- Summary sheet (all departments combined) -----
    summary_ws = wb.create_sheet(title=_unique_sheet_name(wb, "Summary"))
    _build_summary_sheet(
        summary_ws,
        by_dept,
        range_label=range_label,
        today_label=today_label,
        total_reports=total_reports,
        total_employees=total_employees,
        dept_count=dept_count,
    )

    # ----- One sheet per department -----
    sorted_groups = sorted(
        by_dept.values(),
        key=lambda g: ((g["dept"].name if g["dept"] else "No Department")).lower(),
    )
    for group in sorted_groups:
        d = group["dept"]
        sheet_title = d.name if d else "No Department"
        ws = wb.create_sheet(title=_unique_sheet_name(wb, sheet_title))
        _build_department_sheet(ws, group, range_label=range_label)

    if not wb.sheetnames:
        ws = wb.create_sheet(title="No reports")
        ws["B2"] = "No reports match the current filters."
        ws["B2"].font = _SUBTITLE_FONT

    buf = io.BytesIO()
    wb.save(buf)
    payload = buf.getvalue()

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


def _format_range_label(start, end) -> str:
    if start and end:
        if start == end:
            return start.strftime("%d %b %Y")
        return f"{start.strftime('%d %b %Y')} → {end.strftime('%d %b %Y')}"
    if start:
        return f"from {start.strftime('%d %b %Y')}"
    if end:
        return f"until {end.strftime('%d %b %Y')}"
    return "all time"


def _put(ws, row, col, value, font=None, fill=None, align=None):
    cell = ws.cell(row=row, column=col, value=value)
    if font is not None:
        cell.font = font
    if fill is not None:
        cell.fill = fill
    if align is not None:
        cell.alignment = align
    return cell


def _merge_title(ws, row, last_col, value, font):
    ws.cell(row=row, column=2, value=value).font = font
    ws.cell(row=row, column=2).alignment = Alignment(horizontal="left", vertical="center")
    ws.merge_cells(start_row=row, start_column=2, end_row=row, end_column=last_col)


def _build_summary_sheet(
    ws,
    by_dept,
    *,
    range_label: str,
    today_label: str,
    total_reports: int,
    total_employees: int,
    dept_count: int,
):
    """Columns: A (margin) | B Employee | C Department | D Role | E Date range | F Reports."""
    last_col = 6
    ws.column_dimensions["A"].width = 3
    ws.column_dimensions["B"].width = 26
    ws.column_dimensions["C"].width = 22
    ws.column_dimensions["D"].width = 22
    ws.column_dimensions["E"].width = 26
    ws.column_dimensions["F"].width = 11

    _merge_title(ws, 1, last_col, f"All Department Employee Summary Report — {range_label}", _TITLE_FONT)
    ws.row_dimensions[1].height = 30
    subtitle = (
        f"Ornate Solar  |  {total_reports} reports  |  {total_employees} employees  "
        f"|  {dept_count} departments  |  Generated: {today_label}"
    )
    _merge_title(ws, 2, last_col, subtitle, _SUBTITLE_FONT)

    headers = ["Employee Name", "Department", "Role / Designation", "Date range", "Reports"]
    for i, h in enumerate(headers, start=2):
        _put(ws, 4, i, h, font=_HEADER_FONT, fill=_HEADER_FILL, align=_CENTER)
    ws.row_dimensions[4].height = 24

    # Sort: department name, then employee name.
    sorted_groups = sorted(
        by_dept.values(),
        key=lambda g: ((g["dept"].name if g["dept"] else "No Department")).lower(),
    )

    row_idx = 5
    for group in sorted_groups:
        d = group["dept"]
        dept_name = d.name if d else "No Department"
        # Sort users within department by name.
        user_entries = []
        for _, user_reports in group["users"].items():
            user_reports_sorted = sorted(user_reports, key=lambda r: r.date or date_type.min)
            name, role, date_range, _, count, _ = _summarise_user_reports(user_reports_sorted)
            user_entries.append((name, role, date_range, count))
        user_entries.sort(key=lambda t: t[0].lower())

        for idx, (name, role, date_range, count) in enumerate(user_entries):
            _put(ws, row_idx, 2, name, font=_NAME_FONT, fill=_ROW_FILL, align=_LEFT_TOP_WRAP)
            if idx == 0:
                _put(ws, row_idx, 3, dept_name, font=_BADGE_FONT, fill=_BADGE_FILL, align=_CENTER)
            else:
                _put(ws, row_idx, 3, "", fill=_ROW_FILL)
            _put(ws, row_idx, 4, role or _DASH,
                 font=_BODY_FONT if role else _MUTED_FONT, fill=_ROW_FILL, align=_LEFT_TOP_WRAP)
            _put(ws, row_idx, 5, date_range, font=_BODY_FONT, fill=_ROW_FILL, align=_LEFT_TOP_WRAP)
            _put(ws, row_idx, 6, count, font=_NAME_FONT, fill=_ROW_FILL, align=_CENTER)
            row_idx += 1

    ws.freeze_panes = "B5"


def _build_department_sheet(ws, group, *, range_label: str):
    """Columns: A (margin) | B Employee | C Role | D Date range | <fields…> | <Reports>."""
    d = group["dept"]
    dept_name = d.name if d else "No Department"
    fields = list(d.report_fields) if d and d.report_fields else []
    n_fields = len(fields)
    # Static cols: margin(A) + Employee(B) + Role(C) + Date range(D) + Reports(last)
    last_col = 4 + n_fields + 1  # +1 for Reports tail

    ws.column_dimensions["A"].width = 3
    ws.column_dimensions["B"].width = 24
    ws.column_dimensions["C"].width = 22
    ws.column_dimensions["D"].width = 24
    for i in range(n_fields):
        ws.column_dimensions[get_column_letter(5 + i)].width = 30
    ws.column_dimensions[get_column_letter(last_col)].width = 11

    _merge_title(
        ws, 1, last_col,
        f"{dept_name} — Employee Report ({range_label})",
        _TITLE_FONT,
    )
    ws.row_dimensions[1].height = 30
    n_users = len(group["users"])
    n_reports = sum(len(rs) for rs in group["users"].values())
    subtitle = f"{n_users} {'employee' if n_users == 1 else 'employees'}  ·  {n_reports} reports submitted"
    _merge_title(ws, 2, last_col, subtitle, _SUBTITLE_FONT)

    # Header row at row 4 (row 3 left as visual breathing space, like the ref).
    headers = ["Employee", "Role", "Date range"] + [
        (f.get("label") or f.get("key") or "") for f in fields
    ] + ["Reports"]
    for i, h in enumerate(headers, start=2):
        _put(ws, 4, i, h, font=_HEADER_FONT, fill=_HEADER_FILL, align=_CENTER)
    ws.row_dimensions[4].height = 26

    # Build employee rows (sorted by name).
    user_rows = []
    for _, user_reports in group["users"].items():
        sorted_reports = sorted(user_reports, key=lambda r: r.date or date_type.min)
        name, role, date_range, field_values, count, _ = _summarise_user_reports(sorted_reports)
        user_rows.append({
            "name": name,
            "role": role,
            "range": date_range,
            "field_values": field_values,
            "count": count,
        })
    user_rows.sort(key=lambda r: r["name"].lower())

    row_idx = 5
    for row in user_rows:
        _put(ws, row_idx, 2, row["name"], font=_NAME_FONT, fill=_ROW_FILL, align=_LEFT_TOP_WRAP)
        _put(ws, row_idx, 3, row["role"] or _DASH,
             font=_BODY_FONT if row["role"] else _MUTED_FONT, fill=_ROW_FILL, align=_LEFT_TOP_WRAP)
        _put(ws, row_idx, 4, row["range"], font=_BODY_FONT, fill=_ROW_FILL, align=_LEFT_TOP_WRAP)
        for fi, f in enumerate(fields):
            v = row["field_values"].get(f.get("key", ""), "")
            display = v if v not in ("", None) else _DASH
            _put(
                ws, row_idx, 5 + fi, display,
                font=_BODY_FONT if v not in ("", None) else _MUTED_FONT,
                fill=_ROW_FILL,
                align=_LEFT_TOP_WRAP,
            )
        _put(ws, row_idx, last_col, row["count"], font=_NAME_FONT, fill=_ROW_FILL, align=_CENTER)
        row_idx += 1

    # Department total row — sum numeric field columns, em-dash for text/empty.
    if user_rows:
        _put(ws, row_idx, 2, "Department Total", font=_TOTAL_FONT, fill=_TOTAL_FILL, align=_LEFT_TOP_WRAP)
        _put(ws, row_idx, 3,
             f"{len(user_rows)} {'employee' if len(user_rows) == 1 else 'employees'}",
             font=_TOTAL_FONT, fill=_TOTAL_FILL, align=_LEFT_TOP_WRAP)
        _put(ws, row_idx, 4, "", fill=_TOTAL_FILL)
        for fi, f in enumerate(fields):
            vals = [r["field_values"].get(f.get("key", ""), "") for r in user_rows]
            numeric = [v for v in vals if isinstance(v, (int, float))]
            if numeric and len(numeric) == len(vals):
                total = sum(numeric)
                display = int(total) if total == int(total) else round(total, 2)
            else:
                display = _DASH
            _put(
                ws, row_idx, 5 + fi, display,
                font=_TOTAL_FONT if isinstance(display, (int, float)) else _MUTED_FONT,
                fill=_TOTAL_FILL,
                align=_CENTER if isinstance(display, (int, float)) else _LEFT_TOP_WRAP,
            )
        _put(ws, row_idx, last_col, sum(r["count"] for r in user_rows),
             font=_TOTAL_FONT, fill=_TOTAL_FILL, align=_CENTER)

    ws.freeze_panes = "B5"


def _empty_xlsx(filename: str) -> Response:
    wb = Workbook()
    ws = wb.active
    ws.title = "No reports"
    ws["B2"] = "No reports match the current filters."
    ws["B2"].font = _SUBTITLE_FONT
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
