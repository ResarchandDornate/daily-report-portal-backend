import csv
import io
import re
from datetime import date as date_type, datetime, timedelta, timezone

from fastapi import APIRouter, Depends, HTTPException, Query, Response, status
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.page import PageMargins
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

# Colour palette mirrors the on-screen "summary table" component so the
# downloaded XLSX is visually identical to the dashboard view.
_TITLE_FONT = Font(bold=True, size=16, color="1B5E8B")          # blue title
_SUBTITLE_FONT = Font(size=10, color="666666", italic=True)
_HEADER_FONT = Font(bold=True, size=10, color="FFFFFF")
_HEADER_FILL = PatternFill("solid", fgColor="1A2A3A")           # navy header
_NAME_FONT = Font(size=10, color="1A1A1A")                       # employee names
_BADGE_FONT = Font(bold=True, size=9, color="1B5E8B")
_BADGE_FILL = PatternFill("solid", fgColor="D6E8F5")
_BODY_FONT = Font(size=10, color="1A1A1A")
_MUTED_FONT = Font(size=10, color="888888")
_ROW_FILL = PatternFill("solid", fgColor="FFFFFF")              # white rows
_TOTAL_FONT = Font(bold=True, size=10, color="1B5E8B")          # navy bold totals
_TOTAL_FILL = PatternFill("solid", fgColor="EAF1F8")            # subtle blue total bar
_LEFT_TOP_WRAP = Alignment(horizontal="left", vertical="top", wrap_text=True)
_CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)


# Field-label classifiers: which column to treat as "Meetings" and which as
# "Revenue" in the new compact report layout (mirrors the manual HR report).
_MEETING_LABEL_RE = re.compile(r"\b(meeting|call|visit|enquir)", re.I)
_REVENUE_LABEL_RE = re.compile(
    r"\b(revenue|amount|invoice|sales|rupees|payment|earning|₹)", re.I
)
# Broad numeric-label hint — any column whose label reads like a counter or
# amount.  Used by the per-dept detail sheet to decide which columns to sum.
_NUMERIC_LABEL_RE = re.compile(
    r"\b(no\.?|number|count|total|amount|revenue|calls?|meetings?|enquiries|"
    r"leads?|companies|visits?|orders?|hours?|sum|rate|₹|rs\.?|inr|amt|pis?|"
    r"picked|closed|lost|following|invoice|sent|shared|received|made|type|works?)\b",
    re.I,
)
# Number token: handles thousand-separator commas, optional decimals, and
# detects trailing ordinal suffix so "25th" / "1st" can be skipped.
_NUMBER_TOKEN_RE = re.compile(
    r"(\d{1,3}(?:,\d{3})+(?:\.\d+)?|\d+(?:\.\d+)?)(?:\s*(st|nd|rd|th)\b)?",
    re.I,
)


def _extract_numbers_text(s):
    """Pull every number out of a free-text cell.  Handles thousand-separator
    commas, skips year-looking integers (1900-2100), skips ordinal dates
    ("25th", "1st").  So "Monthly Sales- 1,287.32" → [1287.32], and
    "4, 2, 25th May 2026, 3" → [4, 2, 3].
    """
    if s is None:
        return []
    out = []
    for m in _NUMBER_TOKEN_RE.finditer(str(s)):
        if m.group(2):
            continue  # ordinal suffix — skip
        try:
            n = float(m.group(1).replace(",", ""))
        except ValueError:
            continue
        if n.is_integer() and 1900 <= int(n) <= 2100:
            continue  # year-like — skip
        out.append(n)
    return out


# Department display order: Sales first, Inside Sales second, then every
# other department alphabetically.  Applied to both the Summary sheet and
# the per-department sheet creation order so HR sees the same priority in
# both surfaces.
_DEPT_PRIORITY = {"sales": 0, "insideSales": 1}


def _dept_sort_key(group):
    d = group["dept"]
    slug = getattr(d, "slug", "") if d else ""
    name = (d.name if d else "No Department").lower()
    return (_DEPT_PRIORITY.get(slug, 99), name)


def _condense_summary(text, *, max_chars: int = 200) -> str:
    """Shrink a long, multi-paragraph block into a short, readable line.

    Produces clean prose with NO truncation markers — no "…", no
    "(+N more)".  Each surviving phrase is whole; if a phrase doesn't fit
    the budget, it's dropped entirely so the cell stays clean.

    Strategy (deterministic, no LLM):
      1. Split on newlines / periods / semicolons / bullets so each
         "thought" becomes its own fragment.
      2. Strip ISO date prefixes (`2026-05-04:`) and leading list numbers
         (`1. `, `2) `) so fragments compare cleanly.
      3. Dedupe by **keyword overlap** — two fragments are considered
         the same when ≥ 50% of their content words match.  Catches near
         duplicates like "Visit Kamala site" and "visit kamala site again".
      4. For each surviving fragment, keep only the first 5–6 words so it
         reads like a tight headline ("Module installation at Mavikalan").
      5. Pack the headlines into `max_chars`, comma-joined.  Stop cleanly
         when the next phrase would overflow — no trailing marker.
    """
    if not text:
        return ""
    raw = re.split(r"[\n.;•]+|(?<=\s)[-*]\s+", str(text))
    seen_keywords: list[frozenset[str]] = []
    headlines: list[str] = []
    for frag in raw:
        s = frag.strip()
        if not s:
            continue
        s = re.sub(r"^\d{4}-\d{2}-\d{2}\s*[:\-]?\s*", "", s)
        s = re.sub(r"^\d+[.\)]\s*", "", s)
        s = s.strip(" ,")
        if not s:
            continue
        # Keyword set = words ≥ 3 chars, lowercased — used for dedupe.
        kw = frozenset(re.findall(r"[a-z]{3,}", s.lower()))
        if not kw:
            continue
        is_dup = False
        for prev in seen_keywords:
            shared = len(kw & prev)
            if shared and shared >= max(2, min(len(kw), len(prev)) // 2):
                is_dup = True
                break
        if is_dup:
            continue
        seen_keywords.append(kw)
        # Keep only the first 6 words so each headline is tight.
        words = s.split()
        head = " ".join(words[:6]) if len(words) > 6 else s
        headlines.append(head)
    return _compact_join(headlines, max_chars=max_chars)


def _compact_join(values, *, max_chars: int = 120) -> str:
    """Pack whole phrases (comma-joined) into `max_chars`, stopping
    cleanly when the next phrase wouldn't fit.  NO "…", NO "(+N more)".
    """
    if not values:
        return ""
    out = ""
    for v in values:
        s = str(v or "").strip()
        if not s:
            continue
        candidate = s if not out else f"{out}, {s}"
        if len(candidate) > max_chars:
            # Stop — but if we haven't taken anything yet, take one whole
            # phrase even if it slightly exceeds the budget (so the cell
            # isn't empty).
            if not out:
                return s
            break
        out = candidate
    return out


def _format_revenue(n):
    """Render a revenue value as `₹X,XXX,XXX` (no decimal if whole)."""
    if n is None:
        return _DASH
    if isinstance(n, float) and not n.is_integer():
        return f"₹{n:,.2f}"
    return f"₹{int(n):,}"


def _summarise_user_reports(reports):
    """Condense one employee's set of daily reports into a single summary row:

      (full_name, role, key_activities, meetings_total, revenue_total, n_reports)

    `key_activities` is a multi-line "Field Label: distinct values…" string
    that consolidates every text/free-form field for the date range.  Fields
    whose label looks like a meeting/call counter feed the `meetings_total`
    bucket; fields that look like revenue feed `revenue_total`.  Both
    numbers are `None` when no values were filled.
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

    dept = user.department if user else None
    fields = list(dept.report_fields) if dept and dept.report_fields else []

    meetings_total = None
    revenue_total = None
    # Pool every non-empty text value from every text/non-counter field into
    # one big buffer, then run a single `_condense_summary` pass over it.
    # Result: Key Activities reads as ONE unified summary of what the
    # employee actually did — no "Field Name: …" prefixes.  Meeting and
    # revenue columns are still routed into their own numeric buckets
    # (Meetings / Revenue) so they don't show up twice.
    pooled_text_parts: list[str] = []

    for f in fields:
        label = (f.get("label") or f.get("key") or "").strip()
        key = f.get("key") or ""
        non_empty = []
        for r in reports:
            v = (r.data or {}).get(key, "")
            s = ("" if v is None else str(v)).strip()
            if not s or s.lower() == "on leave":
                continue
            non_empty.append((r.date, s))
        if not non_empty:
            continue

        if _MEETING_LABEL_RE.search(label) and meetings_total is None:
            total = sum(n for _, v in non_empty for n in _extract_numbers_text(v))
            meetings_total = total if total > 0 else None
            continue
        if _REVENUE_LABEL_RE.search(label) and revenue_total is None:
            total = sum(n for _, v in non_empty for n in _extract_numbers_text(v))
            revenue_total = total if total > 0 else None
            continue

        # Pure-numeric counter fields contribute their numbers to the pool as
        # a short phrase, but don't bloat the narrative with raw digits.
        # Text/narrative fields are added verbatim.  `_condense_summary`
        # below dedupes near-duplicates by keyword overlap so we can safely
        # dump everything in here.
        for _, v in non_empty:
            pooled_text_parts.append(v)

    big_block = "\n".join(pooled_text_parts)
    key_activities = _condense_summary(big_block, max_chars=220)
    return full_name, role, key_activities, meetings_total, revenue_total, len(reports)


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

    # Three-sheet layout — matches HR's manual Numbers report:
    #   1. Sales — Detail        (per-field columns, only Sales employees)
    #   2. Inside Sales — Detail (per-field columns, only Inside Sales)
    #   3. Detailed Summary      (all OTHER departments, one row per
    #                             employee with consolidated Key Activities)
    sales_group = None
    inside_sales_group = None
    service_group = None
    project_group = None
    other_groups = []
    for group in by_dept.values():
        slug = getattr(group["dept"], "slug", "") if group["dept"] else ""
        if slug == "sales":
            sales_group = group
        elif slug == "insideSales":
            inside_sales_group = group
        elif slug == "service":
            service_group = group
        elif slug == "project":
            project_group = group
        else:
            other_groups.append(group)

    if sales_group:
        ws = wb.create_sheet(title=_unique_sheet_name(wb, "Sales — Detail"))
        _build_dept_detail_sheet(ws, sales_group, range_label=range_label)

    if inside_sales_group:
        ws = wb.create_sheet(title=_unique_sheet_name(wb, "Inside Sales — Detail"))
        # Drop the columns HR doesn't want on the Inside Sales tab.
        _build_dept_detail_sheet(
            ws, inside_sales_group, range_label=range_label,
            exclude_keys=("dataCalledType", "mailSent", "whatsappSent", "otherWorks"),
        )

    if service_group:
        ws = wb.create_sheet(title=_unique_sheet_name(wb, "Service — Detail"))
        _build_service_detail_sheet(ws, service_group, range_label=range_label)

    if project_group:
        ws = wb.create_sheet(title=_unique_sheet_name(wb, "Project — Detail"))
        _build_project_detail_sheet(ws, project_group, range_label=range_label)

    if other_groups:
        ws = wb.create_sheet(title=_unique_sheet_name(wb, "Detailed Summary"))
        _build_combined_summary_sheet(
            ws, other_groups,
            range_label=range_label,
            today_label=today_label,
        )

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


def _set_print_landscape_fit(ws):
    """Configure the sheet so Excel's print/PDF lands on a single landscape
    page wide and lets rows flow to as many pages as needed."""
    ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_margins = PageMargins(left=0.4, right=0.4, top=0.5, bottom=0.5)
    ws.print_options.horizontalCentered = True


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
    """All-employees summary — 6 columns matching HR's manual report:

    A (margin) | B Employee Name | C Department | D Role | E Key Activities | F Meetings | G Revenue (₹)
    """
    last_col = 7
    ws.column_dimensions["A"].width = 3
    ws.column_dimensions["B"].width = 24   # Employee
    ws.column_dimensions["C"].width = 20   # Department
    ws.column_dimensions["D"].width = 22   # Role
    ws.column_dimensions["E"].width = 70   # Key Activities (wide narrative)
    ws.column_dimensions["F"].width = 12   # Meetings
    ws.column_dimensions["G"].width = 16   # Revenue

    _merge_title(
        ws, 1, last_col,
        f"All Department Employee Summary Report — {range_label}",
        _TITLE_FONT,
    )
    ws.row_dimensions[1].height = 30
    subtitle = (
        f"Ornate Solar  |  {total_reports} reports  |  {total_employees} employees  "
        f"|  {dept_count} departments  |  Generated: {today_label}"
    )
    _merge_title(ws, 2, last_col, subtitle, _SUBTITLE_FONT)

    headers = [
        "Employee Name",
        "Department",
        "Role / Designation",
        f"Key Activities ({range_label})",
        "Meetings",
        "Revenue (₹)",
    ]
    for i, h in enumerate(headers, start=2):
        _put(ws, 4, i, h, font=_HEADER_FONT, fill=_HEADER_FILL, align=_CENTER)
    ws.row_dimensions[4].height = 28

    sorted_groups = sorted(by_dept.values(), key=_dept_sort_key)

    row_idx = 5
    for group in sorted_groups:
        d = group["dept"]
        dept_name = d.name if d else "No Department"
        user_entries = []
        for _, user_reports in group["users"].items():
            sorted_reports = sorted(user_reports, key=lambda r: r.date or date_type.min)
            user_entries.append(_summarise_user_reports(sorted_reports))
        user_entries.sort(key=lambda t: t[0].lower())

        for idx, (name, role, key_acts, meetings, revenue, _count) in enumerate(user_entries):
            _put(ws, row_idx, 2, name, font=_NAME_FONT, fill=_ROW_FILL, align=_LEFT_TOP_WRAP)
            if idx == 0:
                _put(ws, row_idx, 3, dept_name, font=_BADGE_FONT, fill=_BADGE_FILL, align=_CENTER)
            else:
                _put(ws, row_idx, 3, "", fill=_ROW_FILL)
            _put(ws, row_idx, 4, role or _DASH,
                 font=_BODY_FONT if role else _MUTED_FONT,
                 fill=_ROW_FILL, align=_LEFT_TOP_WRAP)
            _put(ws, row_idx, 5, key_acts or _DASH,
                 font=_BODY_FONT if key_acts else _MUTED_FONT,
                 fill=_ROW_FILL, align=_LEFT_TOP_WRAP)
            if meetings is not None:
                disp = int(meetings) if float(meetings).is_integer() else round(meetings, 2)
                _put(ws, row_idx, 6, disp, font=_NAME_FONT, fill=_ROW_FILL, align=_CENTER)
            else:
                _put(ws, row_idx, 6, _DASH, font=_MUTED_FONT, fill=_ROW_FILL, align=_CENTER)
            if revenue is not None:
                _put(ws, row_idx, 7, _format_revenue(revenue),
                     font=Font(bold=True, size=10, color="155F00"),
                     fill=_ROW_FILL, align=_CENTER)
            else:
                _put(ws, row_idx, 7, _DASH, font=_MUTED_FONT, fill=_ROW_FILL, align=_CENTER)
            row_idx += 1

    ws.freeze_panes = "B5"
    _set_print_landscape_fit(ws)


def _build_department_sheet(ws, group, *, range_label: str):
    """Per-department sheet — 5 columns matching HR's manual report:

    A (margin) | B Employee | C Role | D Key Activities | E Meetings | F Revenue (₹)
    """
    d = group["dept"]
    dept_name = d.name if d else "No Department"
    last_col = 6

    ws.column_dimensions["A"].width = 3
    ws.column_dimensions["B"].width = 24   # Employee
    ws.column_dimensions["C"].width = 22   # Role
    ws.column_dimensions["D"].width = 80   # Key Activities (wide, narrative)
    ws.column_dimensions["E"].width = 12   # Meetings
    ws.column_dimensions["F"].width = 16   # Revenue

    _merge_title(
        ws, 1, last_col,
        f"{dept_name} — Employee Report ({range_label})",
        _TITLE_FONT,
    )
    ws.row_dimensions[1].height = 30
    n_users = len(group["users"])
    n_reports = sum(len(rs) for rs in group["users"].values())
    subtitle = (
        f"{n_users} {'employee' if n_users == 1 else 'employees'}  ·  "
        f"{n_reports} reports submitted"
    )
    _merge_title(ws, 2, last_col, subtitle, _SUBTITLE_FONT)

    headers = ["Employee", "Role", f"Key Activities ({range_label})", "Meetings", "Revenue (₹)"]
    for i, h in enumerate(headers, start=2):
        _put(ws, 4, i, h, font=_HEADER_FONT, fill=_HEADER_FILL, align=_CENTER)
    ws.row_dimensions[4].height = 28

    user_rows = []
    for _, user_reports in group["users"].items():
        sorted_reports = sorted(user_reports, key=lambda r: r.date or date_type.min)
        name, role, key_acts, meetings, revenue, count = _summarise_user_reports(sorted_reports)
        user_rows.append({
            "name": name, "role": role, "key_acts": key_acts,
            "meetings": meetings, "revenue": revenue, "count": count,
        })
    user_rows.sort(key=lambda r: r["name"].lower())

    row_idx = 5
    for row in user_rows:
        _put(ws, row_idx, 2, row["name"], font=_NAME_FONT, fill=_ROW_FILL, align=_LEFT_TOP_WRAP)
        _put(ws, row_idx, 3, row["role"] or _DASH,
             font=_BODY_FONT if row["role"] else _MUTED_FONT,
             fill=_ROW_FILL, align=_LEFT_TOP_WRAP)
        _put(ws, row_idx, 4, row["key_acts"] or _DASH,
             font=_BODY_FONT if row["key_acts"] else _MUTED_FONT,
             fill=_ROW_FILL, align=_LEFT_TOP_WRAP)
        if row["meetings"] is not None:
            disp = int(row["meetings"]) if float(row["meetings"]).is_integer() else round(row["meetings"], 2)
            _put(ws, row_idx, 5, disp, font=_NAME_FONT, fill=_ROW_FILL, align=_CENTER)
        else:
            _put(ws, row_idx, 5, _DASH, font=_MUTED_FONT, fill=_ROW_FILL, align=_CENTER)
        if row["revenue"] is not None:
            _put(ws, row_idx, 6, _format_revenue(row["revenue"]),
                 font=Font(bold=True, size=10, color="155F00"),
                 fill=_ROW_FILL, align=_CENTER)
        else:
            _put(ws, row_idx, 6, _DASH, font=_MUTED_FONT, fill=_ROW_FILL, align=_CENTER)
        row_idx += 1

    # Department total row.
    if user_rows:
        _put(ws, row_idx, 2, "Department Total",
             font=_TOTAL_FONT, fill=_TOTAL_FILL, align=_LEFT_TOP_WRAP)
        _put(ws, row_idx, 3,
             f"{len(user_rows)} {'employee' if len(user_rows) == 1 else 'employees'}",
             font=_TOTAL_FONT, fill=_TOTAL_FILL, align=_LEFT_TOP_WRAP)
        _put(ws, row_idx, 4, f"{n_reports} reports submitted",
             font=_TOTAL_FONT, fill=_TOTAL_FILL, align=_LEFT_TOP_WRAP)
        meetings_sum = sum(r["meetings"] or 0 for r in user_rows)
        if any(r["meetings"] is not None for r in user_rows):
            disp = int(meetings_sum) if float(meetings_sum).is_integer() else round(meetings_sum, 2)
            _put(ws, row_idx, 5, disp, font=_TOTAL_FONT, fill=_TOTAL_FILL, align=_CENTER)
        else:
            _put(ws, row_idx, 5, _DASH, font=_MUTED_FONT, fill=_TOTAL_FILL, align=_CENTER)
        revenue_sum = sum(r["revenue"] or 0 for r in user_rows)
        if any(r["revenue"] is not None for r in user_rows):
            _put(ws, row_idx, 6, _format_revenue(revenue_sum),
                 font=Font(bold=True, size=10, color="155F00"),
                 fill=_TOTAL_FILL, align=_CENTER)
        else:
            _put(ws, row_idx, 6, _DASH, font=_MUTED_FONT, fill=_TOTAL_FILL, align=_CENTER)

    ws.freeze_panes = "B5"
    _set_print_landscape_fit(ws)


def _build_dept_detail_sheet(ws, group, *, range_label: str, exclude_keys=()):
    """Per-field columns layout for a single department.  Used for Sales and
    Inside Sales detail tabs — Employee + one column per report_field.
    Numeric fields are summed; text fields are joined as distinct values.

    `exclude_keys` skips named fields entirely (used by Inside Sales to drop
    Data Called Type / Mail Sent / WhatsApp Sent / Other Works columns).
    """
    d = group["dept"]
    dept_name = d.name if d else "No Department"
    fields = list(d.report_fields) if d and d.report_fields else []
    if exclude_keys:
        skip = set(exclude_keys)
        fields = [f for f in fields if f.get("key") not in skip]
    n_fields = len(fields)
    last_col = 2 + n_fields  # A margin + B Employee + N field cols

    ws.column_dimensions["A"].width = 3
    ws.column_dimensions["B"].width = 22  # Employee
    for i in range(n_fields):
        ws.column_dimensions[get_column_letter(3 + i)].width = 22

    _merge_title(
        ws, 1, last_col,
        f"{dept_name} — summary table",
        _TITLE_FONT,
    )
    ws.row_dimensions[1].height = 26
    n_users = len(group["users"])
    n_reports = sum(len(rs) for rs in group["users"].values())
    subtitle = (
        f"{range_label}  |  "
        f"{n_users} {'employee' if n_users == 1 else 'employees'}  ·  "
        f"{n_reports} reports submitted"
    )
    _merge_title(ws, 2, last_col, subtitle, _SUBTITLE_FONT)

    # Header row.
    _put(ws, 4, 2, "Employee", font=_HEADER_FONT, fill=_HEADER_FILL, align=_CENTER)
    for i, f in enumerate(fields):
        label = f.get("label") or f.get("key") or ""
        _put(ws, 4, 3 + i, label,
             font=_HEADER_FONT, fill=_HEADER_FILL, align=_CENTER)
    ws.row_dimensions[4].height = 28

    # Build per-employee rows.
    user_rows = []
    for _, user_reports in group["users"].items():
        sorted_reports = sorted(user_reports, key=lambda r: r.date or date_type.min)
        u = sorted_reports[0].user
        full_name = ""
        if u:
            full_name = (
                f"{(u.first_name or '').strip()} {(u.last_name or '').strip()}".strip()
                or u.username
            )
        cells = {}
        for f in fields:
            key = f.get("key") or ""
            label = f.get("label") or key
            pairs = [(r.date, (r.data or {}).get(key, "")) for r in sorted_reports]
            non_empty = [
                (d_, str(v).strip()) for d_, v in pairs
                if v and str(v).strip() and str(v).strip().lower() != "on leave"
            ]
            if not non_empty:
                cells[key] = (_DASH, "muted")
                continue
            # Decide numeric vs text by label hint OR strict parse.
            if _NUMERIC_LABEL_RE.search(label) or _MEETING_LABEL_RE.search(label) or _REVENUE_LABEL_RE.search(label):
                total = sum(n for _, v in non_empty for n in _extract_numbers_text(v))
                is_money = bool(_REVENUE_LABEL_RE.search(label))
                if is_money:
                    cells[key] = (_format_revenue(total) if total else "₹0", "money" if total else "muted")
                else:
                    disp = int(total) if float(total).is_integer() else round(total, 2)
                    cells[key] = (disp, "num")
            else:
                # Strict-parse check.
                nums, all_num = [], True
                for _, v in non_empty:
                    cleaned = v.replace("₹", "").replace("$", "").replace(",", "").replace(" ", "")
                    try:
                        nums.append(float(cleaned))
                    except ValueError:
                        all_num = False
                        break
                if all_num and nums:
                    total = sum(nums)
                    disp = int(total) if float(total).is_integer() else round(total, 2)
                    cells[key] = (disp, "num")
                else:
                    # Text — compact join: distinct values, headline-only,
                    # trimmed to fit ~120 chars total with a trailing "…"
                    # when truncated.  No "(+N more)" markers.
                    seen, distinct = set(), []
                    for _, v in non_empty:
                        if v.lower() in seen:
                            continue
                        seen.add(v.lower())
                        distinct.append(v)
                    cells[key] = (_compact_join(distinct, max_chars=120), "body")
        user_rows.append({"name": full_name, "cells": cells})
    user_rows.sort(key=lambda r: r["name"].lower())

    # Matches the on-screen "Inside Sales — summary table" look: clean data
    # rows (no green tint on money), CENTER-aligned numbers, navy bold totals.
    num_font = Font(size=10, color="1A1A1A")
    money_data_font = Font(size=10, color="1A1A1A")
    total_navy_font = Font(bold=True, size=10, color="1B5E8B")
    num_align = Alignment(horizontal="center", vertical="center", wrap_text=False)
    body_left = Alignment(horizontal="left", vertical="top", wrap_text=True)

    row_idx = 5
    for row in user_rows:
        _put(ws, row_idx, 2, row["name"], font=_NAME_FONT, fill=_ROW_FILL, align=body_left)
        for i, f in enumerate(fields):
            v, kind = row["cells"].get(f.get("key", ""), (_DASH, "muted"))
            if kind == "money":
                _put(ws, row_idx, 3 + i, v, font=money_data_font, fill=_ROW_FILL, align=num_align)
            elif kind == "num":
                _put(ws, row_idx, 3 + i, v, font=num_font, fill=_ROW_FILL, align=num_align)
            elif kind == "body":
                _put(ws, row_idx, 3 + i, v, font=_BODY_FONT, fill=_ROW_FILL, align=body_left)
            else:
                _put(ws, row_idx, 3 + i, v, font=_MUTED_FONT, fill=_ROW_FILL, align=_CENTER)
        row_idx += 1

    # Total row — navy bold text on light-blue fill for every cell.
    if user_rows:
        _put(ws, row_idx, 2, "Total",
             font=total_navy_font, fill=_TOTAL_FILL, align=body_left)
        for i, f in enumerate(fields):
            cells = [r["cells"].get(f.get("key", ""), (_DASH, "muted")) for r in user_rows]
            kinds = {k for _, k in cells}
            if "money" in kinds:
                total = sum(
                    n for v, k in cells if k == "money"
                    for n in _extract_numbers_text(str(v).replace("₹", "").replace(",", ""))
                )
                _put(ws, row_idx, 3 + i, _format_revenue(total) if total else "₹0",
                     font=total_navy_font, fill=_TOTAL_FILL, align=num_align)
            elif "num" in kinds and "body" not in kinds:
                total = sum(v for v, k in cells if k == "num" and isinstance(v, (int, float)))
                disp = int(total) if float(total).is_integer() else round(total, 2)
                _put(ws, row_idx, 3 + i, disp,
                     font=total_navy_font, fill=_TOTAL_FILL, align=num_align)
            elif "body" in kinds:
                all_vals = set()
                for v, k in cells:
                    if k == "body" and isinstance(v, str):
                        for piece in v.split(","):
                            p = piece.strip()
                            if p:
                                all_vals.add(p.lower())
                _put(ws, row_idx, 3 + i, len(all_vals) if all_vals else _DASH,
                     font=total_navy_font if all_vals else _MUTED_FONT,
                     fill=_TOTAL_FILL, align=_CENTER)
            else:
                _put(ws, row_idx, 3 + i, _DASH,
                     font=_MUTED_FONT, fill=_TOTAL_FILL, align=_CENTER)

    ws.freeze_panes = "C5"
    _set_print_landscape_fit(ws)


def _build_service_detail_sheet(ws, group, *, range_label: str):
    """Service department — 3 aggregated columns instead of one per
    report_field.  Combines the underlying daily fields so HR sees totals
    for site visits, inverter complaints, and part replacements per
    employee.
    """
    last_col = 5  # A margin + B Employee + C Site Visit + D Inv Complaint + E Parts Repl

    ws.column_dimensions["A"].width = 3
    ws.column_dimensions["B"].width = 24
    ws.column_dimensions["C"].width = 18
    ws.column_dimensions["D"].width = 26
    ws.column_dimensions["E"].width = 26

    _merge_title(ws, 1, last_col, "Service — summary table", _TITLE_FONT)
    ws.row_dimensions[1].height = 26
    n_users = len(group["users"])
    n_reports = sum(len(rs) for rs in group["users"].values())
    subtitle = (
        f"{range_label}  |  "
        f"{n_users} {'employee' if n_users == 1 else 'employees'}  ·  "
        f"{n_reports} reports submitted"
    )
    _merge_title(ws, 2, last_col, subtitle, _SUBTITLE_FONT)

    headers = ["Employee", "Total Site Visit", "Total Inverter Complaint", "Inverter Parts Replacement"]
    for i, h in enumerate(headers, start=2):
        _put(ws, 4, i, h, font=_HEADER_FONT, fill=_HEADER_FILL, align=_CENTER)
    ws.row_dimensions[4].height = 28

    num_font = Font(size=10, color="1A1A1A")
    right_align = Alignment(horizontal="center", vertical="center", wrap_text=False)
    body_left = Alignment(horizontal="left", vertical="top", wrap_text=True)

    # Aggregation rules:
    #   Total Site Visit         = Solar Install visits + Inverter Complaint visits
    #   Total Inverter Complaint = Complaint site visits + Complaint tele/video
    #   Inverter Parts Replace   = Inverter Part Replacement
    SITE_VISIT_KEYS = ("solarInstallationSiteVisit", "inverterComplaintSiteVisit")
    INV_COMPLAINT_KEYS = ("inverterComplaintSiteVisit", "inverterComplaintTeleVideo")
    PARTS_REPL_KEYS = ("inverterPartReplacement",)

    user_rows = []
    for _, user_reports in group["users"].items():
        sorted_reports = sorted(user_reports, key=lambda r: r.date or date_type.min)
        u = sorted_reports[0].user
        full_name = ""
        if u:
            full_name = (
                f"{(u.first_name or '').strip()} {(u.last_name or '').strip()}".strip()
                or u.username
            )

        def sum_keys(keys):
            total = 0.0
            for r in sorted_reports:
                data = r.data or {}
                for k in keys:
                    for n in _extract_numbers_text(data.get(k, "")):
                        total += n
            return total

        user_rows.append({
            "name": full_name,
            "site_visit": sum_keys(SITE_VISIT_KEYS),
            "inv_complaint": sum_keys(INV_COMPLAINT_KEYS),
            "parts_repl": sum_keys(PARTS_REPL_KEYS),
        })
    user_rows.sort(key=lambda r: r["name"].lower())

    def _disp(n):
        return int(n) if float(n).is_integer() else round(n, 2)

    row_idx = 5
    for row in user_rows:
        _put(ws, row_idx, 2, row["name"], font=_NAME_FONT, fill=_ROW_FILL, align=body_left)
        for i, key in enumerate(("site_visit", "inv_complaint", "parts_repl")):
            _put(ws, row_idx, 3 + i, _disp(row[key]),
                 font=num_font, fill=_ROW_FILL, align=right_align)
        row_idx += 1

    if user_rows:
        _put(ws, row_idx, 2, "Total", font=_TOTAL_FONT, fill=_TOTAL_FILL, align=body_left)
        for i, key in enumerate(("site_visit", "inv_complaint", "parts_repl")):
            total = sum(r[key] for r in user_rows)
            _put(ws, row_idx, 3 + i, _disp(total),
                 font=_TOTAL_FONT, fill=_TOTAL_FILL, align=right_align)

    ws.freeze_panes = "C5"
    _set_print_landscape_fit(ws)


def _build_project_detail_sheet(ws, group, *, range_label: str):
    """Project department — 2 condensed columns:
      - Total Site Visit  (count of reports the employee submitted)
      - Work on Site      (distinct snippets from Work Done + Work in Progress,
                           short list)
    """
    last_col = 4  # A margin + B Employee + C Total Site Visit + D Work on Site

    ws.column_dimensions["A"].width = 3
    ws.column_dimensions["B"].width = 24
    ws.column_dimensions["C"].width = 18
    ws.column_dimensions["D"].width = 90

    _merge_title(ws, 1, last_col, "Project — summary table", _TITLE_FONT)
    ws.row_dimensions[1].height = 26
    n_users = len(group["users"])
    n_reports = sum(len(rs) for rs in group["users"].values())
    subtitle = (
        f"{range_label}  |  "
        f"{n_users} {'employee' if n_users == 1 else 'employees'}  ·  "
        f"{n_reports} reports submitted"
    )
    _merge_title(ws, 2, last_col, subtitle, _SUBTITLE_FONT)

    headers = ["Employee", "Total Site Visit", "Work on Site"]
    for i, h in enumerate(headers, start=2):
        _put(ws, 4, i, h, font=_HEADER_FONT, fill=_HEADER_FILL, align=_CENTER)
    ws.row_dimensions[4].height = 28

    num_font = Font(size=10, color="1A1A1A")
    right_align = Alignment(horizontal="center", vertical="center", wrap_text=False)
    body_left = Alignment(horizontal="left", vertical="top", wrap_text=True)

    WORK_KEYS = ("workDone", "workInProgress")

    user_rows = []
    for _, user_reports in group["users"].items():
        sorted_reports = sorted(user_reports, key=lambda r: r.date or date_type.min)
        u = sorted_reports[0].user
        full_name = ""
        if u:
            full_name = (
                f"{(u.first_name or '').strip()} {(u.last_name or '').strip()}".strip()
                or u.username
            )

        # Site visit count = number of reports submitted in range.
        visit_count = len(sorted_reports)

        # Work on Site — compact join of distinct headline phrases from
        # Work Done + Work in Progress, packed into ~150 chars.  Trailing
        # "…" hints at overflow without a "(+N more)" tag.
        seen = set()
        distinct = []
        for r in sorted_reports:
            data = r.data or {}
            for k in WORK_KEYS:
                v = (data.get(k) or "").strip()
                if not v or v.lower() == "on leave":
                    continue
                key_lower = v.lower()
                if key_lower in seen:
                    continue
                seen.add(key_lower)
                distinct.append(v)
        work_text = _compact_join(distinct, max_chars=150)

        user_rows.append({
            "name": full_name,
            "visit_count": visit_count,
            "work_text": work_text or _DASH,
        })
    user_rows.sort(key=lambda r: r["name"].lower())

    row_idx = 5
    for row in user_rows:
        _put(ws, row_idx, 2, row["name"], font=_NAME_FONT, fill=_ROW_FILL, align=body_left)
        _put(ws, row_idx, 3, row["visit_count"], font=num_font, fill=_ROW_FILL, align=right_align)
        _put(ws, row_idx, 4,
             row["work_text"],
             font=_BODY_FONT if row["work_text"] != _DASH else _MUTED_FONT,
             fill=_ROW_FILL, align=body_left)
        row_idx += 1

    if user_rows:
        _put(ws, row_idx, 2, "Total", font=_TOTAL_FONT, fill=_TOTAL_FILL, align=body_left)
        _put(ws, row_idx, 3,
             sum(r["visit_count"] for r in user_rows),
             font=_TOTAL_FONT, fill=_TOTAL_FILL, align=right_align)
        _put(ws, row_idx, 4, "", fill=_TOTAL_FILL)

    ws.freeze_panes = "C5"
    _set_print_landscape_fit(ws)


def _build_combined_summary_sheet(ws, other_groups, *, range_label: str, today_label: str):
    """One-table combined summary for every department except Sales & Inside
    Sales (those get their own detailed tabs).  Columns:
    Department | # | Employee | Focus Area | Key Activities.
    """
    last_col = 6
    ws.column_dimensions["A"].width = 3
    ws.column_dimensions["B"].width = 20   # Department
    ws.column_dimensions["C"].width = 5    # #
    ws.column_dimensions["D"].width = 22   # Employee
    ws.column_dimensions["E"].width = 22   # Focus Area (role)
    ws.column_dimensions["F"].width = 80   # Key Activities (wider since it's the tail)

    n_employees = sum(len(g["users"]) for g in other_groups)
    n_reports = sum(len(rs) for g in other_groups for rs in g["users"].values())

    _merge_title(
        ws, 1, last_col,
        "All Department Employee Summary (other departments)",
        _TITLE_FONT,
    )
    ws.row_dimensions[1].height = 28
    subtitle = (
        f"Date range: {range_label}  |  "
        f"{n_employees} employees across {len(other_groups)} departments  |  "
        f"{n_reports} reports  |  Generated: {today_label}  "
        f"(Inside Sales & Sales shown on their own detailed tabs)"
    )
    _merge_title(ws, 2, last_col, subtitle, _SUBTITLE_FONT)

    headers = [
        "Department", "#", "Employee", "Focus Area",
        f"Key Activities ({range_label})",
    ]
    for i, h in enumerate(headers, start=2):
        _put(ws, 4, i, h, font=_HEADER_FONT, fill=_HEADER_FILL, align=_CENTER)
    ws.row_dimensions[4].height = 28

    sorted_groups = sorted(other_groups, key=_dept_sort_key)

    row_idx = 5
    for group in sorted_groups:
        d = group["dept"]
        dept_name = d.name if d else "No Department"
        user_entries = []
        for _, user_reports in group["users"].items():
            sorted_reports = sorted(user_reports, key=lambda r: r.date or date_type.min)
            user_entries.append(_summarise_user_reports(sorted_reports))
        user_entries.sort(key=lambda t: t[0].lower())

        for idx, (name, role, key_acts, _meetings, _revenue, _count) in enumerate(user_entries, start=1):
            # Department badge only on first row of the dept block.
            if idx == 1:
                _put(ws, row_idx, 2, dept_name,
                     font=_BADGE_FONT, fill=_BADGE_FILL, align=_CENTER)
            else:
                _put(ws, row_idx, 2, "", fill=_ROW_FILL)
            _put(ws, row_idx, 3, idx, font=_BODY_FONT, fill=_ROW_FILL, align=_CENTER)
            _put(ws, row_idx, 4, name, font=_NAME_FONT, fill=_ROW_FILL, align=_LEFT_TOP_WRAP)
            _put(ws, row_idx, 5, role or _DASH,
                 font=_BODY_FONT if role else _MUTED_FONT,
                 fill=_ROW_FILL, align=_LEFT_TOP_WRAP)
            _put(ws, row_idx, 6, key_acts or _DASH,
                 font=_BODY_FONT if key_acts else _MUTED_FONT,
                 fill=_ROW_FILL, align=_LEFT_TOP_WRAP)
            row_idx += 1

    ws.freeze_panes = "D5"
    _set_print_landscape_fit(ws)


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


@router.get("/export.csv")
def export_csv(
    employee: int | None = Query(None, description="Filter by employee id"),
    department: str | None = Query(None, description="Filter by department slug"),
    start: date_type | None = Query(None, description="Inclusive start date"),
    end: date_type | None = Query(None, description="Inclusive end date"),
    db: Session = Depends(get_db),
    _: User = Depends(get_current_user),
):
    """Single-file CSV mirror of the styled XLSX summary.

    Columns: Department, Employee, Role, Key Activities, Meetings, Revenue (₹).
    Rows ordered Sales → Inside Sales → other departments alphabetically.
    Key Activities lines are joined with " | " so each row stays on one line
    in spreadsheet apps that don't auto-wrap.  UTF-8 with BOM so Excel
    detects the encoding correctly on Windows.
    """
    q = db.query(DailyReport).join(User, DailyReport.user_id == User.id)
    if employee is not None:
        q = q.filter(DailyReport.user_id == employee)
    if department:
        dept = db.query(Department).filter(Department.slug == department).first()
        if not dept:
            return _empty_csv("daily-reports.csv")
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

    buf = io.StringIO()
    writer = csv.writer(buf)
    range_label = _format_range_label(start, end)
    writer.writerow([f"All Department Employee Summary Report — {range_label}"])
    writer.writerow([])
    writer.writerow(
        ["Department", "Employee", "Role / Designation", "Key Activities", "Meetings", "Revenue (₹)"]
    )

    for group in sorted(by_dept.values(), key=_dept_sort_key):
        d = group["dept"]
        dept_name = d.name if d else "No Department"
        # Sort users in each dept by name.
        user_entries = []
        for _, user_reports in group["users"].items():
            sorted_reports = sorted(user_reports, key=lambda r: r.date or date_type.min)
            user_entries.append(_summarise_user_reports(sorted_reports))
        user_entries.sort(key=lambda t: t[0].lower())

        for name, role, key_acts, meetings, revenue, _count in user_entries:
            # Collapse multi-line activities into one CSV cell (single line).
            key_acts_line = (key_acts or "").replace("\n", " | ")
            meetings_str = ""
            if meetings is not None:
                meetings_str = (
                    str(int(meetings))
                    if float(meetings).is_integer()
                    else f"{meetings:.2f}"
                )
            revenue_str = _format_revenue(revenue) if revenue is not None else ""
            writer.writerow([
                dept_name,
                name,
                role or "",
                key_acts_line,
                meetings_str,
                revenue_str,
            ])
        # Blank separator row between departments for visual breathing room.
        writer.writerow([])

    # UTF-8 BOM helps Excel for Windows pick the right encoding.
    payload = buf.getvalue().encode("utf-8-sig")

    parts = ["daily-reports"]
    if start:
        parts.append(start.isoformat())
    if end:
        parts.append(end.isoformat())
    if not (start or end):
        parts.append(date_type.today().isoformat())
    filename = "_".join(parts) + ".csv"

    return Response(
        content=payload,
        media_type="text/csv; charset=utf-8",
        headers={
            "Content-Disposition": f'attachment; filename="{filename}"',
            "Content-Length": str(len(payload)),
        },
    )


def _empty_csv(filename: str) -> Response:
    payload = "No reports match the current filters.\n".encode("utf-8-sig")
    return Response(
        content=payload,
        media_type="text/csv; charset=utf-8",
        headers={
            "Content-Disposition": f'attachment; filename="{filename}"',
            "Content-Length": str(len(payload)),
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
