from datetime import date as date_type, datetime, timedelta, timezone

from fastapi import APIRouter, Depends, HTTPException, Query, status
from sqlalchemy import and_
from sqlalchemy.orm import Session

from auth import get_current_user
from database import get_db
from models import DailyReport, Department, User
from schemas import LeaveIn, ReportIn, ReportListOut, ReportOut

router = APIRouter(prefix="/api/reports", tags=["reports"])


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
        if user.role != "hr":
            raise HTTPException(
                status.HTTP_403_FORBIDDEN,
                "Only HR can submit on behalf of other employees",
            )
        # Confirm the target user actually exists so we don't create an orphan FK.
        target = db.query(User).filter(User.id == payload.user_id).first()
        if not target:
            raise HTTPException(
                status.HTTP_404_NOT_FOUND,
                f"No user with id={payload.user_id}",
            )
        target_user_id = target.id

    existing = db.query(DailyReport).filter(
        and_(DailyReport.user_id == target_user_id, DailyReport.date == payload.date)
    ).first()

    now = datetime.now(timezone.utc)
    if existing:
        # Employees submit-once policy: only HR can overwrite an existing
        # report.  Regular employees who try to "edit" by re-submitting for a
        # past date get a 403.  HR users editing their own reports are allowed.
        if user.role != "hr":
            raise HTTPException(
                status.HTTP_403_FORBIDDEN,
                "A report already exists for this date. Contact HR to edit it.",
            )
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
