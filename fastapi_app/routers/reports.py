from datetime import date as date_type, datetime, timezone

from fastapi import APIRouter, Depends, HTTPException, Query, status
from sqlalchemy import and_
from sqlalchemy.orm import Session

from auth import get_current_user
from database import get_db
from models import DailyReport, Department, User
from schemas import ReportIn, ReportOut

router = APIRouter(prefix="/api/reports", tags=["reports"])


@router.get("", response_model=list[ReportOut])
def list_reports(
    employee: int | None = Query(None, description="Filter by employee id"),
    department: str | None = Query(None, description="Filter by department slug"),
    start: date_type | None = Query(None, description="Inclusive start date"),
    end: date_type | None = Query(None, description="Inclusive end date"),
    db: Session = Depends(get_db),
    _: User = Depends(get_current_user),
):
    q = db.query(DailyReport)
    if employee is not None:
        q = q.filter(DailyReport.user_id == employee)
    if department:
        dept = db.query(Department).filter(Department.slug == department).first()
        if not dept:
            return []
        q = q.join(User, DailyReport.user_id == User.id).filter(User.department_id == dept.id)
    if start:
        q = q.filter(DailyReport.date >= start)
    if end:
        q = q.filter(DailyReport.date <= end)
    return q.order_by(DailyReport.date.desc(), DailyReport.user_id).all()


@router.post("", response_model=ReportOut, status_code=status.HTTP_201_CREATED)
def upsert_report(
    payload: ReportIn,
    db: Session = Depends(get_db),
    user: User = Depends(get_current_user),
):
    existing = db.query(DailyReport).filter(
        and_(DailyReport.user_id == user.id, DailyReport.date == payload.date)
    ).first()

    now = datetime.now(timezone.utc)
    if existing:
        existing.data = payload.data
        existing.submitted_at = now
        report = existing
    else:
        report = DailyReport(
            user_id=user.id,
            date=payload.date,
            data=payload.data,
            submitted_at=now,
            created_at=now,
        )
        db.add(report)
    db.commit()
    db.refresh(report)
    return report


@router.get("/missing-today", response_model=list[int])
def missing_today(
    db: Session = Depends(get_db),
    _: User = Depends(get_current_user),
):
    """Returns the list of user ids that have NOT submitted a report today."""
    today = date_type.today()
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
