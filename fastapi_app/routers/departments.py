from datetime import datetime, timezone

from fastapi import APIRouter, Depends, HTTPException, status
from sqlalchemy.orm import Session

from auth import get_current_user
from database import get_db
from models import DailyReport, Department, User
from schemas import DepartmentCreate, DepartmentOut, DepartmentUpdate

router = APIRouter(prefix="/api/departments", tags=["departments"])


@router.get("", response_model=list[DepartmentOut])
def list_departments(db: Session = Depends(get_db)):
    """Public — the signup form needs this before the user has a token."""
    return db.query(Department).order_by(Department.name).all()


@router.post("", response_model=DepartmentOut, status_code=status.HTTP_201_CREATED)
def create_department(
    payload: DepartmentCreate,
    db: Session = Depends(get_db),
    user: User = Depends(get_current_user),
):
    """HR-only: create a new department."""
    if user.role != "hr":
        raise HTTPException(status.HTTP_403_FORBIDDEN, "Only HR can manage departments")

    slug = (payload.slug or "").strip()
    if not slug or not slug.replace("_", "").replace("-", "").isalnum():
        raise HTTPException(
            status.HTTP_400_BAD_REQUEST,
            "Slug must be alphanumeric (letters/digits, optionally underscores or hyphens)",
        )
    if db.query(Department).filter(Department.slug == slug).first():
        raise HTTPException(status.HTTP_409_CONFLICT, f"Department '{slug}' already exists")

    now = datetime.now(timezone.utc)
    dept = Department(
        slug=slug,
        name=(payload.name or "").strip(),
        color=(payload.color or "zinc").strip() or "zinc",
        report_fields=[f.model_dump() for f in payload.report_fields],
        created_at=now,
        updated_at=now,
    )
    db.add(dept)
    db.commit()
    db.refresh(dept)
    return dept


@router.patch("/{slug}", response_model=DepartmentOut)
def update_department(
    slug: str,
    payload: DepartmentUpdate,
    db: Session = Depends(get_db),
    user: User = Depends(get_current_user),
):
    """HR-only: rename a dept, change its color, or update its report_fields."""
    if user.role != "hr":
        raise HTTPException(status.HTTP_403_FORBIDDEN, "Only HR can manage departments")

    dept = db.query(Department).filter(Department.slug == slug).first()
    if not dept:
        raise HTTPException(status.HTTP_404_NOT_FOUND, f"No department '{slug}'")

    if payload.name is not None:
        dept.name = payload.name.strip()
    if payload.color is not None:
        dept.color = payload.color.strip() or "zinc"
    if payload.report_fields is not None:
        dept.report_fields = [f.model_dump() for f in payload.report_fields]
    dept.updated_at = datetime.now(timezone.utc)
    db.commit()
    db.refresh(dept)
    return dept


@router.delete("/{slug}", status_code=status.HTTP_204_NO_CONTENT)
def delete_department(
    slug: str,
    db: Session = Depends(get_db),
    user: User = Depends(get_current_user),
):
    """HR-only: delete a department.  Refuses if any employee still belongs to it."""
    if user.role != "hr":
        raise HTTPException(status.HTTP_403_FORBIDDEN, "Only HR can manage departments")

    dept = db.query(Department).filter(Department.slug == slug).first()
    if not dept:
        raise HTTPException(status.HTTP_404_NOT_FOUND, f"No department '{slug}'")

    emp_count = db.query(User).filter(User.department_id == dept.id).count()
    if emp_count:
        raise HTTPException(
            status.HTTP_409_CONFLICT,
            f"Department has {emp_count} employee(s) — move them out before deleting.",
        )
    # Safety net: refuse if any report still references this department's employees
    # (shouldn't be possible if emp_count == 0, but belt + braces).
    report_count = (
        db.query(DailyReport)
        .join(User, DailyReport.user_id == User.id)
        .filter(User.department_id == dept.id)
        .count()
    )
    if report_count:
        raise HTTPException(
            status.HTTP_409_CONFLICT,
            "Department still has historical reports tied to it — cannot delete.",
        )

    db.delete(dept)
    db.commit()
    return None
