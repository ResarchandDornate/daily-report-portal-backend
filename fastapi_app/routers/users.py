from fastapi import APIRouter, Depends, HTTPException, Query, status
from sqlalchemy.orm import Session

from auth import get_current_user
from database import get_db
from models import Department, User
from schemas import UserOut

router = APIRouter(prefix="/api", tags=["users"])


@router.get("/me", response_model=UserOut)
def me(user: User = Depends(get_current_user)):
    return user


@router.get("/employees", response_model=list[UserOut])
def list_employees(
    department: str | None = Query(None, description="Department slug filter"),
    db: Session = Depends(get_db),
    _: User = Depends(get_current_user),
):
    q = db.query(User).filter(User.is_active.is_(True))
    if department:
        dept = db.query(Department).filter(Department.slug == department).first()
        if not dept:
            return []
        q = q.filter(User.department_id == dept.id)
    return q.order_by(User.first_name).all()


@router.get("/employees/{user_id}", response_model=UserOut)
def get_employee(
    user_id: int,
    db: Session = Depends(get_db),
    _: User = Depends(get_current_user),
):
    user = db.query(User).filter(User.id == user_id).first()
    if not user:
        raise HTTPException(status.HTTP_404_NOT_FOUND, "Employee not found")
    return user
