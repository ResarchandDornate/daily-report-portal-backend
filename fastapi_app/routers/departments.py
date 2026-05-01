from fastapi import APIRouter, Depends
from sqlalchemy.orm import Session

from database import get_db
from models import Department
from schemas import DepartmentOut

router = APIRouter(prefix="/api/departments", tags=["departments"])


@router.get("", response_model=list[DepartmentOut])
def list_departments(db: Session = Depends(get_db)):
    """Public — the signup form needs this before the user has a token."""
    return db.query(Department).order_by(Department.name).all()
