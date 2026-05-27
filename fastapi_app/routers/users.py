from datetime import datetime, timezone

from fastapi import APIRouter, Depends, HTTPException, Query, status
from sqlalchemy.orm import Session

from auth import get_current_user, hash_password
from database import get_db
from models import Department, User
from schemas import EmployeeCreate, EmployeeUpdate, UserOut

router = APIRouter(prefix="/api", tags=["users"])


@router.get("/me", response_model=UserOut)
def me(user: User = Depends(get_current_user)):
    return user


@router.get("/employees", response_model=list[UserOut])
def list_employees(
    department: str | None = Query(None, description="Department slug filter"),
    include_inactive: bool = Query(False, description="Include deactivated employees"),
    db: Session = Depends(get_db),
    _: User = Depends(get_current_user),
):
    q = db.query(User)
    if not include_inactive:
        q = q.filter(User.is_active.is_(True))
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


# ---------- HR-only CRUD ----------

def _require_hr(user: User) -> None:
    if user.role != "hr":
        raise HTTPException(status.HTTP_403_FORBIDDEN, "Only HR can manage employees")


def _resolve_dept(db: Session, slug: str | None) -> Department | None:
    """Translate a department slug (or empty string / None) into a Department row.

    Empty string → None (clears the assignment).  Unknown slug raises 400.
    """
    if slug is None or slug == "":
        return None
    dept = db.query(Department).filter(Department.slug == slug).first()
    if not dept:
        raise HTTPException(status.HTTP_400_BAD_REQUEST, f"Unknown department slug: {slug}")
    return dept


@router.post("/employees", response_model=UserOut, status_code=status.HTTP_201_CREATED)
def create_employee(
    payload: EmployeeCreate,
    db: Session = Depends(get_db),
    actor: User = Depends(get_current_user),
):
    """HR-only: create a new employee account."""
    _require_hr(actor)

    if payload.role not in ("employee", "hr"):
        raise HTTPException(status.HTTP_400_BAD_REQUEST, "role must be 'employee' or 'hr'")

    # Unique email / username
    if db.query(User).filter(User.email == payload.email).first():
        raise HTTPException(status.HTTP_409_CONFLICT, f"Email '{payload.email}' is already in use")

    username = (payload.username or payload.email.split("@")[0]).strip().lower()
    if db.query(User).filter(User.username == username).first():
        raise HTTPException(status.HTTP_409_CONFLICT, f"Username '{username}' is already taken")

    first = (payload.first_name or "").strip()
    if not first:
        raise HTTPException(status.HTTP_400_BAD_REQUEST, "first_name is required")

    raw_password = payload.password or f"{first.lower()}@ornate"

    dept = _resolve_dept(db, payload.department)

    now = datetime.now(timezone.utc)
    new_user = User(
        email=str(payload.email).lower(),
        username=username,
        first_name=first,
        last_name=(payload.last_name or "").strip(),
        password=hash_password(raw_password),
        role=payload.role,
        is_superuser=(payload.role == "hr"),
        is_staff=(payload.role == "hr"),
        is_active=True,
        is_team_head=bool(payload.is_team_head),
        title=(payload.title or "").strip(),
        contact_number=(payload.contact_number or "").strip(),
        department_id=dept.id if dept else None,
        organisation=(payload.organisation or "").strip(),
        reporting_manager=(payload.reporting_manager or "").strip(),
        date_of_joining=payload.date_of_joining,
        date_joined=now,
    )
    db.add(new_user)
    db.commit()
    db.refresh(new_user)
    return new_user


@router.patch("/employees/{user_id}", response_model=UserOut)
def update_employee(
    user_id: int,
    payload: EmployeeUpdate,
    db: Session = Depends(get_db),
    actor: User = Depends(get_current_user),
):
    """HR-only: edit employee profile, move between departments, deactivate, etc."""
    _require_hr(actor)

    target = db.query(User).filter(User.id == user_id).first()
    if not target:
        raise HTTPException(status.HTTP_404_NOT_FOUND, "Employee not found")

    data = payload.model_dump(exclude_unset=True)

    if "department" in data:
        dept = _resolve_dept(db, data.pop("department"))
        target.department_id = dept.id if dept else None

    if "password" in data:
        raw = data.pop("password")
        if raw:
            target.password = hash_password(raw)

    if "role" in data:
        new_role = data["role"]
        if new_role not in ("employee", "hr"):
            raise HTTPException(status.HTTP_400_BAD_REQUEST, "role must be 'employee' or 'hr'")
        target.role = new_role
        # Keep the legacy Django flags coherent with the role we just set.
        target.is_superuser = (new_role == "hr")
        target.is_staff = (new_role == "hr")
        data.pop("role")

    # Whitelist of fields that map directly onto the column of the same name.
    for field in (
        "first_name", "last_name", "email", "title", "contact_number",
        "organisation", "reporting_manager", "date_of_joining", "is_active",
        "is_team_head",
    ):
        if field in data:
            value = data[field]
            if field == "email" and value is not None:
                value = str(value).lower()
            setattr(target, field, value)

    db.commit()
    db.refresh(target)
    return target
