from datetime import datetime, timezone

from fastapi import APIRouter, Depends, HTTPException, status
from sqlalchemy.orm import Session

from auth import (
    create_access_token, create_refresh_token, hash_password, verify_password,
)
from database import get_db
from models import Department, User
from schemas import LoginRequest, SignupRequest, TokenResponse, UserOut

router = APIRouter(prefix="/api/auth", tags=["auth"])


@router.post("/login", response_model=TokenResponse)
def login(payload: LoginRequest, db: Session = Depends(get_db)):
    user = db.query(User).filter(User.email == payload.email).first()
    if not user or not verify_password(payload.password, user.password):
        raise HTTPException(status.HTTP_401_UNAUTHORIZED, "Invalid email or password")
    if not user.is_active:
        raise HTTPException(status.HTTP_403_FORBIDDEN, "Account is disabled")
    return _token_response(user)


@router.post("/signup", response_model=TokenResponse, status_code=status.HTTP_201_CREATED)
def signup(payload: SignupRequest, db: Session = Depends(get_db)):
    if db.query(User).filter((User.email == payload.email) | (User.username == payload.username)).first():
        raise HTTPException(status.HTTP_409_CONFLICT, "User with this email or username already exists")

    dept = None
    if payload.department:
        dept = db.query(Department).filter(Department.slug == payload.department).first()
        if not dept:
            raise HTTPException(status.HTTP_400_BAD_REQUEST, f"Unknown department: {payload.department}")

    now = datetime.now(timezone.utc)
    user = User(
        username=payload.username,
        email=payload.email,
        first_name=payload.first_name,
        last_name=payload.last_name,
        contact_number=payload.contact_number or "",
        password=hash_password(payload.password),
        is_active=True,
        is_staff=False,
        is_superuser=False,
        date_joined=now,
        role="employee",
        title="",
        department_id=dept.id if dept else None,
    )
    db.add(user)
    db.commit()
    db.refresh(user)
    return _token_response(user)


def _token_response(user: User) -> TokenResponse:
    return TokenResponse(
        access_token=create_access_token(user.id),
        refresh_token=create_refresh_token(user.id),
        user=UserOut.model_validate(user),
    )
