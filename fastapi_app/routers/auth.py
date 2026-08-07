from datetime import datetime, timezone

from fastapi import APIRouter, Depends, HTTPException, status
from jose import JWTError, jwt
from sqlalchemy import func
from sqlalchemy.orm import Session

from auth import (
    JWT_ALGORITHM, JWT_SECRET,
    create_access_token, create_refresh_token, hash_password, verify_password,
)
from database import get_db
from models import Department, User
from schemas import LoginRequest, RefreshRequest, SignupRequest, TokenResponse, UserOut

router = APIRouter(prefix="/api/auth", tags=["auth"])


@router.post("/login", response_model=TokenResponse)
def login(payload: LoginRequest, db: Session = Depends(get_db)):
    # Case-insensitive match: accounts are stored lowercased (see
    # create_employee), but nothing normalized the login side — a mobile
    # keyboard autocapitalizing the email, or HR handing out the address
    # with different casing than it was created with, made an exact-case
    # match silently fail with "Invalid email or password" despite the
    # right credentials.
    email = str(payload.email).strip().lower()
    user = db.query(User).filter(func.lower(User.email) == email).first()
    if not user or not verify_password(payload.password, user.password):
        raise HTTPException(status.HTTP_401_UNAUTHORIZED, "Invalid email or password")
    if not user.is_active:
        raise HTTPException(status.HTTP_403_FORBIDDEN, "Account is disabled")
    return _token_response(user)


@router.post("/refresh", response_model=TokenResponse)
def refresh(payload: RefreshRequest, db: Session = Depends(get_db)):
    """Exchange a still-valid refresh token for a fresh access+refresh pair.

    The frontend (`src/lib/api.js`) has always called this on any 401 to
    silently keep the session alive past the 60-minute access-token expiry —
    but this endpoint never existed, so every expiry (or transient 401)
    forced a full logout back to the login page. Mirrors `login`'s response
    shape exactly so the frontend's existing `auth.save(res.data)` call works
    unchanged.
    """
    try:
        claims = jwt.decode(payload.refresh_token, JWT_SECRET, algorithms=[JWT_ALGORITHM])
        if claims.get("type") != "refresh":
            raise HTTPException(status.HTTP_401_UNAUTHORIZED, "Invalid token type")
        user_id = int(claims["sub"])
    except (JWTError, KeyError, ValueError):
        raise HTTPException(status.HTTP_401_UNAUTHORIZED, "Invalid or expired refresh token")

    user = db.query(User).filter(User.id == user_id, User.is_active.is_(True)).first()
    if not user:
        raise HTTPException(status.HTTP_401_UNAUTHORIZED, "User not found")
    return _token_response(user)


@router.post("/signup", status_code=status.HTTP_201_CREATED)
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
        # New self-service signups start INACTIVE — an HR admin must activate
        # the account (Employees page) before it can log in or call any API.
        # `get_current_user` checks `is_active` on every request, so this is
        # a real gate, not just a UI flag: an inactive account's token (if we
        # issued one) would be rejected everywhere. Previously this was
        # `True`, which let anyone who could reach the API self-register and
        # immediately query every employee's daily reports.
        is_active=False,
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
    # No tokens issued — the account can't do anything until HR activates it.
    return {
        "pending_approval": True,
        "message": "Account created. An HR admin needs to activate your account before you can log in.",
        "user": UserOut.model_validate(user).model_dump(),
    }


def _token_response(user: User) -> TokenResponse:
    return TokenResponse(
        access_token=create_access_token(user.id),
        refresh_token=create_refresh_token(user.id),
        user=UserOut.model_validate(user),
    )
