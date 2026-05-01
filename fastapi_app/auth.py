"""JWT helpers + password hashing compatible with Django's PBKDF2 default.

Django stores passwords as `pbkdf2_sha256$<iters>$<salt>$<b64hash>`. We verify
against that format so a user created via Django admin can log in via FastAPI,
and vice-versa.
"""
import base64
import hashlib
import os
import secrets
from datetime import datetime, timedelta, timezone
from typing import Optional

from fastapi import Depends, HTTPException, status
from fastapi.security import OAuth2PasswordBearer
from jose import JWTError, jwt
from sqlalchemy.orm import Session

from database import get_db
from models import User

JWT_SECRET = os.environ.get("JWT_SECRET", "dev-secret-change-me")
JWT_ALGORITHM = os.environ.get("JWT_ALGORITHM", "HS256")
ACCESS_MINUTES = int(os.environ.get("JWT_ACCESS_TOKEN_MINUTES", "60"))
REFRESH_DAYS = int(os.environ.get("JWT_REFRESH_TOKEN_DAYS", "7"))

oauth2_scheme = OAuth2PasswordBearer(tokenUrl="/api/auth/login", auto_error=False)


# ---------- Password (Django-compatible PBKDF2) ----------

def _pbkdf2_hash(password: str, salt: str, iterations: int) -> str:
    h = hashlib.pbkdf2_hmac("sha256", password.encode(), salt.encode(), iterations)
    return base64.b64encode(h).decode().strip()


def hash_password(password: str, iterations: int = 600_000) -> str:
    salt = secrets.token_urlsafe(16)
    digest = _pbkdf2_hash(password, salt, iterations)
    return f"pbkdf2_sha256${iterations}${salt}${digest}"


def verify_password(password: str, encoded: str) -> bool:
    if not encoded or not encoded.startswith("pbkdf2_sha256$"):
        return False
    try:
        _, iter_str, salt, digest = encoded.split("$", 3)
        iterations = int(iter_str)
    except ValueError:
        return False
    return _pbkdf2_hash(password, salt, iterations) == digest


# ---------- JWT ----------

def create_access_token(user_id: int) -> str:
    return jwt.encode(
        {"sub": str(user_id), "type": "access",
         "exp": datetime.now(timezone.utc) + timedelta(minutes=ACCESS_MINUTES)},
        JWT_SECRET, algorithm=JWT_ALGORITHM,
    )


def create_refresh_token(user_id: int) -> str:
    return jwt.encode(
        {"sub": str(user_id), "type": "refresh",
         "exp": datetime.now(timezone.utc) + timedelta(days=REFRESH_DAYS)},
        JWT_SECRET, algorithm=JWT_ALGORITHM,
    )


def get_current_user(
    token: Optional[str] = Depends(oauth2_scheme),
    db: Session = Depends(get_db),
) -> User:
    if not token:
        raise HTTPException(status.HTTP_401_UNAUTHORIZED, "Not authenticated")
    try:
        payload = jwt.decode(token, JWT_SECRET, algorithms=[JWT_ALGORITHM])
        if payload.get("type") != "access":
            raise HTTPException(status.HTTP_401_UNAUTHORIZED, "Invalid token type")
        user_id = int(payload["sub"])
    except (JWTError, KeyError, ValueError):
        raise HTTPException(status.HTTP_401_UNAUTHORIZED, "Invalid token")

    user = db.query(User).filter(User.id == user_id, User.is_active.is_(True)).first()
    if not user:
        raise HTTPException(status.HTTP_401_UNAUTHORIZED, "User not found")
    return user
