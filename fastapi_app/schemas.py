"""Pydantic request/response shapes."""
from datetime import date, datetime
from typing import Optional

from pydantic import BaseModel, EmailStr


# ---------- Auth ----------

class LoginRequest(BaseModel):
    email: EmailStr
    password: str


class SignupRequest(BaseModel):
    email: EmailStr
    username: str
    password: str
    first_name: str
    last_name: str
    contact_number: Optional[str] = ""
    department: Optional[str] = None  # department slug


class TokenResponse(BaseModel):
    access_token: str
    refresh_token: str
    token_type: str = "bearer"
    user: "UserOut"


# ---------- User ----------

class UserOut(BaseModel):
    id: int
    username: str
    email: str
    first_name: str
    last_name: str
    role: str
    title: str
    contact_number: str
    department: Optional["DepartmentOut"] = None

    # Roster fields
    organisation: str = ""
    reporting_manager: str = ""
    date_of_joining: Optional[date] = None

    class Config:
        from_attributes = True


# ---------- Department ----------

class FieldDef(BaseModel):
    key: str
    label: str


class DepartmentOut(BaseModel):
    id: int
    slug: str
    name: str
    color: str
    report_fields: list[FieldDef] = []

    class Config:
        from_attributes = True


# ---------- Daily Report ----------

class ReportIn(BaseModel):
    date: date
    data: dict[str, str] = {}


class ReportOut(BaseModel):
    id: int
    date: date
    user_id: int
    data: dict[str, str] = {}
    submitted_at: datetime

    class Config:
        from_attributes = True


TokenResponse.model_rebuild()
UserOut.model_rebuild()
