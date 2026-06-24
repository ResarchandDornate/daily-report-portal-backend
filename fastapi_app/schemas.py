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
    is_team_head: bool = False
    # The department slug this team head manages, if it overrides their own.
    # Null when the team head manages their own department (default behavior)
    # or when is_team_head is False.
    team_head_dept: Optional[str] = None

    # Roster fields
    organisation: str = ""
    reporting_manager: str = ""
    date_of_joining: Optional[date] = None
    # Filled in when the user was deactivated (sent to "Employees Left").
    date_of_leaving: Optional[date] = None

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


class OrganisationOut(BaseModel):
    id: int
    name: str
    color: str = "zinc"

    class Config:
        from_attributes = True


class OrganisationCreate(BaseModel):
    name: str
    color: str = "zinc"


class OrganisationUpdate(BaseModel):
    name: Optional[str] = None
    color: Optional[str] = None


class DepartmentCreate(BaseModel):
    slug: str
    name: str
    color: str = "zinc"
    report_fields: list[FieldDef] = []


class DepartmentUpdate(BaseModel):
    name: Optional[str] = None
    color: Optional[str] = None
    report_fields: Optional[list[FieldDef]] = None


class EmployeeCreate(BaseModel):
    email: EmailStr
    username: Optional[str] = None  # defaults to local part of email
    password: Optional[str] = None  # defaults to "<firstname>@ornate" pattern
    first_name: str
    last_name: str = ""
    department: Optional[str] = None  # slug
    title: str = ""
    contact_number: str = ""
    role: str = "employee"  # "employee" | "hr"
    is_team_head: bool = False
    organisation: str = ""
    reporting_manager: str = ""
    date_of_joining: Optional[date] = None


class EmployeeUpdate(BaseModel):
    first_name: Optional[str] = None
    last_name: Optional[str] = None
    email: Optional[EmailStr] = None
    department: Optional[str] = None  # slug, or "" to unset
    title: Optional[str] = None
    contact_number: Optional[str] = None
    role: Optional[str] = None
    organisation: Optional[str] = None
    reporting_manager: Optional[str] = None
    date_of_joining: Optional[date] = None
    date_of_leaving: Optional[date] = None
    is_active: Optional[bool] = None
    is_team_head: Optional[bool] = None
    # Slug of the department this team head manages (overrides their own dept).
    # Empty string clears the override.  Ignored if is_team_head is false.
    team_head_dept: Optional[str] = None
    password: Optional[str] = None  # set only if HR wants to reset


# ---------- Daily Report ----------

class ReportIn(BaseModel):
    date: date
    data: dict[str, str] = {}
    # Optional override — only HR users may set this to submit on behalf of
    # another employee.  Non-HR callers must leave it null (default).
    user_id: int | None = None


class LeaveIn(BaseModel):
    start_date: date
    days: int = 1
    reason: str = ""
    user_id: int | None = None  # HR-only: apply leave on behalf of another


class ReportOut(BaseModel):
    id: int
    date: date
    user_id: int
    data: dict[str, str] = {}
    submitted_at: datetime

    class Config:
        from_attributes = True


class ReportListOut(BaseModel):
    items: list[ReportOut]
    total: int
    limit: int
    offset: int


# ---------- Sales Uploads (Inside Sales weekly/monthly Excel) ----------

class SalesUploadOut(BaseModel):
    id: int
    user_id: int
    user_name: str = ""
    period_type: str
    period_start: Optional[date] = None
    period_end: Optional[date] = None
    note: str = ""
    original_filename: str
    file_size_bytes: int
    parsed_summary: dict = {}
    uploaded_at: datetime

    class Config:
        from_attributes = True


class ExpenseBillOut(BaseModel):
    """One attached bill — `filename` is the original upload name, `index`
    is the bill's position in the expense's `bills` list (used by the
    `/api/expenses/{id}/bill/{index}` download endpoint).
    """
    index: int
    filename: str


class ExpenseOut(BaseModel):
    """Single expense row — both employee-facing and admin-facing views use
    this shape.  Bills are exposed via `bills[i].index` + download endpoint
    rather than the raw MinIO key so the frontend doesn't see internals.
    """
    id: int
    user_id: int
    user_name: str = ""
    user_department: str = ""
    date: date
    mode: str = ""
    expense_type: str
    travel_type: str = ""
    amount: int
    remarks: str = ""
    bills: list[ExpenseBillOut] = []
    status: str
    decided_by_id: Optional[int] = None
    decided_by_name: str = ""
    decided_at: Optional[datetime] = None
    decision_note: str = ""
    created_at: datetime

    class Config:
        from_attributes = True


class ExpenseDecideIn(BaseModel):
    """Approval / rejection / on-hold payload from the admin modal."""
    decision: str  # "approved" | "rejected" | "onhold"
    note: str = ""


class ExpensePatchIn(BaseModel):
    """Owner-only edit payload — only used when the expense is still in
    `pending` or `onhold` status.  Every field is optional; only the ones
    that are non-None get applied.  Bills are managed via separate upload /
    delete endpoints (left untouched on edit).
    """
    date: Optional[date] = None
    mode: Optional[str] = None
    expense_type: Optional[str] = None
    travel_type: Optional[str] = None
    amount: Optional[int] = None
    remarks: Optional[str] = None


TokenResponse.model_rebuild()
UserOut.model_rebuild()
