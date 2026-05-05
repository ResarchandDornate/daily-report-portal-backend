"""
SQLAdmin — replacement for Django admin, mounted at /admin on FastAPI.

Auth flow:
  1. User visits /admin → SQLAdmin redirects to /admin/login
  2. They enter email + password (login form's "username" field is treated as email)
  3. We validate against users_user using the same PBKDF2 verifier as the JSON API,
     then check role=='hr' AND is_superuser==True
  4. On success a signed session cookie is set; subsequent /admin pages auth via that

Cookies are signed with JWT_SECRET so we don't introduce a new key to manage.
SQLAdmin auto-installs Starlette's SessionMiddleware when an AuthenticationBackend
is supplied, so we don't add it manually elsewhere.

Note: this admin is intentionally restrictive.  Only HR superusers reach it.
Regular employees + non-HR staff are bounced.
"""
from __future__ import annotations

import os
from typing import Any

from sqladmin import Admin, ModelView
from sqladmin.authentication import AuthenticationBackend
from starlette.requests import Request

from auth import hash_password, verify_password
from database import SessionLocal, engine
from models import DailyReport, Department, User


# ---------------------------------------------------------------------------
# Auth
# ---------------------------------------------------------------------------

class AdminAuth(AuthenticationBackend):
    """Re-uses the same password hashing & user model the JSON API uses."""

    async def login(self, request: Request) -> bool:
        form = await request.form()
        # SQLAdmin's login form labels these "username" + "password".  We treat
        # "username" as the email so HR types their dashboard email.
        email = (form.get("username") or "").strip()
        password = form.get("password") or ""
        if not email or not password:
            return False

        db = SessionLocal()
        try:
            user = db.query(User).filter(User.email == email).first()
            if not user or not verify_password(password, user.password):
                return False
            # Only HR superusers may use the admin.  Regular employees are blocked
            # even if they somehow had role=hr without is_superuser.
            if not (user.is_active and user.is_superuser and user.role == "hr"):
                return False
            request.session["user_id"] = user.id
            return True
        finally:
            db.close()

    async def logout(self, request: Request) -> bool:
        request.session.clear()
        return True

    async def authenticate(self, request: Request) -> bool:
        user_id = request.session.get("user_id")
        if not user_id:
            return False
        # Re-validate per request — protects against deactivated/demoted users
        # whose cookie hasn't expired yet.
        db = SessionLocal()
        try:
            user = db.query(User).filter(User.id == user_id).first()
            if not user or not (user.is_active and user.is_superuser and user.role == "hr"):
                request.session.clear()
                return False
            return True
        finally:
            db.close()


# ---------------------------------------------------------------------------
# Views
# ---------------------------------------------------------------------------

class DepartmentAdmin(ModelView, model=Department):
    name = "Department"
    name_plural = "Departments"
    icon = "fa-solid fa-sitemap"
    category = "Org"

    column_list = [
        Department.id,
        Department.slug,
        Department.name,
        Department.color,
        Department.report_fields,
        Department.created_at,
    ]
    column_searchable_list = [Department.slug, Department.name]
    column_sortable_list = [Department.id, Department.slug, Department.name]
    column_default_sort = [(Department.name, False)]

    form_columns = [
        Department.slug,
        Department.name,
        Department.color,
        Department.report_fields,
    ]


class UserAdmin(ModelView, model=User):
    name = "User"
    name_plural = "Users"
    icon = "fa-solid fa-user"
    category = "Org"

    column_list = [
        User.id,
        User.username,
        User.email,
        User.first_name,
        User.last_name,
        User.role,
        User.department,
        User.title,
        User.organisation,
        User.is_active,
        User.is_superuser,
    ]
    column_searchable_list = [
        User.username,
        User.email,
        User.first_name,
        User.last_name,
        User.reporting_manager,
    ]
    column_sortable_list = [
        User.id,
        User.username,
        User.email,
        User.role,
        User.is_active,
        User.date_joined,
    ]
    column_default_sort = [(User.username, False)]

    form_columns = [
        User.username,
        User.email,
        User.password,  # hashed automatically by on_model_change below
        User.first_name,
        User.last_name,
        User.role,
        User.department,
        User.title,
        User.contact_number,
        User.organisation,
        User.reporting_manager,
        User.date_of_joining,
        User.is_active,
        User.is_staff,
        User.is_superuser,
    ]

    # Show password field as type="password" so it's not visible by default.
    form_widget_args = {
        "password": {"type": "password", "placeholder": "Leave blank to keep current"},
    }

    async def on_model_change(
        self,
        data: dict,
        model: Any,
        is_created: bool,
        request: Request,
    ) -> None:
        """Hash plaintext passwords and apply our HR-promotion rule."""
        new_pwd = (data.get("password") or "").strip()
        if new_pwd:
            if not new_pwd.startswith("pbkdf2_sha256$"):
                # Admin typed a fresh password → hash it
                model.password = hash_password(new_pwd)
            # else: user pasted a hash directly (rare, allowed)
        elif is_created and not model.password:
            # Brand-new user with no password set → mark unusable.  They can be
            # given one later from this form, or via the dashboard signup flow.
            model.password = "!"
        # On edit with blank password, keep the existing hash by doing nothing
        # (SQLAlchemy won't update the column if model.password is unchanged).


class DailyReportAdmin(ModelView, model=DailyReport):
    name = "Daily Report"
    name_plural = "Daily Reports"
    icon = "fa-solid fa-clipboard"
    category = "Reports"

    column_list = [
        DailyReport.id,
        DailyReport.date,
        DailyReport.user,
        DailyReport.data,
        DailyReport.submitted_at,
    ]
    column_searchable_list = []  # JSON column searches don't work via simple LIKE
    column_sortable_list = [DailyReport.id, DailyReport.date, DailyReport.submitted_at]
    column_default_sort = [(DailyReport.date, True)]  # newest first

    form_columns = [
        DailyReport.user,
        DailyReport.date,
        DailyReport.data,
    ]


# ---------------------------------------------------------------------------
# Setup helper called from main.py
# ---------------------------------------------------------------------------

def setup_admin(app) -> Admin:
    """Mount SQLAdmin on the FastAPI app at /admin and register all views."""
    secret = os.environ.get("JWT_SECRET", "dev-secret-change-me")
    auth_backend = AdminAuth(secret_key=secret)

    admin = Admin(
        app=app,
        engine=engine,
        authentication_backend=auth_backend,
        title="Daily Report Portal — Admin",
        base_url="/admin",
    )
    admin.add_view(DepartmentAdmin)
    admin.add_view(UserAdmin)
    admin.add_view(DailyReportAdmin)
    return admin
