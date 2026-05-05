"""
Runtime invariants enforced via SQLAlchemy events.

These mirror the rules currently in `django_app/users/models.py` `def save()`,
so the same guarantees hold whether a user is saved via Django ORM, the FastAPI
JSON API, SQLAdmin, an Alembic data migration, or a raw shell script that uses
SQLAlchemy.

Rules:
  1. Auto-promote — assigning a user to the HR department (slug="hrDept")
     makes them role="hr" + is_superuser=True + is_staff=True.
  2. Hard rule — only superusers may hold role="hr"; any non-superuser
     attempt is silently downgraded to "employee".

Why both layers (Django save + SQLAlchemy events)?  During Phase 1+2 the codebase
still has Django paths.  Once Django is removed in Phase 3, the Django save()
override goes away and these events become the single enforcement point.
"""
from __future__ import annotations

from sqlalchemy import event, text
from sqlalchemy.engine import Connection
from sqlalchemy.orm.mapper import Mapper

from models import User


HR_DEPT_SLUG = "hrDept"


def _apply_user_invariants(
    mapper: Mapper,
    connection: Connection,
    user: User,
) -> None:
    """Runs in the same transaction as the INSERT/UPDATE.  Mutates `user`
    in-place so the eventual SQL reflects the corrected fields.
    """
    # ---- Rule 1: auto-promote HR-department members ----
    if user.department_id is not None:
        # Don't use user.department (relationship) — would emit an extra SELECT
        # outside this transaction's flush.  Use the connection directly.
        row = connection.execute(
            text("SELECT slug FROM departments_department WHERE id = :id"),
            {"id": user.department_id},
        ).fetchone()
        if row and row[0] == HR_DEPT_SLUG:
            user.role = "hr"
            user.is_superuser = True
            user.is_staff = True

    # ---- Rule 2: only superusers may be HR ----
    if user.role == "hr" and not user.is_superuser:
        user.role = "employee"


# Register the same handler on both INSERT and UPDATE.
event.listen(User, "before_insert", _apply_user_invariants)
event.listen(User, "before_update", _apply_user_invariants)
