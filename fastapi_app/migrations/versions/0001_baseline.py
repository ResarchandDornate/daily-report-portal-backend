"""Baseline — captures the schema Django created for our 3 portal tables.

This migration is the starting point.  On a brand-new database, running
`alembic upgrade head` creates these tables.  On the existing DB (where
Django already created them) you run `alembic stamp head` instead, which
records this revision as applied without executing any DDL.

Auxiliary Django tables (auth_*, django_*, users_user_groups, users_user_user_permissions)
are NOT created here — Alembic ignores them via env.py.  Django owns those
during Phase 1+2.  After Django is removed in Phase 3, those tables can be
left in place (read-only) or dropped manually.

Revision ID: 0001_baseline
Revises:
Create Date: 2026-05-02
"""
from typing import Sequence, Union

from alembic import op
import sqlalchemy as sa
from sqlalchemy.dialects import postgresql

revision: str = "0001_baseline"
down_revision: Union[str, None] = None
branch_labels: Union[str, Sequence[str], None] = None
depends_on: Union[str, Sequence[str], None] = None


def upgrade() -> None:
    # ---- departments_department ----
    op.create_table(
        "departments_department",
        sa.Column("id", sa.BigInteger(), autoincrement=True, primary_key=True),
        sa.Column("slug", sa.String(length=32), nullable=False),
        sa.Column("name", sa.String(length=64), nullable=False),
        sa.Column("color", sa.String(length=16), nullable=False, server_default="zinc"),
        sa.Column(
            "report_fields",
            postgresql.JSONB(astext_type=sa.Text()),
            nullable=False,
            server_default=sa.text("'[]'::jsonb"),
        ),
        sa.Column(
            "created_at",
            sa.DateTime(timezone=True),
            nullable=False,
            server_default=sa.text("now()"),
        ),
        sa.Column(
            "updated_at",
            sa.DateTime(timezone=True),
            nullable=False,
            server_default=sa.text("now()"),
        ),
        sa.UniqueConstraint("slug", name="departments_department_slug_key"),
    )
    op.create_index(
        "ix_departments_department_slug",
        "departments_department",
        ["slug"],
        unique=True,
    )

    # ---- users_user ----
    op.create_table(
        "users_user",
        sa.Column("id", sa.BigInteger(), autoincrement=True, primary_key=True),
        sa.Column("password", sa.String(length=128), nullable=False),
        sa.Column("last_login", sa.DateTime(timezone=True), nullable=True),
        sa.Column(
            "is_superuser",
            sa.Boolean(),
            nullable=False,
            server_default=sa.text("false"),
        ),
        sa.Column("username", sa.String(length=150), nullable=False),
        sa.Column(
            "first_name", sa.String(length=150), nullable=False, server_default=""
        ),
        sa.Column(
            "last_name", sa.String(length=150), nullable=False, server_default=""
        ),
        sa.Column("email", sa.String(length=254), nullable=False, server_default=""),
        sa.Column(
            "is_staff", sa.Boolean(), nullable=False, server_default=sa.text("false")
        ),
        sa.Column(
            "is_active", sa.Boolean(), nullable=False, server_default=sa.text("true")
        ),
        sa.Column(
            "date_joined",
            sa.DateTime(timezone=True),
            nullable=False,
            server_default=sa.text("now()"),
        ),
        sa.Column(
            "contact_number",
            sa.String(length=20),
            nullable=False,
            server_default="",
        ),
        sa.Column(
            "role",
            sa.String(length=16),
            nullable=False,
            server_default="employee",
        ),
        sa.Column("title", sa.String(length=64), nullable=False, server_default=""),
        sa.Column("department_id", sa.BigInteger(), nullable=True),
        sa.Column(
            "organisation",
            sa.String(length=64),
            nullable=False,
            server_default="",
        ),
        sa.Column(
            "reporting_manager",
            sa.String(length=128),
            nullable=False,
            server_default="",
        ),
        sa.Column("date_of_joining", sa.Date(), nullable=True),
        sa.ForeignKeyConstraint(
            ["department_id"],
            ["departments_department.id"],
            name="users_user_department_id_fkey",
            ondelete="SET NULL",
        ),
        sa.UniqueConstraint("username", name="users_user_username_key"),
    )
    op.create_index("ix_users_user_username", "users_user", ["username"], unique=True)
    op.create_index("ix_users_user_email", "users_user", ["email"])
    op.create_index(
        "ix_users_user_department_id", "users_user", ["department_id"]
    )

    # ---- reports_dailyreport ----
    op.create_table(
        "reports_dailyreport",
        sa.Column("id", sa.BigInteger(), autoincrement=True, primary_key=True),
        sa.Column("user_id", sa.BigInteger(), nullable=False),
        sa.Column("date", sa.Date(), nullable=False),
        sa.Column(
            "data",
            postgresql.JSONB(astext_type=sa.Text()),
            nullable=False,
            server_default=sa.text("'{}'::jsonb"),
        ),
        sa.Column(
            "submitted_at",
            sa.DateTime(timezone=True),
            nullable=False,
            server_default=sa.text("now()"),
        ),
        sa.Column(
            "created_at",
            sa.DateTime(timezone=True),
            nullable=False,
            server_default=sa.text("now()"),
        ),
        sa.ForeignKeyConstraint(
            ["user_id"],
            ["users_user.id"],
            name="reports_dailyreport_user_id_fkey",
            ondelete="CASCADE",
        ),
        sa.UniqueConstraint("user_id", "date", name="uq_user_date"),
    )
    op.create_index(
        "ix_reports_dailyreport_user_id", "reports_dailyreport", ["user_id"]
    )
    op.create_index(
        "ix_reports_dailyreport_date", "reports_dailyreport", ["date"]
    )
    op.create_index(
        "ix_reports_date_user", "reports_dailyreport", ["date", "user_id"]
    )


def downgrade() -> None:
    op.drop_index("ix_reports_date_user", table_name="reports_dailyreport")
    op.drop_index("ix_reports_dailyreport_date", table_name="reports_dailyreport")
    op.drop_index("ix_reports_dailyreport_user_id", table_name="reports_dailyreport")
    op.drop_table("reports_dailyreport")

    op.drop_index("ix_users_user_department_id", table_name="users_user")
    op.drop_index("ix_users_user_email", table_name="users_user")
    op.drop_index("ix_users_user_username", table_name="users_user")
    op.drop_table("users_user")

    op.drop_index("ix_departments_department_slug", table_name="departments_department")
    op.drop_table("departments_department")
