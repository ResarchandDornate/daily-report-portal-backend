"""Create monthly_expense_notes table — HR's per-employee, per-month
advance + remarks for the expense Monthly Summary modal.

This is the canonical storage for advances (previously kept only in the
admin's browser localStorage) plus a remarks field HR can use to tell an
employee why their payment is on hold for the month.

Revision ID: 0010_monthly_expense_notes
Revises: 0009_expense_advance
Create Date: 2026-06-25
"""
from typing import Sequence, Union

from alembic import op
import sqlalchemy as sa


revision: str = "0010_monthly_expense_notes"
down_revision: Union[str, None] = "0009_expense_advance"
branch_labels: Union[str, Sequence[str], None] = None
depends_on: Union[str, Sequence[str], None] = None


def upgrade() -> None:
    op.create_table(
        "monthly_expense_notes",
        sa.Column("id", sa.BigInteger(), nullable=False, autoincrement=True),
        sa.Column("user_id", sa.BigInteger(), nullable=False),
        sa.Column("year", sa.Integer(), nullable=False),
        sa.Column("month", sa.Integer(), nullable=False),  # 1..12
        sa.Column("advance", sa.Integer(), nullable=False, server_default="0"),
        sa.Column("remark", sa.String(length=2048), nullable=False, server_default=""),
        sa.Column(
            "updated_at",
            sa.DateTime(timezone=True),
            nullable=False,
            server_default=sa.text("now()"),
        ),
        sa.Column("updated_by_id", sa.BigInteger(), nullable=True),
        sa.ForeignKeyConstraint(
            ["user_id"], ["users_user.id"], ondelete="CASCADE"
        ),
        sa.ForeignKeyConstraint(
            ["updated_by_id"], ["users_user.id"], ondelete="SET NULL"
        ),
        sa.PrimaryKeyConstraint("id"),
        sa.UniqueConstraint(
            "user_id", "year", "month", name="uq_monthly_expense_notes_user_period"
        ),
    )
    op.create_index(
        "ix_monthly_expense_notes_period",
        "monthly_expense_notes",
        ["year", "month"],
        unique=False,
    )


def downgrade() -> None:
    op.drop_index("ix_monthly_expense_notes_period", table_name="monthly_expense_notes")
    op.drop_table("monthly_expense_notes")
