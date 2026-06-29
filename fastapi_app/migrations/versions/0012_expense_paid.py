"""Add paid_at + paid_by_id to expenses.

Tracks the finance-team disbursal step — after Tarini approves a claim,
Shivangi clicks Paid to record that the money has actually gone out.
The status transitions to "paid" (terminal); these two columns capture
who marked it and when, kept separate from decided_*  so we preserve
both the approval and disbursal audit trail.

Revision ID: 0012_expense_paid
Revises: 0011_expense_site_name
Create Date: 2026-06-29
"""
from typing import Sequence, Union

from alembic import op
import sqlalchemy as sa


revision: str = "0012_expense_paid"
down_revision: Union[str, None] = "0011_expense_site_name"
branch_labels: Union[str, Sequence[str], None] = None
depends_on: Union[str, Sequence[str], None] = None


def upgrade() -> None:
    op.add_column(
        "expenses",
        sa.Column(
            "paid_by_id", sa.BigInteger(), nullable=True,
        ),
    )
    op.add_column(
        "expenses",
        sa.Column(
            "paid_at", sa.DateTime(timezone=True), nullable=True,
        ),
    )
    op.create_foreign_key(
        "fk_expenses_paid_by_id_users_user",
        "expenses", "users_user",
        ["paid_by_id"], ["id"],
        ondelete="SET NULL",
    )


def downgrade() -> None:
    op.drop_constraint(
        "fk_expenses_paid_by_id_users_user", "expenses", type_="foreignkey",
    )
    op.drop_column("expenses", "paid_at")
    op.drop_column("expenses", "paid_by_id")
