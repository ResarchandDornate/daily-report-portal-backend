"""Add payment_ref to expenses.

Stores the UTR / reference number entered by the finance approver (Shivangi)
when marking an expense as paid.  Also captures the exact payment date she
enters via the Mark-Paid modal, stored as paid_at override.

Revision ID: 0013_expense_payment_ref
Revises: 0012_expense_paid
Create Date: 2026-07-07
"""
from typing import Sequence, Union

from alembic import op
import sqlalchemy as sa


revision: str = "0013_expense_payment_ref"
down_revision: Union[str, None] = "0012_expense_paid"
branch_labels: Union[str, Sequence[str], None] = None
depends_on: Union[str, Sequence[str], None] = None


def upgrade() -> None:
    op.add_column(
        "expenses",
        sa.Column(
            "payment_ref",
            sa.String(255),
            nullable=False,
            server_default="",
        ),
    )


def downgrade() -> None:
    op.drop_column("expenses", "payment_ref")
