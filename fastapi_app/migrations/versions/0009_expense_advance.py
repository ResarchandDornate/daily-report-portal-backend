"""Add `advance` column to the expenses table.

Captures any cash / UPI advance that HR (or another admin) paid the
employee BEFORE the actual expense was incurred — used to compute the
net reimbursement owed to the employee.

Revision ID: 0009_expense_advance
Revises: 0008_date_of_leaving
Create Date: 2026-06-25
"""
from typing import Sequence, Union

from alembic import op
import sqlalchemy as sa


revision: str = "0009_expense_advance"
down_revision: Union[str, None] = "0008_date_of_leaving"
branch_labels: Union[str, Sequence[str], None] = None
depends_on: Union[str, Sequence[str], None] = None


def upgrade() -> None:
    op.add_column(
        "expenses",
        sa.Column(
            "advance",
            sa.Integer(),
            nullable=False,
            server_default="0",
        ),
    )


def downgrade() -> None:
    op.drop_column("expenses", "advance")
