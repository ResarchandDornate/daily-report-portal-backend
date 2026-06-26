"""Add `site_name` to the expenses table.

Optional free-text label that ties an expense to a specific site /
project / customer location.  Empty string for legacy rows.

Revision ID: 0011_expense_site_name
Revises: 0010_monthly_expense_notes
Create Date: 2026-06-26
"""
from typing import Sequence, Union

from alembic import op
import sqlalchemy as sa


revision: str = "0011_expense_site_name"
down_revision: Union[str, None] = "0010_monthly_expense_notes"
branch_labels: Union[str, Sequence[str], None] = None
depends_on: Union[str, Sequence[str], None] = None


def upgrade() -> None:
    op.add_column(
        "expenses",
        sa.Column(
            "site_name",
            sa.String(length=255),
            nullable=False,
            server_default="",
        ),
    )


def downgrade() -> None:
    op.drop_column("expenses", "site_name")
