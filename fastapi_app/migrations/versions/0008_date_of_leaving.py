"""Add `date_of_leaving` to the user table.

Captured when HR (or a named approver) deactivates an employee — the
calendar date their employment ended.  Nullable because every existing
deactivated row was migrated without one, and re-activation should not
auto-populate it.

Revision ID: 0008_date_of_leaving
Revises: 0007_expenses_multi_bills
Create Date: 2026-06-24
"""
from typing import Sequence, Union

from alembic import op
import sqlalchemy as sa


revision: str = "0008_date_of_leaving"
down_revision: Union[str, None] = "0007_expenses_multi_bills"
branch_labels: Union[str, Sequence[str], None] = None
depends_on: Union[str, Sequence[str], None] = None


def upgrade() -> None:
    op.add_column(
        "users_user",
        sa.Column("date_of_leaving", sa.Date(), nullable=True),
    )


def downgrade() -> None:
    op.drop_column("users_user", "date_of_leaving")
