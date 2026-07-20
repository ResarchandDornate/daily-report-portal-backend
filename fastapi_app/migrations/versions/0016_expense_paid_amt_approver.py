"""Add paid_amount + approved_by_name to expenses.

Both are captured on the HR "Issue Advance" form:

* paid_amount      — what was actually disbursed.  Separate from `advance`
                     because a partial payment is allowed (advance ₹5,000,
                     paid ₹4,000).  NULL = nothing paid out yet.
* approved_by_name — free-text record of who authorised the advance, for
                     advances authorised outside the normal approval flow.
                     Not linked to an account (unlike decided_by_id).

Revision ID: 0016_expense_paid_amt_approver
Revises: 0015_advance_sent_by_manager
Create Date: 2026-07-17
"""
from typing import Sequence, Union

from alembic import op
import sqlalchemy as sa


revision: str = "0016_expense_paid_amt_approver"
down_revision: Union[str, None] = "0015_advance_sent_by_manager"
branch_labels: Union[str, Sequence[str], None] = None
depends_on: Union[str, Sequence[str], None] = None


def upgrade() -> None:
    op.add_column(
        "expenses",
        sa.Column("paid_amount", sa.Integer(), nullable=True),
    )
    op.add_column(
        "expenses",
        sa.Column(
            "approved_by_name",
            sa.String(120),
            nullable=False,
            server_default="",
        ),
    )


def downgrade() -> None:
    op.drop_column("expenses", "approved_by_name")
    op.drop_column("expenses", "paid_amount")
