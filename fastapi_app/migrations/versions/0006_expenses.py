"""Create expenses table for employee expense claims with approval workflow.

Each row is a single expense submitted by an employee.  Optional bill
images are stored in MinIO; this row carries the metadata + the approval
status (pending / approved / rejected) plus decision audit fields.

Revision ID: 0006_expenses
Revises: 0005_team_head_dept
Create Date: 2026-06-22
"""
from typing import Sequence, Union

from alembic import op
import sqlalchemy as sa


revision: str = "0006_expenses"
down_revision: Union[str, None] = "0005_team_head_dept"
branch_labels: Union[str, Sequence[str], None] = None
depends_on: Union[str, Sequence[str], None] = None


def upgrade() -> None:
    op.create_table(
        "expenses",
        sa.Column("id", sa.BigInteger(), nullable=False, autoincrement=True),
        sa.Column("user_id", sa.BigInteger(), nullable=False),
        sa.Column("date", sa.Date(), nullable=False),
        sa.Column("mode", sa.String(length=32), nullable=False, server_default=""),
        sa.Column("expense_type", sa.String(length=32), nullable=False),
        sa.Column(
            "travel_type", sa.String(length=32), nullable=False, server_default=""
        ),
        sa.Column(
            "amount", sa.Integer(), nullable=False, server_default="0"
        ),
        sa.Column(
            "remarks", sa.String(length=1024), nullable=False, server_default=""
        ),
        sa.Column(
            "bill_filename",
            sa.String(length=255),
            nullable=False,
            server_default="",
        ),
        sa.Column(
            "minio_object_key",
            sa.String(length=512),
            nullable=False,
            server_default="",
        ),
        sa.Column(
            "status",
            sa.String(length=16),
            nullable=False,
            server_default="pending",
        ),
        sa.Column("decided_by_id", sa.BigInteger(), nullable=True),
        sa.Column("decided_at", sa.DateTime(timezone=True), nullable=True),
        sa.Column(
            "decision_note",
            sa.String(length=512),
            nullable=False,
            server_default="",
        ),
        sa.Column(
            "created_at",
            sa.DateTime(timezone=True),
            nullable=False,
            server_default=sa.text("now()"),
        ),
        sa.ForeignKeyConstraint(
            ["user_id"], ["users_user.id"], ondelete="CASCADE"
        ),
        sa.ForeignKeyConstraint(
            ["decided_by_id"], ["users_user.id"], ondelete="SET NULL"
        ),
        sa.PrimaryKeyConstraint("id"),
    )
    op.create_index("ix_expenses_user_id", "expenses", ["user_id"], unique=False)
    op.create_index("ix_expenses_date", "expenses", ["date"], unique=False)
    op.create_index(
        "ix_expenses_user_date", "expenses", ["user_id", "date"], unique=False
    )
    op.create_index("ix_expenses_status", "expenses", ["status"], unique=False)


def downgrade() -> None:
    op.drop_index("ix_expenses_status", table_name="expenses")
    op.drop_index("ix_expenses_user_date", table_name="expenses")
    op.drop_index("ix_expenses_date", table_name="expenses")
    op.drop_index("ix_expenses_user_id", table_name="expenses")
    op.drop_table("expenses")
