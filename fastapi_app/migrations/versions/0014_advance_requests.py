"""Create advance_requests table for travel advance approvals.

Revision ID: 0014_advance_requests
Revises: 0013_expense_payment_ref
Create Date: 2026-07-08
"""

from alembic import op
import sqlalchemy as sa
from sqlalchemy.dialects.postgresql import JSONB

revision = "0014_advance_requests"
down_revision = "0013_expense_payment_ref"
branch_labels = None
depends_on = None


def upgrade() -> None:
    op.create_table(
        "advance_requests",
        sa.Column("id", sa.Integer, primary_key=True),
        sa.Column("created_by_id", sa.Integer, sa.ForeignKey("users_user.id"), nullable=False),
        sa.Column("employee_names", JSONB, nullable=False, server_default=sa.text("'[]'::jsonb")),
        sa.Column("city", sa.String(255), nullable=False, server_default=""),
        sa.Column("travel_date", sa.Date, nullable=False),
        sa.Column("return_date", sa.Date, nullable=False),
        sa.Column("tour_days", sa.Integer, nullable=False, server_default="1"),
        sa.Column("mode_going", sa.String(100), nullable=False, server_default=""),
        sa.Column("mode_return", sa.String(100), nullable=False, server_default=""),
        sa.Column("purpose", sa.String(1024), nullable=False, server_default=""),
        sa.Column("accommodation_days", sa.Integer, nullable=False, server_default="0"),
        sa.Column("accommodation_rate", sa.Numeric(10, 2), nullable=False, server_default="0"),
        sa.Column("food_amount", sa.Numeric(10, 2), nullable=False, server_default="0"),
        sa.Column("conveyance_days", sa.Integer, nullable=False, server_default="0"),
        sa.Column("conveyance_rate", sa.Numeric(10, 2), nullable=False, server_default="0"),
        sa.Column("total_amount", sa.Numeric(10, 2), nullable=False, server_default="0"),
        sa.Column("status", sa.String(20), nullable=False, server_default="pending"),
        sa.Column("decision_note", sa.String(1024), nullable=False, server_default=""),
        sa.Column("decided_by_id", sa.Integer, sa.ForeignKey("users_user.id"), nullable=True),
        sa.Column("decided_at", sa.DateTime(timezone=True), nullable=True),
        sa.Column("created_at", sa.DateTime(timezone=True), nullable=False, server_default=sa.func.now()),
    )
    op.create_index("ix_advance_requests_created_by_id", "advance_requests", ["created_by_id"])
    op.create_index("ix_advance_requests_status", "advance_requests", ["status"])


def downgrade() -> None:
    op.drop_index("ix_advance_requests_status")
    op.drop_index("ix_advance_requests_created_by_id")
    op.drop_table("advance_requests")
