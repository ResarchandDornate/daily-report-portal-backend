"""Add sent_by_manager column to advance_requests.

Free-text name of the manager who sanctioned the trip — HR asked for
this so they can see who authorised the travel before approving the
advance amount.

Revision ID: 0015_advance_sent_by_manager
Revises: 0014_advance_requests
Create Date: 2026-07-09
"""
from typing import Sequence, Union

from alembic import op
import sqlalchemy as sa


revision: str = "0015_advance_sent_by_manager"
down_revision: Union[str, None] = "0014_advance_requests"
branch_labels: Union[str, Sequence[str], None] = None
depends_on: Union[str, Sequence[str], None] = None


def upgrade() -> None:
    op.add_column(
        "advance_requests",
        sa.Column(
            "sent_by_manager",
            sa.String(length=255),
            nullable=False,
            server_default="",
        ),
    )


def downgrade() -> None:
    op.drop_column("advance_requests", "sent_by_manager")
