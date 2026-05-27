"""Add is_team_head flag to users_user.

A team-head user is a non-HR employee who can submit / edit daily reports
on behalf of any colleague in their own department.  Used for team leads
who file reports for the whole team (e.g. Justina in Sales).

Revision ID: 0004_team_head_flag
Revises: 0003_organisations
Create Date: 2026-05-27
"""
from typing import Sequence, Union

from alembic import op
import sqlalchemy as sa

revision: str = "0004_team_head_flag"
down_revision: Union[str, None] = "0003_organisations"
branch_labels: Union[str, Sequence[str], None] = None
depends_on: Union[str, Sequence[str], None] = None


def upgrade() -> None:
    op.add_column(
        "users_user",
        sa.Column(
            "is_team_head",
            sa.Boolean(),
            nullable=False,
            server_default=sa.text("false"),
        ),
    )


def downgrade() -> None:
    op.drop_column("users_user", "is_team_head")
