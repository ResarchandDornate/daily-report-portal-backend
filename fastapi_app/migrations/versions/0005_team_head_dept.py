"""Add team_head_dept_id override column to users_user.

When set, this team head manages employees in the referenced department
rather than their own.  Used when an employee sits in a "Sales Head"
department but files reports for the Sales team members.

Revision ID: 0005_team_head_dept
Revises: 0004_team_head_flag
Create Date: 2026-05-28
"""
from typing import Sequence, Union

from alembic import op
import sqlalchemy as sa

revision: str = "0005_team_head_dept"
down_revision: Union[str, None] = "0004_team_head_flag"
branch_labels: Union[str, Sequence[str], None] = None
depends_on: Union[str, Sequence[str], None] = None


def upgrade() -> None:
    op.add_column(
        "users_user",
        sa.Column("team_head_dept_id", sa.BigInteger(), nullable=True),
    )
    op.create_foreign_key(
        "fk_users_user_team_head_dept_id",
        "users_user",
        "departments_department",
        ["team_head_dept_id"],
        ["id"],
        ondelete="SET NULL",
    )


def downgrade() -> None:
    op.drop_constraint(
        "fk_users_user_team_head_dept_id", "users_user", type_="foreignkey"
    )
    op.drop_column("users_user", "team_head_dept_id")
