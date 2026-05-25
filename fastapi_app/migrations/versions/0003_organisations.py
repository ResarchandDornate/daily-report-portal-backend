"""Create organisations table.

This is the master list HR uses when assigning an employee's organisation.
The User.organisation column remains a free-text string carrying the
chosen name; this table just exists so HR can manage the list centrally.

On upgrade we also backfill the table from whatever distinct organisation
strings already exist on users_user, so the dropdown isn't empty after
the migration.

Revision ID: 0003_organisations
Revises: 0002_sales_uploads
Create Date: 2026-05-21
"""
from typing import Sequence, Union

from alembic import op
import sqlalchemy as sa

revision: str = "0003_organisations"
down_revision: Union[str, None] = "0002_sales_uploads"
branch_labels: Union[str, Sequence[str], None] = None
depends_on: Union[str, Sequence[str], None] = None


def upgrade() -> None:
    op.create_table(
        "organisations",
        sa.Column("id", sa.BigInteger(), nullable=False, autoincrement=True),
        sa.Column("name", sa.String(length=64), nullable=False),
        sa.Column(
            "color", sa.String(length=16), nullable=False, server_default="zinc"
        ),
        sa.Column(
            "created_at",
            sa.DateTime(timezone=True),
            nullable=False,
            server_default=sa.text("now()"),
        ),
        sa.Column(
            "updated_at",
            sa.DateTime(timezone=True),
            nullable=False,
            server_default=sa.text("now()"),
        ),
        sa.PrimaryKeyConstraint("id"),
        sa.UniqueConstraint("name", name="uq_organisations_name"),
    )

    # Backfill from existing distinct user.organisation values so the list
    # is already populated for HR on first load.
    op.execute(
        """
        INSERT INTO organisations (name, color, created_at, updated_at)
        SELECT DISTINCT TRIM(organisation), 'zinc', NOW(), NOW()
        FROM users_user
        WHERE organisation IS NOT NULL AND TRIM(organisation) <> ''
        ON CONFLICT (name) DO NOTHING;
        """
    )


def downgrade() -> None:
    op.drop_table("organisations")
