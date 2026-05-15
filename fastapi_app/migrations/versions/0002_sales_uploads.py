"""Create sales_uploads table for Inside Sales weekly/monthly Excel uploads.

The file bytes live in MinIO; this row carries metadata and a parsed
summary so the dashboard can render a preview without re-parsing.

Revision ID: 0002_sales_uploads
Revises: 0001_baseline
Create Date: 2026-05-15
"""
from typing import Sequence, Union

from alembic import op
import sqlalchemy as sa
from sqlalchemy.dialects import postgresql

revision: str = "0002_sales_uploads"
down_revision: Union[str, None] = "0001_baseline"
branch_labels: Union[str, Sequence[str], None] = None
depends_on: Union[str, Sequence[str], None] = None


def upgrade() -> None:
    op.create_table(
        "sales_uploads",
        sa.Column("id", sa.BigInteger(), nullable=False, autoincrement=True),
        sa.Column("user_id", sa.BigInteger(), nullable=False),
        sa.Column(
            "period_type",
            sa.String(length=16),
            nullable=False,
            server_default="weekly",
        ),
        sa.Column("period_start", sa.Date(), nullable=True),
        sa.Column("period_end", sa.Date(), nullable=True),
        sa.Column("note", sa.String(length=512), nullable=False, server_default=""),
        sa.Column("original_filename", sa.String(length=255), nullable=False),
        sa.Column("minio_object_key", sa.String(length=512), nullable=False),
        sa.Column(
            "file_size_bytes",
            sa.Integer(),
            nullable=False,
            server_default="0",
        ),
        sa.Column(
            "parsed_summary",
            postgresql.JSONB(astext_type=sa.Text()),
            nullable=False,
            server_default=sa.text("'{}'::jsonb"),
        ),
        sa.Column(
            "uploaded_at",
            sa.DateTime(timezone=True),
            nullable=False,
            server_default=sa.text("now()"),
        ),
        sa.ForeignKeyConstraint(
            ["user_id"], ["users_user.id"], ondelete="CASCADE"
        ),
        sa.PrimaryKeyConstraint("id"),
        sa.UniqueConstraint("minio_object_key", name="uq_sales_uploads_object_key"),
    )
    op.create_index(
        "ix_sales_uploads_user_id", "sales_uploads", ["user_id"], unique=False
    )
    op.create_index(
        "ix_sales_uploads_user_uploaded_at",
        "sales_uploads",
        ["user_id", "uploaded_at"],
        unique=False,
    )


def downgrade() -> None:
    op.drop_index("ix_sales_uploads_user_uploaded_at", table_name="sales_uploads")
    op.drop_index("ix_sales_uploads_user_id", table_name="sales_uploads")
    op.drop_table("sales_uploads")
