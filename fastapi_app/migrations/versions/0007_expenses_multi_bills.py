"""Convert Expense single-bill columns to a JSONB `bills` array.

Drops `bill_filename` + `minio_object_key` (singular) and replaces with a
JSONB `bills` list of `{"filename": str, "object_key": str}` objects so each
expense can carry multiple receipt photos / PDFs.

Any existing row with a non-empty `bill_filename` is migrated into a
single-element `bills` list during upgrade — and unpacked back into the
two singular columns on downgrade.

Revision ID: 0007_expenses_multi_bills
Revises: 0006_expenses
Create Date: 2026-06-22
"""
from typing import Sequence, Union

from alembic import op
import sqlalchemy as sa
from sqlalchemy.dialects import postgresql


revision: str = "0007_expenses_multi_bills"
down_revision: Union[str, None] = "0006_expenses"
branch_labels: Union[str, Sequence[str], None] = None
depends_on: Union[str, Sequence[str], None] = None


def upgrade() -> None:
    op.add_column(
        "expenses",
        sa.Column(
            "bills",
            postgresql.JSONB(astext_type=sa.Text()),
            nullable=False,
            server_default=sa.text("'[]'::jsonb"),
        ),
    )
    # Backfill the new column from the legacy singular columns.
    op.execute(
        """
        UPDATE expenses
           SET bills = jsonb_build_array(
                 jsonb_build_object(
                   'filename', bill_filename,
                   'object_key', minio_object_key
                 )
               )
         WHERE bill_filename <> ''
           AND minio_object_key <> ''
        """
    )
    op.drop_column("expenses", "bill_filename")
    op.drop_column("expenses", "minio_object_key")


def downgrade() -> None:
    op.add_column(
        "expenses",
        sa.Column(
            "bill_filename",
            sa.String(length=255),
            nullable=False,
            server_default="",
        ),
    )
    op.add_column(
        "expenses",
        sa.Column(
            "minio_object_key",
            sa.String(length=512),
            nullable=False,
            server_default="",
        ),
    )
    # Unpack the FIRST bill back into the singular columns (anything beyond
    # the first is lost on downgrade — that's the price of going back).
    op.execute(
        """
        UPDATE expenses
           SET bill_filename    = COALESCE(bills->0->>'filename', ''),
               minio_object_key = COALESCE(bills->0->>'object_key', '')
         WHERE jsonb_array_length(bills) > 0
        """
    )
    op.drop_column("expenses", "bills")
