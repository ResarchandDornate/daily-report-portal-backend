"""
Authoritative SQLAlchemy models for the Daily Report Portal.

Table names match Django's auto-generated names so the same Postgres rows
are read/written by both stacks during the Phase 1/2 migration window.
After Phase 3 (Django removed), Alembic uses these models as the single
source of truth for the schema.

Column types, defaults, NOT-NULL flags, indexes, and unique constraints
are written to match exactly what Django created (see django_app/*/migrations
and the verified DB structure).  Anything that differs from this file once
Django is gone will surface as an Alembic autogenerate diff.

Auxiliary Django tables (auth_group, django_session, users_user_groups, etc.)
are NOT modelled here.  Alembic's env.py is configured to ignore them so
autogenerate doesn't try to drop them later.
"""
from sqlalchemy import (
    BigInteger,
    Boolean,
    Column,
    Date,
    DateTime,
    ForeignKey,
    Index,
    Integer,
    String,
    UniqueConstraint,
    text,
)
from sqlalchemy.dialects.postgresql import JSONB
from sqlalchemy.orm import relationship
from sqlalchemy.sql import func

from database import Base


class Department(Base):
    """Mirrors Django's `departments.Department` (table `departments_department`)."""

    __tablename__ = "departments_department"

    id = Column(BigInteger, primary_key=True, autoincrement=True)
    slug = Column(String(32), unique=True, nullable=False, index=True)
    name = Column(String(64), nullable=False)
    color = Column(String(16), nullable=False, default="zinc", server_default="zinc")
    # Each entry is `{"key": "fieldKey", "label": "Field Label"}`.
    report_fields = Column(
        JSONB, nullable=False, default=list, server_default=text("'[]'::jsonb")
    )
    created_at = Column(
        DateTime(timezone=True), nullable=False, server_default=func.now()
    )
    updated_at = Column(
        DateTime(timezone=True),
        nullable=False,
        server_default=func.now(),
        onupdate=func.now(),
    )

    employees = relationship("User", back_populates="department")

    def __repr__(self) -> str:
        return f"<Department slug={self.slug!r} name={self.name!r}>"


class User(Base):
    """Mirrors Django's custom user model (table `users_user`).

    Django's AbstractUser gives us username/password/email/etc.  Our portal
    extras (role, department FK, organisation, RM, DOJ) sit alongside.
    """

    __tablename__ = "users_user"

    id = Column(BigInteger, primary_key=True, autoincrement=True)

    # ------ Django auth fields ------
    password = Column(String(128), nullable=False)
    last_login = Column(DateTime(timezone=True), nullable=True)
    is_superuser = Column(
        Boolean, nullable=False, default=False, server_default=text("false")
    )
    username = Column(String(150), unique=True, nullable=False, index=True)
    first_name = Column(String(150), nullable=False, default="", server_default="")
    last_name = Column(String(150), nullable=False, default="", server_default="")
    email = Column(String(254), nullable=False, default="", server_default="", index=True)
    is_staff = Column(
        Boolean, nullable=False, default=False, server_default=text("false")
    )
    is_active = Column(
        Boolean, nullable=False, default=True, server_default=text("true")
    )
    date_joined = Column(
        DateTime(timezone=True), nullable=False, server_default=func.now()
    )

    # ------ Portal-specific ------
    contact_number = Column(String(20), nullable=False, default="", server_default="")
    role = Column(
        String(16), nullable=False, default="employee", server_default="employee"
    )
    title = Column(String(64), nullable=False, default="", server_default="")
    department_id = Column(
        BigInteger,
        ForeignKey("departments_department.id", ondelete="SET NULL"),
        nullable=True,
        index=True,
    )

    # ------ Imported from the HR roster spreadsheet ------
    organisation = Column(String(64), nullable=False, default="", server_default="")
    reporting_manager = Column(
        String(128), nullable=False, default="", server_default=""
    )
    date_of_joining = Column(Date, nullable=True)

    department = relationship("Department", back_populates="employees")
    daily_reports = relationship(
        "DailyReport", back_populates="user", cascade="all, delete-orphan"
    )

    def __repr__(self) -> str:
        return f"<User id={self.id} username={self.username!r} role={self.role!r}>"


class DailyReport(Base):
    """One report per (user, date).  `data` is a JSONB blob keyed by the
    department's report-field keys (e.g. {"meeting": "...", "revenue": "..."}).
    """

    __tablename__ = "reports_dailyreport"

    id = Column(BigInteger, primary_key=True, autoincrement=True)
    user_id = Column(
        BigInteger,
        ForeignKey("users_user.id", ondelete="CASCADE"),
        nullable=False,
        index=True,
    )
    date = Column(Date, nullable=False, index=True)
    data = Column(
        JSONB, nullable=False, default=dict, server_default=text("'{}'::jsonb")
    )
    submitted_at = Column(
        DateTime(timezone=True),
        nullable=False,
        server_default=func.now(),
        onupdate=func.now(),
    )
    created_at = Column(
        DateTime(timezone=True), nullable=False, server_default=func.now()
    )

    user = relationship("User", back_populates="daily_reports")

    __table_args__ = (
        UniqueConstraint("user_id", "date", name="uq_user_date"),
        # Composite index mirrors Django's `Meta.indexes = [Index(fields=["date", "user"])]`.
        Index("ix_reports_date_user", "date", "user_id"),
    )

    def __repr__(self) -> str:
        return f"<DailyReport id={self.id} user_id={self.user_id} date={self.date}>"


class SalesUpload(Base):
    """One uploaded Excel sheet (weekly / monthly calling log).

    The file itself lives in MinIO; this row carries the metadata and a
    parsed summary so the dashboard can render a preview without re-parsing
    the bytes every time.
    """

    __tablename__ = "sales_uploads"

    id = Column(BigInteger, primary_key=True, autoincrement=True)
    user_id = Column(
        BigInteger,
        ForeignKey("users_user.id", ondelete="CASCADE"),
        nullable=False,
        index=True,
    )
    period_type = Column(
        String(16), nullable=False, default="weekly", server_default="weekly"
    )  # "weekly" | "monthly" | "adhoc"
    period_start = Column(Date, nullable=True)
    period_end = Column(Date, nullable=True)
    note = Column(String(512), nullable=False, default="", server_default="")
    original_filename = Column(String(255), nullable=False)
    minio_object_key = Column(String(512), nullable=False, unique=True)
    file_size_bytes = Column(Integer, nullable=False, default=0, server_default="0")
    # Parsed summary from openpyxl — columns, row count, optional totals.
    parsed_summary = Column(
        JSONB, nullable=False, default=dict, server_default=text("'{}'::jsonb")
    )
    uploaded_at = Column(
        DateTime(timezone=True), nullable=False, server_default=func.now()
    )

    user = relationship("User")

    __table_args__ = (
        Index("ix_sales_uploads_user_uploaded_at", "user_id", "uploaded_at"),
    )

    def __repr__(self) -> str:
        return (
            f"<SalesUpload id={self.id} user_id={self.user_id} "
            f"file={self.original_filename!r}>"
        )
