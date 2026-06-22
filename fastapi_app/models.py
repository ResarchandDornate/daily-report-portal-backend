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

    # Disambiguate: User has TWO FKs to this table (department_id for the
    # user's own dept, team_head_dept_id for the dept they manage).  This
    # relationship is the membership one — employees IN this department.
    employees = relationship(
        "User",
        back_populates="department",
        foreign_keys="User.department_id",
    )

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
    # When True, this user can submit/edit daily reports on behalf of any
    # employee in the SAME department.  Lets a team lead (e.g. Justina in
    # Sales) file reports for colleagues without giving them full HR powers.
    is_team_head = Column(
        Boolean, nullable=False, default=False, server_default=text("false")
    )
    # Optional override: when set, this team head manages employees in a
    # DIFFERENT department than their own.  Used when an employee sits in
    # a "Sales Head" / "Operations Head" department but is responsible for
    # filing reports for the underlying Sales / Operations team members.
    # When NULL, the team head manages their own department (default).
    team_head_dept_id = Column(
        BigInteger,
        ForeignKey("departments_department.id", ondelete="SET NULL"),
        nullable=True,
    )

    @property
    def team_head_dept(self):
        """Convenience accessor — slug of the managed department or None."""
        managed = getattr(self, "team_head_department", None)
        return managed.slug if managed else None

    # ------ Imported from the HR roster spreadsheet ------
    organisation = Column(String(64), nullable=False, default="", server_default="")
    reporting_manager = Column(
        String(128), nullable=False, default="", server_default=""
    )
    date_of_joining = Column(Date, nullable=True)

    department = relationship(
        "Department",
        back_populates="employees",
        foreign_keys=[department_id],
    )
    team_head_department = relationship(
        "Department",
        foreign_keys=[team_head_dept_id],
    )
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


class Organisation(Base):
    """Master list of organisations HR can pick from when adding employees.

    Currently `User.organisation` is a free-text column carrying the same
    value as `Organisation.name`.  This table just enables central
    management — add / rename / delete — without needing every employee
    row to be updated when an organisation's name changes (renames are
    a follow-up via UPDATE if needed).
    """

    __tablename__ = "organisations"

    id = Column(BigInteger, primary_key=True, autoincrement=True)
    name = Column(String(64), unique=True, nullable=False)
    color = Column(String(16), nullable=False, default="zinc", server_default="zinc")
    created_at = Column(
        DateTime(timezone=True), nullable=False, server_default=func.now()
    )
    updated_at = Column(
        DateTime(timezone=True),
        nullable=False,
        server_default=func.now(),
        onupdate=func.now(),
    )

    def __repr__(self) -> str:
        return f"<Organisation name={self.name!r}>"


class Expense(Base):
    """One expense claim submitted by an employee.

    Bills (receipt images) live in MinIO under `minio_object_key`; this row
    carries the metadata + the approval status.  Approval is decided by HR
    (or by the dedicated approvers TARINI / SMITA — see routers/expenses.py).
    """

    __tablename__ = "expenses"

    id = Column(BigInteger, primary_key=True, autoincrement=True)
    user_id = Column(
        BigInteger,
        ForeignKey("users_user.id", ondelete="CASCADE"),
        nullable=False,
        index=True,
    )
    date = Column(Date, nullable=False, index=True)
    # Payment mode — Cash / UPI / Card / Bank Transfer / Other.
    mode = Column(String(32), nullable=False, default="", server_default="")
    # Expense category — material / food / travel / hotel / fuel / others.
    expense_type = Column(String(32), nullable=False)
    # Sub-type used only when expense_type == "travel"
    # (bus / cab / bike / rapido / car / other).
    travel_type = Column(String(32), nullable=False, default="", server_default="")
    amount = Column(Integer, nullable=False, default=0, server_default="0")
    remarks = Column(String(1024), nullable=False, default="", server_default="")
    # Receipt files — JSONB list of `{"filename": str, "object_key": str}`.
    # Empty list when no bills were attached.  Supports multiple uploads per
    # expense (HR + employees both asked for this).
    bills = Column(
        JSONB, nullable=False, default=list, server_default=text("'[]'::jsonb")
    )
    # Approval lifecycle: pending → approved | rejected.
    status = Column(
        String(16), nullable=False, default="pending", server_default="pending"
    )
    decided_by_id = Column(
        BigInteger,
        ForeignKey("users_user.id", ondelete="SET NULL"),
        nullable=True,
    )
    decided_at = Column(DateTime(timezone=True), nullable=True)
    decision_note = Column(
        String(512), nullable=False, default="", server_default=""
    )
    created_at = Column(
        DateTime(timezone=True), nullable=False, server_default=func.now()
    )

    user = relationship("User", foreign_keys=[user_id])
    decided_by = relationship("User", foreign_keys=[decided_by_id])

    __table_args__ = (
        Index("ix_expenses_user_date", "user_id", "date"),
        Index("ix_expenses_status", "status"),
    )

    def __repr__(self) -> str:
        return (
            f"<Expense id={self.id} user_id={self.user_id} "
            f"type={self.expense_type!r} amount={self.amount} "
            f"status={self.status!r}>"
        )


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
