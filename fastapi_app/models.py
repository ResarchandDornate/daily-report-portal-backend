"""SQLAlchemy models that mirror Django's tables — table names match
Django's auto-generated names so we read/write the same rows.
"""
from sqlalchemy import (
    Boolean, Column, Date, DateTime, ForeignKey, Integer, String, UniqueConstraint,
)
from sqlalchemy.dialects.postgresql import JSONB
from sqlalchemy.orm import relationship

from database import Base


class Department(Base):
    __tablename__ = "departments_department"

    id = Column(Integer, primary_key=True)
    slug = Column(String(32), unique=True, nullable=False)
    name = Column(String(64), nullable=False)
    color = Column(String(16), default="zinc")
    report_fields = Column(JSONB, default=list)
    created_at = Column(DateTime, nullable=False)
    updated_at = Column(DateTime, nullable=False)

    employees = relationship("User", back_populates="department")


class User(Base):
    __tablename__ = "users_user"

    id = Column(Integer, primary_key=True)
    password = Column(String(128), nullable=False)
    last_login = Column(DateTime, nullable=True)
    is_superuser = Column(Boolean, default=False)
    username = Column(String(150), unique=True, nullable=False)
    first_name = Column(String(150), default="")
    last_name = Column(String(150), default="")
    email = Column(String(254), default="")
    is_staff = Column(Boolean, default=False)
    is_active = Column(Boolean, default=True)
    date_joined = Column(DateTime, nullable=False)

    contact_number = Column(String(20), default="")
    role = Column(String(16), default="employee")
    title = Column(String(64), default="")
    department_id = Column(Integer, ForeignKey("departments_department.id"), nullable=True)

    # Roster fields imported from HR spreadsheet
    organisation = Column(String(64), default="")
    reporting_manager = Column(String(128), default="")
    date_of_joining = Column(Date, nullable=True)

    department = relationship("Department", back_populates="employees")
    daily_reports = relationship("DailyReport", back_populates="user", cascade="all, delete")


class DailyReport(Base):
    __tablename__ = "reports_dailyreport"

    id = Column(Integer, primary_key=True)
    user_id = Column(Integer, ForeignKey("users_user.id", ondelete="CASCADE"), nullable=False)
    date = Column(Date, nullable=False, index=True)
    data = Column(JSONB, default=dict)
    submitted_at = Column(DateTime, nullable=False)
    created_at = Column(DateTime, nullable=False)

    user = relationship("User", back_populates="daily_reports")

    __table_args__ = (UniqueConstraint("user_id", "date", name="uq_user_date"),)
