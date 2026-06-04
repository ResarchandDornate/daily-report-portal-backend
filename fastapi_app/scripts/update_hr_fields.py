"""Update the HR department's daily-report form fields to the new 6-tab layout.

Run from inside the FastAPI container on the Mac mini:

    docker exec -i portals-daily-report-api-1 python -m scripts.update_hr_fields

The script is idempotent — re-running with the same fields is a no-op.
It writes ONE row in `departments_department` and prints a summary.
"""
from __future__ import annotations

import sys
from pathlib import Path

# Make `from database import ...` resolvable when run as a module.
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from sqlalchemy.orm.attributes import flag_modified

from database import SessionLocal  # noqa: E402
from models import Department  # noqa: E402


HR_FIELDS = [
    {"key": "recruitmentScreening",      "label": "Recruitment / Screening"},
    {"key": "inductionOnboarding",       "label": "Induction / Onboarding"},
    {"key": "exitProcess",               "label": "Exit Process"},
    {"key": "attendancePayrollCompliance", "label": "Attendance, Payroll & Compliance"},
    {"key": "happayConveyance",          "label": "Happay & Conveyance"},
    {"key": "otherWork",                 "label": "Other Work"},
]


def find_hr_department(db):
    """Locate the HR department by trying common slug variants, then by
    name match.  Returns None if no candidate is found."""
    # Try common slug spellings first.
    for slug in ("hr", "HR", "humanResources", "human_resources", "human-resources"):
        d = db.query(Department).filter(Department.slug == slug).first()
        if d:
            return d
    # Fallback — match by name containing 'hr' / 'human resources'.
    for d in db.query(Department).all():
        name_l = (d.name or "").lower()
        if name_l in ("hr", "human resources") or "human resources" in name_l:
            return d
    # Last resort: exact 'HR' name (case-insensitive).
    for d in db.query(Department).all():
        if (d.name or "").strip().lower() == "hr":
            return d
    return None


def main() -> int:
    db = SessionLocal()
    try:
        dept = find_hr_department(db)
        if dept is None:
            print(
                "ERROR: could not find the HR department.  "
                "Check the slug/name in the admin UI and pass the slug as an arg.",
                file=sys.stderr,
            )
            return 1

        before = list(dept.report_fields or [])
        dept.report_fields = HR_FIELDS
        # JSONB mutations need an explicit dirty flag for SQLAlchemy to flush.
        flag_modified(dept, "report_fields")
        db.commit()
        db.refresh(dept)

        print(f"Updated department: id={dept.id}  slug={dept.slug!r}  name={dept.name!r}")
        print(f"Before ({len(before)} fields):")
        for f in before:
            print(f"  - {f.get('key')!r:30}  {f.get('label')!r}")
        print(f"After  ({len(dept.report_fields)} fields):")
        for f in dept.report_fields:
            print(f"  - {f.get('key')!r:30}  {f.get('label')!r}")
        return 0
    finally:
        db.close()


if __name__ == "__main__":
    sys.exit(main())
