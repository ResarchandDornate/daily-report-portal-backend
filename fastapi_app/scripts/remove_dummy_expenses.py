"""Remove the dummy Expense rows seeded by `scripts.seed_expenses`.

Matches the 8 exact `remarks` strings inserted by the seed so any genuine
expenses an employee submitted through the UI are preserved.

Run from the FastAPI container:

    docker exec -i drp_fastapi python -m scripts.remove_dummy_expenses
"""
from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

import storage  # noqa: E402
from database import SessionLocal  # noqa: E402
from models import Expense  # noqa: E402


DUMMY_REMARKS = [
    "Stationery + printer cartridges for the office.",
    "Uber from Noida office to Gurgaon meeting (Hero MotoCorp).",
    "Team lunch with the design intern on her first day.",
    "Bikaner site stay - one night, single room.",
    "Last-mile Rapido from the metro to the warehouse.",
    "Diesel refuel for the Bikaner site trip - bill pending.",
    "Personal bike fuel claim for two client visits.",
    "Courier charges - dispatch of NOPA papers to vendor.",
]


def main() -> int:
    db = SessionLocal()
    try:
        rows = (
            db.query(Expense)
            .filter(Expense.remarks.in_(DUMMY_REMARKS))
            .all()
        )
        if not rows:
            print("No dummy expenses found — nothing to remove.")
            return 0
        for r in rows:
            # Best-effort cleanup of any attached bill in MinIO (the seeded
            # rows don't have one, but real edits via the UI might have
            # added a file).
            if r.minio_object_key:
                try:
                    storage.delete_object(r.minio_object_key)
                except Exception:
                    pass
            db.delete(r)
        db.commit()
        print(f"Removed {len(rows)} dummy expense row{'' if len(rows) == 1 else 's'}.")
        for r in rows:
            print(f"  - id={r.id}  user_id={r.user_id}  "
                  f"type={r.expense_type!r}  amount={r.amount}  "
                  f"status={r.status!r}")
        return 0
    finally:
        db.close()


if __name__ == "__main__":
    sys.exit(main())
