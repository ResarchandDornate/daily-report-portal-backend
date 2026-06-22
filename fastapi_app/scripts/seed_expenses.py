"""Seed a handful of dummy Expense rows for testing the new feature.

Idempotent (sort of) — re-running will keep adding more rows.  Delete via
the UI or `DELETE FROM expenses` when you're done testing.

Run from the FastAPI container:

    docker exec -i portals-daily-report-api-1 python -m scripts.seed_expenses
"""
from __future__ import annotations

import random
import sys
from datetime import date as date_type, datetime, timedelta, timezone
from pathlib import Path

# Make `from database import ...` resolvable when run as a module.
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from database import SessionLocal  # noqa: E402
from models import Expense, User  # noqa: E402


SAMPLES = [
    # (expense_type, travel_type, mode, amount, remarks, status, days_ago)
    ("material",  "",       "cash", 1450,  "Stationery + printer cartridges for the office.",          "pending",  0),
    ("travel",    "cab",    "upi",  320,   "Uber from Noida office to Gurgaon meeting (Hero MotoCorp).", "pending",  1),
    ("food",      "",       "upi",  680,   "Team lunch with the design intern on her first day.",      "approved", 2),
    ("hotel",     "",       "card", 4200,  "Bikaner site stay - one night, single room.",              "approved", 5),
    ("travel",    "rapido", "cash", 95,    "Last-mile Rapido from the metro to the warehouse.",        "pending",  3),
    ("fuel",      "",       "card", 2000,  "Diesel refuel for the Bikaner site trip - bill pending.",  "rejected", 8),
    ("travel",    "bike",   "cash", 250,   "Personal bike fuel claim for two client visits.",          "pending",  4),
    ("others",    "",       "upi",  150,   "Courier charges - dispatch of NOPA papers to vendor.",     "approved", 6),
]


def _pick_target_users(db, max_users: int = 3) -> list[User]:
    """Pick a few active non-HR users to attach expenses to.

    Preference: any active employee — we don't care which department.
    Fall back to ANY active user if nothing matches the first filter.
    """
    qs = (
        db.query(User)
        .filter(User.is_active.is_(True))
        .filter(User.role != "hr")
        .order_by(User.id)
        .limit(max_users)
        .all()
    )
    if qs:
        return qs
    return db.query(User).filter(User.is_active.is_(True)).limit(1).all()


def _decider(db) -> User | None:
    """Find a plausible decider (HR or one of the named approvers) so
    approved / rejected rows have a `decided_by_id`.  None if no match.
    """
    # Try email match first (TARINI / SMITA).
    for email in ("tarini@ornatesolar.com", "smita@ornatesolar.com"):
        u = (
            db.query(User)
            .filter(User.email.ilike(email))
            .first()
        )
        if u:
            return u
    # Fall back to any HR user.
    return db.query(User).filter(User.role == "hr").first()


def main() -> int:
    db = SessionLocal()
    try:
        users = _pick_target_users(db)
        if not users:
            print("ERROR: no active users found.", file=sys.stderr)
            return 1
        decider = _decider(db)
        if not decider:
            print(
                "WARNING: no HR / approver user found — approved/rejected "
                "rows will have decided_by_id=NULL.",
                file=sys.stderr,
            )

        today = date_type.today()
        rng = random.Random(42)
        added = 0
        for i, (etype, ttype, mode, amount, remarks, status_, days_ago) in enumerate(SAMPLES):
            user = users[i % len(users)]
            d = today - timedelta(days=days_ago)
            decided_at = None
            decided_by_id = None
            decision_note = ""
            if status_ in ("approved", "rejected") and decider:
                # Decided "a day or two after submission".
                offset_hours = rng.randint(8, 36)
                decided_at = datetime.combine(
                    d + timedelta(days=1), datetime.min.time(),
                    tzinfo=timezone.utc,
                ) + timedelta(hours=offset_hours)
                decided_by_id = decider.id
                decision_note = (
                    "Approved — bill matches.  Reimburse with next salary."
                    if status_ == "approved"
                    else "Rejected — bill missing or amount mismatch.  Please resubmit with the original receipt."
                )
            exp = Expense(
                user_id=user.id,
                date=d,
                mode=mode,
                expense_type=etype,
                travel_type=ttype,
                amount=amount,
                remarks=remarks,
                bill_filename="",
                minio_object_key="",
                status=status_,
                decided_by_id=decided_by_id,
                decided_at=decided_at,
                decision_note=decision_note,
            )
            db.add(exp)
            added += 1
        db.commit()
        print(f"Seeded {added} expense row{'' if added == 1 else 's'} "
              f"across {len(users)} user{'' if len(users) == 1 else 's'}.")
        for u in users:
            print(f"  - {u.id}  {u.username!r}  "
                  f"{((u.first_name or '') + ' ' + (u.last_name or '')).strip()}")
        return 0
    finally:
        db.close()


if __name__ == "__main__":
    sys.exit(main())
