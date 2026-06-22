"""Delete EVERY expense row and remove its attached bill files from MinIO.

Use this only on dev / staging to clear out test data — never on production
without first confirming no real expenses exist.

By default it wipes ALL expenses.  Pass `--email <addr>` (env override
`EXPENSE_USER_EMAIL`) to limit the wipe to a single employee.

Run from the FastAPI container:

    docker exec -i drp_fastapi python -m scripts.clear_expenses
    docker exec -i drp_fastapi python -m scripts.clear_expenses --email smita@ornatesolar.com
"""
from __future__ import annotations

import os
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

import storage  # noqa: E402
from database import SessionLocal  # noqa: E402
from models import Expense, User  # noqa: E402


def _parse_args(argv: list[str]) -> str | None:
    """Return the email filter (or None if no filter)."""
    email = os.environ.get("EXPENSE_USER_EMAIL", "").strip().lower() or None
    if "--email" in argv:
        idx = argv.index("--email")
        if idx + 1 < len(argv):
            email = argv[idx + 1].strip().lower()
    if "--all" in argv:
        email = None
    return email


def main(argv: list[str]) -> int:
    target_email = _parse_args(argv)
    db = SessionLocal()
    try:
        q = db.query(Expense).join(User, Expense.user_id == User.id)
        if target_email:
            q = q.filter(User.email.ilike(target_email))
        rows = q.all()
        if not rows:
            scope = (
                f"user {target_email!r}" if target_email else "the entire expenses table"
            )
            print(f"No expenses found for {scope}.")
            return 0

        scope_msg = (
            f"user {target_email!r}" if target_email else "ALL employees"
        )
        print(f"About to remove {len(rows)} expense row(s) for {scope_msg}.")

        n_files_deleted = 0
        for r in rows:
            for bill in (r.bills or []):
                key = (bill.get("object_key") or "")
                if key:
                    try:
                        storage.delete_object(key)
                        n_files_deleted += 1
                    except Exception as e:
                        print(f"  warn: could not delete {key!r}: {e}", file=sys.stderr)
            db.delete(r)
        db.commit()
        print(
            f"Removed {len(rows)} expense row(s).  "
            f"Cleaned up {n_files_deleted} bill file(s) from MinIO."
        )
        return 0
    finally:
        db.close()


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
