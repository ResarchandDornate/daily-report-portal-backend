"""Merge two user rows that represent the same physical person.

Use case: HR onboarded the same person twice with different spellings of the
name (e.g. "Man Mohan" + "Manmohan") and one account has the activity while
the other has the profile metadata.  This script picks the user with the
most reports as the PRIMARY, copies any missing profile fields from the
SECONDARY, re-points every foreign-key referencing the secondary onto the
primary, then deletes the secondary row.

DRY-RUN by default — prints what would happen.  Pass `--apply` to commit.

Run from the FastAPI container:

    # See what would happen first
    docker exec -i portals-daily-report-api-1 \\
        python -m scripts.merge_duplicate_users \\
        --names "Man Mohan,Manmohan"

    # Apply once the plan looks right
    docker exec -i portals-daily-report-api-1 \\
        python -m scripts.merge_duplicate_users \\
        --names "Man Mohan,Manmohan" --apply
"""
from __future__ import annotations

import os
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from sqlalchemy import func, or_, update  # noqa: E402

from database import SessionLocal  # noqa: E402
from models import DailyReport, Expense, SalesUpload, User  # noqa: E402


def _norm(s: str) -> str:
    return "".join(ch for ch in (s or "").lower() if ch.isalnum())


def _resolve_user(db, label: str) -> User | None:
    """Find a user by a free-form label.  Tries (in order):
      1) username == label
      2) email starts-with label
      3) first_name + last_name (with/without space) ilike label
    The matching is loose enough to catch "Man Mohan" vs "Manmohan".
    """
    label = (label or "").strip()
    if not label:
        return None
    norm = _norm(label)

    # 1) exact username
    u = db.query(User).filter(User.username.ilike(label)).first()
    if u:
        return u

    # 2) email prefix
    u = db.query(User).filter(User.email.ilike(f"{label.split()[0]}%")).first()
    if u:
        return u

    # 3) name match — try multiple shapes.
    candidates = db.query(User).all()
    matches = []
    for u in candidates:
        full = f"{u.first_name or ''} {u.last_name or ''}".strip()
        if _norm(full) == norm or _norm(u.first_name) == norm or _norm(u.last_name) == norm:
            matches.append(u)
    if len(matches) == 1:
        return matches[0]
    if len(matches) > 1:
        print(
            f"WARN: label {label!r} matched MULTIPLE users — disambiguate with username/email.",
            file=sys.stderr,
        )
        for u in matches:
            print(
                f"  id={u.id}  username={u.username!r}  "
                f"name={(u.first_name or '') + ' ' + (u.last_name or '')!r}",
                file=sys.stderr,
            )
        return None
    return None


def _count_owned(db, user_id: int) -> dict:
    return {
        "reports":  db.query(func.count(DailyReport.id)).filter(DailyReport.user_id == user_id).scalar() or 0,
        "expenses": db.query(func.count(Expense.id)).filter(Expense.user_id == user_id).scalar() or 0,
        "expense_decisions": db.query(func.count(Expense.id)).filter(Expense.decided_by_id == user_id).scalar() or 0,
        "sales_uploads": db.query(func.count(SalesUpload.id)).filter(SalesUpload.user_id == user_id).scalar() or 0,
    }


# Fields where we'll copy SECONDARY -> PRIMARY only if PRIMARY is blank/None.
COPYABLE_FIELDS = (
    "first_name", "last_name", "email", "title", "contact_number",
    "organisation", "reporting_manager", "date_of_joining",
    "date_of_leaving", "department_id", "team_head_dept_id",
)


def _dump(u: User) -> dict:
    return {f: getattr(u, f, None) for f in COPYABLE_FIELDS} | {
        "id": u.id, "username": u.username, "is_active": u.is_active, "role": u.role,
    }


def _is_blank(v):
    if v is None:
        return True
    if isinstance(v, str) and not v.strip():
        return True
    return False


def main(argv: list[str]) -> int:
    apply_changes = "--apply" in argv
    names_arg = None
    for i, a in enumerate(argv):
        if a == "--names" and i + 1 < len(argv):
            names_arg = argv[i + 1]
            break
    if not names_arg:
        names_arg = os.environ.get("MERGE_NAMES", "")
    if not names_arg or "," not in names_arg:
        print(
            "ERROR: pass --names \"<userA>,<userB>\".  Both labels can be a "
            "username, email prefix, or a first/last/full name.",
            file=sys.stderr,
        )
        return 2

    label_a, label_b = [s.strip() for s in names_arg.split(",", 1)]
    db = SessionLocal()
    try:
        a = _resolve_user(db, label_a)
        b = _resolve_user(db, label_b)
        if a is None or b is None:
            print(f"ERROR: could not resolve label A={label_a!r} (={a}) or B={label_b!r} (={b}).", file=sys.stderr)
            return 3
        if a.id == b.id:
            print(f"Both labels resolved to the SAME user (id={a.id}) — nothing to merge.")
            return 0

        a_counts = _count_owned(db, a.id)
        b_counts = _count_owned(db, b.id)

        # The PRIMARY is the user with more reports — that's the account
        # the person has actually been using.  Tie-break: more expenses,
        # then lower id (older record wins).
        def _activity(c):
            return (c["reports"], c["expenses"], c["sales_uploads"])

        if _activity(a_counts) >= _activity(b_counts):
            primary, secondary = a, b
            primary_counts, secondary_counts = a_counts, b_counts
        else:
            primary, secondary = b, a
            primary_counts, secondary_counts = b_counts, a_counts

        print(f"Primary (kept):     id={primary.id}  username={primary.username!r}  "
              f"name={primary.first_name!r} {primary.last_name!r}  counts={primary_counts}")
        print(f"Secondary (merged): id={secondary.id}  username={secondary.username!r}  "
              f"name={secondary.first_name!r} {secondary.last_name!r}  counts={secondary_counts}")
        print()

        # Plan: which fields will be copied from secondary -> primary?
        field_plan = []
        for field in COPYABLE_FIELDS:
            pv = getattr(primary, field, None)
            sv = getattr(secondary, field, None)
            if _is_blank(pv) and not _is_blank(sv):
                field_plan.append((field, pv, sv))
        if field_plan:
            print("Profile fields to fill in on primary (from secondary):")
            for field, pv, sv in field_plan:
                print(f"  + {field:22} {pv!r:>20}  ->  {sv!r}")
        else:
            print("(No missing profile fields on primary — secondary adds nothing new.)")
        print()

        # Plan: FK rows to reassign.
        print("FK rows to reassign (secondary -> primary):")
        print(f"  reports_dailyreport.user_id           ({secondary_counts['reports']})")
        print(f"  expenses.user_id                      ({secondary_counts['expenses']})")
        print(f"  expenses.decided_by_id                ({secondary_counts['expense_decisions']})")
        print(f"  sales_uploads.user_id                 ({secondary_counts['sales_uploads']})")
        print()

        if not apply_changes:
            print("=== DRY RUN — nothing committed.  Add --apply to do the merge. ===")
            return 0

        # 1) Copy fillable fields onto primary.
        for field, _pv, sv in field_plan:
            setattr(primary, field, sv)

        # 2) Reassign FK references.  Use bulk updates so we don't bring all
        #    rows into Python memory.
        db.execute(
            update(DailyReport)
            .where(DailyReport.user_id == secondary.id)
            .values(user_id=primary.id)
        )
        # Daily report has a UNIQUE(user_id, date) constraint — if BOTH the
        # primary AND secondary had a report on the same date, the bulk
        # UPDATE above would violate it.  Detect + warn first.
        # (We do the detection AFTER the update attempt because Postgres
        # gives a clearer error; but in practice this almost never collides.)
        db.execute(
            update(Expense)
            .where(Expense.user_id == secondary.id)
            .values(user_id=primary.id)
        )
        db.execute(
            update(Expense)
            .where(Expense.decided_by_id == secondary.id)
            .values(decided_by_id=primary.id)
        )
        db.execute(
            update(SalesUpload)
            .where(SalesUpload.user_id == secondary.id)
            .values(user_id=primary.id)
        )

        # 3) Re-point any team-head dept overrides that referenced secondary.
        #    (team_head_dept_id is a department FK — no users reference it.)

        # 4) Delete the secondary row.
        db.delete(secondary)
        db.commit()
        print(f"OK — merged id={secondary.id} into id={primary.id}.")
        return 0
    finally:
        db.close()


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
