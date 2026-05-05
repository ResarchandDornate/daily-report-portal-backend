"""
Auto-bootstrap Alembic before `alembic upgrade head` runs.

Three cases covered:
  1. Fresh DB (no tables) — do nothing; upgrade head creates everything from
     0001_baseline onward.
  2. Existing DB stamped already (alembic_version exists) — do nothing; upgrade
     head is idempotent and just lands on the current head.
  3. Existing DB NOT stamped (Django created tables, alembic_version missing) —
     stamp head so alembic thinks the schema is current and won't try to
     re-create the tables.

This is what makes the FastAPI container start cleanly on every machine —
the very first start on a Django-seeded DB doesn't blow up.
"""
from __future__ import annotations

import os
import subprocess
import sys
from pathlib import Path

import psycopg

APP_ROOT = Path(__file__).resolve().parent.parent  # /app inside the container


def _normalise_url(raw: str) -> str:
    return raw.replace("postgresql+psycopg://", "postgresql://", 1)


def _table_exists(cur, name: str) -> bool:
    cur.execute("SELECT to_regclass(%s)", (f"public.{name}",))
    return cur.fetchone()[0] is not None


def main() -> None:
    url = _normalise_url(os.environ["DATABASE_URL"])
    with psycopg.connect(url) as conn, conn.cursor() as cur:
        has_alembic = _table_exists(cur, "alembic_version")
        has_users = _table_exists(cur, "users_user")

    if has_alembic:
        print("[bootstrap] alembic_version exists — nothing to do.")
        return

    if has_users:
        print(
            "[bootstrap] Existing schema detected without alembic_version. "
            "Stamping head so future migrations build on the current state."
        )
        result = subprocess.run(
            ["alembic", "stamp", "head"],
            cwd=str(APP_ROOT),
        )
        if result.returncode != 0:
            print("[bootstrap] alembic stamp head failed", file=sys.stderr)
            sys.exit(result.returncode)
        return

    print("[bootstrap] Empty DB — upgrade head will create the schema.")


if __name__ == "__main__":
    main()
