"""
Ensure the HR superuser exists on every container start.

Replaces the Django-side auto-superuser-creation that used to live in
docker-compose.yml's `django` service command.  Idempotent — only creates
the user if they're missing.

Reads from environment:
  - DJANGO_SUPERUSER_USERNAME (default: "admin")
  - DJANGO_SUPERUSER_EMAIL    (default: "admin@ornatesolar.com")
  - DJANGO_SUPERUSER_PASSWORD (default: "admin@ornate")

The DJANGO_ prefix is preserved purely for backwards compatibility with the
existing .env file.  After Django is gone these vars are simply "the admin
user's credentials" — feel free to rename to ADMIN_* at your leisure.
"""
from __future__ import annotations

import os
import sys
from datetime import datetime, timezone
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from auth import hash_password  # noqa: E402
from database import SessionLocal  # noqa: E402
from models import User  # noqa: E402
import model_events  # noqa: F401, E402  -- registers user invariants


def main() -> None:
    username = os.environ.get("DJANGO_SUPERUSER_USERNAME", "admin")
    email = os.environ.get("DJANGO_SUPERUSER_EMAIL", "admin@ornatesolar.com")
    password = os.environ.get("DJANGO_SUPERUSER_PASSWORD", "admin@ornate")

    db = SessionLocal()
    try:
        # Match by either username or email — covers admins created by either
        # path during the migration window.
        existing = (
            db.query(User)
            .filter((User.username == username) | (User.email == email))
            .first()
        )
        if existing:
            print(f"[bootstrap_admin] admin user {existing.username!r} already exists.")
            return

        admin = User(
            username=username,
            email=email,
            first_name="Admin",
            last_name="",
            role="hr",
            is_superuser=True,
            is_staff=True,
            is_active=True,
            date_joined=datetime.now(timezone.utc),
            password=hash_password(password),
        )
        db.add(admin)
        db.commit()
        print(f"[bootstrap_admin] created HR superuser {username!r} <{email}>")
    finally:
        db.close()


if __name__ == "__main__":
    main()
