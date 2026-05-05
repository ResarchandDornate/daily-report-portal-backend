"""
Alembic env.py

Connects Alembic to our SQLAlchemy models and the running Postgres.
Reads `DATABASE_URL` from the environment so we don't hard-code creds.

Importantly, this filters out Django's auxiliary tables (auth_*, django_*,
users_user_groups, etc.) so `alembic revision --autogenerate` doesn't try
to drop them while we're still running Django alongside FastAPI (Phase 1+2).
After Phase 3 (Django removed), the filter can stay or go — it's harmless.
"""
import os
import sys
from logging.config import fileConfig
from pathlib import Path

from alembic import context
from sqlalchemy import engine_from_config, pool

# Make `import models` etc. work no matter where alembic is run from.
HERE = Path(__file__).resolve().parent.parent  # fastapi_app/
sys.path.insert(0, str(HERE))

# Import Base so its metadata is populated with our model classes.
from database import Base  # noqa: E402
import models  # noqa: F401, E402  -- registers models on Base.metadata

# Alembic Config object
config = context.config

# Read the same DATABASE_URL the app uses, normalising to psycopg3 driver.
db_url = os.environ.get(
    "DATABASE_URL",
    "postgresql+psycopg://drp_user:drp@postgres:5432/drp",
)
if db_url.startswith("postgresql://"):
    db_url = db_url.replace("postgresql://", "postgresql+psycopg://", 1)
config.set_main_option("sqlalchemy.url", db_url)

# Set up Python logging from alembic.ini if it has a logger config section.
if config.config_file_name is not None:
    fileConfig(config.config_file_name)

# This is the metadata Alembic compares against the live DB during autogenerate.
target_metadata = Base.metadata


# Tables Alembic must NOT touch — Django owns these during the migration window.
# After Phase 3 (Django gone) we'll keep the filter so a stray rerun doesn't drop
# orphan tables that may still hold legacy admin/session data.
DJANGO_AUX_TABLES = {
    "auth_group",
    "auth_group_permissions",
    "auth_permission",
    "django_admin_log",
    "django_content_type",
    "django_migrations",
    "django_session",
    "users_user_groups",
    "users_user_user_permissions",
}


def include_object(object_, name, type_, reflected, compare_to):
    """Skip Django's auxiliary tables during autogenerate diffs."""
    if type_ == "table" and name in DJANGO_AUX_TABLES:
        return False
    # Skip indexes / constraints that belong to skipped tables.
    if type_ in ("index", "unique_constraint", "foreign_key_constraint"):
        owner = getattr(object_, "table", None)
        if owner is not None and owner.name in DJANGO_AUX_TABLES:
            return False
    return True


def run_migrations_offline() -> None:
    """Generate SQL without connecting to the DB.  Useful for review/CI."""
    context.configure(
        url=db_url,
        target_metadata=target_metadata,
        literal_binds=True,
        dialect_opts={"paramstyle": "named"},
        include_object=include_object,
        compare_type=True,
        compare_server_default=True,
    )
    with context.begin_transaction():
        context.run_migrations()


def run_migrations_online() -> None:
    """Connect to the DB and run migrations."""
    connectable = engine_from_config(
        config.get_section(config.config_ini_section, {}),
        prefix="sqlalchemy.",
        poolclass=pool.NullPool,
    )
    with connectable.connect() as connection:
        context.configure(
            connection=connection,
            target_metadata=target_metadata,
            include_object=include_object,
            compare_type=True,
            compare_server_default=True,
        )
        with context.begin_transaction():
            context.run_migrations()


if context.is_offline_mode():
    run_migrations_offline()
else:
    run_migrations_online()
