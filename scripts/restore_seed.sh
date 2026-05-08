#!/usr/bin/env sh
#
# Restore seed.sql into the running Postgres on the Mac mini.
#
# Run this AFTER `docker compose up -d --build` has finished bringing up the
# stack — by then bootstrap_alembic has stamped the schema and bootstrap_admin
# has created an admin user.  This script wipes those starter rows and replaces
# them with the dump from the Windows machine.
#
# Usage:
#   1. Copy seed.sql to this folder (NOT via git — scp / USB / encrypted upload).
#   2. From the project root:
#         sh scripts/restore_seed.sh
#   3. Confirm 'yes' when prompted.
#
# WARNING: TRUNCATE wipes ALL departments, users, and reports.  Run only on
# a freshly-deployed Mac mini, never on a system that already has real data.
#
set -e

DUMP_FILE="${DUMP_FILE:-seed.sql}"

if [ ! -f "$DUMP_FILE" ]; then
    echo "ERROR: $DUMP_FILE not found. Copy it from the source machine first."
    exit 1
fi

echo "Found dump: $(wc -l < "$DUMP_FILE") lines, $(du -h "$DUMP_FILE" | cut -f1)"
echo ""
echo "This will:"
echo "  1. TRUNCATE departments_department, users_user, reports_dailyreport (CASCADE)"
echo "  2. Replay $DUMP_FILE into Postgres"
echo ""
echo "All existing rows in those three tables will be REPLACED."
echo ""
printf "Type 'yes' to continue: "
read -r confirm
[ "$confirm" = "yes" ] || { echo "Aborted."; exit 1; }

if ! docker compose ps postgres 2>/dev/null | grep -q "Up"; then
    echo "ERROR: Postgres container isn't running. Start the stack first:"
    echo "    docker compose up -d"
    exit 1
fi

echo ""
echo "Truncating tables..."
docker compose exec -T postgres psql -U "${POSTGRES_USER:-drp_user}" -d "${POSTGRES_DB:-drp}" <<'SQL'
TRUNCATE reports_dailyreport, users_user, departments_department RESTART IDENTITY CASCADE;
SQL

echo "Restoring dump..."
docker compose exec -T postgres psql -U "${POSTGRES_USER:-drp_user}" -d "${POSTGRES_DB:-drp}" < "$DUMP_FILE"

echo ""
echo "Verifying counts..."
docker compose exec -T postgres psql -U "${POSTGRES_USER:-drp_user}" -d "${POSTGRES_DB:-drp}" -c "
  SELECT 'departments' AS table, COUNT(*) FROM departments_department
  UNION ALL SELECT 'users', COUNT(*) FROM users_user
  UNION ALL SELECT 'reports', COUNT(*) FROM reports_dailyreport;
"

echo ""
echo "Done.  Dashboard logins from the source machine now work on this server."
