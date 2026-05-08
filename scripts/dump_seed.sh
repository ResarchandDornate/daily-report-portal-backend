#!/usr/bin/env sh
#
# Snapshot the live Postgres data into seed.sql for transfer to another server.
#
# Output:
#   seed.sql — INSERT statements for departments_department, users_user,
#              reports_dailyreport.  Schema is NOT included (Alembic creates
#              it on the destination).  Sequences ARE handled (setval calls
#              come baked into pg_dump --inserts output).
#
# This file is gitignored (*.sql in .gitignore).  Transfer manually via scp,
# USB, or encrypted upload — never via git, since it contains real PII.
#
# Usage:
#   sh scripts/dump_seed.sh
#
set -e

OUT="${OUT:-seed.sql}"

if ! docker compose ps postgres 2>/dev/null | grep -q "Up"; then
    echo "ERROR: Postgres container isn't running."
    exit 1
fi

echo "Dumping to $OUT ..."
docker compose exec -T postgres pg_dump \
  -U "${POSTGRES_USER:-drp_user}" \
  -d "${POSTGRES_DB:-drp}" \
  --data-only \
  --inserts \
  --column-inserts \
  -t departments_department \
  -t users_user \
  -t reports_dailyreport \
  > "$OUT"

echo "Done."
echo "  $(wc -l < "$OUT") lines  |  $(du -h "$OUT" | cut -f1)"
echo "  $(grep -c "^INSERT INTO public.departments_department" "$OUT") departments"
echo "  $(grep -c "^INSERT INTO public.users_user" "$OUT") users"
echo "  $(grep -c "^INSERT INTO public.reports_dailyreport" "$OUT") reports"
echo ""
echo "Transfer to the destination server (NOT via git):"
echo "    scp $OUT user@mac-mini.local:~/daily-report-portal-backend/"
echo "Then on the Mac mini:"
echo "    sh scripts/restore_seed.sh"
