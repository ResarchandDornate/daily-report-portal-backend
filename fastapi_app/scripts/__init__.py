"""SQLAlchemy-based seed/import scripts.

These mirror the Django scripts in django_app/ (seed_data.py, seed_reports.py,
import_employees.py) but use SQLAlchemy directly — no Django dependency.

During Phase 1+2 of the migration, both versions exist.  Run whichever stack
you prefer:

  Django side (existing):
    docker compose exec django python manage.py shell -c \\
        "exec(open('/app/seed_reports.py').read())"

  FastAPI side (new — works after Phase 3 too):
    docker compose exec fastapi python -m scripts.seed_reports

After Phase 3 (Django removed), only the SQLAlchemy versions remain.
"""
