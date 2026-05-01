# Daily Report Portal — Backend

Two-service Python backend for the Ornate Solar Daily Report Portal.

| Service  | Stack            | Port | Purpose                                    |
|----------|------------------|------|--------------------------------------------|
| Django   | Django 5 + Admin | 8000 | DB schema (migrations) + admin UI          |
| FastAPI  | FastAPI + JWT    | 8001 | JSON API the Next.js frontend talks to     |
| Postgres | postgres:16      | 5432 | Shared database                            |
| pgAdmin  | dpage/pgadmin4   | 5050 | Browse / edit DB in a browser              |

## Quick start

```bash
# 1. Copy env file and edit any secrets you want to change
cp .env.example .env

# 2. Build and start everything (first run takes ~2 min to download images)
docker compose up --build

# 3. Open the URLs:
#    http://localhost:8000/admin   →  Django admin (login: admin / admin12345)
#    http://localhost:8001/docs    →  FastAPI auto-generated Swagger UI
#    http://localhost:5050         →  pgAdmin (login: admin@ornatesolar.com / admin12345)
```

## What gets created on first run

- Postgres database `drp` with a user `drp_user`
- All Django migrations applied (User, Department, DailyReport tables)
- A Django superuser (so you can sign in to `/admin`)

## Frontend wiring

In your Next.js `.env.local`:

```
NEXT_PUBLIC_API_URL=http://localhost:8001
```

Frontend calls FastAPI (port 8001). FastAPI handles auth, reports CRUD, summary generation. Use Django admin (port 8000) for backstage data fixes.

## Connecting pgAdmin to the DB the first time

In pgAdmin (http://localhost:5050) → right-click "Servers" → Register → Server:

- **General → Name**: `DRP`
- **Connection → Host**: `postgres` *(the Docker service name, not localhost)*
- **Connection → Port**: `5432`
- **Connection → Username**: `drp_user`
- **Connection → Password**: whatever you set in `.env`

## Folder layout

```
.
├── docker-compose.yml        ← orchestrates 4 containers
├── .env / .env.example       ← shared config
├── django_app/               ← admin + ORM + migrations
│   ├── core/                 ← project settings/urls
│   ├── users/                ← custom User model
│   ├── departments/          ← Department model
│   └── reports/              ← DailyReport model (JSON `data` field for per-dept fields)
└── fastapi_app/              ← JSON API for the frontend
    ├── main.py
    ├── database.py           ← SQLAlchemy session
    ├── models.py             ← SQLAlchemy mirrors of Django tables
    ├── schemas.py            ← Pydantic request/response shapes
    ├── auth.py               ← JWT helpers
    └── routers/              ← endpoints by resource
```

## Why two frameworks?

- **Django** is the right tool for "I need a UI to fix data right now" (built-in admin) and for managing schema migrations.
- **FastAPI** is the right tool for serving fast, typed JSON APIs to a SPA. Auto Swagger docs, async support, Pydantic validation.

They don't overlap: Django doesn't expose any REST API, FastAPI doesn't manage the schema. Both connect to the same Postgres.

## Common commands

```bash
# View logs of one service
docker compose logs -f fastapi
docker compose logs -f django

# Run a Django management command
docker compose exec django python manage.py makemigrations
docker compose exec django python manage.py shell

# Stop everything (data persists in Docker volumes)
docker compose down

# Wipe data and start fresh
docker compose down -v
```
