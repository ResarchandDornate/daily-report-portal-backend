"""FastAPI entrypoint — JSON API for the Daily Report Portal frontend."""
import os

from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware

from routers import auth, departments, reports, users

app = FastAPI(
    title="Daily Report Portal API",
    description="JSON API for the Ornate Solar Daily Report Portal. "
                "Schema is owned by the sibling Django service (see /admin on :8000).",
    version="0.1.0",
)

# CORS — allow the Next.js dev/prod origins
origins = [o.strip() for o in os.environ.get("CORS_ORIGINS", "").split(",") if o.strip()]
if not origins:
    origins = ["http://localhost:3000", "http://localhost:3001"]

app.add_middleware(
    CORSMiddleware,
    allow_origins=origins,
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)


@app.get("/", tags=["meta"])
def root():
    return {
        "service": "Daily Report Portal API",
        "docs": "/docs",
        "status": "ok",
    }


@app.get("/health", tags=["meta"])
def health():
    return {"status": "healthy"}


app.include_router(auth.router)
app.include_router(users.router)
app.include_router(departments.router)
app.include_router(reports.router)
