"""HR-managed list of organisations.

The list is shown as a dropdown in the Add / Edit Employee form so HR can
pick a canonical value instead of free-typing.  Reads are open to any
authenticated user (so the form populates the dropdown); writes are HR-only.

The User.organisation column remains a free-text string carrying the
chosen name — this table is just the master list.  Renaming an
organisation here does NOT propagate to existing User rows automatically
(by design, to avoid surprise bulk updates); HR can rerun the rename via
a UPDATE if needed.
"""
from datetime import datetime, timezone

from fastapi import APIRouter, Depends, HTTPException, status
from sqlalchemy.orm import Session

from auth import get_current_user
from database import get_db
from models import Organisation, User
from schemas import OrganisationCreate, OrganisationOut, OrganisationUpdate

router = APIRouter(prefix="/api/organisations", tags=["organisations"])


def _require_hr(user: User) -> None:
    if user.role != "hr":
        raise HTTPException(status.HTTP_403_FORBIDDEN, "Only HR can manage organisations")


@router.get("", response_model=list[OrganisationOut])
def list_organisations(
    db: Session = Depends(get_db),
    _: User = Depends(get_current_user),
):
    return db.query(Organisation).order_by(Organisation.name).all()


@router.post("", response_model=OrganisationOut, status_code=status.HTTP_201_CREATED)
def create_organisation(
    payload: OrganisationCreate,
    db: Session = Depends(get_db),
    actor: User = Depends(get_current_user),
):
    _require_hr(actor)
    name = (payload.name or "").strip()
    if not name:
        raise HTTPException(status.HTTP_400_BAD_REQUEST, "Name is required")
    if db.query(Organisation).filter(Organisation.name == name).first():
        raise HTTPException(status.HTTP_409_CONFLICT, f"Organisation '{name}' already exists")
    now = datetime.now(timezone.utc)
    org = Organisation(
        name=name,
        color=(payload.color or "zinc").strip() or "zinc",
        created_at=now,
        updated_at=now,
    )
    db.add(org)
    db.commit()
    db.refresh(org)
    return org


@router.patch("/{org_id}", response_model=OrganisationOut)
def update_organisation(
    org_id: int,
    payload: OrganisationUpdate,
    db: Session = Depends(get_db),
    actor: User = Depends(get_current_user),
):
    _require_hr(actor)
    org = db.query(Organisation).filter(Organisation.id == org_id).first()
    if not org:
        raise HTTPException(status.HTTP_404_NOT_FOUND, "Organisation not found")
    if payload.name is not None:
        new_name = payload.name.strip()
        if not new_name:
            raise HTTPException(status.HTTP_400_BAD_REQUEST, "Name cannot be empty")
        clash = (
            db.query(Organisation)
            .filter(Organisation.name == new_name, Organisation.id != org_id)
            .first()
        )
        if clash:
            raise HTTPException(status.HTTP_409_CONFLICT, f"Another organisation already uses the name '{new_name}'")
        org.name = new_name
    if payload.color is not None:
        org.color = payload.color.strip() or "zinc"
    org.updated_at = datetime.now(timezone.utc)
    db.commit()
    db.refresh(org)
    return org


@router.delete("/{org_id}", status_code=status.HTTP_204_NO_CONTENT)
def delete_organisation(
    org_id: int,
    db: Session = Depends(get_db),
    actor: User = Depends(get_current_user),
):
    _require_hr(actor)
    org = db.query(Organisation).filter(Organisation.id == org_id).first()
    if not org:
        raise HTTPException(status.HTTP_404_NOT_FOUND, "Organisation not found")
    # Block delete if any employee still references this org by name —
    # forces HR to reassign first, so we never leave employees pointing at
    # an organisation that no longer exists.
    in_use = (
        db.query(User)
        .filter(User.organisation == org.name)
        .count()
    )
    if in_use:
        raise HTTPException(
            status.HTTP_409_CONFLICT,
            f"Organisation has {in_use} employee(s) — reassign them before deleting.",
        )
    db.delete(org)
    db.commit()
    return None
