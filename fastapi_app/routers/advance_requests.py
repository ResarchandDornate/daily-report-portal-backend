"""Travel advance request workflow.

Endpoints:
  POST   /api/advance-requests/              — employee submits advance request
  GET    /api/advance-requests/              — list (own for employee, all for HR/approver)
  POST   /api/advance-requests/{id}/decide   — HR/approver approves or rejects
  DELETE /api/advance-requests/{id}          — owner can withdraw pending; HR can delete any
"""
from __future__ import annotations

from datetime import datetime, timezone

from fastapi import APIRouter, Depends, HTTPException, status
from pydantic import BaseModel
from sqlalchemy.orm import Session

from auth import get_current_user
from database import get_db
from models import AdvanceRequest, User

router = APIRouter(prefix="/api/advance-requests", tags=["advance-requests"])


# ── helpers ──────────────────────────────────────────────────────────────────

def _is_approver(user: User) -> bool:
    if user.role == "hr":
        return True
    local = (user.email or "").lower().split("@")[0]
    return any(
        local == p or local.startswith(p + ".") or local.startswith(p + "_")
        for p in ("tarini", "smita")
    )


def _full_name(u: User | None) -> str:
    if not u:
        return ""
    return f"{u.first_name} {u.last_name}".strip() or u.username


def _to_out(req: AdvanceRequest) -> dict:
    return {
        "id": req.id,
        "created_by_id": req.created_by_id,
        "created_by_name": _full_name(req.created_by),
        "employee_names": req.employee_names or [],
        "city": req.city,
        "travel_date": req.travel_date.isoformat() if req.travel_date else None,
        "return_date": req.return_date.isoformat() if req.return_date else None,
        "tour_days": req.tour_days,
        "mode_going": req.mode_going,
        "mode_return": req.mode_return,
        "purpose": req.purpose,
        "accommodation_days": req.accommodation_days,
        "accommodation_rate": float(req.accommodation_rate),
        "food_amount": float(req.food_amount),
        "conveyance_days": req.conveyance_days,
        "conveyance_rate": float(req.conveyance_rate),
        "total_amount": float(req.total_amount),
        "status": req.status,
        "decision_note": req.decision_note,
        "decided_by_name": _full_name(req.decided_by) if req.decided_by else None,
        "decided_at": req.decided_at.isoformat() if req.decided_at else None,
        "created_at": req.created_at.isoformat() if req.created_at else None,
    }


# ── schemas ───────────────────────────────────────────────────────────────────

class AdvanceRequestIn(BaseModel):
    employee_names: list[str]
    city: str
    travel_date: str   # YYYY-MM-DD
    return_date: str
    tour_days: int
    mode_going: str
    mode_return: str
    purpose: str = ""
    accommodation_days: int = 0
    accommodation_rate: float = 0
    food_amount: float = 0
    conveyance_days: int = 0
    conveyance_rate: float = 0
    total_amount: float = 0


class DecideIn(BaseModel):
    action: str          # "approve" | "reject"
    note: str = ""


# ── endpoints ────────────────────────────────────────────────────────────────

@router.post("", status_code=status.HTTP_201_CREATED)
def create_advance_request(
    body: AdvanceRequestIn,
    db: Session = Depends(get_db),
    me: User = Depends(get_current_user),
):
    from datetime import date as date_type
    req = AdvanceRequest(
        created_by_id=me.id,
        employee_names=body.employee_names,
        city=body.city,
        travel_date=date_type.fromisoformat(body.travel_date),
        return_date=date_type.fromisoformat(body.return_date),
        tour_days=body.tour_days,
        mode_going=body.mode_going,
        mode_return=body.mode_return,
        purpose=body.purpose,
        accommodation_days=body.accommodation_days,
        accommodation_rate=body.accommodation_rate,
        food_amount=body.food_amount,
        conveyance_days=body.conveyance_days,
        conveyance_rate=body.conveyance_rate,
        total_amount=body.total_amount,
        status="pending",
    )
    db.add(req)
    db.commit()
    db.refresh(req)
    return _to_out(req)


@router.get("")
def list_advance_requests(
    db: Session = Depends(get_db),
    me: User = Depends(get_current_user),
):
    q = db.query(AdvanceRequest)
    if not _is_approver(me):
        q = q.filter(AdvanceRequest.created_by_id == me.id)
    rows = q.order_by(AdvanceRequest.created_at.desc()).all()
    return [_to_out(r) for r in rows]


@router.patch("/{req_id}")
def update_advance_request(
    req_id: int,
    body: AdvanceRequestIn,
    db: Session = Depends(get_db),
    me: User = Depends(get_current_user),
):
    """Owner-only edit of a pending advance request.  Once approved or
    rejected the row is locked; HR can still delete via the DELETE route
    if they need to clear a mistake."""
    from datetime import date as date_type
    req = db.get(AdvanceRequest, req_id)
    if not req:
        raise HTTPException(status_code=404, detail="Not found.")
    if req.created_by_id != me.id:
        raise HTTPException(status_code=403, detail="You can only edit your own advance requests.")
    if req.status != "pending":
        raise HTTPException(status_code=400, detail="A decided request can no longer be edited.")
    req.employee_names = body.employee_names
    req.city = body.city
    req.travel_date = date_type.fromisoformat(body.travel_date)
    req.return_date = date_type.fromisoformat(body.return_date)
    req.tour_days = body.tour_days
    req.mode_going = body.mode_going
    req.mode_return = body.mode_return
    req.purpose = body.purpose
    req.accommodation_days = body.accommodation_days
    req.accommodation_rate = body.accommodation_rate
    req.food_amount = body.food_amount
    req.conveyance_days = body.conveyance_days
    req.conveyance_rate = body.conveyance_rate
    req.total_amount = body.total_amount
    db.commit()
    db.refresh(req)
    return _to_out(req)


@router.post("/{req_id}/decide")
def decide_advance_request(
    req_id: int,
    body: DecideIn,
    db: Session = Depends(get_db),
    me: User = Depends(get_current_user),
):
    if not _is_approver(me):
        raise HTTPException(status_code=403, detail="Not authorised to approve advance requests.")
    req = db.get(AdvanceRequest, req_id)
    if not req:
        raise HTTPException(status_code=404, detail="Not found.")
    if req.status != "pending":
        raise HTTPException(status_code=400, detail="Request is already decided.")
    if body.action not in ("approve", "reject"):
        raise HTTPException(status_code=422, detail="action must be 'approve' or 'reject'.")
    req.status = "approved" if body.action == "approve" else "rejected"
    req.decision_note = body.note
    req.decided_by_id = me.id
    req.decided_at = datetime.now(timezone.utc)
    db.commit()
    db.refresh(req)
    return _to_out(req)


@router.delete("/{req_id}", status_code=status.HTTP_204_NO_CONTENT)
def delete_advance_request(
    req_id: int,
    db: Session = Depends(get_db),
    me: User = Depends(get_current_user),
):
    req = db.get(AdvanceRequest, req_id)
    if not req:
        raise HTTPException(status_code=404, detail="Not found.")
    if req.created_by_id != me.id and not _is_approver(me):
        raise HTTPException(status_code=403, detail="Not authorised.")
    if req.status != "pending" and not _is_approver(me):
        raise HTTPException(status_code=400, detail="Cannot delete a decided request.")
    db.delete(req)
    db.commit()
