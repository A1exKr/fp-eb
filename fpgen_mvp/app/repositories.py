"""Data-access layer backed by SQLAlchemy.

Replaces the previous JSON-file reads/writes. Serialization helpers return the
same shapes the API and frontend already expect, so existing endpoints and
templates keep working. Per-role rates resolve from the ``unit_rates`` table
(the source of truth), with optional per-assignment overrides.
"""
import uuid
from datetime import date

from sqlalchemy import select
from sqlalchemy.orm import Session, selectinload

from app.models import (
    Asset,
    Personnel,
    Proposal,
    ReferenceProject,
    TeamAssignment,
    TeamPreset,
    UnitRate,
)


# --------------------------------------------------------------------------- #
# Serialization helpers (match the legacy JSON shapes)
# --------------------------------------------------------------------------- #
def personnel_to_dict(p: Personnel) -> dict:
    return {
        "id": p.id,
        "name": p.name,
        "title": p.title or "",
        "roles": p.roles or [],
        "cv_asset_id": p.cv_asset_id,
    }


def rate_to_dict(r: UnitRate) -> dict:
    return {
        "id": r.id,
        "role": r.role,
        "rate": r.rate,
        "currency": r.currency or "USD",
        "effective_from": r.effective_from.isoformat() if r.effective_from else None,
    }


def preset_to_dict(preset: TeamPreset, rate_map: dict[str, float]) -> dict:
    return {
        "id": preset.id,
        "name": preset.name,
        "types": preset.types or [],
        "assignments": [
            {
                "role": a.role,
                "person_id": a.person_id,
                "rate": a.rate if a.rate is not None else rate_map.get(a.role),
            }
            for a in preset.assignments
        ],
    }


def reference_to_dict(r: ReferenceProject) -> dict:
    return {
        "id": r.id,
        "name": r.name,
        "project_type": r.project_type or "",
        "location": r.location or "",
        "keywords": r.keywords or [],
        "summary": r.summary or "",
    }


def asset_to_dict(a: Asset) -> dict:
    return {
        "id": a.id,
        "kind": a.kind,
        "role": a.role,
        "reference_project_id": a.reference_project_id,
        "filename": a.filename,
        "storage_ref": a.storage_ref,
        "mime_type": a.mime_type,
        "size_bytes": a.size_bytes,
        "uploaded_by": a.uploaded_by,
    }


def proposal_to_wrapped(p: Proposal) -> dict:
    return {
        "proposal_id": p.proposal_id,
        "created_at": p.created_at.isoformat() if p.created_at else None,
        "updated_at": p.updated_at.isoformat() if p.updated_at else None,
        "payload": p.payload or {},
    }


# --------------------------------------------------------------------------- #
# Unit rates
# --------------------------------------------------------------------------- #
def rate_map(db: Session, currency: str = "USD") -> dict[str, float]:
    rows = db.scalars(select(UnitRate).where(UnitRate.currency == currency)).all()
    return {r.role: r.rate for r in rows}


def list_unit_rates(db: Session) -> list[dict]:
    rows = db.scalars(select(UnitRate).order_by(UnitRate.role)).all()
    return [rate_to_dict(r) for r in rows]


def upsert_unit_rate(db: Session, data: dict) -> dict:
    role = data["role"]
    currency = data.get("currency") or "USD"
    row = db.scalar(
        select(UnitRate).where(UnitRate.role == role, UnitRate.currency == currency)
    )
    eff = data.get("effective_from")
    if isinstance(eff, str) and eff:
        eff = date.fromisoformat(eff)
    if row is None:
        row = UnitRate(role=role, currency=currency)
        db.add(row)
    row.rate = float(data["rate"])
    row.effective_from = eff or None
    db.commit()
    db.refresh(row)
    return rate_to_dict(row)


def delete_unit_rate(db: Session, rate_id: int) -> bool:
    row = db.get(UnitRate, rate_id)
    if row is None:
        return False
    db.delete(row)
    db.commit()
    return True


# --------------------------------------------------------------------------- #
# Personnel
# --------------------------------------------------------------------------- #
def list_personnel(db: Session) -> list[dict]:
    rows = db.scalars(select(Personnel).order_by(Personnel.id)).all()
    return [personnel_to_dict(p) for p in rows]


def upsert_personnel(db: Session, data: dict) -> dict:
    pid = data["id"]
    row = db.get(Personnel, pid)
    if row is None:
        row = Personnel(id=pid)
        db.add(row)
    row.name = data.get("name", row.name or "")
    row.title = data.get("title", row.title or "")
    row.roles = data.get("roles", row.roles or [])
    row.cv_asset_id = data.get("cv_asset_id", row.cv_asset_id)
    db.commit()
    db.refresh(row)
    return personnel_to_dict(row)


def delete_personnel(db: Session, pid: str) -> bool:
    row = db.get(Personnel, pid)
    if row is None:
        return False
    db.delete(row)
    db.commit()
    return True


# --------------------------------------------------------------------------- #
# Team presets (+ assignments)
# --------------------------------------------------------------------------- #
def _presets_query():
    return select(TeamPreset).options(selectinload(TeamPreset.assignments)).order_by(TeamPreset.id)


def list_team_presets(db: Session) -> list[dict]:
    rmap = rate_map(db)
    rows = db.scalars(_presets_query()).all()
    return [preset_to_dict(p, rmap) for p in rows]


def get_team_preset(db: Session, preset_id: str) -> dict | None:
    row = db.scalar(_presets_query().where(TeamPreset.id == preset_id))
    return preset_to_dict(row, rate_map(db)) if row else None


def upsert_team_preset(db: Session, data: dict) -> dict:
    pid = data["id"]
    row = db.scalar(_presets_query().where(TeamPreset.id == pid))
    if row is None:
        row = TeamPreset(id=pid)
        db.add(row)
    row.name = data.get("name", row.name or "")
    row.types = data.get("types", row.types or [])
    if "assignments" in data:
        row.assignments.clear()
        for a in data["assignments"]:
            row.assignments.append(
                TeamAssignment(
                    role=a["role"],
                    person_id=a.get("person_id"),
                    rate=a.get("rate"),
                )
            )
    db.commit()
    return get_team_preset(db, pid)


def delete_team_preset(db: Session, preset_id: str) -> bool:
    row = db.get(TeamPreset, preset_id)
    if row is None:
        return False
    db.delete(row)
    db.commit()
    return True


# --------------------------------------------------------------------------- #
# Reference projects
# --------------------------------------------------------------------------- #
def list_reference_projects(db: Session) -> list[dict]:
    rows = db.scalars(select(ReferenceProject).order_by(ReferenceProject.id)).all()
    return [reference_to_dict(r) for r in rows]


def upsert_reference_project(db: Session, data: dict) -> dict:
    rid = data["id"]
    row = db.get(ReferenceProject, rid)
    if row is None:
        row = ReferenceProject(id=rid)
        db.add(row)
    row.name = data.get("name", row.name or "")
    row.project_type = data.get("project_type", row.project_type or "")
    row.location = data.get("location", row.location or "")
    row.keywords = data.get("keywords", row.keywords or [])
    row.summary = data.get("summary", row.summary or "")
    db.commit()
    db.refresh(row)
    return reference_to_dict(row)


def delete_reference_project(db: Session, rid: str) -> bool:
    row = db.get(ReferenceProject, rid)
    if row is None:
        return False
    db.delete(row)
    db.commit()
    return True


# --------------------------------------------------------------------------- #
# Assets (CV / experience metadata)
# --------------------------------------------------------------------------- #
def list_assets(db: Session, kind: str | None = None) -> list[dict]:
    stmt = select(Asset).order_by(Asset.id)
    if kind:
        stmt = stmt.where(Asset.kind == kind)
    return [asset_to_dict(a) for a in db.scalars(stmt).all()]


def get_asset(db: Session, asset_id: str) -> dict | None:
    row = db.get(Asset, asset_id)
    return asset_to_dict(row) if row else None


def upsert_asset(db: Session, data: dict) -> dict:
    aid = data["id"]
    row = db.get(Asset, aid)
    if row is None:
        row = Asset(id=aid)
        db.add(row)
    for field in (
        "kind",
        "role",
        "reference_project_id",
        "filename",
        "storage_ref",
        "mime_type",
        "size_bytes",
        "uploaded_by",
    ):
        if field in data:
            setattr(row, field, data[field])
    db.commit()
    db.refresh(row)
    return asset_to_dict(row)


def delete_asset(db: Session, asset_id: str) -> bool:
    row = db.get(Asset, asset_id)
    if row is None:
        return False
    db.delete(row)
    db.commit()
    return True


# --------------------------------------------------------------------------- #
# Proposals
# --------------------------------------------------------------------------- #
def save_proposal(db: Session, payload: dict) -> tuple[str, dict]:
    proposal_id = str(uuid.uuid4())
    row = Proposal(proposal_id=proposal_id, payload=payload)
    db.add(row)
    db.commit()
    db.refresh(row)
    return proposal_id, proposal_to_wrapped(row)


def load_proposal(db: Session, proposal_id: str) -> dict | None:
    row = db.get(Proposal, proposal_id)
    return proposal_to_wrapped(row) if row else None


def update_proposal(db: Session, proposal_id: str, payload: dict) -> dict | None:
    row = db.get(Proposal, proposal_id)
    if row is None:
        return None
    row.payload = payload
    db.commit()
    db.refresh(row)
    return proposal_to_wrapped(row)


def list_proposals(db: Session) -> list[dict]:
    rows = db.scalars(select(Proposal).order_by(Proposal.created_at.desc())).all()
    return [proposal_to_wrapped(p) for p in rows]
