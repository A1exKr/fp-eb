"""Proposal storage — now backed by the database.

Thin wrappers kept for backward compatibility with non-request callers
(scripts, tests). FastAPI endpoints use ``app.repositories`` directly with the
request-scoped session. Each function here opens a short-lived session.
"""
from app import repositories
from app.db import SessionLocal


def save_proposal(payload: dict) -> tuple[str, dict]:
    with SessionLocal() as db:
        return repositories.save_proposal(db, payload)


def load_proposal(proposal_id: str) -> dict | None:
    with SessionLocal() as db:
        return repositories.load_proposal(db, proposal_id)


def update_proposal(proposal_id: str, payload: dict) -> dict | None:
    with SessionLocal() as db:
        return repositories.update_proposal(db, proposal_id, payload)
