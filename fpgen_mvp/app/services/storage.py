import json
import uuid
from datetime import datetime, timezone

from app.config import settings


def save_proposal(payload: dict) -> tuple[str, dict]:
    settings.storage_dir.mkdir(parents=True, exist_ok=True)
    proposal_id = str(uuid.uuid4())
    wrapped = {
        "proposal_id": proposal_id,
        "created_at": datetime.now(timezone.utc).isoformat(),
        "payload": payload,
    }
    out_file = settings.storage_dir / f"{proposal_id}.json"
    out_file.write_text(json.dumps(wrapped, ensure_ascii=False, indent=2), encoding="utf-8")
    return proposal_id, wrapped


def load_proposal(proposal_id: str) -> dict | None:
    in_file = settings.storage_dir / f"{proposal_id}.json"
    if not in_file.exists():
        return None
    return json.loads(in_file.read_text(encoding="utf-8"))


def update_proposal(proposal_id: str, payload: dict) -> dict | None:
    in_file = settings.storage_dir / f"{proposal_id}.json"
    if not in_file.exists():
        return None

    wrapped = json.loads(in_file.read_text(encoding="utf-8"))
    wrapped["payload"] = payload
    wrapped["updated_at"] = datetime.now(timezone.utc).isoformat()
    in_file.write_text(json.dumps(wrapped, ensure_ascii=False, indent=2), encoding="utf-8")
    return wrapped
