"""Seed the database from the legacy JSON master-data files (idempotent).

Run from the ``fpgen_mvp`` directory:

    python scripts/import_json_to_db.py

Safe to re-run: every record is upserted. Unit rates are derived from the
per-role rates found in team presets (first rate seen per role).
"""
import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from app import repositories  # noqa: E402
from app.config import settings  # noqa: E402
from app.db import SessionLocal, init_db  # noqa: E402
from app.models import Proposal  # noqa: E402


def _read_json(path: Path, default):
    if not path.exists() or path.stat().st_size == 0:
        return default
    try:
        # utf-8-sig tolerates a UTF-8 BOM (some source files have one).
        return json.loads(path.read_text(encoding="utf-8-sig"))
    except Exception as exc:  # pragma: no cover - surfaced to the operator
        print(f"WARNING: could not parse {path}: {exc}", file=sys.stderr)
        return default


def import_personnel(db) -> int:
    data = _read_json(settings.personnel_path, [])
    for person in data:
        repositories.upsert_personnel(db, person)
    return len(data)


def import_rates_and_presets(db) -> tuple[int, int]:
    presets = _read_json(settings.team_presets_path, [])
    seen: dict[str, float] = {}
    for preset in presets:
        for assignment in preset.get("assignments", []):
            role, rate = assignment.get("role"), assignment.get("rate")
            if role and rate is not None and role not in seen:
                seen[role] = rate
    for role, rate in seen.items():
        repositories.upsert_unit_rate(
            db, {"role": role, "rate": rate, "currency": settings.default_currency}
        )
    for preset in presets:
        repositories.upsert_team_preset(db, preset)
    return len(seen), len(presets)


def import_reference(db) -> int:
    data = _read_json(settings.reference_projects_path, [])
    for project in data:
        repositories.upsert_reference_project(db, project)
    return len(data)


def import_assets(db) -> int:
    registry = _read_json(settings.assets_registry_path, {})
    count = 0
    for role, asset_id in (registry.get("cvs") or {}).items():
        repositories.upsert_asset(db, {"id": asset_id, "kind": "cv", "role": role})
        count += 1
    for rp_id, asset_id in (registry.get("experience") or {}).items():
        repositories.upsert_asset(
            db, {"id": asset_id, "kind": "experience", "reference_project_id": rp_id}
        )
        count += 1
    return count


def import_proposals(db) -> int:
    folder = settings.storage_dir
    if not folder.is_dir():
        return 0
    count = 0
    for path in folder.glob("*.json"):
        wrapped = _read_json(path, None)
        if not wrapped:
            continue
        pid = wrapped.get("proposal_id") or path.stem
        payload = wrapped.get("payload", wrapped)
        if repositories.load_proposal(db, pid) is None:
            db.add(Proposal(proposal_id=pid, payload=payload))
            db.commit()
        count += 1
    return count


def main() -> None:
    init_db()
    with SessionLocal() as db:
        n_personnel = import_personnel(db)
        n_rates, n_presets = import_rates_and_presets(db)
        n_ref = import_reference(db)
        n_assets = import_assets(db)
        n_prop = import_proposals(db)
    print(
        f"Imported: personnel={n_personnel}, unit_rates={n_rates}, presets={n_presets}, "
        f"reference_projects={n_ref}, assets={n_assets}, proposals={n_prop}"
    )


if __name__ == "__main__":
    main()
