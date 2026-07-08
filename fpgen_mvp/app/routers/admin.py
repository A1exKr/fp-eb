"""Admin / setup CRUD endpoints for master data.

All routes require authentication; writes are role-gated. Rates additionally
allow the Finance group. On exaBase these sit behind the oauth2-proxy, so the
forwarded identity headers drive the role checks in ``app.auth``.
"""
from fastapi import APIRouter, Depends, File, Form, HTTPException, UploadFile
from sqlalchemy.orm import Session

from app import repositories
from app.auth import CurrentUser, require_admin, require_finance
from app.db import get_db
from app.schemas import (
    AssetIn,
    PersonnelIn,
    ReferenceProjectIn,
    TeamPresetIn,
    UnitRateIn,
)
from app.services.file_storage import save_asset_file

router = APIRouter(prefix="/v1/admin", tags=["admin"])


# --------------------------------------------------------------------------- #
# Personnel / team members
# --------------------------------------------------------------------------- #
@router.get("/personnel")
def list_personnel(
    db: Session = Depends(get_db), _: CurrentUser = Depends(require_admin)
) -> list:
    return repositories.list_personnel(db)


@router.post("/personnel")
def upsert_personnel(
    body: PersonnelIn, db: Session = Depends(get_db), _: CurrentUser = Depends(require_admin)
) -> dict:
    return repositories.upsert_personnel(db, body.model_dump())


@router.delete("/personnel/{pid}")
def delete_personnel(
    pid: str, db: Session = Depends(get_db), _: CurrentUser = Depends(require_admin)
) -> dict:
    if not repositories.delete_personnel(db, pid):
        raise HTTPException(status_code=404, detail="Personnel not found")
    return {"deleted": pid}


# --------------------------------------------------------------------------- #
# Unit rates (Finance or admin)
# --------------------------------------------------------------------------- #
@router.get("/rates")
def list_rates(
    db: Session = Depends(get_db), _: CurrentUser = Depends(require_finance)
) -> list:
    return repositories.list_unit_rates(db)


@router.post("/rates")
def upsert_rate(
    body: UnitRateIn, db: Session = Depends(get_db), _: CurrentUser = Depends(require_finance)
) -> dict:
    return repositories.upsert_unit_rate(db, body.model_dump())


@router.delete("/rates/{rate_id}")
def delete_rate(
    rate_id: int, db: Session = Depends(get_db), _: CurrentUser = Depends(require_finance)
) -> dict:
    if not repositories.delete_unit_rate(db, rate_id):
        raise HTTPException(status_code=404, detail="Rate not found")
    return {"deleted": rate_id}


# --------------------------------------------------------------------------- #
# Team presets
# --------------------------------------------------------------------------- #
@router.get("/presets")
def list_presets(
    db: Session = Depends(get_db), _: CurrentUser = Depends(require_admin)
) -> list:
    return repositories.list_team_presets(db)


@router.get("/presets/{preset_id}")
def get_preset(
    preset_id: str, db: Session = Depends(get_db), _: CurrentUser = Depends(require_admin)
) -> dict:
    preset = repositories.get_team_preset(db, preset_id)
    if not preset:
        raise HTTPException(status_code=404, detail="Preset not found")
    return preset


@router.post("/presets")
def upsert_preset(
    body: TeamPresetIn, db: Session = Depends(get_db), _: CurrentUser = Depends(require_admin)
) -> dict:
    return repositories.upsert_team_preset(db, body.model_dump())


@router.delete("/presets/{preset_id}")
def delete_preset(
    preset_id: str, db: Session = Depends(get_db), _: CurrentUser = Depends(require_admin)
) -> dict:
    if not repositories.delete_team_preset(db, preset_id):
        raise HTTPException(status_code=404, detail="Preset not found")
    return {"deleted": preset_id}


# --------------------------------------------------------------------------- #
# Reference projects (relevant experience)
# --------------------------------------------------------------------------- #
@router.get("/reference-projects")
def list_reference_projects(
    db: Session = Depends(get_db), _: CurrentUser = Depends(require_admin)
) -> list:
    return repositories.list_reference_projects(db)


@router.post("/reference-projects")
def upsert_reference_project(
    body: ReferenceProjectIn, db: Session = Depends(get_db), _: CurrentUser = Depends(require_admin)
) -> dict:
    return repositories.upsert_reference_project(db, body.model_dump())


@router.delete("/reference-projects/{rid}")
def delete_reference_project(
    rid: str, db: Session = Depends(get_db), _: CurrentUser = Depends(require_admin)
) -> dict:
    if not repositories.delete_reference_project(db, rid):
        raise HTTPException(status_code=404, detail="Reference project not found")
    return {"deleted": rid}


# --------------------------------------------------------------------------- #
# Assets (CV / experience metadata + file upload)
# --------------------------------------------------------------------------- #
@router.get("/assets")
def list_assets(
    kind: str | None = None,
    db: Session = Depends(get_db),
    _: CurrentUser = Depends(require_admin),
) -> list:
    return repositories.list_assets(db, kind)


@router.post("/assets")
def upsert_asset(
    body: AssetIn, db: Session = Depends(get_db), _: CurrentUser = Depends(require_admin)
) -> dict:
    return repositories.upsert_asset(db, body.model_dump())


@router.post("/assets/upload")
async def upload_asset(
    asset_id: str = Form(...),
    kind: str = Form("cv"),
    role: str | None = Form(None),
    reference_project_id: str | None = Form(None),
    file: UploadFile = File(...),
    db: Session = Depends(get_db),
    user: CurrentUser = Depends(require_admin),
) -> dict:
    if not file.filename:
        raise HTTPException(status_code=400, detail="A file is required")
    content = await file.read()
    info = save_asset_file(asset_id, file.filename, content)
    data = {
        "id": asset_id,
        "kind": kind,
        "role": role,
        "reference_project_id": reference_project_id,
        "filename": file.filename,
        "uploaded_by": user.user,
        **info,
    }
    return repositories.upsert_asset(db, data)


@router.delete("/assets/{asset_id}")
def delete_asset(
    asset_id: str, db: Session = Depends(get_db), _: CurrentUser = Depends(require_admin)
) -> dict:
    if not repositories.delete_asset(db, asset_id):
        raise HTTPException(status_code=404, detail="Asset not found")
    return {"deleted": asset_id}
