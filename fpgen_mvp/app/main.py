import json
import traceback
from contextlib import asynccontextmanager
from datetime import datetime
from pathlib import Path

from fastapi import Depends, FastAPI, File, Form, HTTPException, Request, UploadFile
from fastapi.responses import FileResponse, JSONResponse, StreamingResponse
from fastapi.staticfiles import StaticFiles
from sqlalchemy.orm import Session

from app.config import settings
from app.db import IS_SQLITE, get_db, init_db
from app import repositories
from app.auth import CurrentUser, get_current_user
from app.routers import admin
from app.schemas import (
    FileParseResponse,
    GenerateRequest,
    InDesignExportResponse,
    JsxExportRequest,
    ParseRequest,
    ParseResponse,
    ProposalResponse,
    ProposalUpdateRequest,
    FeeInput,
)
from app.services.indd_exporter import InDesignExportError, export_proposal_to_indd, get_indd_export_capability
from app.services.fee_engine import calculate_fee
from app.services.jsx_exporter import build_jsx_bundle
from app.services.parser import parse_rfp
from app.services.proposal_builder import build_proposal_payload, update_proposal_sections
from app.services.relevant_selector import select_relevant_projects
from app.services.rfp_file_service import extract_rfp_text


@asynccontextmanager
async def lifespan(app: FastAPI):
    # Local/dev convenience: ensure SQLite tables exist. On exaBase PostgreSQL,
    # Alembic migrations manage the schema at container startup.
    if IS_SQLITE:
        init_db()
    yield


app = FastAPI(title="FP-GEN MVP API", version="0.1.0", lifespan=lifespan)
STATIC_DIR = Path(__file__).resolve().parent / "static"
ADMIN_DIR = STATIC_DIR / "admin"

app.include_router(admin.router)
if ADMIN_DIR.is_dir():
    app.mount("/admin", StaticFiles(directory=str(ADMIN_DIR), html=True), name="admin")


@app.get("/v1/me")
def whoami(user: CurrentUser = Depends(get_current_user)) -> dict:
    return {"user": user.user, "email": user.email, "groups": user.groups}


@app.exception_handler(Exception)
async def global_exception_handler(request: Request, exc: Exception) -> JSONResponse:
    traceback.print_exc()
    return JSONResponse(
        status_code=500,
        content={"detail": str(exc) or "Internal server error"},
    )


@app.get("/")
def index() -> FileResponse:
    return FileResponse(STATIC_DIR / "index.html")


@app.get("/ui")
def ui() -> FileResponse:
    return FileResponse(STATIC_DIR / "index.html")


@app.get("/health")
def health() -> dict:
    return {
        "status": "ok",
        "capabilities": {
            "indd_export": get_indd_export_capability(),
        },
    }


@app.get("/v1/capabilities")
def capabilities() -> dict:
    return {
        "indd_export": get_indd_export_capability(),
    }


@app.get("/v1/reference-projects")
def reference_projects(db: Session = Depends(get_db)) -> list:
    return repositories.list_reference_projects(db)


@app.get("/v1/personnel")
def personnel_endpoint(db: Session = Depends(get_db)) -> list:
    return repositories.list_personnel(db)


@app.get("/v1/team-presets")
def team_presets_endpoint(db: Session = Depends(get_db)) -> list:
    return repositories.list_team_presets(db)


@app.get("/proposals/{proposal_id}/review")
def review_page(proposal_id: str) -> FileResponse:
    return FileResponse(STATIC_DIR / "review.html")


@app.get("/v1/assets")
def list_assets() -> dict:
    result: dict[str, list[dict]] = {"template": [], "cvs": [], "experience": []}
    for category in result:
        folder = settings.assets_dir / category
        if folder.is_dir():
            result[category] = [
                {"asset_id": p.stem, "filename": p.name}
                for p in sorted(folder.glob("*.indd"))
            ]
    return result


@app.post("/v1/assets/upload")
async def upload_asset(
    category: str = Form(...),
    asset_id: str = Form(...),
    file: UploadFile = File(...),
) -> dict:
    allowed_categories = {"template", "cvs", "experience"}
    if category not in allowed_categories:
        raise HTTPException(status_code=400, detail=f"category must be one of {allowed_categories}")
    if not asset_id or not asset_id.replace("-", "").replace("_", "").isalnum():
        raise HTTPException(status_code=400, detail="asset_id must be kebab-case alphanumeric")
    if not file.filename or not file.filename.lower().endswith(".indd"):
        raise HTTPException(status_code=400, detail="Only .indd files are accepted")

    dest_dir = settings.assets_dir / category
    dest_dir.mkdir(parents=True, exist_ok=True)
    dest_path = dest_dir / f"{asset_id}.indd"
    content = await file.read()
    dest_path.write_bytes(content)
    return {"uploaded": str(dest_path.relative_to(settings.assets_dir)), "size_bytes": len(content)}


@app.post("/v1/parse", response_model=ParseResponse)
def parse_endpoint(req: ParseRequest) -> ParseResponse:
    parsed = parse_rfp(rfp_text=req.rfp_text, project_hint=req.project_hint)
    return ParseResponse(parsed=parsed)


@app.post("/v1/parse/file", response_model=FileParseResponse)
def parse_file_endpoint(
    rfp_file: UploadFile = File(...),
    project_hint: str | None = Form(default=None),
) -> FileParseResponse:
    try:
        filename, extracted_text = extract_rfp_text(rfp_file)
    except ValueError as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc

    parsed = parse_rfp(rfp_text=extracted_text, project_hint=project_hint)
    return FileParseResponse(filename=filename, extracted_text=extracted_text, parsed=parsed)


@app.post("/v1/fee/calculate")
def fee_endpoint(fee_input: FeeInput) -> dict:
    fee = calculate_fee(fee_input)
    return {"fee": fee}


@app.post("/v1/proposals/generate", response_model=ProposalResponse)
def generate_endpoint(req: GenerateRequest, db: Session = Depends(get_db)) -> ProposalResponse:
    parsed = parse_rfp(rfp_text=req.rfp_text, project_hint=req.overrides.get("project_name") or req.overrides.get("name"))

    if req.overrides:
        proj = parsed.setdefault("project", {})
        cli = parsed.setdefault("client", {})
        # Accept both old key names (backward compat) and new schema key names
        if v := req.overrides.get("name") or req.overrides.get("project_name"):
            proj["name"] = v
        if v := req.overrides.get("location"):
            proj["location"] = v
        if v := req.overrides.get("type") or req.overrides.get("project_type"):
            proj["type"] = v
        if v := req.overrides.get("siteArea") or req.overrides.get("site_area"):
            proj["siteArea"] = v
        if v := req.overrides.get("client_name"):
            cli["name"] = v

    fee = calculate_fee(req.fee_input)
    relevant = select_relevant_projects(db, parsed, req.selected_reference_ids)
    proposal_payload = build_proposal_payload(
        parsed=parsed,
        fee=fee,
        relevant=relevant,
        fee_input=req.fee_input,
    )
    proposal_id, _ = repositories.save_proposal(db, proposal_payload)

    return ProposalResponse(proposal_id=proposal_id, proposal=proposal_payload)


@app.post("/v1/proposals/generate/file", response_model=ProposalResponse)
def generate_file_endpoint(
    rfp_file: UploadFile = File(...),
    fee_input_json: str = Form(default="{}"),
    selected_reference_ids_json: str = Form(default="[]"),
    overrides_json: str = Form(default="{}"),
    db: Session = Depends(get_db),
) -> ProposalResponse:
    try:
        _, extracted_text = extract_rfp_text(rfp_file)
        fee_payload = json.loads(fee_input_json or "{}")
        selected_reference_ids = json.loads(selected_reference_ids_json or "[]")
        overrides = json.loads(overrides_json or "{}")
    except ValueError as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc
    except json.JSONDecodeError as exc:
        raise HTTPException(status_code=400, detail=f"Invalid JSON in form fields: {exc}") from exc

    fee_input = FeeInput.model_validate(fee_payload)
    parsed = parse_rfp(rfp_text=extracted_text, project_hint=overrides.get("project_name") or overrides.get("name"))

    if overrides:
        proj = parsed.setdefault("project", {})
        cli = parsed.setdefault("client", {})
        if v := overrides.get("name") or overrides.get("project_name"):
            proj["name"] = v
        if v := overrides.get("location"):
            proj["location"] = v
        if v := overrides.get("type") or overrides.get("project_type"):
            proj["type"] = v
        if v := overrides.get("siteArea") or overrides.get("site_area"):
            proj["siteArea"] = v
        if v := overrides.get("client_name"):
            cli["name"] = v

    fee = calculate_fee(fee_input)
    relevant = select_relevant_projects(db, parsed, selected_reference_ids)
    proposal_payload = build_proposal_payload(
        parsed=parsed,
        fee=fee,
        relevant=relevant,
        fee_input=fee_input,
    )
    proposal_id, _ = repositories.save_proposal(db, proposal_payload)

    return ProposalResponse(proposal_id=proposal_id, proposal=proposal_payload)


@app.get("/v1/proposals/{proposal_id}", response_model=ProposalResponse)
def get_proposal(proposal_id: str, db: Session = Depends(get_db)) -> ProposalResponse:
    proposal = repositories.load_proposal(db, proposal_id)
    if not proposal:
        raise HTTPException(status_code=404, detail="Proposal not found")
    return ProposalResponse(proposal_id=proposal_id, proposal=proposal["payload"])


@app.put("/v1/proposals/{proposal_id}", response_model=ProposalResponse)
def update_proposal_endpoint(
    proposal_id: str, req: ProposalUpdateRequest, db: Session = Depends(get_db)
) -> ProposalResponse:
    proposal = repositories.load_proposal(db, proposal_id)
    if not proposal:
        raise HTTPException(status_code=404, detail="Proposal not found")

    updated_payload = update_proposal_sections(proposal["payload"], req.sections)
    wrapped = repositories.update_proposal(db, proposal_id, updated_payload)
    if not wrapped:
        raise HTTPException(status_code=404, detail="Proposal not found")

    return ProposalResponse(proposal_id=proposal_id, proposal=wrapped["payload"])


@app.post("/v1/proposals/{proposal_id}/export/jsx")
def export_jsx(proposal_id: str, req: JsxExportRequest, db: Session = Depends(get_db)) -> StreamingResponse:
    proposal = repositories.load_proposal(db, proposal_id)
    if not proposal:
        raise HTTPException(status_code=404, detail="Proposal not found")

    zip_bytes = build_jsx_bundle(
        proposal=proposal["payload"],
        cv_assignments=req.cv_assignments,
        experience_ids=req.experience_ids,
        template_id=req.template_id,
        assets_dir=settings.assets_dir,
    )
    project_name = proposal["payload"].get("project", {}).get("name", proposal_id)
    safe = "".join(c if c.isalnum() or c in "-_" else "_" for c in project_name).strip() or "proposal"
    stamp = datetime.now().strftime("%Y%m%d_%H%M")
    filename = f"{safe[:60]}_{stamp}.zip"
    return StreamingResponse(
        iter([zip_bytes]),
        media_type="application/zip",
        headers={"Content-Disposition": f'attachment; filename="{filename}"'},
    )


@app.post("/v1/proposals/{proposal_id}/export/indd", response_model=InDesignExportResponse)
def export_indd(proposal_id: str, db: Session = Depends(get_db)) -> InDesignExportResponse:
    proposal = repositories.load_proposal(db, proposal_id)
    if not proposal:
        raise HTTPException(status_code=404, detail="Proposal not found")

    try:
        out_path = export_proposal_to_indd(proposal_id=proposal_id, proposal=proposal["payload"])
    except InDesignExportError as exc:
        raise HTTPException(status_code=500, detail=str(exc)) from exc

    return InDesignExportResponse(
        proposal_id=proposal_id,
        export_path=str(out_path),
        download_url=f"/v1/proposals/{proposal_id}/export/indd/download",
    )


@app.get("/v1/proposals/{proposal_id}/export/indd/download")
def download_indd(proposal_id: str, db: Session = Depends(get_db)) -> FileResponse:
    proposal = repositories.load_proposal(db, proposal_id)
    if not proposal:
        raise HTTPException(status_code=404, detail="Proposal not found")

    project_name = proposal["payload"].get("project", {}).get("project_name", proposal_id)
    safe_name = "".join(ch if ch.isalnum() or ch in " ._-" else "_" for ch in project_name).strip() or "proposal"
    out_path = settings.export_dir / proposal_id / f"{safe_name[:80]}.indd"
    if not out_path.exists():
        raise HTTPException(status_code=404, detail="INDD export not found. Generate it first.")
    return FileResponse(out_path, media_type="application/octet-stream", filename=out_path.name)
