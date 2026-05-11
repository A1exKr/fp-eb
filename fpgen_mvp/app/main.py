import json
from pathlib import Path

from fastapi import FastAPI, File, Form, HTTPException, UploadFile
from fastapi.responses import FileResponse

from app.config import settings
from app.schemas import (
    FileParseResponse,
    GenerateRequest,
    InDesignExportResponse,
    ParseRequest,
    ParseResponse,
    ProposalResponse,
    ProposalUpdateRequest,
    FeeInput,
)
from app.services.indd_exporter import InDesignExportError, export_proposal_to_indd, get_indd_export_capability
from app.services.fee_engine import calculate_fee
from app.services.parser import parse_rfp
from app.services.proposal_builder import build_proposal_payload, update_proposal_sections
from app.services.relevant_selector import select_relevant_projects
from app.services.rfp_file_service import extract_rfp_text
from app.services.storage import load_proposal, save_proposal, update_proposal


app = FastAPI(title="FP-GEN MVP API", version="0.1.0")
STATIC_DIR = Path(__file__).resolve().parent / "static"


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
def generate_endpoint(req: GenerateRequest) -> ProposalResponse:
    parsed = parse_rfp(rfp_text=req.rfp_text, project_hint=req.overrides.get("project_name"))

    if req.overrides:
        parsed_project = parsed.setdefault("project", {})
        for key, value in req.overrides.items():
            if key in {"project_name", "client_name", "location", "project_type", "site_area", "duration"}:
                parsed_project[key] = value

    fee = calculate_fee(req.fee_input)
    relevant = select_relevant_projects(parsed, req.selected_reference_ids)
    proposal_payload = build_proposal_payload(
        parsed=parsed,
        fee=fee,
        relevant=relevant,
        fee_input=req.fee_input,
    )
    proposal_id, _ = save_proposal(proposal_payload)

    return ProposalResponse(proposal_id=proposal_id, proposal=proposal_payload)


@app.post("/v1/proposals/generate/file", response_model=ProposalResponse)
def generate_file_endpoint(
    rfp_file: UploadFile = File(...),
    fee_input_json: str = Form(default="{}"),
    selected_reference_ids_json: str = Form(default="[]"),
    overrides_json: str = Form(default="{}"),
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
    parsed = parse_rfp(rfp_text=extracted_text, project_hint=overrides.get("project_name"))

    if overrides:
        parsed_project = parsed.setdefault("project", {})
        for key, value in overrides.items():
            if key in {"project_name", "client_name", "location", "project_type", "site_area", "duration"}:
                parsed_project[key] = value

    fee = calculate_fee(fee_input)
    relevant = select_relevant_projects(parsed, selected_reference_ids)
    proposal_payload = build_proposal_payload(
        parsed=parsed,
        fee=fee,
        relevant=relevant,
        fee_input=fee_input,
    )
    proposal_id, _ = save_proposal(proposal_payload)

    return ProposalResponse(proposal_id=proposal_id, proposal=proposal_payload)


@app.get("/v1/proposals/{proposal_id}", response_model=ProposalResponse)
def get_proposal(proposal_id: str) -> ProposalResponse:
    proposal = load_proposal(proposal_id)
    if not proposal:
        raise HTTPException(status_code=404, detail="Proposal not found")
    return ProposalResponse(
        proposal_id=proposal["proposal_id"],
        proposal=proposal["payload"],
    )


@app.put("/v1/proposals/{proposal_id}", response_model=ProposalResponse)
def update_proposal_endpoint(proposal_id: str, req: ProposalUpdateRequest) -> ProposalResponse:
    proposal = load_proposal(proposal_id)
    if not proposal:
        raise HTTPException(status_code=404, detail="Proposal not found")

    updated_payload = update_proposal_sections(proposal["payload"], req.sections)
    wrapped = update_proposal(proposal_id=proposal_id, payload=updated_payload)
    if not wrapped:
        raise HTTPException(status_code=404, detail="Proposal not found")

    return ProposalResponse(proposal_id=proposal_id, proposal=wrapped["payload"])


@app.post("/v1/proposals/{proposal_id}/export/indd", response_model=InDesignExportResponse)
def export_indd(proposal_id: str) -> InDesignExportResponse:
    proposal = load_proposal(proposal_id)
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
def download_indd(proposal_id: str) -> FileResponse:
    proposal = load_proposal(proposal_id)
    if not proposal:
        raise HTTPException(status_code=404, detail="Proposal not found")

    project_name = proposal["payload"].get("project", {}).get("project_name", proposal_id)
    safe_name = "".join(ch if ch.isalnum() or ch in " ._-" else "_" for ch in project_name).strip() or "proposal"
    out_path = settings.export_dir / proposal_id / f"{safe_name[:80]}.indd"
    if not out_path.exists():
        raise HTTPException(status_code=404, detail="INDD export not found. Generate it first.")
    return FileResponse(out_path, media_type="application/octet-stream", filename=out_path.name)
