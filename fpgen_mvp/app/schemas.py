from typing import Any
from pydantic import BaseModel, Field


class ParseRequest(BaseModel):
    rfp_text: str = Field(min_length=20)
    project_hint: str | None = None


class ParseResponse(BaseModel):
    parsed: dict[str, Any]


class FeeRoleInput(BaseModel):
    role: str
    rate: float
    hours_by_phase: dict[str, float] = Field(default_factory=dict)


class SubconsultantInput(BaseModel):
    name: str
    fee: float
    included_in_lump_sum: bool = True


class TravelInput(BaseModel):
    trips: int = 0
    people_per_trip: int = 0
    unit_cost: float | None = None
    include_in_lump_sum: bool = True


class FeeInput(BaseModel):
    roles: list[FeeRoleInput] = Field(default_factory=list)
    overhead_pct: float | None = None
    subconsultants: list[SubconsultantInput] = Field(default_factory=list)
    travel: TravelInput = Field(default_factory=TravelInput)
    misc_reimbursables: float = 0


class GenerateRequest(BaseModel):
    rfp_text: str = Field(min_length=20)
    fee_input: FeeInput = Field(default_factory=FeeInput)
    selected_reference_ids: list[str] = Field(default_factory=list)
    overrides: dict[str, Any] = Field(default_factory=dict)


class ProposalResponse(BaseModel):
    proposal_id: str
    proposal: dict[str, Any]


class ProposalUpdateRequest(BaseModel):
    sections: dict[str, str] = Field(default_factory=dict)


class SectionRegenerateRequest(BaseModel):
    instruction: str | None = Field(default=None, max_length=2000)
    # Echoed back from a preview response so Apply commits the previewed result
    # without spending a second LLM call.
    input_patch: dict[str, Any] | None = None
    apply_text: str | None = None
    commit: bool = False


class ExperienceReselectRequest(BaseModel):
    selected_reference_ids: list[str] = Field(default_factory=list)
    auto: bool = False
    limit: int = Field(default=3, ge=1, le=10)
    commit: bool = False


class FeeRecalculateRequest(BaseModel):
    fee_input: FeeInput
    commit: bool = False


class RegenerationResponse(BaseModel):
    proposal_id: str
    committed: bool
    changed_sections: list[str] = Field(default_factory=list)
    stale_sections: list[dict[str, Any]] = Field(default_factory=list)
    proposal: dict[str, Any]
    input_patch: dict[str, Any] | None = None
    notice: str | None = None


class FileParseResponse(BaseModel):
    filename: str
    extracted_text: str
    parsed: dict[str, Any]


class InDesignExportResponse(BaseModel):
    proposal_id: str
    export_path: str
    download_url: str


class JsxExportRequest(BaseModel):
    cv_assignments: dict[str, str] = Field(default_factory=dict)
    experience_ids: list[str] = Field(default_factory=list)
    template_id: str = "commercial"


# --------------------------------------------------------------------------- #
# Admin / setup payloads
# --------------------------------------------------------------------------- #
class PersonnelIn(BaseModel):
    id: str = Field(min_length=1)
    name: str = Field(min_length=1)
    title: str = ""
    roles: list[str] = Field(default_factory=list)
    cv_asset_id: str | None = None


class UnitRateIn(BaseModel):
    role: str = Field(min_length=1)
    rate: float = Field(ge=0)
    currency: str = "USD"
    effective_from: str | None = None


class TeamAssignmentIn(BaseModel):
    role: str = Field(min_length=1)
    person_id: str | None = None
    rate: float | None = None


class TeamPresetIn(BaseModel):
    id: str = Field(min_length=1)
    name: str = Field(min_length=1)
    types: list[str] = Field(default_factory=list)
    assignments: list[TeamAssignmentIn] = Field(default_factory=list)


class ReferenceProjectIn(BaseModel):
    id: str = Field(min_length=1)
    name: str = Field(min_length=1)
    project_type: str = ""
    location: str = ""
    keywords: list[str] = Field(default_factory=list)
    summary: str = ""


class AssetIn(BaseModel):
    id: str = Field(min_length=1)
    kind: str = "cv"
    role: str | None = None
    reference_project_id: str | None = None
    filename: str | None = None
    storage_ref: str | None = None
    mime_type: str | None = None
    size_bytes: int | None = None
