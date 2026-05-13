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
