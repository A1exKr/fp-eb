"""RFP parser 窶・uses the full Nikken prompt when OpenAI is available, regex fallback otherwise."""

import json
import re

from app.config import settings
from app.services.openai_service import OpenAIService

# ---------------------------------------------------------------------------
# System prompt (Prompt_all-parse.txt)
# ---------------------------------------------------------------------------

_SYSTEM_PROMPT = """\
You are an experienced architect/master planner, licensed by both the AIA and JIA, \
and have worked on international projects for many years.
Please analyze the provided RFP text and create a well-structured JSON object that can \
be used directly by the FP-GEN Fee Proposal Generator.
Ensure all keys and string values are properly quoted, numeric fields are plain numbers \
(not strings), and the JSON is fully valid and well-formed.
If any item is not clearly specified, infer it professionally from context 窶・never write \
"not specified." Be concise and direct.

Return only one JSON object that follows exactly this structure:

{
  "project": { "name": "", "type": "", "location": "", "siteArea": "" },
  "client": { "name": "" },
  "understanding": { "understanding": "" },
  "methodology": { "text": "" },
  "scope": { "scopeList": [], "deliverablesList": [] },
  "schedule": { "totalWeeks": 0, "milestones": [] },
  "team": {
      "principal": {"name": "", "title": ""},
      "pm": {"name": "", "title": ""}
  },
  "fee": {
      "currency": "",
      "rates": {},
      "effortByPhase": {},
      "overheadPct": 0,
      "travel": {},
      "subconsultants": []
  },
  "experience": [],
  "assumptions": { "defaultText": "" }
}

Guidelines:

project 竊・Extract project name, type, location, site area.

client 竊・Extract client name or organisation.

understanding 竊・Write a clear, concise 150-200 word executive summary covering: project \
objectives, site context, client goals, key challenges, and expected outcomes.

methodology.text 竊・Choose and customise either Architecture Design Process Methodology \
or Master Planning Process Methodology. Include all six numbered steps as detailed prose.

scope 竊・Derive scopeList (key design or study tasks) and deliverablesList (expected outputs). \
Be specific and actionable.

schedule 竊・Estimate totalWeeks realistically and list key milestones (e.g. ["Kick-off","WS2","WS3","Final"]).

team 竊・Include principal and project manager names/titles if known; infer sensible titles if not.

fee.currency 竊・"USD" by default.
fee.rates 竊・flat object mapping role name to hourly rate in USD, e.g. {"Urban Planner": 180, "Architect": 200}.
fee.effortByPhase 竊・NESTED object: outer keys are phase names, inner keys are role names \
(matching fee.rates keys), values are estimated integer hours. Every role must appear in every phase. \
Example:
{
  "Kick-off":            { "Urban Planner": 20, "Architect": 10 },
  "Concept Development": { "Urban Planner": 60, "Architect": 50 },
  "Finalization":        { "Urban Planner": 30, "Architect": 25 }
}
fee.overheadPct 竊・decimal fraction, e.g. 0.10 for 10%.
fee.travel 竊・object with keys: trips (int), peoplePerTrip (int), unitCostUSD (number), includedInLumpSum (bool).
fee.subconsultants 竊・array of {"name": "", "feeUSD": 0, "includedInLumpSum": true}.

experience 竊・List 2-3 relevant Nikken projects with "name", "location", "summary".

assumptions.defaultText 竊・Professional commercial disclaimer covering taxes, reimbursables, \
optional services, specialty consultants.

Cover letter template (embed in assumptions.defaultText or a separate fee.notes field):
Dear Sirs,\\n\\nWe are sincerely honored to be invited to submit our proposal for your esteemed \
project. Based on our in-depth understanding of the project, we believe we can establish a solid \
design solution well-suited for the site. To tackle the complex nature of this project, we have \
assembled a highly skilled multidisciplinary team under the leadership of our top management. \
The project will be overseen personally by myself, Senior Executive Officer of Nikken Sekkei Ltd.\
\\n\\nWe truly hope our proposal meets your expectations and look forward to serving you.\\n\\n\
Sincerely yours,\\n\\nWataru TANAKA\\nSenior Executive Officer\\nNikken Sekkei Ltd.

Inside strings use \\' instead of ". Use \\n for line breaks in long text fields.

Output Rule: Return only the JSON object, nothing else. Response must start with { and end with }.
"""


# ---------------------------------------------------------------------------
# Template suggestion (server-side, not AI)
# ---------------------------------------------------------------------------

_TYPE_TO_TEMPLATE: list[tuple[list[str], str]] = [
    (["master plan", "urban", "district", "masterplan", "city"], "master-plan"),
    (["technical", "infrastructure", "industrial", "engineering", "feasibility"], "technical"),
    (["commercial", "mixed-use", "office", "retail", "hotel", "hospitality", "mixed use", "residential", "housing", "architecture", "design", "building"], "commercial"),
]


def _suggest_template(project_type: str) -> str:
    lower = project_type.lower()
    for keywords, template_id in _TYPE_TO_TEMPLATE:
        if any(kw in lower for kw in keywords):
            return template_id
    return "commercial"


# ---------------------------------------------------------------------------
# CV suggestions (server-side, not AI)
# ---------------------------------------------------------------------------

def _suggest_cvs(rates: dict, assets_registry: dict) -> list[str]:
    cv_map: dict = assets_registry.get("cvs", {})
    suggestions: list[str] = []
    for role in rates:
        asset_id = cv_map.get(role) or cv_map.get(role.lower())
        if asset_id and asset_id not in suggestions:
            suggestions.append(asset_id)
    return suggestions


# ---------------------------------------------------------------------------
# Suggested reference IDs (server-side, not AI)
# ---------------------------------------------------------------------------

def _load_reference_projects() -> list[dict]:
    path = settings.reference_projects_path
    if not path.exists():
        return []
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except Exception:
        return []


def _load_assets_registry() -> dict:
    path = settings.assets_registry_path
    if not path.exists():
        return {}
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except Exception:
        return {}


def _suggest_reference_ids(experience: list[dict]) -> list[str]:
    ref_projects = _load_reference_projects()
    if not ref_projects or not experience:
        return []

    suggested: list[str] = []
    for exp in experience:
        exp_name = (exp.get("name") or "").lower()
        exp_summary = (exp.get("summary") or "").lower()
        exp_tokens = set(re.split(r"\W+", f"{exp_name} {exp_summary}")) - {"", "the", "a", "and", "for", "of"}

        best_id: str | None = None
        best_score = 0
        for rp in ref_projects:
            rp_name = (rp.get("name") or "").lower()
            rp_keywords = {kw.lower() for kw in rp.get("keywords", [])}
            rp_tokens = set(re.split(r"\W+", rp_name)) | rp_keywords
            score = len(exp_tokens & rp_tokens)
            if score > best_score:
                best_score = score
                best_id = rp.get("id")

        if best_id and best_score >= 1 and best_id not in suggested:
            suggested.append(best_id)

    return suggested[:3]


# ---------------------------------------------------------------------------
# Computed additions (post-AI)
# ---------------------------------------------------------------------------

def _add_computed_fields(parsed: dict) -> dict:
    project_type = parsed.get("project", {}).get("type", "")
    parsed["template_suggestion"] = _suggest_template(project_type)
    registry = _load_assets_registry()
    rates = parsed.get("fee", {}).get("rates", {})
    parsed["cv_suggestions"] = _suggest_cvs(rates, registry)
    parsed["suggested_reference_ids"] = _suggest_reference_ids(parsed.get("experience", []))
    return parsed


# ---------------------------------------------------------------------------
# Merge helper
# ---------------------------------------------------------------------------

def _safe_dict(value: object) -> dict:
    return value if isinstance(value, dict) else {}


def _safe_list(value: object) -> list:
    return value if isinstance(value, list) else []


def _merge_missing(ai: dict, fallback: dict) -> dict:
    for top_key in ("project", "client", "understanding", "methodology", "scope", "schedule", "team", "fee", "assumptions"):
        if not ai.get(top_key):
            ai[top_key] = fallback.get(top_key, {})
        elif isinstance(ai.get(top_key), dict) and isinstance(fallback.get(top_key), dict):
            for sub_key, sub_val in fallback[top_key].items():
                if not ai[top_key].get(sub_key):
                    ai[top_key][sub_key] = sub_val

    if not _safe_list(ai.get("experience")):
        ai["experience"] = fallback.get("experience", [])

    # Ensure effortByPhase is properly nested (not flat {phase: total})
    ebp = _safe_dict(ai.get("fee", {})).get("effortByPhase", {})
    if ebp:
        first_val = next(iter(ebp.values()), None)
        if not isinstance(first_val, dict):
            rates = _safe_dict(ai["fee"].get("rates")) or _safe_dict(fallback.get("fee", {}).get("rates", {}))
            roles = list(rates.keys()) or ["Urban Planner", "Architect"]
            nested: dict = {}
            for phase, total in ebp.items():
                per_role = round(float(total) / max(len(roles), 1))
                nested[phase] = {r: per_role for r in roles}
            ai["fee"]["effortByPhase"] = nested

    return ai


# ---------------------------------------------------------------------------
# Fallback (no OpenAI)
# ---------------------------------------------------------------------------

def _fallback_parse(rfp_text: str, project_hint: str | None = None) -> dict:
    def labeled(label: str) -> str:
        pattern = rf"(?im)^\s*{re.escape(label)}\s*:\s*(.+?)\s*$"
        m = re.search(pattern, rfp_text)
        return m.group(1).strip() if m else ""

    project_name = project_hint or labeled("Project") or "Untitled Project"
    lower = rfp_text.lower()
    if "master plan" in lower:
        project_type = "Master Plan"
    elif "design" in lower:
        project_type = "Design"
    else:
        project_type = "General Consulting"

    phases: list[str] = []
    if "kick-off" in lower:
        phases.append("Kick-off")
    if any(kw in lower for kw in ["concept", "option"]):
        phases.append("Concept Development")
    if any(kw in lower for kw in ["final", "presentation"]):
        phases.append("Finalization")
    if not phases:
        phases = ["Kick-off", "Development", "Finalization"]

    default_roles = ["Urban Planner", "Architect"]
    default_rates = {"Urban Planner": 180, "Architect": 200}
    effort_by_phase = {phase: {role: 20 for role in default_roles} for phase in phases}
    scope_summary = " ".join(l.strip() for l in rfp_text.splitlines() if l.strip())[:800]

    parsed = {
        "project": {
            "name": project_name,
            "type": project_type,
            "location": labeled("Location") or "TBD",
            "siteArea": labeled("Site Area"),
        },
        "client": {"name": labeled("Client") or "Unknown Client"},
        "understanding": {"understanding": scope_summary},
        "methodology": {"text": "1. Kick-off\n2. Analysis\n3. Concept Development\n4. Refinement\n5. Finalization\n6. Quality Control"},
        "scope": {
            "scopeList": ["Site analysis", "Concept design", "Final deliverables"],
            "deliverablesList": ["Proposal document", "Presentation materials"],
        },
        "schedule": {"totalWeeks": 16, "milestones": phases},
        "team": {
            "principal": {"name": "Wataru Tanaka", "title": "Senior Executive Officer"},
            "pm": {"name": "TBD", "title": "Project Manager"},
        },
        "fee": {
            "currency": "USD",
            "rates": default_rates,
            "effortByPhase": effort_by_phase,
            "overheadPct": 0.10,
            "travel": {"trips": 2, "peoplePerTrip": 2, "unitCostUSD": 6000, "includedInLumpSum": True},
            "subconsultants": [],
        },
        "experience": [],
        "assumptions": {
            "defaultText": (
                "Fees exclude all applicable taxes and levies unless stated otherwise. "
                "Travel costs are included only when explicitly listed in the financial breakdown. "
                "Optional services such as animation and physical models are excluded from the lump sum. "
                "Specialty consultants not listed under included costs are excluded from this fee."
            ),
        },
    }
    return _add_computed_fields(parsed)


# ---------------------------------------------------------------------------
# Public entry point
# ---------------------------------------------------------------------------

def parse_rfp(rfp_text: str, project_hint: str | None = None) -> dict:
    openai = OpenAIService()
    fallback = _fallback_parse(rfp_text, project_hint)

    if not openai.enabled:
        return fallback

    user_prompt = f"RFP TEXT:\n{rfp_text}"
    if project_hint:
        user_prompt = f"Project hint: {project_hint}\n\n{user_prompt}"

    try:
        ai_result = openai.json_completion(
            system_prompt=_SYSTEM_PROMPT,
            user_prompt=user_prompt,
        )
        merged = _merge_missing(ai=ai_result, fallback=fallback)
        return _add_computed_fields(merged)
    except Exception:
        return fallback

