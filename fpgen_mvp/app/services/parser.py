import re

from app.services.openai_service import OpenAIService


def _extract_labeled_value(rfp_text: str, label: str) -> str:
    pattern = rf"(?im)^\s*{re.escape(label)}\s*:\s*(.+?)\s*$"
    match = re.search(pattern, rfp_text)
    return match.group(1).strip() if match else ""


def _split_items(text: str) -> list[str]:
    normalized = re.sub(r"\band\b", ",", text, flags=re.IGNORECASE)
    return [item.strip(" .;:-") for item in normalized.split(",") if item.strip(" .;:-")]


def _extract_summary(rfp_text: str) -> str:
    lines = [line.strip() for line in rfp_text.splitlines()]
    filtered = [
        line for line in lines
        if line and not re.match(r"(?i)^(client|project|location|duration|site area)\s*:", line)
    ]
    return " ".join(filtered)[:1200]


def _extract_deliverables(rfp_text: str) -> list[str]:
    patterns = [
        r"(?is)deliverables?\s*(?:include|:)\s*(.+?)(?:\.|\n)",
        r"(?is)expected to prepare\s*(.+?)(?:\.|\n)",
    ]
    for pattern in patterns:
        match = re.search(pattern, rfp_text)
        if match:
            items = _split_items(match.group(1))
            if items:
                return items
    return ["Proposal document", "Presentation materials"]


def _extract_milestones(rfp_text: str) -> list[str]:
    match = re.search(r"(?is)workshops?.*?during\s+(.+?)\s+stages?", rfp_text)
    if not match:
        return []
    return _split_items(match.group(1))


def _extract_disciplines(rfp_text: str) -> list[str]:
    match = re.search(r"(?is)required disciplines? include\s+(.+?)(?:\.|\n)", rfp_text)
    if match:
        items = _split_items(match.group(1))
        if items:
            return items
    return ["Planning", "Architecture"]


def _infer_project_type(rfp_text: str) -> str:
    lower_text = rfp_text.lower()
    if "master plan" in lower_text:
        return "Master Plan"
    if "design" in lower_text:
        return "Design"
    return "General Consulting"


def _merge_missing(parsed: dict, fallback: dict) -> dict:
    merged = {
        **fallback,
        **parsed,
        "project": {
            **fallback.get("project", {}),
            **parsed.get("project", {}),
        },
    }

    for key in ["deliverables", "phases", "milestones", "required_disciplines", "keywords"]:
        if not merged.get(key):
            merged[key] = fallback.get(key, [])

    if not merged.get("scope_summary"):
        merged["scope_summary"] = fallback.get("scope_summary", "")

    project = merged.get("project", {})
    fallback_project = fallback.get("project", {})
    for key, value in fallback_project.items():
        if not project.get(key):
            project[key] = value
    merged["project"] = project
    return merged


def _fallback_parse(rfp_text: str, project_hint: str | None = None) -> dict:
    title = project_hint or _extract_labeled_value(rfp_text, "Project") or "Untitled Project"
    project_type = _infer_project_type(rfp_text)
    location = _extract_labeled_value(rfp_text, "Location")
    duration = _extract_labeled_value(rfp_text, "Duration")
    site_area = _extract_labeled_value(rfp_text, "Site Area")
    client_name = _extract_labeled_value(rfp_text, "Client") or "Unknown Client"
    deliverables = _extract_deliverables(rfp_text)
    milestones = _extract_milestones(rfp_text)
    required_disciplines = _extract_disciplines(rfp_text)

    phases = []
    if "kick-off" in rfp_text.lower():
        phases.append("Kick-off")
    if any(keyword in rfp_text.lower() for keyword in ["concept", "option", "review"]):
        phases.append("Concept Development")
    if any(keyword in rfp_text.lower() for keyword in ["final", "presentation", "submission"]):
        phases.append("Finalization")
    if not phases:
        phases = ["Kick-off", "Development", "Finalization"]

    return {
        "project": {
            "project_name": title,
            "client_name": client_name,
            "location": location or "Unknown",
            "project_type": project_type,
            "site_area": site_area,
            "duration": duration,
        },
        "scope_summary": _extract_summary(rfp_text),
        "deliverables": deliverables,
        "phases": phases,
        "milestones": milestones,
        "required_disciplines": required_disciplines,
        "keywords": [project_type.lower(), "proposal", *[discipline.lower() for discipline in required_disciplines[:3]]],
    }


def parse_rfp(rfp_text: str, project_hint: str | None = None) -> dict:
    openai = OpenAIService()

    if not openai.enabled:
        return _fallback_parse(rfp_text, project_hint)

    system_prompt = (
        "You extract structured RFP information for fee proposal generation. "
        "Return valid JSON only."
    )
    user_prompt = f"""
Extract these keys from the RFP text below:
- project.project_name
- project.client_name
- project.location
- project.project_type
- project.site_area
- project.duration
- scope_summary (1-2 paragraphs)
- deliverables (array of strings)
- phases (array of strings)
- milestones (array of strings)
- required_disciplines (array of strings)
- keywords (array of strings)

If unknown, use empty string or empty list.
Project hint: {project_hint or ""}

RFP TEXT:
{rfp_text}
"""

    try:
        parsed = openai.json_completion(system_prompt=system_prompt, user_prompt=user_prompt)
        return _merge_missing(parsed=parsed, fallback=_fallback_parse(rfp_text, project_hint))
    except Exception:
        return _fallback_parse(rfp_text, project_hint)
