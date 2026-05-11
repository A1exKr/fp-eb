from app.config import settings
from app.schemas import FeeInput
from app.services.openai_service import OpenAIService


def _build_team_section(parsed: dict, fee_input: FeeInput | None) -> str:
    lines: list[str] = []
    seen_roles: set[str] = set()

    for role_input in (fee_input.roles if fee_input else []):
        if role_input.role not in seen_roles:
            seen_roles.add(role_input.role)
            lines.append(f"- {role_input.role}")

    for discipline in parsed.get("required_disciplines", []):
        if discipline not in seen_roles:
            seen_roles.add(discipline)
            lines.append(f"- {discipline}")

    for subconsultant in (fee_input.subconsultants if fee_input else []):
        label = "included" if subconsultant.included_in_lump_sum else "excluded"
        lines.append(f"- {subconsultant.name} (specialty consultant, {label} in fee)")

    if not lines:
        return "- Team structure to be finalized during review phase."

    return "\n".join(lines)


def _build_schedule_section(parsed: dict, fee: dict) -> str:
    lines: list[str] = []
    duration = parsed.get("project", {}).get("duration", "")
    currency = fee.get("currency", settings.default_currency)
    emitted_phases: set[str] = set()

    if duration:
        lines.append(f"Overall Duration: {duration}")

    phases = parsed.get("phases") or []
    phase_totals = fee.get("phase_totals", {})
    if phases:
        lines.append("Phase Plan:")
        for phase in phases:
            if phase in phase_totals:
                emitted_phases.add(phase)
                lines.append(f"- {phase}: estimated labor fee {phase_totals[phase]:.2f} {currency}")
            else:
                lines.append(f"- {phase}")
    elif phase_totals:
        lines.append("Phase Plan:")
        for phase, amount in phase_totals.items():
            emitted_phases.add(phase)
            lines.append(f"- {phase}: estimated labor fee {amount:.2f} {currency}")

    for phase, amount in phase_totals.items():
        if phase not in emitted_phases:
            if "Phase Plan:" not in lines:
                lines.append("Phase Plan:")
            lines.append(f"- {phase}: estimated labor fee {amount:.2f} {currency}")

    milestones = parsed.get("milestones") or []
    if milestones:
        lines.append("Milestones:")
        for milestone in milestones:
            lines.append(f"- {milestone}")

    if not lines:
        return "Schedule to be finalized with client milestones."

    return "\n".join(lines)


def _build_financial_section(fee: dict, fee_input: FeeInput | None) -> str:
    currency = fee.get("currency", settings.default_currency)
    lines = [
        "Fee Breakdown:",
        f"- Labor: {fee.get('labor_total', 0):.2f} {currency}",
        f"- Overhead ({fee.get('overhead_pct', 0) * 100:.1f}%): {fee.get('overhead_total', 0):.2f} {currency}",
        f"- Included Subconsultants: {fee.get('subconsultants_included_total', 0):.2f} {currency}",
    ]

    travel = fee.get("travel", {})
    if travel:
        travel_mode = "included in lump sum" if travel.get("included_in_lump_sum") else "excluded from lump sum"
        lines.append(
            f"- Travel: {travel.get('total', 0):.2f} {currency} "
            f"({travel.get('trips', 0)} trips x {travel.get('people_per_trip', 0)} people x {travel.get('unit_cost', 0):.2f}; {travel_mode})"
        )

    lines.append(f"- Other Reimbursables: {fee.get('misc_reimbursables', 0):.2f} {currency}")
    lines.append(f"- Total Reimbursables Included: {fee.get('reimbursables_total', 0):.2f} {currency}")
    lines.append(f"- Total Lump Sum Fee: {fee.get('lump_sum_total', 0):.2f} {currency}")

    phase_totals = fee.get("phase_totals", {})
    if phase_totals:
        lines.append("Phase Labor Breakdown:")
        for phase, amount in phase_totals.items():
            lines.append(f"- {phase}: {amount:.2f} {currency}")

    if fee_input and fee_input.roles:
        lines.append("Role Effort Basis:")
        for role_input in fee_input.roles:
            total_hours = sum(role_input.hours_by_phase.values())
            lines.append(
                f"- {role_input.role}: {total_hours:.2f} hours at {role_input.rate:.2f} {currency}"
            )

    if fee.get("subconsultants_excluded"):
        lines.append("Excluded Costs:")
        for subconsultant in fee["subconsultants_excluded"]:
            lines.append(
                f"- {subconsultant['name']}: {subconsultant['fee']:.2f} {currency} excluded from lump sum"
            )

    lines.append("Commercial Notes:")
    lines.append("- Fees exclude taxes and levies unless stated otherwise.")
    lines.append("- Payment terms and invoicing schedule to be confirmed during commercial review.")

    return "\n".join(lines)


def _section_text(parsed: dict, fee: dict, relevant: list[dict], fee_input: FeeInput | None = None) -> dict:
    project = parsed.get("project", {})
    project_name = project.get("project_name", "Project")
    client_name = project.get("client_name", "Client")
    location = project.get("location", "")
    project_type = project.get("project_type", "")
    site_area = project.get("site_area", "")

    disciplines = parsed.get("required_disciplines") or []
    discipline_text = ", ".join(disciplines[:4]) if disciplines else "the required disciplines"

    cover_letter = (
        f"Dear {client_name},\n\n"
        f"We appreciate the opportunity to submit our proposal for {project_name}. "
        f"Our multidisciplinary team is prepared to deliver this {project_type} engagement"
        f"{' in ' + location if location else ''}"
        f"{' for the ' + site_area + ' site' if site_area else ''}.\n\n"
        f"The proposed team brings together {discipline_text} under coordinated project leadership.\n\n"
        "We commit senior oversight, rigorous coordination, and practical execution to meet your goals."
    )

    project_understanding = (
        parsed.get("scope_summary")
        or f"We understand that {project_name} requires a structured and outcome-focused approach."
    )

    methodology_steps = parsed.get("phases") or ["Kick-off", "Analysis", "Delivery"]
    methodology = "\n".join([f"- {step}" for step in methodology_steps])

    deliverables = parsed.get("deliverables") or []
    scope_deliverables = "\n".join([f"- {item}" for item in deliverables])
    if not scope_deliverables:
        scope_deliverables = "- Scope items to be finalized during review."

    schedule = _build_schedule_section(parsed=parsed, fee=fee)
    team = _build_team_section(parsed=parsed, fee_input=fee_input)
    financial = _build_financial_section(fee=fee, fee_input=fee_input)

    relevant_lines = []
    for item in relevant:
        relevant_lines.append(
            f"- {item.get('name')} ({item.get('location')}): {item.get('summary')}"
        )
    relevant_experience = "\n".join(relevant_lines) if relevant_lines else "- No relevant projects selected yet."

    assumptions_exclusions = (
        "- Fees exclude taxes and levies unless stated otherwise.\n"
        "- Travel is included only when explicitly listed in the financial breakdown.\n"
        "- Optional services (animation/model) are excluded from lump sum.\n"
        "- Specialty consultants not listed in included costs are excluded."
    )

    return {
        "cover_letter": cover_letter,
        "project_understanding": project_understanding,
        "methodology": methodology,
        "scope_deliverables": scope_deliverables,
        "schedule": schedule,
        "team": team,
        "financial": financial,
        "relevant_experience": relevant_experience,
        "assumptions_exclusions": assumptions_exclusions,
    }


def _to_markdown(parsed: dict, sections: dict) -> str:
    project_name = parsed.get("project", {}).get("project_name", "Untitled Project")
    ordered = [
        ("Cover Letter", "cover_letter"),
        ("Project Understanding", "project_understanding"),
        ("Methodology", "methodology"),
        ("Scope and Deliverables", "scope_deliverables"),
        ("Schedule", "schedule"),
        ("Team Structure", "team"),
        ("Financial Proposal", "financial"),
        ("Relevant Experience", "relevant_experience"),
        ("Assumptions and Exclusions", "assumptions_exclusions"),
    ]

    parts = [f"# Fee Proposal: {project_name}"]
    for title, key in ordered:
        parts.append(f"\n## {title}\n{sections.get(key, '')}")
    return "\n".join(parts)


def build_proposal_payload(
    parsed: dict,
    fee: dict,
    relevant: list[dict],
    fee_input: FeeInput | None = None,
) -> dict:
    sections = _section_text(parsed=parsed, fee=fee, relevant=relevant, fee_input=fee_input)

    if settings.enable_openai_synthesis:
        openai = OpenAIService()
        if openai.enabled:
            try:
                system_prompt = "You improve proposal language while preserving all facts. Return markdown text only."
                user_prompt = (
                    "Improve this cover letter for professional proposal tone without adding new facts:\n\n"
                    + sections["cover_letter"]
                )
                sections["cover_letter"] = openai.text_completion(system_prompt, user_prompt)
            except Exception:
                pass

    markdown = _to_markdown(parsed=parsed, sections=sections)

    return {
        "project": parsed.get("project", {}),
        "parsed": parsed,
        "sections": sections,
        "financial": fee,
        "relevant_experience": relevant,
        "markdown": markdown,
    }


def update_proposal_sections(payload: dict, section_overrides: dict[str, str]) -> dict:
    sections = dict(payload.get("sections", {}))
    for key, value in section_overrides.items():
        if key in sections:
            sections[key] = value

    updated_payload = dict(payload)
    updated_payload["sections"] = sections
    updated_payload["markdown"] = _to_markdown(parsed=updated_payload.get("parsed", {}), sections=sections)
    return updated_payload
