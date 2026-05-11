from datetime import datetime
from pathlib import Path
import re

from app.config import settings

try:
    import pythoncom
    from win32com.client import Dispatch, DispatchEx
except ImportError:  # pragma: no cover - Windows-only dependency
    pythoncom = None
    Dispatch = None
    DispatchEx = None


class InDesignExportError(RuntimeError):
    pass


NEVER_INTERACT = 1699640946
APPENDIX_BOUNDS = ["15mm", "15mm", "282mm", "405mm"]


def _safe_stem(value: str) -> str:
    stem = re.sub(r"[^A-Za-z0-9._ -]+", "_", value).strip()
    return stem[:80] or "proposal"


def get_indd_export_capability() -> dict[str, str | bool]:
    if not settings.enable_indd_export:
        return {
            "enabled": False,
            "reason": "InDesign export is disabled for this deployment.",
        }

    if Dispatch is None or pythoncom is None:
        return {
            "enabled": False,
            "reason": "InDesign export requires Windows with Adobe InDesign and pywin32.",
        }

    return {
        "enabled": True,
        "reason": "",
    }


def _section_pages(proposal: dict) -> list[tuple[str, str]]:
    sections = proposal.get("sections", {})
    ordered = [
        ("Cover Letter", sections.get("cover_letter", "")),
        ("Project Understanding", sections.get("project_understanding", "")),
        ("Methodology", sections.get("methodology", "")),
        ("Scope and Deliverables", sections.get("scope_deliverables", "")),
        ("Schedule", sections.get("schedule", "")),
        ("Team Structure", sections.get("team", "")),
        ("Financial Proposal", sections.get("financial", "")),
        ("Relevant Experience", sections.get("relevant_experience", "")),
        ("Assumptions and Exclusions", sections.get("assumptions_exclusions", "")),
    ]

    pages: list[tuple[str, str]] = []
    for title, content in ordered:
        body = content.strip() or "To be completed during proposal review."
        if len(body) <= 3200:
            pages.append((title, body))
            continue

        chunks = [body[index:index + 3200] for index in range(0, len(body), 3200)]
        for index, chunk in enumerate(chunks, start=1):
            suffix = "" if index == 1 else f" (cont. {index})"
            pages.append((f"{title}{suffix}", chunk))

    return pages


def _currency_amount(value: float, currency: str) -> str:
    return f"{value:,.2f} {currency}"


def _commercial_summary(proposal: dict) -> str:
    project = proposal.get("project", {})
    parsed = proposal.get("parsed", {})
    financial = proposal.get("financial", {})
    travel = financial.get("travel", {})
    currency = financial.get("currency", settings.default_currency)

    lines = [
        f"Client: {project.get('client_name', '').strip() or 'TBD'}",
        f"Project: {project.get('project_name', '').strip() or 'TBD'}",
        f"Location: {project.get('location', '').strip() or 'TBD'}",
        f"Project Type: {project.get('project_type', '').strip() or 'TBD'}",
    ]

    scope_summary = parsed.get("scope_summary", "").strip()
    if scope_summary:
        lines.extend(["", "Scope Summary:", scope_summary])

    lump_sum_total = float(financial.get("lump_sum_total") or 0)
    reimbursables_total = float(financial.get("reimbursables_total") or 0)
    labor_total = float(financial.get("labor_total") or 0)
    travel_total = float(travel.get("total") or 0)
    lines.extend(
        [
            "",
            "Commercial Snapshot:",
            f"- Labor: {_currency_amount(labor_total, currency)}",
            f"- Reimbursables: {_currency_amount(reimbursables_total, currency)}",
            f"- Travel: {_currency_amount(travel_total, currency)}",
            f"- Lump Sum Fee: {_currency_amount(lump_sum_total, currency)}",
        ]
    )

    return "\r".join(lines)


def _template_pages(proposal: dict) -> list[tuple[str, str]]:
    return [("Commercial Summary", _commercial_summary(proposal)), *_section_pages(proposal)]


def _template_replacements(proposal: dict) -> list[tuple[str, str]]:
    project = proposal.get("project", {})
    project_name = project.get("project_name", "").strip()
    client_name = project.get("client_name", "").strip()
    location = project.get("location", "").strip()
    today = datetime.now().strftime("%Y.%m.%d")

    replacements = [
        ("PROJECT NAME PROJECT NAME PROJECT NAME", project_name),
        ("PROJECT NAME", project_name),
        ("COMPANY NAME", client_name),
        ("Mr. OOOOOOO", client_name),
        ("ADDRESS", location),
        ("Rev 0 / 2024.04.12", f"Rev 0 / {today}"),
    ]
    return [(source, target) for source, target in replacements if source and target]


def _replace_in_template(document, replacements: list[tuple[str, str]]) -> None:
    for page_index in range(1, document.Pages.Count + 1):
        page = document.Pages.Item(page_index)
        for frame_index in range(1, page.TextFrames.Count + 1):
            frame = page.TextFrames.Item(frame_index)
            try:
                contents = str(frame.Contents or "")
            except Exception:
                continue

            updated = contents
            for source, target in replacements:
                updated = updated.replace(source, target)

            if updated != contents:
                frame.Contents = updated


def _append_pages(document, project_name: str, pages: list[tuple[str, str]]) -> None:
    for title, body in pages:
        page = document.Pages.Add()
        frame = page.TextFrames.Add()
        frame.GeometricBounds = APPENDIX_BOUNDS
        frame.Contents = f"{project_name}\r{title}\r\r{body}"


def _export_with_template(app, proposal: dict, out_path: Path, template_path: Path) -> None:
    project_name = proposal.get("project", {}).get("project_name", out_path.stem)
    document = app.Open(str(template_path))
    _replace_in_template(document, _template_replacements(proposal))
    _append_pages(document, project_name, _template_pages(proposal))
    document.Save(str(out_path))
    document.Close()


def _export_simple_document(app, proposal: dict, out_path: Path) -> None:
    project_name = proposal.get("project", {}).get("project_name", out_path.stem)
    document = app.Documents.Add()
    document.DocumentPreferences.FacingPages = False
    document.DocumentPreferences.PageWidth = "210mm"
    document.DocumentPreferences.PageHeight = "297mm"

    pages = _section_pages(proposal)
    for index, (title, body) in enumerate(pages, start=1):
        page = document.Pages.Item(index) if index == 1 else document.Pages.Add()
        frame = page.TextFrames.Add()
        frame.GeometricBounds = ["15mm", "15mm", "282mm", "195mm"]
        frame.Contents = f"{project_name}\r{title}\r\r{body}"

    document.Save(str(out_path))
    document.Close()


def export_proposal_to_indd(proposal_id: str, proposal: dict) -> Path:
    capability = get_indd_export_capability()
    if not capability["enabled"]:
        raise InDesignExportError(str(capability["reason"]))

    project_name = proposal.get("project", {}).get("project_name", proposal_id)
    export_dir = settings.export_dir / proposal_id
    export_dir.mkdir(parents=True, exist_ok=True)
    out_path = export_dir / f"{_safe_stem(project_name)}.indd"
    if out_path.exists():
        out_path.unlink()

    pythoncom.CoInitialize()
    app = None
    previous_interaction_level = None
    try:
        app = DispatchEx("InDesign.Application") if DispatchEx is not None else Dispatch("InDesign.Application")
        previous_interaction_level = app.ScriptPreferences.UserInteractionLevel
        app.ScriptPreferences.UserInteractionLevel = NEVER_INTERACT
        template_path = settings.indd_commercial_template_path
        if template_path.exists():
            _export_with_template(app, proposal, out_path, template_path)
        else:
            _export_simple_document(app, proposal, out_path)
        app.ScriptPreferences.UserInteractionLevel = previous_interaction_level
        app.Quit()
        return out_path
    except Exception as exc:  # pragma: no cover - COM automation depends on local app state
        raise InDesignExportError(f"Failed to export proposal to InDesign: {exc}") from exc
    finally:  # pragma: no cover - COM lifecycle depends on local app state
        if app is not None:
            try:
                if previous_interaction_level is not None:
                    app.ScriptPreferences.UserInteractionLevel = previous_interaction_level
            except Exception:
                pass
        pythoncom.CoUninitialize()