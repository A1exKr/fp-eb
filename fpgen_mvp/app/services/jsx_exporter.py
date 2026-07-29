"""JSX bundle exporter — builds a self-contained ZIP containing an InDesign ExtendScript
assembly script + all required INDD asset files.

The ZIP is fully self-contained: the JSX locates assets relative to its own path using
File($.fileName).parent, so the user just extracts anywhere and double-clicks the JSX
in InDesign's Scripts panel.

Missing assets insert a styled placeholder page rather than failing.

Architecture note
-----------------
Asset I/O is isolated behind the ``AssetStore`` abstract interface so that a future
``NotionAssetStore`` implementation can slot in without touching this module.
"""

from __future__ import annotations

import io
import json
import re
import zipfile
from abc import ABC, abstractmethod
from datetime import datetime
from pathlib import Path
from typing import Optional

from app.config import settings
from app.services.design_profile import get_profile
from app.services.jsx_modea import render_modea_jsx


# ---------------------------------------------------------------------------
# AssetStore interface
# ---------------------------------------------------------------------------

class AssetStore(ABC):
    @abstractmethod
    def get_file(self, category: str, asset_id: str) -> Optional[bytes]:
        """Return file bytes or None if not found."""


class LocalFileAssetStore(AssetStore):
    def __init__(self, assets_dir: Path) -> None:
        self._dir = assets_dir

    def get_file(self, category: str, asset_id: str) -> Optional[bytes]:
        path = self._dir / category / f"{asset_id}.indd"
        if path.is_file():
            return path.read_bytes()
        return None


# ---------------------------------------------------------------------------
# JSX template
# ---------------------------------------------------------------------------

_JSX_TEMPLATE = r"""// FP-GEN — InDesign Assembly Script
// Generated: {generated_date}
// Run from InDesign: Scripts panel → double-click this file
// The .indd output is saved next to this script.

(function () {
    var ROOT = File($.fileName).parent;

    // -----------------------------------------------------------------------
    // Proposal data
    // -----------------------------------------------------------------------
    var PROJECT_NAME = {project_name_js};
    var CLIENT_NAME  = {client_name_js};
    var LOCATION     = {location_js};
    var TODAY        = {today_js};
    var FEE_SNAPSHOT = {fee_snapshot_js};

    // -----------------------------------------------------------------------
    // Asset manifest
    // -----------------------------------------------------------------------
    var TEMPLATE_FILE = new File(ROOT + "/assets/template/{template_filename}");

    // CVs: [{role, file, label}, ...]  — file is null when not bundled
    var CV_MANIFEST = {cv_manifest_js};

    // Experience: [{id, file, label}, ...]
    var EXP_MANIFEST = {exp_manifest_js};

    // -----------------------------------------------------------------------
    // Proposal sections (in order)
    // -----------------------------------------------------------------------
    var SECTIONS = {sections_js};

    // -----------------------------------------------------------------------
    // Helpers
    // -----------------------------------------------------------------------

    function safeContents(frame, text) {
        try { frame.contents = text; } catch (e) {}
    }

    function addPage(doc) {
        doc.pages.add();
        return doc.pages.item(doc.pages.length - 1);
    }

    function addTextFrame(page, bounds, text) {
        var frame = page.textFrames.add();
        frame.geometricBounds = bounds;
        safeContents(frame, text);
        return frame;
    }

    var FULL_BOUNDS  = ["15mm", "15mm", "282mm", "195mm"];
    var TITLE_BOUNDS = ["15mm", "15mm",  "35mm", "195mm"];
    var BODY_BOUNDS  = ["38mm", "15mm", "282mm", "195mm"];

    function appendSectionPage(doc, title, body) {
        var page = addPage(doc);
        addTextFrame(page, TITLE_BOUNDS, PROJECT_NAME + "\r" + title);
        addTextFrame(page, BODY_BOUNDS,  body);
    }

    function appendPlaceholderPage(doc, label) {
        var page = addPage(doc);
        addTextFrame(page, FULL_BOUNDS,
            "[ PLACEHOLDER ]\r\r" +
            "Attach document for:\r" + label + "\r\r" +
            "Replace this page with the appropriate INDD content.");
    }

    function copyAllPages(srcFile, targetDoc, label) {
        if (!srcFile || !srcFile.exists) {
            appendPlaceholderPage(targetDoc, label);
            return;
        }
        try {
            app.scriptPreferences.userInteractionLevel =
                UserInteractionLevels.NEVER_INTERACT;
            var srcDoc = app.open(srcFile, false);
            var pageCount = srcDoc.pages.length;
            for (var i = 0; i < pageCount; i++) {
                var srcPage = srcDoc.pages.item(i);
                srcPage.duplicate(LocationOptions.AT_END, targetDoc);
            }
            srcDoc.close(SaveOptions.NO);
        } catch (e) {
            appendPlaceholderPage(targetDoc, label + " (open failed: " + e.message + ")");
        }
    }

    // -----------------------------------------------------------------------
    // Main assembly
    // -----------------------------------------------------------------------

    app.scriptPreferences.userInteractionLevel =
        UserInteractionLevels.NEVER_INTERACT;

    var doc;
    if (TEMPLATE_FILE.exists) {
        doc = app.open(TEMPLATE_FILE);
    } else {
        // Fallback: create a blank A4 document
        doc = app.documents.add();
        doc.documentPreferences.facingPages = false;
        doc.documentPreferences.pageWidth  = "210mm";
        doc.documentPreferences.pageHeight = "297mm";
    }

    // Fill template placeholders on existing pages
    var replacements = [
        ["PROJECT NAME PROJECT NAME PROJECT NAME", PROJECT_NAME],
        ["PROJECT NAME", PROJECT_NAME],
        ["COMPANY NAME",  CLIENT_NAME],
        ["Mr. OOOOOOO",   CLIENT_NAME],
        ["ADDRESS",       LOCATION],
        ["Rev 0 / 2024.04.12", "Rev 0 / " + TODAY]
    ];

    for (var p = 1; p <= doc.pages.length; p++) {
        var pg = doc.pages.item(p - 1);
        for (var f = 0; f < pg.textFrames.length; f++) {
            var tf = pg.textFrames.item(f);
            try {
                var orig = tf.contents;
                var updated = orig;
                for (var r = 0; r < replacements.length; r++) {
                    // Replace all occurrences
                    while (updated.indexOf(replacements[r][0]) !== -1) {
                        updated = updated.replace(replacements[r][0], replacements[r][1]);
                    }
                }
                if (updated !== orig) { tf.contents = updated; }
            } catch (e) {}
        }
    }

    // Append fee snapshot page
    appendSectionPage(doc, "Commercial Snapshot", FEE_SNAPSHOT);

    // Append narrative proposal sections
    for (var s = 0; s < SECTIONS.length; s++) {
        appendSectionPage(doc, SECTIONS[s][0], SECTIONS[s][1]);
    }

    // Append CV pages after team section
    for (var c = 0; c < CV_MANIFEST.length; c++) {
        var cvEntry = CV_MANIFEST[c];
        var cvFile  = cvEntry.file ? new File(ROOT + "/assets/cvs/" + cvEntry.file) : null;
        copyAllPages(cvFile, doc, "CV: " + cvEntry.label);
    }

    // Append experience pages
    for (var x = 0; x < EXP_MANIFEST.length; x++) {
        var expEntry = EXP_MANIFEST[x];
        var expFile  = expEntry.file ? new File(ROOT + "/assets/experience/" + expEntry.file) : null;
        copyAllPages(expFile, doc, "Experience: " + expEntry.label);
    }

    // Save output next to this script
    var safeName = PROJECT_NAME.replace(/[^A-Za-z0-9.\- ]/g, "_").replace(/\s+/g, "_");
    var outFile  = new File(ROOT + "/" + safeName + ".indd");
    doc.save(outFile);
    doc.close();

    app.scriptPreferences.userInteractionLevel =
        UserInteractionLevels.INTERACT_WITH_ALERTS;

    alert("FP-GEN assembly complete.\n\nSaved: " + outFile.fsName);
}());
"""

_README_TEMPLATE = """\
FP-GEN — InDesign Assembly Bundle
===================================
Generated: {generated_date}
Project:   {project_name}
Client:    {client_name}

INSTRUCTIONS
------------
1. Extract this ZIP to any folder on your Windows machine.
2. Open Adobe InDesign (CC 2019 or later recommended).
3. Open the Scripts panel:  Window → Utilities → Scripts
4. Click "User" scripts folder or navigate to where you extracted this ZIP.
5. Double-click  assemble_proposal.jsx
6. InDesign will assemble the document and save  {safe_name}.indd  in the same folder.

ASSET STATUS
------------
{asset_status}

NOTES
-----
- Missing assets listed above are replaced by placeholder pages in the output.
- The Nikken template (assets/template/{template_filename}) must be present to apply
  corporate styling; if missing, a blank A4 document is used as fallback.
- Do not rename or move the  assets/  subfolder relative to the .jsx file.
"""


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _js_string(value: str) -> str:
    """Encode a Python string as a safe JS double-quoted string."""
    escaped = value.replace("\\", "\\\\").replace('"', '\\"').replace("\n", "\\n").replace("\r", "\\r")
    return f'"{escaped}"'


def _js_array_of_pairs(pairs: list[tuple[str, str]]) -> str:
    items = ", ".join(f'[{_js_string(t)}, {_js_string(b)}]' for t, b in pairs)
    return f"[{items}]"


def _safe_stem(value: str) -> str:
    stem = re.sub(r"[^A-Za-z0-9.\- ]+", "_", value).strip()
    return stem[:60] or "proposal"


def _fee_snapshot(proposal: dict) -> str:
    financial = proposal.get("financial", {})
    currency = financial.get("currency", "USD")
    lines = [
        f"Client:    {proposal.get('client', {}).get('name', '')}",
        f"Project:   {proposal.get('project', {}).get('name', '')}",
        f"Location:  {proposal.get('project', {}).get('location', '')}",
        "",
        f"Labor:          {financial.get('labor_total', 0):,.2f} {currency}",
        f"Overhead:       {financial.get('overhead_total', 0):,.2f} {currency}",
        f"Reimbursables:  {financial.get('reimbursables_total', 0):,.2f} {currency}",
        f"LUMP SUM FEE:   {financial.get('lump_sum_total', 0):,.2f} {currency}",
    ]
    return "\n".join(lines)


# ---------------------------------------------------------------------------
# Mode A data shaping (parse sections -> styled blocks, compute fee tables)
# ---------------------------------------------------------------------------

def _fmt(value) -> str:
    try:
        return f"{float(value):,.2f}"
    except (TypeError, ValueError):
        return "0.00"


def _clean(text: str) -> str:
    return (text or "").replace("**", "").replace("__", "").replace("`", "").strip()


def _parse_blocks(md: str) -> list[dict]:
    """Split a markdown/plain section string into styled paragraph blocks."""
    blocks: list[dict] = []
    for raw in (md or "").splitlines():
        stripped = raw.strip()
        if not stripped:
            continue
        if stripped.startswith("###"):
            blocks.append({"text": _clean(stripped.lstrip("#")), "style": "Caption"})
        elif stripped.startswith("#"):
            blocks.append({"text": _clean(stripped.lstrip("#")), "style": "SectionHeading"})
        elif re.match(r"^[-*+]\s+", stripped):
            blocks.append({"text": _clean(re.sub(r"^[-*+]\s+", "", stripped)), "style": "ListItem"})
        elif re.match(r"^\d+[.)]\s+", stripped):
            blocks.append({"text": _clean(stripped), "style": "ListItem"})
        else:
            blocks.append({"text": _clean(stripped), "style": "Body"})
    return blocks


def _fee_tables(financial: dict) -> dict:
    """Pre-compute formatted fee tables from the ``financial`` block."""
    currency = financial.get("currency", "USD")
    amount_col = f"Amount ({currency})"
    overhead_pct = financial.get("overhead_pct") or 0

    summary_rows = [
        ["Professional Fees (Labour)", _fmt(financial.get("labor_total", 0))],
        [f"Overhead ({round(overhead_pct * 100)}%)", _fmt(financial.get("overhead_total", 0))],
    ]
    if financial.get("subconsultants_included_total"):
        summary_rows.append(["Subconsultants (included)", _fmt(financial["subconsultants_included_total"])])
    summary_rows.append(["Reimbursables", _fmt(financial.get("reimbursables_total", 0))])

    tables: dict = {
        "summary": {
            "title": "Fee Summary",
            "columns": ["Item", amount_col],
            "rows": summary_rows,
            "total": ["LUMP SUM FEE", _fmt(financial.get("lump_sum_total", 0))],
        },
        "breakdown": None,
        "travel": None,
        "payment_schedule": None,
    }

    phase_totals = financial.get("phase_totals") or {}
    phase_sum = sum(phase_totals.values())
    if phase_totals and phase_sum:
        tables["breakdown"] = {
            "title": "Fee Breakdown by Stage",
            "columns": ["Stage", amount_col, "% of Total"],
            "rows": [[phase, _fmt(amt), f"{round(amt / phase_sum * 100)}%"] for phase, amt in phase_totals.items()],
            "total": ["Total", _fmt(phase_sum), "100%"],
        }
        cumulative = 0
        payment_rows = []
        for phase, amt in phase_totals.items():
            pct = round(amt / phase_sum * 100)
            cumulative += pct
            payment_rows.append([phase, f"{pct}%", _fmt(amt), f"{cumulative}%"])
        tables["payment_schedule"] = {
            "title": "Payment Schedule",
            "columns": ["Milestone", "%", amount_col, "Cumulative"],
            "rows": payment_rows,
            "total": None,
        }

    travel = financial.get("travel") or {}
    if travel.get("trips"):
        tables["travel"] = {
            "title": "Travel Expenses",
            "columns": ["Item", "Trips", "People/Trip", "Unit Cost", amount_col],
            "rows": [[
                "Site visits / meetings",
                str(travel.get("trips", 0)),
                str(travel.get("people_per_trip", 0)),
                _fmt(travel.get("unit_cost", 0)),
                _fmt(travel.get("total", 0)),
            ]],
            "total": ["Travel total", "", "", "", _fmt(travel.get("total", 0))],
        }
    return tables


def _experience_items(proposal: dict) -> list[dict]:
    items: list[dict] = []
    for rp in proposal.get("relevant_experience", []) or []:
        meta = " \u00b7 ".join([m for m in (rp.get("project_type", ""), rp.get("location", "")) if m])
        items.append({
            "name": rp.get("name", "Project"),
            "meta": meta,
            "summary_blocks": _parse_blocks(rp.get("summary", "")),
        })
    return items


def _signatory(proposal: dict) -> dict:
    principal = ((proposal.get("parsed", {}) or {}).get("team", {}) or {}).get("principal", {}) or {}
    return {
        "name": principal.get("name") or "Wataru Tanaka",
        "title": principal.get("title") or "Senior Executive Officer",
    }


# ---------------------------------------------------------------------------
# Public entry point
# ---------------------------------------------------------------------------

def build_jsx_bundle(
    proposal: dict,
    cv_assignments: dict[str, str],
    experience_ids: list[str],
    template_id: str,
    assets_dir: Path,
) -> bytes:
    """Build a self-contained ZIP bundle and return raw bytes."""

    store = LocalFileAssetStore(assets_dir)
    today = datetime.now().strftime("%Y.%m.%d")
    generated_date = datetime.now().strftime("%Y-%m-%d %H:%M")

    project = proposal.get("project", {})
    project_name = project.get("name") or project.get("project_name") or "Proposal"
    client_name = proposal.get("client", {}).get("name") or project.get("client_name") or "Client"
    safe_name = _safe_stem(project_name)
    financial = proposal.get("financial", {})
    currency = financial.get("currency", "USD")
    sections = proposal.get("sections", {})

    # CV / experience .indd pages appended whole (pre-designed Mode B assets).
    cv_entries: list[dict] = []
    for role, asset_id in cv_assignments.items():
        data_bytes = store.get_file("cvs", asset_id) if asset_id else None
        cv_entries.append({
            "asset_id": asset_id,
            "file": f"assets/cvs/{asset_id}.indd" if data_bytes else None,
            "label": f"CV: {role}",
            "bytes": data_bytes,
        })
    exp_entries: list[dict] = []
    for exp_id in experience_ids:
        data_bytes = store.get_file("experience", exp_id)
        exp_entries.append({
            "id": exp_id,
            "file": f"assets/experience/{exp_id}.indd" if data_bytes else None,
            "label": f"Experience: {exp_id}",
            "bytes": data_bytes,
        })
    append_pages = [{"file": e["file"], "label": e["label"]} for e in (cv_entries + exp_entries)]

    # Design profile + bundled brand assets (disposition = bundle-with-jsx).
    profile = get_profile(template_id)
    brand_dir = settings.brand_assets_dir
    brand_map: dict[str, str] = {}
    brand_files: dict[str, bytes] = {}
    for role, filename in (profile.get("brand_assets") or {}).items():
        candidate = (brand_dir / filename) if brand_dir else None
        if candidate and candidate.is_file():
            rel = f"assets/brand/{filename}"
            brand_map[role] = rel
            brand_files[rel] = candidate.read_bytes()

    signatory = _signatory(proposal)
    letter_blocks = (
        [{"text": today, "style": "Body"},
         {"text": f"Re: Commercial Proposal \u2014 {project_name}", "style": "Caption"}]
        + _parse_blocks(sections.get("cover_letter", ""))
        + [{"text": signatory["name"], "style": "Caption"},
           {"text": signatory["title"], "style": "Body"}]
    )

    data = {
        "generated_date": generated_date,
        "today": today,
        "output_name": safe_name,
        "currency": currency,
        "project": {
            "name": project_name,
            "type": project.get("type", ""),
            "location": project.get("location", ""),
            "siteArea": project.get("siteArea", ""),
        },
        "client": {"name": client_name},
        "signatory": signatory,
        "assets": {
            "logo": brand_map.get("logo"),
            "signature": brand_map.get("signature"),
            "divider": brand_map.get("divider"),
            "client_logo": None,
        },
        "letter_blocks": letter_blocks,
        "section_blocks": {
            key: _parse_blocks(sections.get(key, ""))
            for key in ("project_understanding", "methodology", "scope_deliverables",
                        "schedule", "team", "financial", "assumptions_exclusions")
        },
        "fee_tables": _fee_tables(financial),
        "experience": _experience_items(proposal),
        "append_pages": append_pages,
    }

    jsx_content = render_modea_jsx(profile, data)

    # README asset status
    asset_lines: list[str] = [f"  [OK]     {rel}" for rel in sorted(brand_files)]
    for e in cv_entries + exp_entries:
        asset_lines.append(f"  {'[OK]    ' if e['bytes'] else '[MISSING]'} {e.get('file') or e['label']}")
    readme_content = _README_TEMPLATE.format(
        generated_date=generated_date,
        project_name=project_name,
        client_name=client_name,
        safe_name=safe_name,
        template_filename=f"{template_id}.indd",
        asset_status="\n".join(asset_lines) if asset_lines else "  (no assets bundled)",
    )

    # Package ZIP
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, mode="w", compression=zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("assemble_proposal.jsx", jsx_content.encode("utf-8"))
        zf.writestr("proposal_data.json", json.dumps(proposal, indent=2, default=str).encode("utf-8"))
        zf.writestr("README.txt", readme_content.encode("utf-8"))
        for rel, data_bytes in brand_files.items():
            zf.writestr(rel, data_bytes)
        for e in cv_entries:
            if e["bytes"]:
                zf.writestr(f"assets/cvs/{e['asset_id']}.indd", e["bytes"])
        for e in exp_entries:
            if e["bytes"]:
                zf.writestr(f"assets/experience/{e['id']}.indd", e["bytes"])

    return buf.getvalue()
