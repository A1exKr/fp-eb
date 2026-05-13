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
    location = project.get("location", "")
    safe_name = _safe_stem(project_name)

    # Template
    template_filename = f"{template_id}.indd"
    template_bytes = store.get_file("template", template_id)

    # Sections
    sections_dict = proposal.get("sections", {})
    section_order = [
        ("Cover Letter",           "cover_letter"),
        ("Project Understanding",  "project_understanding"),
        ("Methodology",            "methodology"),
        ("Scope and Deliverables", "scope_deliverables"),
        ("Schedule",               "schedule"),
        ("Team Structure",         "team"),
        ("Financial Proposal",     "financial"),
        ("Relevant Experience",    "relevant_experience"),
        ("Assumptions and Exclusions", "assumptions_exclusions"),
    ]
    sections_pairs = [(title, sections_dict.get(key, "")) for title, key in section_order]

    # CV manifest
    cv_manifest_entries: list[dict] = []
    for role, asset_id in cv_assignments.items():
        cv_bytes = store.get_file("cvs", asset_id) if asset_id else None
        cv_manifest_entries.append({
            "role": role,
            "asset_id": asset_id,
            "file": f"{asset_id}.indd" if cv_bytes else None,
            "label": role,
            "bytes": cv_bytes,
        })

    # Experience manifest
    exp_manifest_entries: list[dict] = []
    for exp_id in experience_ids:
        exp_bytes = store.get_file("experience", exp_id)
        exp_manifest_entries.append({
            "id": exp_id,
            "file": f"{exp_id}.indd" if exp_bytes else None,
            "label": exp_id,
            "bytes": exp_bytes,
        })

    # Build JSX
    cv_manifest_js = (
        "[\n"
        + ",\n".join(
            f'  {{file: {_js_string(e["file"]) if e["file"] else "null"}, label: {_js_string(e["label"])}}}'
            for e in cv_manifest_entries
        )
        + "\n]"
    )
    exp_manifest_js = (
        "[\n"
        + ",\n".join(
            f'  {{file: {_js_string(e["file"]) if e["file"] else "null"}, label: {_js_string(e["label"])}}}'
            for e in exp_manifest_entries
        )
        + "\n]"
    )

    jsx_content = _JSX_TEMPLATE.format(
        generated_date=generated_date,
        project_name_js=_js_string(project_name),
        client_name_js=_js_string(client_name),
        location_js=_js_string(location),
        today_js=_js_string(today),
        fee_snapshot_js=_js_string(_fee_snapshot(proposal)),
        template_filename=template_filename,
        cv_manifest_js=cv_manifest_js,
        exp_manifest_js=exp_manifest_js,
        sections_js=_js_array_of_pairs(sections_pairs),
    )

    # Build README asset status
    asset_lines: list[str] = []
    if template_bytes:
        asset_lines.append(f"  [OK]     template/{template_filename}")
    else:
        asset_lines.append(f"  [MISSING] template/{template_filename}  → fallback blank document")
    for e in cv_manifest_entries:
        status = "[OK]    " if e["bytes"] else "[MISSING]"
        asset_lines.append(f"  {status} cvs/{e['asset_id']}.indd  ({e['label']})")
    for e in exp_manifest_entries:
        status = "[OK]    " if e["bytes"] else "[MISSING]"
        asset_lines.append(f"  {status} experience/{e['id']}.indd")

    readme_content = _README_TEMPLATE.format(
        generated_date=generated_date,
        project_name=project_name,
        client_name=client_name,
        safe_name=safe_name,
        template_filename=template_filename,
        asset_status="\n".join(asset_lines) if asset_lines else "  (no assets selected)",
    )

    # Package ZIP
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, mode="w", compression=zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("assemble_proposal.jsx", jsx_content.encode("utf-8"))
        zf.writestr("proposal_data.json", json.dumps(proposal, indent=2, default=str).encode("utf-8"))
        zf.writestr("README.txt", readme_content.encode("utf-8"))

        if template_bytes:
            zf.writestr(f"assets/template/{template_filename}", template_bytes)

        for e in cv_manifest_entries:
            if e["bytes"]:
                zf.writestr(f"assets/cvs/{e['asset_id']}.indd", e["bytes"])

        for e in exp_manifest_entries:
            if e["bytes"]:
                zf.writestr(f"assets/experience/{e['id']}.indd", e["bytes"])

    return buf.getvalue()
