"""Mode A InDesign proposal generator (build-from-scratch, config-driven).

``render_modea_jsx(profile, data)`` returns a complete InDesign ExtendScript
(``.jsx``) string that, when run in InDesign, **creates a new document** and
builds a branded fee proposal from:

* ``profile`` — a design profile (see :mod:`app.services.design_profile`): page
  setup, swatches, fonts, paragraph styles, table style, brand assets, and the
  ordered section plan.
* ``data`` — the proposal content already shaped for layout by the caller
  (parsed narrative *blocks*, pre-computed *fee tables*, resolved *asset* paths).

Design notes
------------
* **No ``str.format``.** ``profile`` and ``data`` are injected as JSON object
  literals (``var CONFIG = {...}; var DATA = {...};``) and the ExtendScript body
  is concatenated verbatim — the literal ``{`` / ``$`` in ExtendScript never
  collide with Python formatting.
* **Roman text frames + named paragraph styles only** — never CJK frame grids.
  This is the fix for the greeked / letter-spaced "green grid" output.
* Missing fonts/images degrade to styled placeholders; the script never hard-fails.
"""
from __future__ import annotations

import json


# --------------------------------------------------------------------------- #
# ExtendScript body (pure JS; consumes the injected CONFIG + DATA globals)
# --------------------------------------------------------------------------- #
_JSX_BODY = r"""
var DOC, ROOT, TA;

main();

function main() {
    ROOT = File($.fileName).parent;
    // Optional override injected by the deploy script so a copy living in InDesign's
    // Scripts Panel can still resolve the bundle's assets/ from another folder.
    try { if (typeof ASSET_ROOT !== "undefined" && ASSET_ROOT) ROOT = new Folder(String(ASSET_ROOT)); } catch (e) {}
    app.scriptPreferences.userInteractionLevel = UserInteractionLevels.NEVER_INTERACT;

    DOC = app.documents.add();
    setupDoc();
    setupSwatches();
    setupStyles();

    var plan = CONFIG.sections;
    for (var i = 0; i < plan.length; i++) {
        var sec = plan[i];
        try {
            if (sec.type === "cover")            buildCover(sec);
            else if (sec.type === "letter")      buildLetter(sec);
            else if (sec.type === "fee")         buildFee(sec);
            else if (sec.type === "experience")  buildExperience(sec);
            else                                  buildNarrative(sec);   // narrative / tandc
        } catch (e) {
            try { buildNarrative({ title: (sec.title || "Section"),
                                   _blocks: [{ text: "[Section could not be generated: " + e.message + "]",
                                               style: "Body" }] }); } catch (e2) {}
        }
    }
    try { appendPages(); } catch (e) {}

    saveDoc();
    app.scriptPreferences.userInteractionLevel = UserInteractionLevels.INTERACT_WITH_ALERTS;
    try { alert("FP-GEN proposal assembled.\n\nSaved: " + outFile().fsName); } catch (e) {}
}

/* ----------------------------------------------------------------------- */
/* Document / swatches / styles                                            */
/* ----------------------------------------------------------------------- */

function setupDoc() {
    DOC.viewPreferences.horizontalMeasurementUnits = MeasurementUnits.MILLIMETERS;
    DOC.viewPreferences.verticalMeasurementUnits   = MeasurementUnits.MILLIMETERS;

    var p = CONFIG.page;
    DOC.documentPreferences.pageWidth  = p.width_mm;
    DOC.documentPreferences.pageHeight = p.height_mm;
    DOC.documentPreferences.facingPages = !!p.facing;
    try {
        DOC.documentPreferences.documentBleedTopOffset = p.bleed_mm;
        DOC.documentPreferences.documentBleedBottomOffset = p.bleed_mm;
        DOC.documentPreferences.documentBleedInsideOrLeftOffset = p.bleed_mm;
        DOC.documentPreferences.documentBleedOutsideOrRightOffset = p.bleed_mm;
    } catch (e) {}

    var m = p.margin;
    try {
        DOC.marginPreferences.top = m.top;
        DOC.marginPreferences.bottom = m.bottom;
        DOC.marginPreferences.left = m.inner;
        DOC.marginPreferences.right = m.outer;
    } catch (e) {}

    // Content text area [top, left, bottom, right] in mm.
    TA = [m.top, m.inner, p.height_mm - m.bottom, p.width_mm - m.outer];

    // Never show the composition/layout grid (greeked-text prevention is via
    // roman frames + styles; this just keeps the artboard clean).
    try { DOC.gridPreferences.showBaselineGrid = false; } catch (e) {}
    try { app.activeWindow.showFrameEdges = false; } catch (e) {}

    // Remove the auto-created first page's default; we add pages explicitly.
}

function mkColor(name, cmyk) {
    var c = DOC.colors.itemByName(name);
    if (!c.isValid) {
        c = DOC.colors.add();
        c.name = name;
        c.model = ColorModel.PROCESS;
        c.space = ColorSpace.CMYK;
        c.colorValue = cmyk;
    }
    return c;
}

function setupSwatches() {
    for (var key in CONFIG.colors) {
        if (CONFIG.colors.hasOwnProperty(key)) {
            mkColor("SW." + key, CONFIG.colors[key].cmyk);
        }
    }
}

function colorOf(name) {
    var c = DOC.colors.itemByName("SW." + name);
    return c.isValid ? c : DOC.colors.itemByName("Black");
}

function just(a) {
    a = (a || "left").toLowerCase();
    if (a === "center")  return Justification.CENTER_ALIGN;
    if (a === "right")   return Justification.RIGHT_ALIGN;
    if (a === "justify") return Justification.LEFT_JUSTIFIED;
    return Justification.LEFT_ALIGN;
}

function setupStyles() {
    var specs = CONFIG.paragraph_styles;
    for (var name in specs) {
        if (specs.hasOwnProperty(name)) ensurePS(name, specs[name]);
    }
}

function ensurePS(name, spec) {
    var ps = DOC.paragraphStyles.itemByName(name);
    if (!ps.isValid) ps = DOC.paragraphStyles.add({ name: name });
    try { ps.appliedFont = CONFIG.font_family; } catch (e) {}
    try { ps.fontStyle = spec.font_style || "Regular"; } catch (e) {}
    try { ps.pointSize = spec.size; } catch (e) {}
    try { if (spec.leading != null) ps.leading = spec.leading; } catch (e) {}
    try { ps.justification = just(spec.align); } catch (e) {}
    try { if (spec.space_before != null) ps.spaceBefore = spec.space_before; } catch (e) {}
    try { if (spec.space_after != null)  ps.spaceAfter  = spec.space_after; } catch (e) {}
    try { if (spec.left_indent != null)  ps.leftIndent  = spec.left_indent; } catch (e) {}
    try { if (spec.first_line_indent != null) ps.firstLineIndent = spec.first_line_indent; } catch (e) {}
    try { if (spec.color) ps.fillColor = colorOf(spec.color); } catch (e) {}
    try { if (spec.all_caps) ps.capitalization = Capitalization.ALL_CAPS; } catch (e) {}
    return ps;
}

function PS(name) {
    var ps = DOC.paragraphStyles.itemByName(name);
    return ps.isValid ? ps : DOC.paragraphStyles.itemByName("[Basic Paragraph]");
}

/* ----------------------------------------------------------------------- */
/* Page + frame helpers                                                     */
/* ----------------------------------------------------------------------- */

function addPage() { return DOC.pages.add(); }

function frameAt(page, t, l, b, r) {
    var tf = page.textFrames.add();
    tf.geometricBounds = [t, l, b, r];
    return tf;
}

function addFolio(page) {
    try {
        var f = frameAt(page, TA[2] + 4, TA[3] - 20, TA[2] + 12, TA[3]);
        f.parentStory.contents = "" + DOC.pages.length;
        f.parentStory.paragraphs[0].appliedParagraphStyle = PS("Folio");
    } catch (e) {}
}

// Fill a text frame with an array of {text, style} blocks, threading onto new
// pages while the content overflows.
function renderBlocks(blocks, headingTitle) {
    var page = addPage();
    var frame = frameAt(page, TA[0], TA[1], TA[2], TA[3]);
    var story = frame.parentStory;

    if (headingTitle) pushBlock(story, headingTitle, "SectionHeading");
    for (var i = 0; i < blocks.length; i++) {
        pushBlock(story, blocks[i].text, blocks[i].style || "Body");
    }

    var tail = frame, guard = 0;
    while (tail.overflows && guard < 60) {
        guard++;
        var np = addPage();
        var nf = frameAt(np, TA[0], TA[1], TA[2], TA[3]);
        tail.nextTextFrame = nf;
        addFolio(np);
        tail = nf;
    }
    addFolio(page);
    return page;
}

function pushBlock(story, text, styleName) {
    var ip = story.insertionPoints[-1];
    ip.contents = (story.contents.length ? "\r" : "") + (text == null ? "" : text);
    story.paragraphs[-1].appliedParagraphStyle = PS(styleName);
}

function placeImage(page, t, l, b, r, rel, label, fill) {
    var rect = page.rectangles.add();
    rect.geometricBounds = [t, l, b, r];
    rect.strokeWeight = 0;
    var f = rel ? new File(ROOT + "/" + rel) : null;
    if (f && f.exists) {
        try {
            rect.fillColor = DOC.swatches.itemByName("None");
            rect.place(f);
            rect.fit(fill ? FitOptions.FILL_PROPORTIONALLY : FitOptions.PROPORTIONALLY);
            rect.fit(FitOptions.CENTER_CONTENT);
            return rect;
        } catch (e) {}
    }
    // Placeholder
    try {
        rect.fillColor = colorOf("silver");
        rect.fillTint = 15;
    } catch (e) {}
    try {
        var cap = frameAt(page, (t + b) / 2 - 5, l, (t + b) / 2 + 5, r);
        cap.parentStory.contents = label || "IMAGE";
        cap.parentStory.paragraphs[0].appliedParagraphStyle = PS("Caption");
        cap.parentStory.paragraphs[0].justification = Justification.CENTER_ALIGN;
        cap.textFramePreferences.verticalJustification = VerticalJustification.CENTER_ALIGN;
    } catch (e) {}
    return rect;
}

/* ----------------------------------------------------------------------- */
/* Section builders                                                         */
/* ----------------------------------------------------------------------- */

function buildCover() {
    var page = addPage();
    var p = CONFIG.page, W = p.width_mm;

    // Brand logo (top-left)
    placeImage(page, 15, 15, 32, 70, (DATA.assets && DATA.assets.logo) || null,
               "NIKKEN", false);
    // Client logo placeholder (top-right)
    placeImage(page, 15, W - 60, 35, W - 15, (DATA.assets && DATA.assets.client_logo) || null,
               "CLIENT LOGO", false);

    // Title
    var t = frameAt(page, 95, 20, 150, W - 15);
    var s = t.parentStory;
    pushBlock(s, DATA.project.name, "CoverTitle");
    pushBlock(s, "Commercial Proposal", "CoverSubtitle");
    pushBlock(s, DATA.today, "CoverSubtitle");
}

function buildLetter(sec) {
    var blocks = DATA.letter_blocks || [];
    var page = renderBlocks(blocks, null);
    // Signature image near the bottom of the (first) letter page.
    var sig = (DATA.assets && DATA.assets.signature) || null;
    if (sig) placeImage(page, TA[2] - 40, TA[1], TA[2] - 22, TA[1] + 55, sig, "", false);
}

function buildNarrative(sec) {
    var blocks = sec._blocks
        || (DATA.section_blocks && DATA.section_blocks[sec.source])
        || [];
    if (!blocks.length) blocks = [{ text: "Content to be provided.", style: "Body" }];
    renderBlocks(blocks, sec.title || null);
}

function buildFee(sec) {
    var page = addPage();
    var frame = frameAt(page, TA[0], TA[1], TA[0] + 8, TA[3]);
    pushBlock(frame.parentStory, sec.title || "Financial Proposal", "SectionHeading");

    var y = TA[0] + 14;
    var intro = (DATA.section_blocks && DATA.section_blocks.financial) || [];
    if (intro.length) {
        var infr = frameAt(page, y, TA[1], y + 30, TA[3]);
        for (var k = 0; k < intro.length; k++) pushBlock(infr.parentStory, intro[k].text, intro[k].style || "Body");
        y += 34;
    }

    var order = ["summary", "breakdown", "travel", "payment_schedule"];
    for (var i = 0; i < order.length; i++) {
        var spec = DATA.fee_tables ? DATA.fee_tables[order[i]] : null;
        if (!spec) continue;
        var estH = (1 + spec.rows.length + (spec.total ? 1 : 0)) * 7 + 8;
        if (y + estH > TA[2]) { addFolio(page); page = addPage(); y = TA[0]; }
        if (spec.title) { var hf = frameAt(page, y, TA[1], y + 7, TA[3]);
                          pushBlock(hf.parentStory, spec.title, "TableHeader"); y += 8; }
        buildTable(page, y, spec);
        y += estH;
    }
    addFolio(page);
}

function buildTable(page, top, spec) {
    var nCols = spec.columns.length;
    var rows = [];
    rows.push({ cells: spec.columns, kind: "header" });
    for (var i = 0; i < spec.rows.length; i++) rows.push({ cells: spec.rows[i], kind: "body" });
    if (spec.total) rows.push({ cells: spec.total, kind: "total" });

    var estH = rows.length * 7 + 4;
    var frame = frameAt(page, top, TA[1], top + estH, TA[3]);
    var tbl = frame.parentStory.insertionPoints[-1].tables.add();
    tbl.columnCount = nCols;
    tbl.bodyRowCount = rows.length;

    var st = CONFIG.table;
    for (var r = 0; r < rows.length; r++) {
        var row = rows[r];
        for (var c = 0; c < nCols; c++) {
            var cell = tbl.rows[r].cells[c];
            cell.contents = (row.cells[c] == null ? "" : "" + row.cells[c]);
            var styleName = row.kind === "header" ? "TableHeader"
                          : row.kind === "total"  ? "TableTotal" : "TableCell";
            try { cell.texts[0].appliedParagraphStyle = PS(styleName); } catch (e) {}
            try { cell.texts[0].justification = (c === 0 ? Justification.LEFT_ALIGN : Justification.RIGHT_ALIGN); } catch (e) {}
            try {
                if (row.kind === "header") { cell.fillColor = colorOf(st.header_fill); cell.fillTint = st.header_tint; }
                else if (row.kind === "total") { cell.fillColor = colorOf(st.total_fill); cell.fillTint = st.total_tint; }
            } catch (e) {}
        }
    }
    // Horizontal rules only.
    try {
        tbl.rows.everyItem().topEdgeStrokeWeight = 0;
        tbl.rows.everyItem().bottomEdgeStrokeWeight = st.row_rule_weight;
        tbl.rows.everyItem().bottomEdgeStrokeColor = colorOf(st.row_rule);
        tbl.columns.everyItem().leftEdgeStrokeWeight = 0;
        tbl.columns.everyItem().rightEdgeStrokeWeight = 0;
    } catch (e) {}
    return top + estH;
}

function buildExperience(sec) {
    var items = DATA.experience || [];
    var blocks = [];
    if (!items.length) blocks.push({ text: "Relevant experience available on request.", style: "Body" });
    for (var i = 0; i < items.length; i++) {
        var it = items[i];
        blocks.push({ text: it.name, style: "Caption" });
        if (it.meta) blocks.push({ text: it.meta, style: "Body" });
        var sb = it.summary_blocks || [];
        for (var j = 0; j < sb.length; j++) blocks.push(sb[j]);
    }
    var page = renderBlocks(blocks, sec.title || "Relevant Experience");
    var div = (DATA.assets && DATA.assets.divider) || null;
    if (div) {
        // decorative divider band at the top of the section's first page
        try { placeImage(page, TA[0], TA[1], TA[0] + 45, TA[3], div, "", true); } catch (e) {}
    }
}

/* Append pre-designed CV / experience .indd pages whole (Mode B assets). */
function appendPages() {
    var list = DATA.append_pages || [];
    for (var i = 0; i < list.length; i++) {
        var entry = list[i];
        var f = entry.file ? new File(ROOT + "/" + entry.file) : null;
        if (f && f.exists) {
            try {
                var src = app.open(f, false, OpenOptions.OPEN_COPY);
                var n = src.pages.length;
                for (var p = 0; p < n; p++) src.pages.item(p).duplicate(LocationOptions.AT_END, DOC);
                src.close(SaveOptions.NO);
                continue;
            } catch (e) {}
        }
        var pg = addPage();
        var fr = frameAt(pg, TA[0], TA[1], TA[2], TA[3]);
        pushBlock(fr.parentStory, "[ Attach: " + (entry.label || "asset") + " ]", "Caption");
    }
}

/* ----------------------------------------------------------------------- */
/* Save                                                                     */
/* ----------------------------------------------------------------------- */

function outFile() {
    var base = (typeof OUT_NAME !== "undefined" && OUT_NAME) ? OUT_NAME : (DATA.output_name || "proposal");
    var name = String(base).replace(/[^A-Za-z0-9._\- ]/g, "_");
    return new File(ROOT + "/" + name + ".indd");
}

function saveDoc() {
    var out = outFile();
    try {
        for (var i = 0; i < app.documents.length; i++) {
            var d = app.documents.item(i);
            if (d !== DOC && d.saved && d.fullName && String(d.fullName) === String(out)) { d.close(SaveOptions.NO); break; }
        }
    } catch (e) {}
    DOC.save(out);
}
"""


def render_modea_jsx(profile: dict, data: dict) -> str:
    """Return a complete Mode A ExtendScript (.jsx) as a string."""
    config_json = json.dumps(profile, ensure_ascii=True)
    data_json = json.dumps(data, ensure_ascii=True, default=str)
    header = (
        "#target indesign\r\n"
        "// FP-GEN Mode A proposal generator (config-driven). Generated: "
        + str(data.get("generated_date", "")) + "\r\n"
        "// Run from InDesign: Scripts panel -> double-click. Output .indd is saved next to this script.\r\n"
    )
    return header + "var CONFIG = " + config_json + ";\r\n" + "var DATA = " + data_json + ";\r\n" + _JSX_BODY
