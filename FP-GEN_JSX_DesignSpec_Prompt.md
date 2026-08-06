# FP-GEN — Prompt: Produce a JSX Generation Spec from the Sample Proposal

> **How to use this prompt.** Paste everything below into a vision‑capable LLM and attach the
> **sample fee‑proposal InDesign document** (`.indd`), its exported **PDF**, the **`Links/`**
> folder (all placed images/illustrations), and the **`Document fonts/`** folder. The LLM's job is
> **not** to write code — it is to return a single, exact **Design & Generation Specification** that a
> separate AI coding agent will use to program the generator. Keep it factual and measured.

---

## ROLE

You are a **senior InDesign production designer and ExtendScript (JSX) architect**. You reverse‑engineer
finished InDesign documents into precise, reproducible build specifications.

## YOUR TASK

Analyze the attached **sample fee proposal** and produce a complete **Design & Generation Specification**
(the "Spec"). A downstream AI coding agent will use your Spec — with **no further access to the sample** —
to implement a Python routine (`jsx_exporter.py`) that **emits an InDesign ExtendScript (`.jsx`)**. When a
user runs that `.jsx` in Adobe InDesign (Windows), it must assemble a **branded, print‑ready proposal that
matches the sample's design**, populated automatically from structured JSON data (defined below).

Your Spec must be exact enough that the coding agent never has to make a design decision. Prefer **tables
and concrete numbers** (mm, pt, CMYK/RGB, exact font/style names). Do **not** output `.jsx` code.

---

## SYSTEM CONTEXT (so your Spec fits the pipeline)

- **Pipeline:** RFP text → parsed → a **proposal JSON payload** → `jsx_exporter.py` emits one self‑contained
  `.jsx` **plus** bundles asset `.indd` files into a ZIP → the user runs the `.jsx` from InDesign's Scripts
  panel to produce the final `.indd`. There is **no manual design step** afterward.
- **Runtime:** Adobe InDesign 2019–2025, Windows. The script runs unattended (Scripts panel double‑click or
  COM `DoScript`). Fonts and linked images ship **inside the bundle**; the `.jsx` must reference them by a
  path **relative to itself** (`File($.fileName).parent`).
- **Current implementation is inadequate (must be replaced).** Today the generator opens the commercial
  template and dumps plain text frames over it. The result (see the attached "current output" if provided):
  - text is **greeked/letter‑spaced** — each character sits in a grid cell (e.g. `C l i e n t :`), and a
    **green grid** shows through. This is the classic symptom of **CJK frame‑grid (組みフレーム) composition
    / a visible layout grid**, not roman text frames with paragraph styles.
  - **no real paragraph/character styles**, **no tables** for the fee breakdown, **no images or
    illustrations**, and content **overlaps the template's guide grid**.
  Your Spec must explicitly define how to **avoid** these problems (roman text frames + named styles, real
  tables, proper image frames, grid/guides hidden or non‑printing).

---

## THE DATA MODEL (what drives the layout)

The generator receives this JSON (illustrative shape; field **names** are stable, **values** vary). Your
Spec must map **every** relevant field to an InDesign element + style.

```jsonc
{
  "project":  { "name": "Morgans Point Development", "type": "Commercial / Mixed-use",
                 "location": "Southampton Parish, Bermuda", "siteArea": "..." },
  "client":   { "name": "Morgans Point Development Company Limited (MPDC)" },

  // Narrative sections — each value is a block of plain/markdown text (headings + paragraphs, sometimes
  // bulleted). Order is the proposal order:
  "sections": {
    "cover_letter":            "…",
    "project_understanding":   "…",
    "methodology":             "…",   // ~6 numbered process steps
    "scope_deliverables":      "…",   // scope list + deliverables list
    "schedule":                "…",   // phases + milestones + durations
    "team":                    "…",   // roles / named principals
    "financial":               "…",   // narrative accompanying the fee table
    "relevant_experience":     "…",
    "assumptions_exclusions":  "…"
  },

  // Structured fee figures (drive the FEE TABLE — this is central):
  "financial": {
    "currency": "USD",
    "phase_totals": { "Kick-off": 0.0, "Concept Development": 0.0, "Finalization": 0.0 },
    "labor_total": 464750.00,
    "overhead_pct": 0.10,
    "overhead_total": 46475.00,
    "subconsultants_included_total": 12000.00,
    "subconsultants_excluded": [ { "name": "…", "fee": 0.0 } ],
    "travel": { "trips": 2, "people_per_trip": 2, "unit_cost": 6000, "total": 24000.00,
                "included_in_lump_sum": true },
    "misc_reimbursables": 3000.00,
    "reimbursables_total": 24000.00,
    "lump_sum_total": 535225.00
  },

  // Selected reference projects for the "Relevant Experience" section:
  "relevant_experience": [ { "id": "rp-001", "name": "…", "project_type": "…",
                              "location": "…", "summary": "…", "keywords": ["…"] } ],

  // Asset manifests: pre-designed .indd pages to place (CVs of key staff, project experience sheets),
  // and the chosen template id. Files ship in the bundle under assets/{template,cvs,experience}/.
  "cv_manifest":  [ { "role": "Principal", "file": "cv-principal.indd", "label": "…" } ],
  "exp_manifest": [ { "id": "rp-001", "file": "rp-001.indd", "label": "…" } ],
  "template_id":  "commercial"     // one of: commercial | master-plan | technical
}
```

Also available if the Spec calls for it: a **personnel** table (names/titles/roles), **unit rates**, and the
full **`markdown`** rendering of the proposal.

---

## WHAT TO EXTRACT FROM THE SAMPLE (inspection checklist)

Open the sample `.indd` (and study the PDF) and record, with **exact values**:

1. **Document setup** — page size + orientation; facing pages/spreads vs single; margins (T/B/inner/outer);
   column count + gutter; bleed/slug; intent (Print); transparency/color space; and **whether the document
   uses CJK frame‑grid composition or roman text frames** (decisive for the greeked‑text bug).
2. **Master pages** — how many, their names, and contents (logo, running header/footer, page numbers,
   section markers, background art, guides).
3. **Swatches** — every named color: model (CMYK/RGB/Spot/Pantone), values, and tints used.
4. **Fonts** — exact **family + style** names as shown in InDesign's Font menu, per text role; list all fonts
   in `Document fonts/`.
5. **Paragraph styles** — full catalog with attributes: font, size/leading, alignment, space before/after,
   indents, color, hyphenation, keep options, based‑on/next‑style, and any style **groups** they live in.
6. **Character styles** — for emphasis, figures/units, links, small caps, etc.
7. **Object styles** — for image frames, callout/highlight boxes, sidebars.
8. **Table & cell styles** — column structure, header row, body rows, alternating fills, borders, number/
   currency alignment, totals‑row emphasis. Capture the **fee table** design in full detail.
9. **Per‑page layout** — for every page: which master, all frames (position + size in mm), their style,
   what content type they hold (heading / body / table / image / caption / callout / signature), and
   **image placements** (frame size, fitting option, crop).
10. **Illustrations/images** — cover art, section dividers, project photos, diagrams, maps: where each
    appears, size, fitting, and whether decorative (fixed) or data‑driven (from project/experience assets).
    **Also inventory the actual image files** in the sample's `Links/` (and any embedded images): exact
    filename, format, pixel dimensions, effective DPI, color space, file size, and the page/frame each is used on.
11. **Placeholder patterns** — any literal placeholder strings in the sample (e.g. `PROJECT NAME`,
    `COMPANY NAME`, `ADDRESS`, `Mr. OOOOOOO`, dates, revision tags) that map to data fields.
12. **Section order & pagination** — the exact page/section sequence; where the TOC is (if any) and the
    frame/style that identifies it.

---

## REQUIRED CONTENTS OF YOUR SPEC (deliverable)

Return a single **Markdown** document with these numbered sections. Be exhaustive and exact.

1. **Document Setup** — all values from checklist #1, plus an explicit directive: use roman text frames with
   paragraph styles (grids hidden/non‑printing) unless the sample genuinely requires CJK grids.
2. **Color Catalog** — table: swatch name → model → values → where used.
3. **Font Catalog** — table: text role → family → style → fallback; note licensing/packaging needs.
4. **Paragraph Style Catalog** — table with every attribute the coding agent must set to recreate each style.
5. **Character / Object / Table / Cell Style Catalogs** — same treatment; the fee table gets its own subsection.
6. **Master Pages** — each master: name, purpose, exact contents and coordinates.
7. **Per‑Page Layout Maps** — for each page/section, an ordered list of frames: `{purpose, x, y, w, h (mm),
   paragraph/object style, data source, image fitting}`. Include a cover page, section dividers, the fee page,
   relevant experience, team/CV pages, and assumptions/T&C.
8. **Fee Table Spec** — columns, widths, headers, row model (phases, labor, overhead, subconsultants, travel,
   reimbursables, **LUMP SUM total**), number/currency formatting, and totals styling — mapped to the
   `financial` fields above.
9. **Illustration / Image Spec + Image Asset Inventory** — two parts:
   - **Image slots:** every slot's source (fixed brand asset vs data‑driven from `cv_manifest`/`exp_manifest`/
     project photos), frame + fitting, and the **placeholder** shown when the asset is missing (so output never breaks).
   - **Image Asset Inventory (required):** a table of the **actual image files used in the sample** — exact
     filename, format, px dimensions, effective DPI, color space, file size, the page/frame(s) it appears on,
     and a **disposition** for each: either **`bundle-with-jsx`** (brand/decorative art reused in every
     proposal → shipped in the ZIP next to the `.jsx` and referenced relatively) or **`upload-to-FP-GEN`**
     (project‑specific imagery that varies per proposal → uploaded through FP‑GEN's asset upload and supplied
     per job). State the exact set of files the user must **extract from the sample** — via InDesign
     **File ▸ Package** (collects all links) or directly from `Links/` — so they can be uploaded to FP‑GEN
     and/or downloaded together with the generated `.jsx`.
10. **Data → Element Mapping** — the master table: **every** payload field → target page/frame/style/table cell.
    This is the contract I implement against. Include how to convert each narrative `sections.*` string into
    styled heading + body paragraphs (and bulleted lists where applicable).
11. **Assembly Strategy & Reflow Rules** — recommend **Mode A** (build the document + styles from scratch) or
    **Mode B** (open a clean template with named placeholder frames/styles and populate). Justify the choice
    for repeatable, good‑looking, data‑driven output. Define how variable‑length text reflows (threaded frames,
    auto‑add pages), how to keep the fee table intact on its page, and how bundled CV/experience `.indd` pages
    are appended.
12. **ExtendScript Implementation Notes** — encode the invariants the coding agent must follow (see below),
    and **explicitly** state the fix for the greeked‑text/green‑grid problem observed in the current output.
13. **Open Questions / Assumptions** — anything ambiguous in the sample, with your recommended default.

---

## EXTENDSCRIPT INVARIANTS TO ENCODE IN THE SPEC (apply, don't re‑derive)

- **Composition:** create real **roman `TextFrame`s** with applied **paragraph styles**; do **not** place
  narrative text into **frame grids** (CJK 組みフレーム) — that is what produced the letter‑spaced/greeked text
  and visible green grid. Ensure the document/layout grid is **hidden and non‑printing**.
- **Group‑aware style lookup:** styles may be nested in style **groups**; look up via a helper that also scans
  `paragraphStyleGroups` before falling back to `[Basic Paragraph]`.
- **Open as copy (Mode B):** `app.open(file, true, OpenOptions.OPEN_COPY)` — never rely on the 2nd arg as a copy flag.
- **Paragraph indexing (ID 2022+):** an empty story has **0** paragraphs; use `paragraphs[-1]` (last inserted).
- **Scoped find/change:** use `doc.changeText()` (not `app.`), resetting find/change prefs before and after.
- **Preserve non‑text template items:** when clearing a template page, delete **`TextFrame`s only** — keep
  rectangles, placed images, and PDFs.
- **Fonts & links are relative:** resolve every asset from `File($.fileName).parent`; missing asset ⇒ styled
  **placeholder** frame, never a hard failure.
- **Unattended run:** set `app.scriptPreferences.userInteractionLevel = UserInteractionLevels.NEVER_INTERACT`
  during open/save; include `#target indesign` for Scripts‑panel use (guard it for COM `DoScript`).
- **Tables:** build with real `Table` objects + table/cell styles and right‑aligned currency columns — not
  tab‑spaced text.

---

## OUTPUT FORMAT

- One Markdown document, the numbered sections above, tables preferred, exact numeric values throughout.
- **No `.jsx` code.** Pseudo‑layout and field‑mapping tables are welcome.
- You **cannot emit binary image files** — instead, in the **Image Asset Inventory**, name every file exactly
  and tell the user how to collect them (InDesign **File ▸ Package**, or copy from `Links/`) for upload to
  FP‑GEN and/or bundling with the `.jsx`.
- If a needed value cannot be determined from the sample, list it under **Open Questions** with a recommended
  default rather than omitting it.

---

### Reference (for the human running this prompt)

- Current template/sample assets in the repo live at `fpgen_mvp/examples/INDD/cp/` — `I. Commercial.indd`,
  `I. Commercial.pdf`, `Instructions.txt`, `Links/`, `Document fonts/`. Attach your **preferred** sample if it
  differs from this one, plus its Links and fonts.
- The coding agent will implement the returned Spec in `fpgen_mvp/app/services/jsx_exporter.py` (Python that
  emits the `.jsx`) and its ExtendScript template.
