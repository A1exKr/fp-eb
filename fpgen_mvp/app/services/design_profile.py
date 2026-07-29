"""Design profiles for the InDesign (JSX) proposal generator.

A **design profile** is a plain, JSON-serialisable dict that fully describes the
*look* of a proposal: page setup, colour swatches, fonts, paragraph styles, a
table style, the bundled brand assets, and the ordered list of sections to build.

The generator ([app.services.jsx_modea][]) consumes a profile + the proposal data
and emits an InDesign ExtendScript. Keeping the profile separate from the
generation logic is deliberate:

* **Today** the profile is the hardcoded :data:`DEFAULT_COMMERCIAL_PROFILE`
  below (the Nikken "I. Commercial" spec, with the design‑spec's recommended
  defaults for values that could not be read from the binary ``.indd``).
* **Future** — a "use this FP as a sample" feature can analyse an uploaded
  ``.indd``/PDF + its linked files and emit a profile of the *same shape*, which
  the generator will consume without any code change. That is the whole reason
  the design is data‑driven rather than hardcoded into the ExtendScript.

All numeric values marked in the design spec as "[unverified — recommended
default]" live here so a designer (or the ingestion feature) can tune them in one
place. Colours are CMYK ``[c, m, y, k]`` (0–100).
"""
from __future__ import annotations

from copy import deepcopy


# Attribute keys understood by the ExtendScript style applier (see jsx_modea):
#   font_style, size, leading, align(left|center|right|justify),
#   space_before, space_after, left_indent, first_line_indent,
#   color(name of a swatch), all_caps(bool)
DEFAULT_COMMERCIAL_PROFILE: dict = {
    "id": "commercial",
    "page": {
        "width_mm": 210.0,
        "height_mm": 297.0,
        # v1 builds single pages for geometric robustness; the sample uses reader
        # spreads. Flip to true once facing-page coordinates are tuned in InDesign.
        "facing": False,
        "margin": {"top": 15.0, "bottom": 18.0, "inner": 20.0, "outer": 15.0},
        "bleed_mm": 3.0,
    },
    "colors": {
        # Values are recommended defaults; real swatch numbers are a 2‑minute
        # lookup in InDesign's Swatches panel (design spec §2, §13).
        "ink": {"cmyk": [0, 0, 0, 90]},
        "silver": {"cmyk": [0, 0, 0, 35]},
        "accent": {"cmyk": [0, 0, 0, 100]},  # no verified brand colour → black
    },
    "font_family": "Century Gothic",
    "fallback_font_family": "Arial",
    "paragraph_styles": {
        "CoverTitle": {"font_style": "Bold", "size": 28, "leading": 32, "align": "left",
                        "space_after": 4, "color": "ink", "all_caps": True},
        "CoverSubtitle": {"font_style": "Regular", "size": 15, "leading": 19, "align": "left",
                           "space_after": 0, "color": "ink", "all_caps": True},
        "SectionHeading": {"font_style": "Bold", "size": 18, "leading": 22, "align": "left",
                            "space_after": 6, "color": "ink", "all_caps": True},
        "Body": {"font_style": "Regular", "size": 9, "leading": 13, "align": "left",
                  "space_after": 3, "color": "ink"},
        "BodyClause": {"font_style": "Regular", "size": 8.5, "leading": 12.5, "align": "justify",
                        "space_after": 2, "left_indent": 5, "first_line_indent": -5, "color": "ink"},
        "ListItem": {"font_style": "Regular", "size": 9, "leading": 13, "align": "left",
                      "space_after": 1.5, "left_indent": 5, "first_line_indent": -5, "color": "ink"},
        "Folio": {"font_style": "Regular", "size": 9, "leading": 9, "align": "right", "color": "ink"},
        "Caption": {"font_style": "Bold", "size": 10, "leading": 12, "color": "ink"},
        "Signature": {"font_style": "Regular", "size": 9, "leading": 12, "color": "ink"},
        "TableHeader": {"font_style": "Bold", "size": 9, "leading": 11, "align": "left", "color": "ink"},
        "TableCell": {"font_style": "Regular", "size": 9, "leading": 11, "align": "left", "color": "ink"},
        "TableTotal": {"font_style": "Bold", "size": 9.5, "leading": 11.5, "align": "left", "color": "ink"},
    },
    "table": {
        "header_fill": "silver", "header_tint": 30,
        "total_fill": "silver", "total_tint": 50,
        "row_rule": "silver", "row_rule_weight": 0.25,
        "total_rule_weight": 0.75,
    },
    # Bundled brand art (design spec §9, disposition = bundle-with-jsx). The
    # generator ships whichever of these it finds in the brand assets dir; any
    # missing file degrades to a styled placeholder frame.
    "brand_assets": {
        "logo": "170703_Nikken_Brandmark_Tagline.ai",
        "signature": "wataru_tanaka_signature.png",
        "divider": "Podium.jpg",
    },
    # Ordered build plan. Each entry dispatches to a builder in jsx_modea.
    # source = key in the proposal "sections" dict; title = section heading.
    "sections": [
        {"type": "cover"},
        {"type": "letter", "source": "cover_letter"},
        {"type": "narrative", "source": "project_understanding", "title": "Project Understanding"},
        {"type": "narrative", "source": "methodology", "title": "Methodology"},
        {"type": "narrative", "source": "scope_deliverables", "title": "Scope & Deliverables"},
        {"type": "narrative", "source": "schedule", "title": "Schedule"},
        {"type": "narrative", "source": "team", "title": "Team Structure"},
        {"type": "fee", "title": "Financial Proposal", "source": "financial"},
        {"type": "experience", "title": "Relevant Experience", "source": "relevant_experience"},
        {"type": "narrative", "source": "assumptions_exclusions", "title": "Assumptions & Exclusions"},
    ],
}


def get_profile(template_id: str | None = None) -> dict:
    """Return a deep copy of the design profile for ``template_id``.

    Only ``commercial`` exists today; ``master-plan`` / ``technical`` fall back to
    it until their own sample-derived profiles are added.
    """
    return deepcopy(DEFAULT_COMMERCIAL_PROFILE)
