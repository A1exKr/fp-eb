```markdown
# fp-gen

FP‑GEN — Fee Proposal Generator (VBA)

This repository contains exported VBA modules, userforms and helpers that implement a small pipeline:
- parse RFP text or PDFs via an LLM API
- map AI output to a canonical review model
- open a Review & Finalize editor for each section
- generate PowerPoint or InDesign fee proposals

What’s included
---------------
- src/ — exported .bas and .cls modules (VBA source)
- forms/ — exported .frm and .frx userform files
- formats/ — PPTX/INDD templates (not included by default)
- docs/ — design notes and instructions
- test/ — smoke tests and sample payloads (rfp_mock.json)

Quick install (developer)
1. Export VBA modules from the VBE (Alt+F11) using File ▸ Export File…:
   - Standard modules → save as `.bas`
   - Class modules → save as `.cls`
   - UserForms → save as `.frm` and `.frx` (export both)
2. Copy exported files into `src/` and `forms/` respectively.
3. Ensure `JsonConverter.bas` (VBA-JSON) is present in `src/`.
4. Recommended reference (Tools ▸ References…): Microsoft Scripting Runtime (optional, but helpful).

How to open the Review editor from your existing form
After your parser returns `rfpAnalysisJson As String`, call:
    modReviewBridge.OpenReviewWithJson rfpAnalysisJson

Smoke tests
- Test_OpenReviewWithInlineMock (modSmokeTest)
- Test_OpenReviewFromFile (reads `rfp_mock.json` from repo/test/ or current directory)

Security note
- Do NOT commit raw API keys. The project expects an `apikey.dat` file (encrypted) — keep it out of git (.gitignore includes it).
- Remove any hard-coded secrets before pushing.

Repo layout (recommended)
- README.md
- .gitignore
- .gitattributes
- src/
  - *.bas
  - *.cls
- forms/
  - *.frm
  - *.frx
- docs/
  - design_notes.md
- formats/
  - FeeProposalFormat.pptx (template; optional)
- test/
  - rfp_mock.json

If you want, I can also:
- provide a zip of the prepared repo contents (you'll need to tell me if you want me to render all provided module files as repo files in the chat),
- or generate a commit-ready set of files you can copy into a new repo (I can output file blocks of each exported module on request).