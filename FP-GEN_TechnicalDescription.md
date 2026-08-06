# FP-GEN — Technical Description

**System Name:** FP-GEN (Fee Proposal Generator)
**Version:** MVP v0.1.0
**Owner / Operating Department:** AIL — Business Process Transformation Team
**Platform:** exaBase (internal containerised service)
**Last Updated:** 2026-08-06

---

## 1. Purpose and Overview

FP-GEN is an internal web application that automates the drafting of fee proposals in response to Requests for Proposal (RFPs). A staff member uploads an RFP document; the system extracts its requirements using a Large Language Model (LLM) API, assembles the full proposal content across all standard sections, calculates the fee schedule, and produces a reviewable draft. The output can be exported as a structured JSON record and, where Adobe InDesign is available on the host, as a formatted `.indd` file.

The goal is to eliminate the manual effort of translating an RFP into a proposal skeleton, reduce inconsistency between proposals, and shorten the turn-around time from brief receipt to first draft.

---

## 2. Operating Department

| Item | Detail |
|---|---|
| Department | AIL (Business Process Transformation Team) |
| Connected Platform | exaBase |
| Deployment Unit | Two pods — `fpgen-app-api` (FastAPI backend) + `fpgen-app-nginx` (reverse proxy) |
| Container Image | `ghcr.io/a1exkr/fpgen-mvp-api:latest` |

---

## 3. Architecture

```
┌─────────────────────────────────────────────────────────┐
│                        exaBase                          │
│                                                         │
│  ┌──────────────┐        ┌────────────────────────┐    │
│  │   Browser    │◄──────►│  nginx  (reverse proxy) │    │
│  │  (SPA / UI)  │        └──────────┬─────────────┘    │
│  └──────────────┘                   │                   │
│                                     ▼                   │
│                         ┌───────────────────────┐       │
│                         │  FastAPI Application  │       │
│                         │   (Python 3.12)       │       │
│                         │                       │       │
│                         │  ┌─────────────────┐  │       │
│                         │  │  RFP File Svc   │  │       │
│                         │  │ (PDF/DOCX/TXT)  │  │       │
│                         │  └────────┬────────┘  │       │
│                         │           │            │       │
│                         │  ┌────────▼────────┐  │       │
│                         │  │  Parser (LLM)   │◄─┼───────┼──► OpenAI API
│                         │  └────────┬────────┘  │       │    (or LiteLLM)
│                         │           │            │       │
│                         │  ┌────────▼────────┐  │       │
│                         │  │  Fee Engine     │  │       │
│                         │  └────────┬────────┘  │       │
│                         │           │            │       │
│                         │  ┌────────▼────────┐  │       │
│                         │  │ Proposal Builder│  │       │
│                         │  └────────┬────────┘  │       │
│                         │           │            │       │
│                         │  ┌────────▼────────┐  │       │
│                         │  │  Storage (JSON) │  │       │
│                         │  └─────────────────┘  │       │
│                         │                       │       │
│                         │  ┌─────────────────┐  │       │
│                         │  │  InDD Exporter  │  │       │
│                         │  │ (Windows only)  │  │       │
│                         │  └─────────────────┘  │       │
│                         └───────────────────────┘       │
│                                                         │
│  Persistent Volume: /app/data  (proposals, exports)     │
└─────────────────────────────────────────────────────────┘
```

---

## 4. Technology Stack

| Layer | Technology |
|---|---|
| Runtime | Python 3.12 |
| Web Framework | FastAPI 0.116 + Uvicorn 0.35 |
| Data Validation | Pydantic v2 |
| LLM Client | `openai` SDK v1 (OpenAI or LiteLLM proxy) |
| LLM Model | `gpt-4.1-mini` (configurable via `OPENAI_MODEL`) |
| PDF Extraction | `pypdf` + `pdfplumber` |
| DOCX Extraction | `python-docx` |
| InDesign Export | `pywin32` / COM automation (Windows only) |
| Containerisation | Docker (python:3.12-slim base) |
| Reverse Proxy | nginx |
| Persistence | PostgreSQL via SQLAlchemy 2 + Alembic (SQLite fallback for local dev); proposals stored as a JSON payload column |

---

## 5. Data Processing Flow

```
1. User uploads RFP file  (PDF / Word / TXT / MD / JSON)
        │
        ▼
2. rfp_file_service  —  extracts plain text from the file
        │
        ▼
3. parser (LLM call)  —  sends text to the configured chat model with a structured
   system prompt; receives a normalised JSON object containing:
   project, client, understanding, methodology, scope, schedule,
   team, fee (rates / effortByPhase / travel / subconsultants),
   experience, assumptions
        │
        ▼
4. fee_engine  —  computes:
   labor total (role × hours-per-phase), overhead,
   included subconsultant fees, travel reimbursables,
   lump-sum total
        │
        ▼
5. relevant_selector  —  scores reference projects from the
   internal registry against project type, location, and
   keyword overlap; returns top-3 matches
        │
        ▼
6. proposal_builder  —  assembles nine proposal sections:
   Cover Letter / Project Understanding / Methodology /
   Scope & Deliverables / Schedule / Team Structure /
   Financial Proposal / Relevant Experience /
   Assumptions & Exclusions
        │
        ▼
7. storage  —  persists the proposal payload to PostgreSQL under a UUID key
        │
        ├─► REST API response (JSON)
        ├─► InDesign export  (.indd, Windows hosts only)
        ├─► JSX ExtendScript bundle (.zip)
        │
        ▼
8. review loop (§8)  —  an authorised reviewer re-runs individual pipeline
   stages on the saved proposal: reselect reference projects, rebuild a
   single section, or recalculate the fee. Every stage previews before it
   commits, and the cascade respects each section's provenance.
```

---

## 6. API Endpoints

`Auth` column: **none** = unauthenticated; **user** = any authenticated identity; **admin** =
`FPGEN_ADMIN_ROLES` (default `fpgen_admin`); **admin/finance** = `FPGEN_FINANCE_ROLES`
(default `fpgen_admin,Finance`). See §14.

| Method | Path | Auth | Description |
|---|---|---|---|
| `GET` | `/health` | none | Health check; reports InDesign export capability |
| `GET` | `/v1/capabilities` | none | Capability flags |
| `GET` | `/v1/me` | user | Current identity, groups, and UI capability flags |
| `GET` | `/` , `/ui` | user | Single-page application shell |
| `GET` | `/proposals/{id}/review` | user | Proposal review and regeneration page |
| `GET` | `/admin`, `/v1/admin/*` | admin/finance | Master-data console and CRUD (personnel, rates, presets, reference projects, assets) |
| `GET` | `/v1/reference-projects` | none | List available reference projects |
| `GET` | `/v1/personnel` | none | List personnel registry |
| `GET` | `/v1/team-presets` | none | List saved team presets |
| `POST` | `/v1/parse` | none | Parse raw RFP text via LLM |
| `POST` | `/v1/parse/file` | none | Upload RFP file; extract text and parse |
| `POST` | `/v1/fee/calculate` | none | Calculate fee from role/phase/rate inputs |
| `POST` | `/v1/proposals/generate` | user | Full pipeline: parse → fee → select → build → save |
| `POST` | `/v1/proposals/generate/file` | user | Same pipeline, RFP supplied as file upload |
| `GET` | `/v1/proposals/{id}` | user | Retrieve a saved proposal |
| `PUT` | `/v1/proposals/{id}` | user | Overwrite proposal section text (marks sections as `edited`) |
| `POST` | `/v1/proposals/{id}/sections/{key}/regenerate` | admin | Rebuild one section; optional instruction (§8) |
| `POST` | `/v1/proposals/{id}/experience/reselect` | admin | Re-run or override reference-project selection |
| `POST` | `/v1/proposals/{id}/fee/recalculate` | admin/finance | Re-run the fee engine and cascade |
| `POST` | `/v1/proposals/{id}/export/jsx` | user | Build the InDesign JSX ExtendScript ZIP bundle |
| `POST` | `/v1/proposals/{id}/export/indd` | user | Trigger InDesign COM export |
| `GET` | `/v1/proposals/{id}/export/indd/download` | user | Download exported `.indd` file |
| `GET` | `/v1/assets` | none | List uploaded InDesign asset files |
| `POST` | `/v1/assets/upload` | none | Upload a `.indd` template, CV, or experience asset |

---

## 7. Proposal Sections Generated

The system auto-drafts the following standard proposal sections:

1. **Cover Letter** — personalised salutation and project commitment statement
2. **Project Understanding** — 150–200 word executive summary of objectives, site context, client goals, key challenges, and expected outcomes
3. **Methodology** — step-by-step architectural / master-planning design process
4. **Scope and Deliverables** — itemised scope tasks and expected output list
5. **Schedule** — total duration, phase plan with estimated labor fees, and milestones
6. **Team Structure** — principal, project manager, specialist roles, and subconsultants
7. **Financial Proposal** — fee breakdown: labor, overhead, subconsultants, travel, reimbursables, and lump-sum total
8. **Relevant Experience** — up to three reference projects matched to the RFP
9. **Assumptions and Exclusions** — commercial disclaimer covering taxes, reimbursables, optional services, and specialty consultants

---

## 8. Review and Regeneration Model

A generated proposal is not final. The review page (`/proposals/{id}/review`) lets an authorised user
re-run individual pipeline stages against the saved proposal. The design constraint throughout is that
**the LLM stays a text function and the human stays in the loop** — no autonomous agent, no model access
to fee arithmetic, and no silent overwriting of authored content.

### 8.1 Stored proposal payload

| Key | Contents |
|---|---|
| `parsed` | Structured RFP extraction (the source data for most sections) |
| `fee_input` | Roles, rates, hours-by-phase, overhead, subconsultants, travel, misc — persisted so the fee engine can be re-run |
| `financial` | Output of `calculate_fee(fee_input)` |
| `relevant_experience` | Selected reference-project records |
| `experience_selection_mode` | `auto` (scored) or `manual` (explicit picks) |
| `sections` | The nine rendered section texts |
| `section_state` | Per section: `{origin, stale_reason}` (§8.4) |
| `markdown` | Rendered document, recomputed on every mutation |

### 8.2 Regeneration endpoints

All three accept `commit: false` to return a **preview** without writing. This is required: a commit
would otherwise silently overwrite unsaved manual edits in the review editor. The preview response is
echoed back on apply (`apply_text` / `input_patch`) so no second LLM call is spent.

| Endpoint | Behaviour |
|---|---|
| `sections/{key}/regenerate` | Rebuilds one section; optional free-text `instruction` |
| `experience/reselect` | Re-runs `select_relevant_projects`, or applies an explicit `selected_reference_ids` list |
| `fee/recalculate` | Re-runs `calculate_fee` and cascades |

Responses carry `changed_sections`, `stale_sections`, the applied `input_patch`, and a human-readable
`notice`.

### 8.3 How an instruction is applied

Only the **Cover Letter** is rewritten as text by the LLM. For every other section the instruction is
converted into a **type-checked JSON patch of that section's source inputs in `parsed`**, and the section
is then re-rendered deterministically from the patched data. The model therefore never writes rendered
output, never touches fee arithmetic, and cannot introduce a field the exporters read.

Allowlisted input paths (`SECTION_INPUT_FIELDS`):

| Section | Editable inputs |
|---|---|
| Project Understanding | `understanding.understanding` |
| Methodology | `methodology.text` |
| Scope and Deliverables | `scope.scopeList`, `scope.deliverablesList` |
| Schedule | `project.duration`, `schedule.totalWeeks`, `schedule.milestones` |
| Team Structure | `team.principal.name/title`, `team.pm.name/title` |
| Assumptions and Exclusions | `assumptions.defaultText` |

**Financial Proposal** and **Relevant Experience** reject an instruction with HTTP 400 — they are changed
only through `fee/recalculate` and `experience/reselect` respectively, both of which are pure
server-side Python. Any patch key outside the allowlist is dropped; values are type- and length-checked
(`str`, `int`, `list[str]`; 6 000 characters, 60 list items).

When no LLM connection is configured, or a call fails, regeneration falls back to the deterministic
render and returns an explanatory `notice` rather than failing.

### 8.4 Dependency cascade and provenance

The section dependency graph is declared once in `SECTION_SOURCES` (what each section reads) and
`INPUT_FIELD_SOURCES` (which source each allowlisted path belongs to). `propagate()` computes the
downstream closure of any change, so no endpoint hardcodes its own rebuild list.

The edges are one-directional. Nothing in the Team or Schedule allowlists is an input to
`calculate_fee`, so editing them cannot move the financials; the cascade runs the other way:

```
fee_input ──calculate_fee──> financial ──> [financial, schedule] sections
fee_input ──────────────────────────────-> [team, financial] sections
parsed.project.identity + keywords ──select_relevant_projects──> relevant_experience
parsed.project.identity + fee.rates ────────────────────────────> cover_letter
```

Source keys are finer-grained than the `parsed` blocks — `parsed.project.duration` is distinct from
`parsed.project.identity` — so a schedule edit does not falsely invalidate the cover letter.

What happens to a downstream section depends on its **provenance**, not its name:

| `origin` | Meaning | On upstream change |
|---|---|---|
| `derived` | Deterministic render of stored data | Rebuilt silently — it is a projection, nothing is lost |
| `edited` | Human edit via the review editor | **Never overwritten**; flagged stale with a reason |
| `llm` | Produced by `text_completion` | **Never overwritten**; re-running costs money, is non-deterministic, and would destroy applied polish |

`PUT /v1/proposals/{id}` marks any section whose text actually changed as `edited` and does not cascade —
that is the human authoring. Proposals predating `section_state` default to `derived`, since their edit
history is unknown.

A special case applies to reference projects: when `experience_selection_mode` is `auto`, a change to
project identity or keywords invalidates the **selection** rather than the rendered text, so the section
is flagged with a "re-run Auto-select" reason instead of being rebuilt. A `manual` selection is never
replaced automatically.

### 8.5 Review page behaviour

Controls appear only when `GET /v1/me` reports the matching capability, and the same check is enforced
server-side. Stale sections are shown with an amber card, an "Out of date — <reason>" note, and a
**Rebuild** action; the status bar and preview dialog report `Rebuilt: …` and `N sections need review: …`.
A reviewer with unsaved manual edits is warned before an apply overwrites them.

The reviewer is also warned when the RFP's planned phases (`parsed.fee.effortByPhase`) and the costed
phases (`fee_input` hours) diverge. The mismatch is reported through `notice` rather than written into
the client-facing text.

---

## 9. Reference Data

The following reference datasets are bundled in the container image under `/app/appdata/` (read-only; not overwritten by the persistent volume):

| File | Contents |
|---|---|
| `reference_projects.json` | Past project portfolio used for relevant-experience matching |
| `personnel.json` | Staff names, titles, and roles available for team assignment |
| `team_presets.json` | Pre-configured team compositions |
| `assets_registry.json` | Index of available InDesign asset files (templates, CVs, experience sheets) |

Proposal payloads are stored in PostgreSQL (`FPGEN_DB_SCHEMA`, managed by Alembic). Master data
(personnel, unit rates, team presets, reference projects, asset records) is also database-backed and
editable through the admin console; the bundled JSON files above are the seed set. InDesign export
artefacts are written to the persistent volume at `/app/data/`.

---

## 10. Supported Input File Types

| Extension | Parser |
|---|---|
| `.pdf` | `pypdf` (text layer) with `pdfplumber` fallback |
| `.docx` | `python-docx` paragraph extraction |
| `.txt`, `.md`, `.json` | Direct text decode (UTF-8 / CP1252) |

---

## 11. LLM Integration

The parser calls the OpenAI Chat Completions API (model `gpt-4.1-mini` by default). Alternatively, a LiteLLM proxy URL can be set via `LITELLM_URL` / `LITELLM_MASTER_KEY` environment variables, allowing the system to route through an internally hosted LLM gateway without a direct public internet dependency.

There are exactly three LLM call sites, all funnelled through `services/openai_service.py`:

| Call site | Purpose |
|---|---|
| `parser.parse_rfp` | RFP text → structured JSON (`json_completion`) |
| `proposal_builder.regenerate_cover_letter` | Cover-letter polish (`text_completion`) |
| `proposal_builder.propose_input_patch` | Instruction → allowlisted input patch (`json_completion`, §8.3) |

Fee arithmetic, reference-project selection, persistence and both export paths are pure server-side code;
no model performs calculations or writes files.

The system prompt instructs the model to return a single, strictly-typed JSON object. A regex-based fallback parser is used if the LLM API is unavailable or returns an unparseable response.

No RFP content is stored externally; the only external call is the LLM API request containing the extracted RFP text. Responses are processed in-memory and discarded after the proposal is built.

---

## 12. InDesign Export

The system supports two independent paths for producing a formatted `.indd` proposal document.

### 12a. Server-side COM Automation (Windows only)

When the FastAPI service is running on a Windows host with Adobe InDesign installed, it can drive InDesign directly via COM automation (`pywin32`). The server opens the corporate template, performs text-frame substitutions (project name, client name, date), appends paginated section spreads, and saves the output `.indd` file to the export volume. Long sections are automatically split across continuation pages. The file is then available for download via `GET /v1/proposals/{id}/export/indd/download`.

This capability is **disabled by default** in the Docker image (`FPGEN_ENABLE_INDD_EXPORT=false`) and is intended for use on the Windows workstation deployment only.

### 12b. Client-side JSX ExtendScript Bundle

The `POST /v1/proposals/{id}/export/jsx` endpoint produces a **self-contained ZIP bundle** that does not require pywin32 or a server-side InDesign installation. The ZIP contains:

- `assemble_proposal.jsx` — an InDesign ExtendScript (JavaScript) assembly script
- `assets/template/` — the bundled corporate `.indd` template
- `assets/cvs/` — selected CV `.indd` files for assigned team members
- `assets/experience/` — selected experience-sheet `.indd` files
- `README.txt` — step-by-step extraction and execution instructions

The user extracts the ZIP to any folder on a local Windows machine, opens the Scripts panel in Adobe InDesign (CC 2019 or later), and double-clicks `assemble_proposal.jsx`. The script assembles the full document, applies template substitutions, appends section pages, inserts CV and experience spreads (inserting a styled placeholder page for any missing asset), and saves `<ProjectName>.indd` in the same folder. This path works on any Windows workstation with InDesign without requiring server-side COM access.

---

## 13. Environment Variables

| Variable | Default | Description |
|---|---|---|
| `OPENAI_API_KEY` | — | OpenAI API key |
| `OPENAI_MODEL` | `gpt-4.1-mini` | LLM model name |
| `LLM_PROVIDER` | `openai` | `openai` or `litellm` |
| `LITELLM_URL` | — | LiteLLM proxy base URL |
| `LITELLM_MASTER_KEY` | — | LiteLLM key (a scoped virtual key, not the gateway master key) |
| `ENABLE_OPENAI_SYNTHESIS` | `true` | Enable/disable LLM call |
| `FPGEN_ENABLE_INDD_EXPORT` | `false` | Enable InDesign COM export |
| `DATABASE_URL` | SQLite file | PostgreSQL connection string |
| `FPGEN_DB_SCHEMA` | `fpgen` | PostgreSQL schema for the application tables |
| `FPGEN_AUTH_ENABLED` | `false` | Trust forwarded identity headers; `false` yields a synthetic dev admin |
| `FPGEN_AUTH_DEV_USER` | `dev-admin@example.com` | Identity used when auth is disabled |
| `FPGEN_AUTH_DEV_GROUPS` | `fpgen_admin,Finance` | Groups used when auth is disabled |
| `FPGEN_ADMIN_ROLES` | `fpgen_admin` | Groups permitted to administer and regenerate |
| `FPGEN_FINANCE_ROLES` | `fpgen_admin,Finance` | Groups permitted to edit unit rates and recalculate fees |
| `FPGEN_STORAGE_DIR` | `/app/data/proposals` | Legacy proposal file path |
| `FPGEN_EXPORT_DIR` | `/app/data/exports` | InDesign export output path |
| `DEFAULT_CURRENCY` | `USD` | Default fee currency |
| `DEFAULT_OVERHEAD_PCT` | `0.10` | Default overhead rate (10%) |
| `TRAVEL_UNIT_COST_USD` | `6000` | Default per-trip travel cost |

---

## 14. Security Considerations

### 14.1 Authentication and authorisation

The application does not present a login page. It trusts `X-Forwarded-User` / `X-Forwarded-Email` /
`X-Forwarded-Groups` from an oauth2-proxy sitting in front of it, which performs the Keycloak OIDC
handshake. `FPGEN_AUTH_ENABLED=false` bypasses this for local development by returning a synthetic
admin identity.

| Surface | Requirement |
|---|---|
| SPA shell, proposal read/write, both export paths | Any authenticated identity |
| Admin console and `/v1/admin/*` | `FPGEN_ADMIN_ROLES` or `FPGEN_FINANCE_ROLES` |
| Section regeneration, experience reselection | `FPGEN_ADMIN_ROLES` |
| Fee recalculation | `FPGEN_FINANCE_ROLES` |

`GET /v1/me` advertises `can_setup`, `can_regenerate` and `can_recalculate_fee` so the UI can hide
controls, but every gate is enforced server-side regardless.

> **Deployment note:** if the oauth2-proxy front door is absent, setting `FPGEN_AUTH_ENABLED=true`
> makes every request return `401` with no login prompt. See `EXABASE_REQUESTS.md` (EB-006).

### 14.2 Untrusted input and prompt injection

RFP files are untrusted input, and their extracted text reaches the model. The controls are:

- The model is never given tools, file access, or network access — only `chat.completions`.
- Model output is never executed and never written to the exporters directly. For every section except
  the cover letter, output is constrained to an **allowlist of typed input fields** (§8.3); unknown keys
  are dropped and values are type- and length-checked before being applied.
- Fee arithmetic, reference-project selection and both export paths contain no model output.
- System prompts instruct the model to treat existing field values as data and ignore instructions found
  inside them.

### 14.3 General

- The `OPENAI_API_KEY` and `LITELLM_MASTER_KEY` are injected via environment variables and are never written to disk or returned by any endpoint.
- Uploaded RFP files are processed in-memory; no raw uploads are persisted.
- Asset uploads are restricted to `.indd` files with alphanumeric `asset_id` validation to prevent path traversal.
- All file paths are resolved relative to configured base directories; no user-supplied paths are accepted directly.

---

## 15. Connected Systems

| System | Integration Type | Purpose |
|---|---|---|
| OpenAI API (or LiteLLM proxy) | HTTPS REST | RFP requirement extraction via LLM |
| Adobe InDesign (COM) | Windows COM automation | Formatted proposal document export |
| exaBase | Internal container platform | Hosting, routing, persistent storage |

> **Note (Notion integration):** The proposal sections listed in §7 are the content units intended for future Notion page storage. Each section maps 1-to-1 to a Notion page block for draft management and review via the Notion API. An integration verification page should be provisioned in Notion covering these nine sections before enabling that workflow.
