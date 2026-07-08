# FP-GEN — Technical Description

**System Name:** FP-GEN (Fee Proposal Generator)
**Version:** MVP v0.1.0
**Owner / Operating Department:** AIL — Business Process Transformation Team
**Platform:** exaBase (internal containerised service)
**Last Updated:** 2026-05-28

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
| LLM Model | `gpt-5.4-mini` (configurable via `OPENAI_MODEL`) |
| PDF Extraction | `pypdf` + `pdfplumber` |
| DOCX Extraction | `python-docx` |
| InDesign Export | `pywin32` / COM automation (Windows only) |
| Containerisation | Docker (python:3.12-slim base) |
| Reverse Proxy | nginx |
| Storage | JSON files on a persistent volume |

---

## 5. Data Processing Flow

```
1. User uploads RFP file  (PDF / Word / TXT / MD / JSON)
        │
        ▼
2. rfp_file_service  —  extracts plain text from the file
        │
        ▼
3. parser (LLM call)  —  sends text to OpenAI GPT-5.4-mini with a structured
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
7. storage  —  persists proposal as a UUID-keyed JSON file
        │
        ├─► REST API response (JSON)
        └─► InDesign export  (.indd, Windows hosts only)
```

---

## 6. API Endpoints

| Method | Path | Description |
|---|---|---|
| `GET` | `/health` | Health check; reports InDesign export capability |
| `GET` | `/v1/capabilities` | Capability flags |
| `GET` | `/v1/reference-projects` | List available reference projects |
| `GET` | `/v1/personnel` | List personnel registry |
| `GET` | `/v1/team-presets` | List saved team presets |
| `POST` | `/v1/parse` | Parse raw RFP text via LLM |
| `POST` | `/v1/parse/file` | Upload RFP file; extract text and parse |
| `POST` | `/v1/fee/calculate` | Calculate fee from role/phase/rate inputs |
| `POST` | `/v1/proposals/generate` | Full pipeline: parse → fee → select → build → save |
| `POST` | `/v1/proposals/generate/file` | Same pipeline, RFP supplied as file upload |
| `GET` | `/v1/proposals/{id}` | Retrieve a saved proposal |
| `PATCH` | `/v1/proposals/{id}` | Update proposal sections |
| `GET` | `/v1/proposals/{id}/export/markdown` | Export proposal as Markdown |
| `POST` | `/v1/proposals/{id}/export/indd` | Trigger InDesign export |
| `GET` | `/v1/proposals/{id}/export/indd/download` | Download exported `.indd` file |
| `GET` | `/v1/assets` | List uploaded InDesign asset files |
| `POST` | `/v1/assets/upload` | Upload a `.indd` template, CV, or experience asset |

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

## 8. Reference Data

The following reference datasets are bundled in the container image under `/app/appdata/` (read-only; not overwritten by the persistent volume):

| File | Contents |
|---|---|
| `reference_projects.json` | Past project portfolio used for relevant-experience matching |
| `personnel.json` | Staff names, titles, and roles available for team assignment |
| `team_presets.json` | Pre-configured team compositions |
| `assets_registry.json` | Index of available InDesign asset files (templates, CVs, experience sheets) |

Proposal JSON files and export artefacts are written to the persistent volume at `/app/data/`.

---

## 9. Supported Input File Types

| Extension | Parser |
|---|---|
| `.pdf` | `pypdf` (text layer) with `pdfplumber` fallback |
| `.docx` | `python-docx` paragraph extraction |
| `.txt`, `.md`, `.json` | Direct text decode (UTF-8 / CP1252) |

---

## 10. LLM Integration

The parser calls the OpenAI Chat Completions API (model `gpt-5.4-mini` by default). Alternatively, a LiteLLM proxy URL can be set via `LITELLM_URL` / `LITELLM_MASTER_KEY` environment variables, allowing the system to route through an internally hosted LLM gateway without a direct public internet dependency.

The system prompt instructs the model to return a single, strictly-typed JSON object. A regex-based fallback parser is used if the LLM API is unavailable or returns an unparseable response.

No RFP content is stored externally; the only external call is the LLM API request containing the extracted RFP text. Responses are processed in-memory and discarded after the proposal is built.

---

## 11. InDesign Export

The system supports two independent paths for producing a formatted `.indd` proposal document.

### 11a. Server-side COM Automation (Windows only)

When the FastAPI service is running on a Windows host with Adobe InDesign installed, it can drive InDesign directly via COM automation (`pywin32`). The server opens the corporate template, performs text-frame substitutions (project name, client name, date), appends paginated section spreads, and saves the output `.indd` file to the export volume. Long sections are automatically split across continuation pages. The file is then available for download via `GET /v1/proposals/{id}/export/indd/download`.

This capability is **disabled by default** in the Docker image (`FPGEN_ENABLE_INDD_EXPORT=false`) and is intended for use on the Windows workstation deployment only.

### 11b. Client-side JSX ExtendScript Bundle

The `POST /v1/proposals/{id}/export/jsx` endpoint produces a **self-contained ZIP bundle** that does not require pywin32 or a server-side InDesign installation. The ZIP contains:

- `assemble_proposal.jsx` — an InDesign ExtendScript (JavaScript) assembly script
- `assets/template/` — the bundled corporate `.indd` template
- `assets/cvs/` — selected CV `.indd` files for assigned team members
- `assets/experience/` — selected experience-sheet `.indd` files
- `README.txt` — step-by-step extraction and execution instructions

The user extracts the ZIP to any folder on a local Windows machine, opens the Scripts panel in Adobe InDesign (CC 2019 or later), and double-clicks `assemble_proposal.jsx`. The script assembles the full document, applies template substitutions, appends section pages, inserts CV and experience spreads (inserting a styled placeholder page for any missing asset), and saves `<ProjectName>.indd` in the same folder. This path works on any Windows workstation with InDesign without requiring server-side COM access.

---

## 12. Environment Variables

| Variable | Default | Description |
|---|---|---|
| `OPENAI_API_KEY` | — | OpenAI API key |
| `OPENAI_MODEL` | `gpt-5.4-mini` | LLM model name |
| `LLM_PROVIDER` | `openai` | `openai` or `litellm` |
| `LITELLM_URL` | — | LiteLLM proxy base URL |
| `LITELLM_MASTER_KEY` | — | LiteLLM master key |
| `ENABLE_OPENAI_SYNTHESIS` | `true` | Enable/disable LLM call |
| `FPGEN_ENABLE_INDD_EXPORT` | `false` | Enable InDesign COM export |
| `FPGEN_STORAGE_DIR` | `/app/data/proposals` | Proposal JSON storage path |
| `FPGEN_EXPORT_DIR` | `/app/data/exports` | InDesign export output path |
| `DEFAULT_CURRENCY` | `USD` | Default fee currency |
| `DEFAULT_OVERHEAD_PCT` | `0.10` | Default overhead rate (10%) |
| `TRAVEL_UNIT_COST_USD` | `6000` | Default per-trip travel cost |

---

## 13. Security Considerations

- The API does not implement authentication at the application layer; access control is delegated to the exaBase platform and the nginx reverse proxy.
- The `OPENAI_API_KEY` and `LITELLM_MASTER_KEY` are injected via environment variables and are never written to disk or returned by any endpoint.
- Uploaded RFP files are processed in-memory; no raw uploads are persisted.
- Asset uploads are restricted to `.indd` files with alphanumeric `asset_id` validation to prevent path traversal.
- All file paths are resolved relative to configured base directories; no user-supplied paths are accepted directly.

---

## 14. Connected Systems

| System | Integration Type | Purpose |
|---|---|---|
| OpenAI API (or LiteLLM proxy) | HTTPS REST | RFP requirement extraction via LLM |
| Adobe InDesign (COM) | Windows COM automation | Formatted proposal document export |
| exaBase | Internal container platform | Hosting, routing, persistent storage |

> **Note (Notion integration):** The proposal sections listed in §7 are the content units intended for future Notion page storage. Each section maps 1-to-1 to a Notion page block for draft management and review via the Notion API. An integration verification page should be provisioned in Notion covering these nine sections before enabling that workflow.
