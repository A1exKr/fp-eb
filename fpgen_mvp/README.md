# FP-GEN MVP for exaBase

> **Status: deployed and running.**
> Image: `ghcr.io/a1exkr/fpgen-mvp-api:latest`
> Repo: https://github.com/A1exKr/fp-eb
> Pods: `fpgen-app-api` + `fpgen-app-nginx` — both Running.
> Recommended import: `../FP-GEN_exaBase_import_v2.json`

This folder contains the Python backend for an inside-exaBase MVP of FP-GEN.

It is not the whole original FP-GEN project. Your existing VB/Excel implementation already exists at:

`C:\Users\03669\Desktop\Trud\matls\created\FP-GEN`

That root contains assets such as:

- `FP-GEN.xlsm`
- `code/`
- `APP/`
- `templates/`
- `REFERENCE/`
- `prompts/`
- `formats/`
- `FEE/`
- `Presentation/`

## What this MVP is

This MVP is the backend service that exaBase should run internally.

It provides:

- RFP parsing
- file-based RFP intake for `.txt`, `.md`, `.json`, `.pdf`, and `.docx`
- fee calculation
- proposal section generation
- relevant-project selection
- JSON persistence
- Markdown proposal output
- local `.indd` export through Adobe InDesign on Windows

Main code:

- `app/main.py`
- `app/config.py`
- `app/services/parser.py`
- `app/services/fee_engine.py`
- `app/services/proposal_builder.py`
- `app/services/relevant_selector.py`
- `app/services/storage.py`

## What this MVP is not

This folder does not replace your existing VB UI.

Your VB UI should currently be treated as:

- business logic reference
- UI/flow reference
- source of templates, prompts, and calculation rules
- source for data migration into the Python/exaBase version

For the first exaBase MVP, do not try to put the entire VB application into the container image.

That would create unnecessary risk.

The safe boundary is:

- exaBase image contains only the Python backend from this `fpgen_mvp` folder
- VB/Excel assets remain outside the image unless you deliberately copy selected files into the MVP

## Recommended architecture

For now, use one internal exaBase service only:

- one sideapp running this FastAPI backend

That one service should expose:

- `GET /health`
- `POST /v1/parse`
- `POST /v1/fee/calculate`
- `POST /v1/proposals/generate`
- `GET /v1/proposals/{proposal_id}`

The revised full canvas file already assumes this pattern:

- `../FP-GEN_exaBase_import.json`

## Relationship to your VB UI

Your current VB UI in `C:\Users\03669\Desktop\Trud\matls\created\FP-GEN` should guide what the exaBase app eventually does.

Use it for:

- section names
- fee logic rules
- export expectations
- field mapping
- prompts and templates
- reference-project data

Do not use it as the build context for the container unless you intentionally need a file from it.

Reason:

- the image should stay small
- the runtime should be predictable
- Excel/VBA files are not needed to run the Python API
- including the whole legacy project will complicate deployment

If you need a legacy asset, copy only that asset into `fpgen_mvp/data/` or another explicit folder inside this MVP.

Examples:

- copy reference-project records into `data/reference_projects.json`
- copy text prompts into a dedicated `data/prompts/` folder
- copy export templates only after the backend actually needs them

## Local setup

Use the existing Python installation only once to create the project-local virtual environment:

`C:\Users\03669\Desktop\Trud\matls\created\PQQ-Automation\python\python.exe`

From this folder:

```powershell
Set-Location "C:\Users\03669\Desktop\Trud\matls\created\FP-GEN\exaBase\fpgen_mvp"
Copy-Item .env.example .env
& "C:\Users\03669\Desktop\Trud\matls\created\PQQ-Automation\python\python.exe" -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
```

## LLM configuration

For now, the recommended path is direct OpenAI.

Reason:

- you do not currently have `LITELLM_MASTER_KEY`
- it removes one moving part during initial validation
- you can switch to LiteLLM later without changing the outer API shape

### Option 1: direct OpenAI

```env
LLM_PROVIDER=openai
OPENAI_API_KEY=your_key_here
OPENAI_MODEL=gpt-4.1-mini
```

### Option 2: inside exaBase LiteLLM

```env
LLM_PROVIDER=litellm
LITELLM_URL=http://fpgen-ai-litellm:8080
LITELLM_MASTER_KEY=your_litellm_admin_or_gateway_key
OPENAI_MODEL=gpt-4.1-mini
```

Common settings:

```env
ENABLE_OPENAI_SYNTHESIS=true
DEFAULT_CURRENCY=USD
DEFAULT_OVERHEAD_PCT=0.10
TRAVEL_UNIT_COST_USD=6000
FPGEN_STORAGE_DIR=/app/data/proposals
FPGEN_REFERENCE_PROJECTS=/app/data/reference_projects.json
```

## Run locally without Docker

Yes, this can be done with Python only.

You do not need Docker and you do not need exaBase for the first local test.

Create a project-local virtual environment first:

```powershell
Set-Location "C:\Users\03669\Desktop\Trud\matls\created\FP-GEN\exaBase\fpgen_mvp"
& "C:\Users\03669\Desktop\Trud\matls\created\PQQ-Automation\python\python.exe" -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
```

When you run `uvicorn`, that is your local HTTP server.

In other words:

- Python starts the API server
- your browser talks to that local server
- the local server calls OpenAI directly

```powershell
Set-Location "C:\Users\03669\Desktop\Trud\matls\created\FP-GEN\exaBase\fpgen_mvp"
.\.venv\Scripts\Activate.ps1
python -m uvicorn app.main:app --host 0.0.0.0 --port 8080
```

Built-in test UI:

- open `http://127.0.0.1:8080/`
- or open `http://127.0.0.1:8080/ui`

Built-in file workflow:

- upload an RFP file in `.txt`, `.md`, `.json`, `.pdf`, or `.docx`
- use `Parse RFP File` to extract text and parse it through the OpenAI-backed parser
- use `Generate From File` to build and save the proposal directly from the uploaded file
- use `Export INDD` after generation to create a local `.indd` file via Adobe InDesign

Local INDD export prerequisites:

- Windows
- Adobe InDesign installed locally
- `pywin32` available in the Python environment

The current export uses the packaged commercial `.indd` template at `examples/INDD/cp/I. Commercial.indd` when available, fills the stable text placeholders it can identify, and appends generated proposal pages into that template. If the template is unavailable, it falls back to the simple generated document export.

Smoke check:

```powershell
Invoke-RestMethod -Method Get -Uri "http://127.0.0.1:8080/health"
```

File-based parse smoke check:

```powershell
$form = @{ 
	rfp_file = Get-Item .\examples\sample_rfp.txt
	project_hint = "HCMC Mixed-Use Master Plan"
}
Invoke-RestMethod -Method Post -Uri "http://127.0.0.1:8080/v1/parse/file" -Form $form
```

File-based proposal generation:

```powershell
$feeJson = Get-Content .\examples\generate_request.json -Raw | ConvertFrom-Json
$genForm = @{
	rfp_file = Get-Item .\examples\sample_rfp.txt
	fee_input_json = ($feeJson.fee_input | ConvertTo-Json -Depth 8 -Compress)
	selected_reference_ids_json = "[]"
	overrides_json = (@{
		project_name = "HCMC Mixed-Use Master Plan"
		client_name = "Example Client"
		location = "Ho Chi Minh City, Vietnam"
		project_type = "Master Plan"
	} | ConvertTo-Json -Compress)
}
Invoke-RestMethod -Method Post -Uri "http://127.0.0.1:8080/v1/proposals/generate/file" -Form $genForm
```

Local INDD export:

```powershell
Invoke-RestMethod -Method Post -Uri "http://127.0.0.1:8080/v1/proposals/<proposal_id>/export/indd"
```

## How to create the image

This is the important part.

exaBase will not build this folder automatically.

You must build an image yourself, push it to a registry that exaBase can pull from, then paste that image URL into the `api` sideapp in `../FP-GEN_exaBase_import.json`.

So there are two separate stages:

### Stage A: local proof of concept

Use Python only.

You can fully test the backend and the built-in UI locally with:

- Python
- `uvicorn`
- direct OpenAI API

No Docker is required for that stage.

### Stage B: inside exaBase deployment

If you want the backend to run inside exaBase as a sideapp, the current import expects a pullable container image for the `api` node.

So for exaBase deployment, Docker or some equivalent image-build process is still needed unless you later switch to an exaBase-native code-edit/base-container workflow.

Right now, this repository is prepared for the image-based path.

### Build context

Build from this folder only:

`C:\Users\03669\Desktop\Trud\matls\created\FP-GEN\exaBase\fpgen_mvp`

Do not build from the root FP-GEN folder — it contains the legacy VB/Excel project and is much larger than needed.

### How the image is built

Local Docker is not required. The image is built and published automatically via GitHub Actions.

### GitHub Actions build (current approach)

Repository preparation:

- put the `exaBase` folder into a GitHub repository
- keep the current folder structure so that `fpgen_mvp/Dockerfile` remains valid
- commit `.github/workflows/build-fpgen-image.yml`

What the workflow does:

- runs on `push` to `main` or `master`
- can also be started manually with `workflow_dispatch`
- builds from `fpgen_mvp/`
- publishes to GitHub Container Registry as `ghcr.io/a1exkr/fpgen-mvp-api`

Repo: https://github.com/A1exKr/fp-eb

What you need in GitHub:

- Actions enabled for the repository
- Packages enabled for the repository or organization
- permission for `GITHUB_TOKEN` to write packages

Published image tags:

```text
ghcr.io/a1exkr/fpgen-mvp-api:latest
ghcr.io/a1exkr/fpgen-mvp-api:main
ghcr.io/a1exkr/fpgen-mvp-api:sha-<commit>
```

exaBase pulls from the published image. The GHCR package is kept **private** (not visible to anyone else). Because it is private, exaBase Studio must authenticate to GHCR with a `read:packages` pull credential configured as an image pull secret in the workspace — see `DEPLOY.md`, Option A, Step 6. The recommended build path is now a local Docker Engine (WSL2) build; GitHub Actions remains a fallback and also publishes a private package.

## What to paste into exaBase

The image is already set in all import files:

```text
ghcr.io/a1exkr/fpgen-mvp-api:latest
```

The `api` sideapp receives these envs (all already configured in `FP-GEN_exaBase_import_v2.json`):

### Required

- `LLM_PROVIDER=openai`
- `OPENAI_API_KEY=${{SECRET.OPENAI_API_KEY}}`
- `OPENAI_MODEL=gpt-4.1-mini`
- `FPGEN_ENABLE_INDD_EXPORT=false`

### Canvas variables (set before import)

- `DEFAULT_CURRENCY` (default: USD)
- `DEFAULT_OVERHEAD_PCT` (default: 0.10)
- `TRAVEL_UNIT_COST_USD` (default: 6000)
- `FPGEN_STORAGE_DIR` (default: /app/data/proposals)
- `FPGEN_REFERENCE_PROJECTS_PATH` (default: /app/data/reference_projects.json)

### Not required

- `APP_URL` — only needed for Keycloak auth stack, not used in v2 canvas
- `AUTH_URL` — only needed for Keycloak auth stack, not used in v2 canvas
- `CONFIG_CRYPTO_KEY` — only needed for Config Center sideapp, not used in v2 canvas
- `LITELLM_URL` / `LITELLM_MASTER_KEY` — using direct OpenAI instead
- `DB_USER` / `DB_PASSWORD` — no PostgreSQL in v2 canvas

The sideapp also needs persistent storage mounted at `/app/data` for proposal JSON output.

Important for exaBase sideapps:

- the container runs on Linux, so Adobe InDesign automation is not available there
- INDD export is disabled by default (`FPGEN_ENABLE_INDD_EXPORT=false`)
- proposal generation, review, JSON storage, and download endpoints are all available

## exaBase import variables

For `FP-GEN_exaBase_import_v2.json` (recommended):

| Variable | Default | Required |
|---|---|---|
| `DEFAULT_CURRENCY` | USD | yes |
| `DEFAULT_OVERHEAD_PCT` | 0.10 | yes |
| `TRAVEL_UNIT_COST_USD` | 6000 | yes |
| `FPGEN_API_URL` | http://fpgen-app-api:8080 | yes (internal only) |
| `FPGEN_STORAGE_DIR` | /app/data/proposals | yes |
| `FPGEN_REFERENCE_PROJECTS_PATH` | /app/data/reference_projects.json | yes |
| `APP_URL` | — | **not needed** |
| `AUTH_URL` | — | **not needed** |
| `CONFIG_CRYPTO_KEY` | — | **not needed** |
| `FPGEN_LITELLM_URL` | — | **not needed** |

## exaBase import step by step

### Available import files

| File | What it contains | When to use |
|---|---|---|
| `../FP-GEN_exaBase_import_v2.json` | API + nginx proxy. No auth, no DB. | **Recommended.** Use this. |
| `../FP-GEN_exaBase_import_mvp_public.json` | API + nginx proxy (original version) | Legacy, superseded by v2 |
| `../FP-GEN_exaBase_import_mvp.json` | API only, no proxy | Minimal, no public URL |

### 1. Choose the import file

Recommended import:

- `../FP-GEN_exaBase_import_v2.json`

This canvas contains only the two components needed: the FastAPI runtime and an nginx reverse proxy. No auth stack, no PostgreSQL, no LiteLLM.

### 2. Decide which image tag to deploy

The checked-in import files currently point to:

```text
ghcr.io/a1exkr/fpgen-mvp-api:latest
```

That is acceptable for a first import.

If you want a pinned image for a stable deployment, replace it with the published SHA tag before import, for example:

```text
ghcr.io/a1exkr/fpgen-mvp-api:sha-97f661b
```

### 3. Prepare exaBase secrets

Create this secret in exaBase before deploy:

- `OPENAI_API_KEY`

The current import files already expect:

```text
${{SECRET.OPENAI_API_KEY}}
```

### 4. Import the canvas JSON

In exaBase Studio:

1. Create or open the target workspace.
2. Choose the canvas import action.
3. Import `../FP-GEN_exaBase_import_mvp.json` first.
4. Confirm the `fpgen-mvp-api` sideapp appears.

### 5. Verify the sideapp image and env values

For the v2 import, the `api` sideapp should already contain:

- `image=ghcr.io/a1exkr/fpgen-mvp-api:latest`
- `LLM_PROVIDER=openai`
- `OPENAI_API_KEY=${{SECRET.OPENAI_API_KEY}}`
- `OPENAI_MODEL=gpt-4.1-mini`
- `DEFAULT_CURRENCY=$[[DEFAULT_CURRENCY]]`
- `DEFAULT_OVERHEAD_PCT=$[[DEFAULT_OVERHEAD_PCT]]`
- `TRAVEL_UNIT_COST_USD=$[[TRAVEL_UNIT_COST_USD]]`
- `FPGEN_STORAGE_DIR=$[[FPGEN_STORAGE_DIR]]`
- `FPGEN_REFERENCE_PROJECTS=$[[FPGEN_REFERENCE_PROJECTS_PATH]]`
- `FPGEN_ENABLE_INDD_EXPORT=false`

Do not enable INDD export. The container runs on Linux and cannot automate desktop Adobe InDesign.

### 6. Confirm storage is mounted

The sideapp needs persistent storage at:

```text
/app/data
```

This is required so proposal JSON output can persist under:

```text
/app/data/proposals
```

### 7. Deploy the canvas

After import:

1. Save the canvas.
2. Start deployment.
3. Wait until the `fpgen-mvp-api` sideapp is in a healthy running state.

Do not debug UI flow yet if the sideapp itself is not healthy. First prove the API container is up.

### 8. Find the public URL and run the first health check

In the exaBase canvas, click the endpoint node (port 8080 / uri `/`) at the top of the workspace. The assigned public HTTPS URL appears in the panel.

All traffic enters through the nginx proxy pod, which forwards to the API pod internally.

Test these once the URL is known:

```text
GET <endpoint-url>/health
GET <endpoint-url>/ui
GET <endpoint-url>/docs
```

Expected health response:

```json
{"status":"ok","capabilities":{"indd_export":{"enabled":false,"reason":"..."}}}
```

`indd_export.enabled=false` is expected inside exaBase.

### 9. Run the first functional checks

After health succeeds, test these endpoints in order:

1. `POST /v1/parse`
2. `POST /v1/proposals/generate`
3. `GET /v1/proposals/{proposal_id}`

The goal is to verify:

- OpenAI access works through the provided secret
- proposal JSON is saved to storage
- the generated proposal can be loaded back

### 10. If you need more components later

The v2 canvas is intentionally minimal — API + nginx only.

If you later need to add auth, Config Center, or vector search, see `../FP-GEN_exaBase_import.json` (the original full canvas). Before deploying it you will need:

- `APP_URL` — public HTTPS URL assigned by exaBase to the app workspace
- `AUTH_URL` — public HTTPS URL assigned by exaBase to the auth workspace
- `CONFIG_CRYPTO_KEY` — a secure random string (any 32+ char random value)
- `DB_USER` / `DB_PASSWORD` — for PostgreSQL
- `LITE_LLM_ADMIN_PASSWORD` secret — for LiteLLM gateway

For MVP testing, none of these are needed.

### 11. First failure points to check

If import succeeds but deploy fails, check these first:

- the sideapp image tag exists in GHCR and the workspace has a valid GHCR pull credential (private image)
- `OPENAI_API_KEY` secret exists and is spelled exactly the same as in the JSON
- `/app/data` storage is attached
- the sideapp logs show the container started and bound to port `8080`
- `/health` is reachable before trying any proposal generation call

Only after the API sideapp is stable should you spend time recreating more of the legacy VB behavior inside exaBase.

## If you want to reuse more of the VB project later

Do it selectively.

Good candidates to migrate next:

- prompts
- reference data
- fee tables
- template mappings
- export rules

Bad candidates for the first container image:

- the whole root project folder
- Excel macros
- portable tools like PDFTK or Tesseract bundles
- old output folders

## Current limitations

- no DOCX generation yet
- no PDF generation yet
- no PostgreSQL persistence yet
- no direct import of the VB UI itself
- JSON and Markdown are the current output artifacts
- detailed page-specific commercial field mapping inside the packaged InDesign template is still partial; stable placeholders are filled automatically and generated proposal pages are appended for review

## Quick smoke test after deploy

Replace `<endpoint-url>` with the public URL shown on the endpoint node in the exaBase canvas.

```powershell
# Health check
Invoke-RestMethod -Method Get -Uri "https://<endpoint-url>/health"

# Browser UI
Start-Process "https://<endpoint-url>/ui"

# FastAPI docs
Start-Process "https://<endpoint-url>/docs"

# Parse test
$parseBody = @{ rfp_text = "RFP for mixed-use master plan in Ho Chi Minh City with 16-week timeline." } | ConvertTo-Json
Invoke-RestMethod -Method Post -Uri "https://<endpoint-url>/v1/parse" -ContentType "application/json" -Body $parseBody

# Proposal generation test
$genBody = Get-Content .\examples\generate_request.json -Raw
Invoke-RestMethod -Method Post -Uri "https://<endpoint-url>/v1/proposals/generate" -ContentType "application/json" -Body $genBody
```

For local dev only (without exaBase):

```powershell
Invoke-RestMethod -Method Get -Uri "http://127.0.0.1:8080/health"
```

Expected results:

- health returns `ok`
- parse returns a `parsed` object
- generate returns `proposal_id` and `proposal.markdown`
