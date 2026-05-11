# FP-GEN MVP for exaBase

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

Do not build from:

`C:\Users\03669\Desktop\Trud\matls\created\FP-GEN`

Why:

- that root contains the legacy VB/Excel project
- it is much larger than needed
- it mixes runtime code with historical assets
- the current Dockerfile only expects the Python MVP folder structure

### Step-by-step image build

1. Open Docker Desktop or make sure Docker Engine is running.
2. Open PowerShell.
3. Move into the MVP folder:

```powershell
Set-Location "C:\Users\03669\Desktop\Trud\matls\created\FP-GEN\exaBase\fpgen_mvp"
```

4. Build the image:

```powershell
docker build -t fpgen-mvp-api:0.1.0 .
```

5. Confirm the image exists:

```powershell
docker images fpgen-mvp-api
```

6. Test-run it locally:

```powershell
docker run --rm -p 8080:8080 fpgen-mvp-api:0.1.0
```

7. In another terminal, test the health endpoint:

```powershell
Invoke-RestMethod -Method Get -Uri "http://127.0.0.1:8080/health"
```

8. Stop the local container after the test.

### Tag and push the image

After the local test passes, tag it for your registry.

Generic form:

```powershell
docker tag fpgen-mvp-api:0.1.0 <your-registry>/fpgen-mvp-api:0.1.0
```

Example:

```powershell
docker tag fpgen-mvp-api:0.1.0 registry.example.com/fpgen/fpgen-mvp-api:0.1.0
```

Log in:

```powershell
docker login <your-registry>
```

Push:

```powershell
docker push <your-registry>/fpgen-mvp-api:0.1.0
```

The final value exaBase needs is the pushed image URL, for example:

```text
registry.example.com/fpgen/fpgen-mvp-api:0.1.0
```

### Alternative: build the image in GitHub Actions

If you cannot run Docker locally, you can build and publish the image from GitHub instead.

Repository preparation:

- put the `exaBase` folder into a GitHub repository
- keep the current folder structure so that `fpgen_mvp/Dockerfile` remains valid
- commit `.github/workflows/build-fpgen-image.yml`

What the workflow does:

- runs on `push` to `main` or `master`
- can also be started manually with `workflow_dispatch`
- builds from `fpgen_mvp/`
- publishes to GitHub Container Registry as `ghcr.io/<owner>/fpgen-mvp-api`

What you need in GitHub:

- Actions enabled for the repository
- Packages enabled for the repository or organization
- permission for `GITHUB_TOKEN` to write packages

After the first workflow run, the exaBase image value can be:

```text
ghcr.io/<github-owner>/fpgen-mvp-api:latest
```

or a SHA tag produced by the workflow.

Important:

- this workspace is not currently a git repository, so the workflow will not run until the files are placed in a real GitHub repo
- exaBase still pulls and runs the resulting image in its own environment; GitHub Actions only builds and publishes it

## What to paste into exaBase

In the full canvas file:

- `../FP-GEN_exaBase_import.json`

The `api` sideapp image field is intentionally a placeholder.

Replace it with your actual pushed image URL.

The sideapp should then receive these envs:

### Required in LiteLLM mode

- `LLM_PROVIDER=litellm`
- `LITELLM_URL=http://fpgen-ai-litellm:8080`
- `LITELLM_MASTER_KEY=<secret>`
- `OPENAI_MODEL=gpt-4.1-mini`

### Required in direct OpenAI mode

- `LLM_PROVIDER=openai`
- `OPENAI_API_KEY=<secret>`
- `OPENAI_MODEL=gpt-4.1-mini`
- `FPGEN_ENABLE_INDD_EXPORT=false`

### Common runtime envs

- `DEFAULT_CURRENCY`
- `DEFAULT_OVERHEAD_PCT`
- `TRAVEL_UNIT_COST_USD`
- `FPGEN_STORAGE_DIR`
- `FPGEN_REFERENCE_PROJECTS`

The sideapp also needs persistent storage mounted for proposal JSON output.

Important for exaBase sideapps:

- the container runs on Linux, so Adobe InDesign automation is not available there
- the provided Dockerfile and import files now disable INDD export by default for exaBase deployment
- proposal generation, review, JSON storage, and download endpoints remain deployable inside exaBase

## exaBase import variables

If you import the full canvas, set these variables:

- `FPGEN_API_URL=http://fpgen-mvp-api:8080`
- `FPGEN_LITELLM_URL=http://fpgen-ai-litellm:8080`
- `FPGEN_STORAGE_DIR=/app/data/proposals`
- `FPGEN_REFERENCE_PROJECTS_PATH=/app/data/reference_projects.json`
- `APP_URL=https://<your-exabase-app-url>/`
- `AUTH_URL=https://<your-exabase-auth-url>/`
- `CONFIG_CRYPTO_KEY=<secure-random-key>`
- `DEFAULT_CURRENCY=USD`
- `DEFAULT_OVERHEAD_PCT=0.10`
- `TRAVEL_UNIT_COST_USD=6000`

## Recommended first deployment path

1. Keep your VB UI as reference only.
2. Run the backend locally with Python only.
3. Open the built-in UI at `http://127.0.0.1:8080/` and test parse/generate end to end.
4. Adjust prompts, fee rules, and reference data until the behavior is acceptable.
5. Only then build the image from this folder.
6. Push the image to a registry.
7. Import `../FP-GEN_exaBase_import.json`.
8. Replace the `api` sideapp image with your pushed image.
9. Set secrets and env vars.
10. Deploy.
11. Wait until Pods are `Running`.
12. Test `/health`.
13. Test `/v1/proposals/generate`.

Only after that should you start porting or recreating the VB UI behavior inside exaBase.

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

```powershell
Invoke-RestMethod -Method Get -Uri "http://127.0.0.1:8080/health"

$parseBody = @{ rfp_text = "RFP for mixed-use master plan in Ho Chi Minh City with 16-week timeline." } | ConvertTo-Json
Invoke-RestMethod -Method Post -Uri "http://127.0.0.1:8080/v1/parse" -ContentType "application/json" -Body $parseBody

$genBody = Get-Content .\examples\generate_request.json -Raw
Invoke-RestMethod -Method Post -Uri "http://127.0.0.1:8080/v1/proposals/generate" -ContentType "application/json" -Body $genBody
```

Expected results:

- health returns `ok`
- parse returns a `parsed` object
- generate returns `proposal_id` and `proposal.markdown`
