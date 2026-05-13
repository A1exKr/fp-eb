# FP-GEN Deployment Guide

## How to deploy a new image to exaBase

**Always use Option B — GitHub Actions.** Docker Desktop does not work on this machine (no WSL2, no admin rights). Never try to build or push the Docker image locally.

---

## Option B: Push to GitHub → GitHub Actions builds + pushes

The repo `https://github.com/verv0lk/fp-eb` has a GitHub Actions workflow that automatically:
1. Builds the Docker image from the Dockerfile
2. Tags it as `ghcr.io/verv0lk/fpgen-mvp-api:latest`
3. Pushes it to GitHub Container Registry (GHCR)

### Step 1 — Commit and push changes

In a terminal from the repo root:

```powershell
git add .
git commit -m "deploy: <brief description of changes>"
git push origin main
```

GitHub Actions will trigger automatically on push to `main`.

### Step 2 — Verify the build

Go to: https://github.com/verv0lk/fp-eb/actions

Wait for the workflow run to complete (green tick). If it fails, check the logs there.

### Step 3 — Restart the pod in exaBase

In exaBase Studio, go to workspace `fpgen-app` and restart the `fpgen-app-api` pod to pull the new `:latest` image.

---

## Files to include when deploying UI / backend changes

| Changed area | Files |
|---|---|
| Frontend UI | `app/static/index.html`, `app/static/review.html` |
| Backend routes | `app/main.py` |
| Config / env | `app/config.py` |
| Services | `app/services/*.py` |
| Data files | `data/personnel.json`, `data/team_presets.json`, `data/reference_projects.json` |
| Dependencies | `requirements.txt` |

---

## Environment variables required in exaBase pod

Set these in the `fpgen-app-api` pod environment (NOT committed to the repo):

| Variable | Description |
|---|---|
| `OPENAI_API_KEY` | OpenAI API key |
| `LLM_PROVIDER` | `openai` |
| `OPENAI_MODEL` | `gpt-4.1-mini` |
| `ENABLE_OPENAI_SYNTHESIS` | `true` |
| `OPENAI_SSL_VERIFY` | `false` (corporate proxy) |
| `DEFAULT_CURRENCY` | `USD` |
| `DEFAULT_OVERHEAD_PCT` | `0.10` |
| `TRAVEL_UNIT_COST_USD` | `6000` |
| `FPGEN_ENABLE_INDD_EXPORT` | `false` (Linux — no COM) |

---

## Image details

- **Registry**: `ghcr.io/verv0lk/fpgen-mvp-api`
- **Tags**: `:latest` (always), optionally `:YYYY-MM-DD` for versioned rollback
- **Source repo**: https://github.com/verv0lk/fp-eb
- **Base image**: `python:3.12-slim`
- **Port**: 8080 (Uvicorn)
