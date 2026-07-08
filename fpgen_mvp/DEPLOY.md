# FP-GEN Deployment Guide

## How to deploy a new image to exaBase

The image `ghcr.io/a1exkr/fpgen-mvp-api` is a **private** GHCR package. It is **not** visible to anyone without access to the `A1exKr` account. Because it is private, exaBase Studio must authenticate to GHCR with a pull credential (see Option A, Step 6).

Two ways to build and push:

- **Option A — Local build with Docker Engine (WSL2).** Use this on a machine with admin rights + WSL2. Full control, no CI wait.
- **Option B — GitHub Actions.** Fallback. Also pushes a private package by default.

> Keep the GHCR package set to **Private** in both cases. Never switch it to Public.

---

## Option A: Local build with Docker Engine (WSL2)

### Step 0 — One-time: install WSL2 + Docker Engine

Run PowerShell **as Administrator**:

```powershell
wsl --install
```

Reboot when prompted. Then, inside the default Ubuntu (WSL) shell, install Docker Engine (not Docker Desktop):

```bash
# Ubuntu (WSL2)
sudo apt-get update
sudo apt-get install -y ca-certificates curl
sudo install -m 0755 -d /etc/apt/keyrings
sudo curl -fsSL https://download.docker.com/linux/ubuntu/gpg -o /etc/apt/keyrings/docker.asc
sudo chmod a+r /etc/apt/keyrings/docker.asc
echo "deb [arch=$(dpkg --print-architecture) signed-by=/etc/apt/keyrings/docker.asc] https://download.docker.com/linux/ubuntu $(. /etc/os-release && echo $VERSION_CODENAME) stable" | sudo tee /etc/apt/sources.list.d/docker.list > /dev/null
sudo apt-get update
sudo apt-get install -y docker-ce docker-ce-cli containerd.io docker-buildx-plugin
sudo service docker start
sudo usermod -aG docker $USER   # log out/in of WSL to take effect
```

Verify: `docker version` should print both Client and Server.

### Step 1 — Get the code into WSL

From the WSL shell, clone (or `cd` into) the repo, then move into the build context:

```bash
git clone https://github.com/A1exKr/fp-eb.git
cd fp-eb/fpgen_mvp
```

> The Dockerfile does `COPY data ./data`. Confirm the `data/` folder with `personnel.json`, `team_presets.json`, `reference_projects.json`, and `assets_registry.json` is present before building, or the build fails.

### Step 2 — Log in to GHCR

Create a GitHub Personal Access Token (classic) with scope `write:packages` (this also grants `read:packages`). Then:

```bash
echo "<YOUR_PAT>" | docker login ghcr.io -u A1exKr --password-stdin
```

> Do not paste the token into chat. Enter it directly in the terminal.

### Step 3 — Build the image

```bash
docker build -t ghcr.io/a1exkr/fpgen-mvp-api:latest \
             -t ghcr.io/a1exkr/fpgen-mvp-api:$(date +%Y-%m-%d) .
```

### Step 4 — Push to GHCR

```bash
docker push ghcr.io/a1exkr/fpgen-mvp-api:latest
docker push ghcr.io/a1exkr/fpgen-mvp-api:$(date +%Y-%m-%d)
```

### Step 5 — Confirm the package is Private

In the browser: `https://github.com/users/A1exKr/packages/container/fpgen-mvp-api/settings`
→ **Danger Zone / Change visibility** must show **Private**. A freshly created GHCR package is Private by default — leave it that way.

### Step 6 — Give exaBase Studio a pull credential (required for private images)

Because the image is private, the Studio cluster cannot pull it anonymously. Provide GHCR credentials in Studio:

1. Create a GitHub PAT (classic) with **only** `read:packages` (dedicated, minimal-scope token for the cluster).
2. In exaBase Studio, add a **container registry credential / image pull secret** for registry `ghcr.io`:
   - **Registry**: `ghcr.io`
   - **Username**: `A1exKr`
   - **Password/Token**: the `read:packages` PAT
3. Attach that pull secret to the `api` sideapp (workspace `fpgen-app`) so it is used when pulling `ghcr.io/a1exkr/fpgen-mvp-api:latest`.

> If your Studio tenant does not expose an image-pull-secret / private-registry option, ask ExaWizards support to register the GHCR credential for the workspace. Without it, a private image cannot be pulled.

### Step 7 — Redeploy in exaBase

In exaBase Studio, go to workspace `fpgen-app` and restart the `fpgen-app-api` pod to pull the new `:latest` image.

---

## Option B: Push to GitHub → GitHub Actions builds + pushes

The repo `https://github.com/A1exKr/fp-eb` has a GitHub Actions workflow that automatically:
1. Builds the Docker image from the Dockerfile
2. Tags it as `ghcr.io/a1exkr/fpgen-mvp-api:latest`
3. Pushes it to GitHub Container Registry (GHCR) as a **private** package

### Step 1 — Commit and push changes

In a terminal from the repo root:

```powershell
git add .
git commit -m "deploy: <brief description of changes>"
git push origin main
```

GitHub Actions will trigger automatically on push to `main`.

### Step 2 — Verify the build

Go to: https://github.com/A1exKr/fp-eb/actions

Wait for the workflow run to complete (green tick). If it fails, check the logs there. Confirm the package stays **Private** (Step 5 above).

### Step 3 — Restart the pod in exaBase

In exaBase Studio, go to workspace `fpgen-app` and restart the `fpgen-app-api` pod to pull the new `:latest` image. The pull credential from Option A, Step 6 must be in place.

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

- **Registry**: `ghcr.io/a1exkr/fpgen-mvp-api`
- **Tags**: `:latest` (always), optionally `:YYYY-MM-DD` for versioned rollback
- **Source repo**: https://github.com/A1exKr/fp-eb
- **Base image**: `python:3.12-slim`
- **Port**: 8080 (Uvicorn)

---

## Troubleshooting: 502 Bad Gateway right after deploy

**Symptom:** the endpoint returns `502 Bad Gateway` (nginx) even though both pods show Running. The `nginx` pod log shows `connect() failed (111: Connection refused) while connecting to upstream`.

**Cause:** startup race. nginx resolves `fpgen-app-api:8080` once at startup and caches the pod IP. If nginx starts before the `api` pod is ready, it caches a dead/old upstream and keeps returning 502 even after `api` comes up.

**Fix (restart nginx only, leaving `api` running):**
1. Open the canvas in **編集 (Edit)** mode and select the **FP-GEN Reverse Proxy (nginx)** node.
2. Set **replicas 1 → 0**, click **デプロイ (Deploy)**.
3. Set **replicas 0 → 1**, click **デプロイ** again.
4. When the new nginx pod is Running, refresh the endpoint URL.

Because `api` is left untouched, the fresh nginx pod resolves the live `api` and the 502 clears. This is the accepted operational workaround; a permanent fix (runtime DNS re-resolution) would require hardcoding the cluster CoreDNS resolver IP in the nginx injection.
