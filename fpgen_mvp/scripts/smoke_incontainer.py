"""In-container smoke check for local Docker (Route C) validation.

Runs INSIDE the api container (python + httpx are present) against
``http://localhost:8080`` so it does not depend on Windows<->WSL port
forwarding. Complements ``smoke_live.ps1`` by exercising the auth-disabled
admin surface — master-data writes and CV/experience upload — plus the JSX
export fix. Exits non-zero on any failure.

Usage (from the workspace root, via WSL Docker):
    docker cp fpgen_mvp/scripts/smoke_incontainer.py fpgen-local:/tmp/smoke.py
    docker exec fpgen-local python /tmp/smoke.py
"""
import sys

import httpx

BASE = "http://localhost:8080"

_passed = 0
_failed = 0


def check(name: str, ok: bool, detail: str = "") -> None:
    global _passed, _failed
    if ok:
        _passed += 1
        print(f"  [PASS] {name}")
    else:
        _failed += 1
        print(f"  [FAIL] {name}{(' -> ' + detail) if detail else ''}")


def main() -> int:
    client = httpx.Client(base_url=BASE, timeout=60.0)

    # --- health / identity (auth disabled => dev-admin) ---
    r = client.get("/health")
    check("GET /health -> 200", r.status_code == 200, str(r.status_code))
    cap = r.json().get("capabilities", {}).get("indd_export") if r.status_code == 200 else None
    enabled = cap.get("enabled") if isinstance(cap, dict) else cap
    check("health.indd_export disabled", enabled is False, str(cap))

    r = client.get("/v1/me")
    check("GET /v1/me -> 200 (auth disabled)", r.status_code == 200, r.text[:120])
    check("me.user present", r.status_code == 200 and bool(r.json().get("user")))

    # --- seeded master-data reads ---
    for path in ("/v1/personnel", "/v1/team-presets", "/v1/reference-projects", "/v1/assets"):
        rr = client.get(path)
        check(f"GET {path} -> 200", rr.status_code == 200, str(rr.status_code))

    # --- generate -> JSX export (validates the str.format brace fix) ---
    gen = client.post(
        "/v1/proposals/generate",
        json={
            "rfp_text": "Client requests a mixed-use master plan in Ho Chi Minh City over 16 weeks, with concept options and a final master plan report.",
            "overrides": {"project_name": "HCMC Master Plan", "client_name": "Example Client"},
        },
    )
    check("POST /v1/proposals/generate -> 200", gen.status_code == 200, gen.text[:160])
    proposal_id = gen.json().get("proposal_id", "") if gen.status_code == 200 else ""

    jsx = client.post(
        f"/v1/proposals/{proposal_id}/export/jsx",
        json={"cv_assignments": {}, "experience_ids": [], "template_id": "commercial"},
    )
    ctype = jsx.headers.get("content-type", "")
    check(
        "POST /export/jsx -> 200 zip (JSX fix)",
        jsx.status_code == 200 and "zip" in ctype,
        f"status={jsx.status_code} type={ctype} body={jsx.text[:140] if jsx.status_code != 200 else ''}",
    )
    check(
        "jsx zip non-empty + PK header",
        jsx.status_code == 200 and len(jsx.content) > 200 and jsx.content[:2] == b"PK",
    )

    # --- admin write + experience upload (auth disabled) ---
    r = client.post(
        "/v1/admin/personnel",
        json={"id": "persist-check", "name": "Persist Check", "title": "QA", "roles": ["Architect"]},
    )
    check("POST /v1/admin/personnel -> 200", r.status_code == 200, r.text[:160])

    up = client.post(
        "/v1/admin/assets/upload",
        data={"asset_id": "exp-smoke", "kind": "experience", "reference_project_id": "rp-x"},
        files={"file": ("exp.txt", b"Reference experience: 42ha riverside mixed-use master plan.", "text/plain")},
    )
    check("POST /v1/admin/assets/upload (experience) -> 200", up.status_code == 200, up.text[:200])
    if up.status_code == 200:
        check("experience upload has storage_ref", bool(up.json().get("storage_ref")), str(up.json()))

    print(f"\nPASS={_passed} FAIL={_failed}")
    return 1 if _failed else 0


def verify_persist() -> int:
    """Read-only: confirm records written by a previous run survived an api
    container recreate (i.e. they live in PostgreSQL, not container-local state)."""
    client = httpx.Client(base_url=BASE, timeout=30.0)
    personnel = client.get("/v1/admin/personnel").json()
    assets = client.get("/v1/admin/assets").json()
    check(
        f"persist-check personnel survived api recreate (n={len(personnel)})",
        any(x.get("id") == "persist-check" for x in personnel),
    )
    check(
        f"exp-smoke asset row survived api recreate (n={len(assets)})",
        any(x.get("id") == "exp-smoke" for x in assets),
    )
    print(f"\nPASS={_passed} FAIL={_failed}")
    return 1 if _failed else 0


def dump_jsx() -> int:
    """Generate a proposal, export the JSX bundle, and print its structure."""
    import io
    import zipfile

    client = httpx.Client(base_url=BASE, timeout=60.0)
    gen = client.post(
        "/v1/proposals/generate",
        json={
            "rfp_text": "Client requests a commercial mixed-use proposal in Bermuda over 16 weeks with concept options and a final report.",
            "overrides": {"project_name": "Morgans Point Development", "client_name": "MPDC"},
        },
    )
    pid = gen.json().get("proposal_id", "") if gen.status_code == 200 else ""
    r = client.post(
        f"/v1/proposals/{pid}/export/jsx",
        json={"cv_assignments": {}, "experience_ids": [], "template_id": "commercial"},
    )
    print(f"export/jsx -> status={r.status_code} type={r.headers.get('content-type')} bytes={len(r.content)}")
    if r.status_code != 200:
        print(r.text[:400])
        return 1
    zf = zipfile.ZipFile(io.BytesIO(r.content))
    print("ZIP entries:", zf.namelist())
    jsx = zf.read("assemble_proposal.jsx").decode("utf-8", "replace")
    print(f"JSX length: {len(jsx)}")
    print("----- JSX head -----")
    for line in jsx.splitlines()[:8]:
        print(line[:160])
    with open("/tmp/fpgen_bundle.zip", "wb") as fh:
        fh.write(r.content)
    print("saved /tmp/fpgen_bundle.zip")
    return 0


if __name__ == "__main__":
    if "--verify-persist" in sys.argv:
        sys.exit(verify_persist())
    if "--dump-jsx" in sys.argv:
        sys.exit(dump_jsx())
    sys.exit(main())
