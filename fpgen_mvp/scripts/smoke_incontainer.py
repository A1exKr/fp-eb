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
import json
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
            "fee_input": {
                "roles": [
                    {"role": "Urban Planner", "rate": 180, "hours_by_phase": {"Kick-off": 30, "Concept Development": 80, "Finalization": 40}},
                    {"role": "Architect", "rate": 200, "hours_by_phase": {"Kick-off": 20, "Concept Development": 70, "Finalization": 35}},
                ],
                "overhead_pct": 0.1,
                "subconsultants": [{"name": "Local Traffic Consultant", "fee": 12000, "included_in_lump_sum": True}],
                "travel": {"trips": 2, "people_per_trip": 2, "unit_cost": 6000, "include_in_lump_sum": True},
                "misc_reimbursables": 3000,
            },
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


def regen_smoke() -> int:
    """Phase 1 (Option C+): fee_input persistence + the three regeneration endpoints."""
    client = httpx.Client(base_url=BASE, timeout=120.0)
    fee_input = {
        "roles": [
            {"role": "Urban Planner", "rate": 180, "hours_by_phase": {"Kick-off": 30, "Concept Development": 80}},
            {"role": "Architect", "rate": 200, "hours_by_phase": {"Kick-off": 20, "Concept Development": 70}},
        ],
        "overhead_pct": 0.1,
        "subconsultants": [{"name": "Local Traffic Consultant", "fee": 12000, "included_in_lump_sum": True}],
        "travel": {"trips": 2, "people_per_trip": 2, "unit_cost": 6000, "include_in_lump_sum": True},
        "misc_reimbursables": 3000,
    }
    gen = client.post(
        "/v1/proposals/generate",
        json={
            "rfp_text": "Client requests a mixed-use master plan in Ho Chi Minh City over 16 weeks, with concept options and a final master plan report.",
            "fee_input": fee_input,
            "overrides": {"project_name": "Regen Smoke", "client_name": "Example Client"},
        },
    )
    check("generate -> 200", gen.status_code == 200, gen.text[:200])
    if gen.status_code != 200:
        print(f"\nPASS={_passed} FAIL={_failed}")
        return 1
    pid = gen.json()["proposal_id"]
    payload = gen.json()["proposal"]

    # 1. fee_input is now persisted with the proposal
    check("payload carries fee_input", isinstance(payload.get("fee_input"), dict), str(payload.get("fee_input"))[:120])
    stored = client.get(f"/v1/proposals/{pid}").json()["proposal"]
    check("fee_input survives a reload", (stored.get("fee_input") or {}).get("roles") is not None)

    # 2. exports must be unchanged for an untouched proposal
    jsx_before = client.post(
        f"/v1/proposals/{pid}/export/jsx",
        json={"cv_assignments": {}, "experience_ids": [], "template_id": "commercial"},
    ).content

    # 3. section regeneration: preview must not write
    baseline_schedule = stored["sections"]["schedule"]
    prev = client.post(f"/v1/proposals/{pid}/sections/schedule/regenerate", json={"commit": False})
    check("section preview -> 200", prev.status_code == 200, prev.text[:200])
    check("preview reports committed=false", prev.status_code == 200 and prev.json()["committed"] is False)
    after_preview = client.get(f"/v1/proposals/{pid}").json()["proposal"]
    check("preview did not write to the DB", after_preview["sections"]["schedule"] == baseline_schedule)

    # 4. instruction is refused on the fee/experience sections
    locked = client.post(f"/v1/proposals/{pid}/sections/financial/regenerate", json={"instruction": "make it cheaper", "commit": False})
    check("instruction locked on financial -> 400", locked.status_code == 400, str(locked.status_code))

    bad = client.post(f"/v1/proposals/{pid}/sections/not_a_section/regenerate", json={"commit": False})
    check("unknown section -> 404", bad.status_code == 404, str(bad.status_code))

    # 5. an allowlisted input patch changes the source data, section re-renders deterministically
    patched = client.post(
        f"/v1/proposals/{pid}/sections/schedule/regenerate",
        json={"input_patch": {"project.duration": "30 weeks", "schedule.milestones": ["Design review — week 12"], "lump_sum_total": 1}, "commit": True},
    )
    check("input_patch commit -> 200", patched.status_code == 200, patched.text[:200])
    if patched.status_code == 200:
        body = patched.json()
        check("out-of-allowlist key dropped", "lump_sum_total" not in (body.get("input_patch") or {}), str(body.get("input_patch")))
        check("schedule text reflects the patch", "30 weeks" in body["proposal"]["sections"]["schedule"], body["proposal"]["sections"]["schedule"][:160])
        check("parsed source data updated", body["proposal"]["parsed"]["project"]["duration"] == "30 weeks")
        check("markdown rebuilt", "30 weeks" in body["proposal"]["markdown"])

    # 6. cover letter: falls back to the deterministic template when no LLM is configured
    cover = client.post(f"/v1/proposals/{pid}/sections/cover_letter/regenerate", json={"commit": False})
    check("cover_letter regenerate -> 200", cover.status_code == 200, cover.text[:200])
    if cover.status_code == 200:
        cj = cover.json()
        check("cover_letter text non-empty", bool(cj["proposal"]["sections"]["cover_letter"].strip()))
        print(f"    notice: {cj.get('notice')}")

    # 7. experience reselect
    refs = client.get("/v1/reference-projects").json()
    resel = client.post(f"/v1/proposals/{pid}/experience/reselect", json={"auto": True, "limit": 2, "commit": True})
    check("experience reselect -> 200", resel.status_code == 200, resel.text[:200])
    if resel.status_code == 200:
        rel = resel.json()["proposal"]["relevant_experience"]
        check("reselect honours limit", len(rel) <= 2, str(len(rel)))
        check("reselect matches available references", len(rel) == min(2, len(refs)), f"{len(rel)} vs {len(refs)}")

    if refs:
        manual = client.post(
            f"/v1/proposals/{pid}/experience/reselect",
            json={"selected_reference_ids": [refs[0]["id"]], "limit": 3, "commit": True},
        )
        check("manual reselect -> 200", manual.status_code == 200, manual.text[:200])
        if manual.status_code == 200:
            chosen = manual.json()["proposal"]["relevant_experience"]
            check("manual reselect picks the chosen id", len(chosen) == 1 and chosen[0]["id"] == refs[0]["id"], str(chosen)[:160])

    # 8. fee recalculation must equal the fee engine exactly
    new_fee = json.loads(json.dumps(fee_input))
    new_fee["roles"][0]["hours_by_phase"]["Concept Development"] = 120
    new_fee["misc_reimbursables"] = 4500
    expected = client.post("/v1/fee/calculate", json=new_fee).json()["fee"]
    recalc = client.post(f"/v1/proposals/{pid}/fee/recalculate", json={"fee_input": new_fee, "commit": True})
    check("fee recalculate -> 200", recalc.status_code == 200, recalc.text[:200])
    if recalc.status_code == 200:
        rb = recalc.json()
        check("financial equals calculate_fee output", rb["proposal"]["financial"] == expected,
              f"{rb['proposal']['financial'].get('lump_sum_total')} vs {expected.get('lump_sum_total')}")
        check("fee_input persisted on recalculate", rb["proposal"]["fee_input"]["misc_reimbursables"] == 4500)
        check("dependent sections rebuilt", set(rb["changed_sections"]) == {"financial", "schedule", "team"}, str(rb["changed_sections"]))
        check("financial section quotes the new total",
              f"{expected['lump_sum_total']:.2f}" in rb["proposal"]["sections"]["financial"])
        reloaded = client.get(f"/v1/proposals/{pid}").json()["proposal"]
        check("recalculated fee persisted", reloaded["financial"]["lump_sum_total"] == expected["lump_sum_total"])

    # 9. an untouched proposal still exports byte-identical JSX
    control = client.post(
        "/v1/proposals/generate",
        json={
            "rfp_text": "Client requests a mixed-use master plan in Ho Chi Minh City over 16 weeks, with concept options and a final master plan report.",
            "fee_input": fee_input,
            "overrides": {"project_name": "Regen Smoke", "client_name": "Example Client"},
        },
    )
    if control.status_code == 200:
        cpid = control.json()["proposal_id"]
        jsx_control = client.post(
            f"/v1/proposals/{cpid}/export/jsx",
            json={"cv_assignments": {}, "experience_ids": [], "template_id": "commercial"},
        ).content
        check("untouched proposal exports a JSX bundle of the same size", len(jsx_control) == len(jsx_before),
              f"{len(jsx_control)} vs {len(jsx_before)}")

    print(f"\nPASS={_passed} FAIL={_failed}")
    return 1 if _failed else 0


if __name__ == "__main__":
    if "--verify-persist" in sys.argv:
        sys.exit(verify_persist())
    if "--dump-jsx" in sys.argv:
        sys.exit(dump_jsx())
    if "--regen" in sys.argv:
        sys.exit(regen_smoke())
    sys.exit(main())
