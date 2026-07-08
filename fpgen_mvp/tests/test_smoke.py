"""Smoke tests for the database-backed data layer and read endpoints."""
from fastapi.testclient import TestClient

from app import repositories
from app.db import SessionLocal, init_db
from app.main import app

init_db()
client = TestClient(app)


def _seed():
    with SessionLocal() as db:
        repositories.upsert_unit_rate(db, {"role": "Architect", "rate": 200, "currency": "USD"})
        repositories.upsert_unit_rate(db, {"role": "Principal", "rate": 350, "currency": "USD"})
        repositories.upsert_personnel(
            db,
            {"id": "p-003", "name": "Aiko Suzuki", "title": "Senior Architect",
             "roles": ["Architect"], "cv_asset_id": "cv-architect"},
        )
        repositories.upsert_team_preset(
            db,
            {"id": "preset-commercial", "name": "Commercial", "types": ["Commercial"],
             "assignments": [
                 {"role": "Principal", "person_id": "p-001", "rate": 350},
                 {"role": "Architect", "person_id": "p-003"},  # rate resolves from unit_rates
             ]},
        )
        repositories.upsert_reference_project(
            db,
            {"id": "rp-001", "name": "Central District", "project_type": "Commercial",
             "location": "Tokyo", "keywords": ["mixed-use"], "summary": "A project."},
        )


def test_personnel_endpoint():
    _seed()
    r = client.get("/v1/personnel")
    assert r.status_code == 200
    assert any(p["id"] == "p-003" for p in r.json())


def test_preset_rate_resolution():
    _seed()
    r = client.get("/v1/team-presets")
    assert r.status_code == 200
    preset = next(p for p in r.json() if p["id"] == "preset-commercial")
    rates = {a["role"]: a["rate"] for a in preset["assignments"]}
    assert rates["Principal"] == 350  # explicit override
    assert rates["Architect"] == 200  # resolved from unit_rates


def test_reference_projects_endpoint():
    _seed()
    r = client.get("/v1/reference-projects")
    assert r.status_code == 200
    assert any(p["id"] == "rp-001" for p in r.json())


def test_proposal_roundtrip():
    with SessionLocal() as db:
        pid, wrapped = repositories.save_proposal(db, {"project": {"name": "X"}, "sections": {"a": "1"}})
    r = client.get(f"/v1/proposals/{pid}")
    assert r.status_code == 200
    assert r.json()["proposal"]["project"]["name"] == "X"

    r = client.put(f"/v1/proposals/{pid}", json={"sections": {"cover_letter": "Hello"}})
    assert r.status_code == 200


def test_unit_rates_listing():
    _seed()
    with SessionLocal() as db:
        rates = repositories.list_unit_rates(db)
    assert {r["role"] for r in rates} >= {"Architect", "Principal"}


# --- Admin API (auth disabled in tests => synthetic admin) --- #
def test_me_dev_admin():
    r = client.get("/v1/me")
    assert r.status_code == 200
    assert "fpgen_admin" in r.json()["groups"]


def test_admin_personnel_crud():
    r = client.post(
        "/v1/admin/personnel",
        json={"id": "p-999", "name": "Zed", "title": "Eng", "roles": ["MEP Engineer"]},
    )
    assert r.status_code == 200
    assert any(p["id"] == "p-999" for p in client.get("/v1/admin/personnel").json())
    assert client.delete("/v1/admin/personnel/p-999").status_code == 200


def test_admin_rate_crud():
    r = client.post("/v1/admin/rates", json={"role": "MEP Engineer", "rate": 175, "currency": "USD"})
    assert r.status_code == 200
    assert any(x["role"] == "MEP Engineer" for x in client.get("/v1/admin/rates").json())


def test_admin_asset_upload():
    files = {"file": ("cv.txt", b"hello", "text/plain")}
    data = {"asset_id": "cv-test", "kind": "cv", "role": "Architect"}
    r = client.post("/v1/admin/assets/upload", data=data, files=files)
    assert r.status_code == 200
    body = r.json()
    assert body["id"] == "cv-test"
    assert body["size_bytes"] == 5


def test_admin_preset_crud():
    _seed()
    r = client.post(
        "/v1/admin/presets",
        json={
            "id": "preset-test",
            "name": "Test",
            "types": ["X"],
            "assignments": [{"role": "Architect"}],
        },
    )
    assert r.status_code == 200
    # Architect rate should resolve from unit_rates seeded above.
    assert r.json()["assignments"][0]["rate"] == 200


def test_static_pages_served():
    assert client.get("/").status_code == 200
    admin_index = client.get("/admin/")
    assert admin_index.status_code == 200
    assert "FP-GEN" in admin_index.text
    assert client.get("/admin/admin.js").status_code == 200


