"""API tests. They run the real app against a temporary database."""
from __future__ import annotations

import os
import tempfile
from pathlib import Path

import pytest
from fastapi.testclient import TestClient

SAMPLES = Path(__file__).resolve().parent.parent / "samples"


@pytest.fixture(scope="module")
def client():
    tmp = tempfile.mkdtemp(prefix="gridlint-test-")
    os.environ["GRIDLINT_DATA"] = tmp
    import importlib

    from gridlint import server
    importlib.reload(server)
    with TestClient(server.app) as c:
        yield c
    os.environ.pop("GRIDLINT_DATA", None)


def test_health_reports_the_configured_model(client):
    r = client.get("/api/health")
    assert r.status_code == 200
    assert r.json()["ok"] and "provider" in r.json()["model"]


def test_rules_endpoint_lists_every_rule(client):
    rules = client.get("/api/rules").json()
    codes = {r["code"] for r in rules}
    assert {"R001", "R002", "R005", "R012", "R013"} <= codes
    assert all(r["why"] for r in rules)


def test_demo_needs_no_account_and_no_key(client):
    r = client.get("/api/demo")
    assert r.status_code == 200
    d = r.json()
    assert d["counts"]["critical"] >= 1
    assert d["engine"]["trustworthy"]
    assert d["findings"][0]["fix_verified"] is True


def test_demo_rejects_a_file_outside_the_samples_folder(client):
    assert client.get("/api/demo", params={"file": "../gridlint/server.py"}).status_code == 404


def test_me_is_200_for_an_anonymous_visitor(client):
    r = client.get("/api/me")
    assert r.status_code == 200 and r.json() == {"signed_in": False}


def test_check_uploads_and_audits_without_storing(client):
    data = (SAMPLES / "board-model.xlsx").read_bytes()
    r = client.post("/api/check", files={"file": ("board-model.xlsx", data, "application/vnd.ms-excel")})
    assert r.status_code == 200
    assert r.json()["findings"][0]["rule"] == "R001"


def test_a_non_xlsx_upload_gets_a_useful_message(client):
    r = client.post("/api/check", files={"file": ("notes.csv", b"a,b,c\n1,2,3\n", "text/csv")})
    assert r.status_code == 400
    assert ".xlsx" in r.json()["detail"]


def test_a_renamed_non_zip_is_rejected(client):
    r = client.post("/api/check", files={"file": ("fake.xlsx", b"not a zip at all", "application/vnd.ms-excel")})
    assert r.status_code == 400


def test_a_corrupt_xlsx_is_reported_not_crashed(client):
    r = client.post("/api/check", files={"file": ("broken.xlsx", b"PK\x03\x04broken", "application/vnd.ms-excel")})
    assert r.status_code == 422
    assert "Could not read" in r.json()["detail"]


def test_workspace_endpoints_require_a_session(client):
    assert client.get("/api/workbooks").status_code == 401


def test_signup_upload_recheck_share_flow(client):
    r = client.post("/api/signup", data={"email": "a@example.com", "password": "correct horse",
                                         "workspace": "Finance"})
    assert r.status_code == 200

    me = client.get("/api/me").json()
    assert me["signed_in"] and me["email"] == "a@example.com"

    data = (SAMPLES / "board-model.xlsx").read_bytes()
    up = client.post("/api/workbooks", files={"file": ("board-model.xlsx", data, "application/vnd.ms-excel")})
    assert up.status_code == 200
    wb_id, run_id = up.json()["workbook_id"], up.json()["run_id"]

    listed = client.get("/api/workbooks").json()
    assert len(listed) == 1 and listed[0]["id"] == wb_id and listed[0]["critical"] >= 1

    again = client.post(f"/api/workbooks/{wb_id}/recheck")
    assert again.status_code == 200
    assert len(client.get(f"/api/workbooks/{wb_id}/runs").json()) == 2

    token = client.post(f"/api/runs/{run_id}/share").json()["token"]
    anon = TestClient(client.app)                       # a browser with no cookies
    shared = anon.get(f"/api/shared/{token}")
    assert shared.status_code == 200
    assert shared.json()["report"]["findings"][0]["rule"] == "R001"

    client.delete(f"/api/runs/{run_id}/share")
    assert anon.get(f"/api/shared/{token}").status_code == 404


def test_signup_rejects_a_short_password(client):
    r = client.post("/api/signup", data={"email": "b@example.com", "password": "short"})
    assert r.status_code == 400


def test_login_rejects_a_wrong_password(client):
    client.post("/api/signup", data={"email": "c@example.com", "password": "correct horse"})
    fresh = TestClient(client.app)
    assert fresh.post("/api/login", data={"email": "c@example.com", "password": "wrong pass"}).status_code == 401
    assert fresh.post("/api/login", data={"email": "c@example.com", "password": "correct horse"}).status_code == 200


def test_one_workspace_cannot_read_another(client):
    data = (SAMPLES / "runway.xlsx").read_bytes()
    a = TestClient(client.app)
    a.post("/api/signup", data={"email": "owner@example.com", "password": "correct horse"})
    wb_id = a.post("/api/workbooks", files={"file": ("runway.xlsx", data, "x")}).json()["workbook_id"]

    b = TestClient(client.app)
    b.post("/api/signup", data={"email": "other@example.com", "password": "correct horse"})
    assert b.get(f"/api/workbooks/{wb_id}/runs").status_code == 404
    assert b.post(f"/api/workbooks/{wb_id}/recheck").status_code == 404
    assert b.delete(f"/api/workbooks/{wb_id}").status_code == 404


def test_download_of_the_corrected_file_returns_a_workbook(client):
    import io

    import openpyxl

    data = (SAMPLES / "runway.xlsx").read_bytes()
    s = TestClient(client.app)
    s.post("/api/signup", data={"email": "fix@example.com", "password": "correct horse"})
    up = s.post("/api/workbooks", files={"file": ("runway.xlsx", data, "x")}).json()
    r = s.post(f"/api/workbooks/{up['workbook_id']}/fixed", json={})
    assert r.status_code == 200
    wb = openpyxl.load_workbook(io.BytesIO(r.content))
    assert wb["Model"]["C16"].value == "=SUM(C11:C15)"
    assert wb["Model"]["C16"].comment is not None, "the original formula should be kept in a comment"


def test_explain_endpoint_returns_recorded_notes_without_a_key(client):
    report = client.get("/api/demo").json()
    r = client.post("/api/explain", json={"findings": report["findings"][:3], "report": report})
    assert r.status_code == 200
    body = r.json()
    assert len(body["explanations"]) == 3
    assert any(e["explanation"] for e in body["explanations"]), "fixtures should replay"
    assert all(e["rejected_numbers"] is None for e in body["explanations"])
    assert body["review_note"]


def test_index_and_static_assets_are_served(client):
    assert "Gridlint" in client.get("/").text
    assert client.get("/static/app.css").status_code == 200
    assert client.get("/static/app.js").status_code == 200
