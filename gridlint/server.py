"""The web app: a workspace where a team checks the same workbooks every month.

Design notes a reviewer might want:
  * The uploaded file never leaves the server. Plain-English notes are optional
    and send only finding metadata to the model, never the workbook.
  * A report can be shared with a read-only link, which is how a reviewer sends
    "here is what I found" to whoever owns the file.
  * `/api/demo` runs the bundled sample so the app is usable with no account and
    no API key, which is how a judge should meet it.
"""
from __future__ import annotations

import hashlib
import hmac
import json
import os
import secrets
import tempfile
from pathlib import Path
from typing import Any

from fastapi import Cookie, Depends, FastAPI, File, Form, HTTPException, Request, Response, UploadFile
from fastapi.responses import FileResponse, HTMLResponse, JSONResponse
from fastapi.staticfiles import StaticFiles

from . import __version__
from .audit import audit
from .db import Store
from .rules import registry

ROOT = Path(__file__).resolve().parent
WEB = ROOT / "web"
SAMPLES = ROOT.parent / "samples"
DATA_DIR = Path(os.environ.get("GRIDLINT_DATA", ROOT.parent / "data"))
MAX_UPLOAD_BYTES = int(os.environ.get("GRIDLINT_MAX_UPLOAD", 15 * 1024 * 1024))

app = FastAPI(title="Gridlint", version=__version__, docs_url="/api/docs", redoc_url=None)
store = Store(DATA_DIR / "gridlint.db", DATA_DIR / "workbooks")


# --------------------------------------------------------------------- helpers

def hash_password(password: str, salt: str | None = None) -> str:
    salt = salt or secrets.token_hex(16)
    digest = hashlib.pbkdf2_hmac("sha256", password.encode(), salt.encode(), 240_000)
    return f"pbkdf2$240000${salt}${digest.hex()}"


def verify_password(password: str, stored: str) -> bool:
    try:
        _algo, iters, salt, digest = stored.split("$")
        want = hashlib.pbkdf2_hmac("sha256", password.encode(), salt.encode(), int(iters)).hex()
        return hmac.compare_digest(want, digest)
    except ValueError:
        return False


def current_member(gridlint_session: str | None = Cookie(default=None)):
    member = store.member_for_token(gridlint_session)
    if member is None:
        raise HTTPException(status_code=401, detail="Sign in to continue.")
    return member


def _safe_name(name: str) -> str:
    base = Path(name or "workbook.xlsx").name
    return "".join(ch for ch in base if ch.isalnum() or ch in " ._-()[]")[:120] or "workbook.xlsx"


def _check_upload(file: UploadFile, data: bytes) -> None:
    if len(data) > MAX_UPLOAD_BYTES:
        raise HTTPException(413, f"File is larger than {MAX_UPLOAD_BYTES // (1024 * 1024)} MB.")
    if not data[:2] == b"PK":
        raise HTTPException(400, "That does not look like an .xlsx file. "
                                 "Save as Excel Workbook (.xlsx) and try again; .xls and .csv are not supported yet.")
    if not (file.filename or "").lower().endswith((".xlsx", ".xlsm")):
        raise HTTPException(400, "Only .xlsx and .xlsm files can be checked.")


def _audit_bytes(data: bytes, name: str) -> dict[str, Any]:
    with tempfile.TemporaryDirectory(prefix="gridlint-") as tmp:
        p = Path(tmp) / _safe_name(name)
        p.write_bytes(data)
        try:
            report = audit(p)
        except Exception as e:
            raise HTTPException(422, f"Could not read that workbook: {type(e).__name__}: {e}") from e
    d = report.to_dict()
    d["path"] = name
    return d


# ------------------------------------------------------------------ public API

@app.get("/api/health")
def health() -> dict[str, Any]:
    from .llm import provider_info

    return {"ok": True, "version": __version__, "model": provider_info()}


@app.get("/api/rules")
def rules() -> list[dict[str, str]]:
    return [{"code": m.code, "name": m.name, "why": m.why, "severity": m.default_severity}
            for m, _fn in registry()]


@app.get("/api/demo")
def demo(file: str = "board-model.xlsx") -> JSONResponse:
    """Audit a bundled sample. No account, no key, no upload."""
    allowed = {p.name for p in SAMPLES.glob("*.xlsx")}
    if file not in allowed:
        raise HTTPException(404, f"Unknown sample. Available: {sorted(allowed)}")
    report = audit(SAMPLES / file)
    return JSONResponse(report.to_dict())


@app.get("/api/samples")
def samples() -> list[dict[str, Any]]:
    out = []
    for p in sorted(SAMPLES.glob("*.xlsx")):
        out.append({"name": p.name, "size": p.stat().st_size,
                    "label": {"board-model.xlsx": "Board deck operating model (has defects)",
                              "clean-model.xlsx": "The same model, healthy",
                              "conformance.xlsx": "Engine conformance suite",
                              "runway.xlsx": "Small runway model (has a defect)"}.get(p.name, p.stem)})
    return out


@app.post("/api/check")
async def check(file: UploadFile = File(...)) -> JSONResponse:
    """Audit an uploaded workbook without storing it. Used by the try-it-now box."""
    data = await file.read()
    _check_upload(file, data)
    return JSONResponse(_audit_bytes(data, _safe_name(file.filename or "workbook.xlsx")))


@app.post("/api/explain")
def explain_endpoint(payload: dict = None) -> JSONResponse:
    """Add plain-English notes to a report the client already has.

    Only the finding metadata is sent to the model, never the workbook.
    """
    from . import explain as ex
    from .explain import _finding_from_dict

    payload = payload or {}
    findings = payload.get("findings") or []
    report_name = (payload.get("report") or {}).get("path", "workbook.xlsx")
    out = []
    for f in findings[:10]:
        finding = _finding_from_dict(f)
        # The context is derived from the report, never from the client, so the
        # same finding always yields the same prompt and the same cache key.
        ctx = ex.sheet_context(report_name, finding.sheet)
        text, guard = ex.explain_finding(finding, sheet_context=ctx)
        out.append({"cell": f.get("cell"), "rule": f.get("rule"), "explanation": text,
                    "rejected_numbers": None if guard.ok else guard.offending})
    note = None
    if payload.get("report"):
        note = ex.review_note(payload["report"])
    return JSONResponse({"explanations": out, "review_note": note})


# --------------------------------------------------------------------- accounts

@app.post("/api/signup")
def signup(response: Response, email: str = Form(...), password: str = Form(...),
           workspace: str = Form("My workspace")) -> dict[str, Any]:
    if len(password) < 8:
        raise HTTPException(400, "Use a password of at least 8 characters.")
    if store.member_by_email(email):
        raise HTTPException(409, "That email already has an account.")
    ws_id, m_id = store.create_workspace(workspace.strip() or "My workspace", email, hash_password(password))
    token = store.start_session(m_id)
    _set_cookie(response, token)
    return {"workspace_id": ws_id, "email": email.lower()}


@app.post("/api/login")
def login(response: Response, email: str = Form(...), password: str = Form(...)) -> dict[str, Any]:
    member = store.member_by_email(email)
    if member is None or not verify_password(password, member["password_hash"]):
        raise HTTPException(401, "Email or password is not right.")
    _set_cookie(response, store.start_session(member["id"]))
    return {"workspace_id": member["workspace_id"], "email": member["email"]}


@app.post("/api/logout")
def logout(response: Response, gridlint_session: str | None = Cookie(default=None)) -> dict[str, bool]:
    if gridlint_session:
        store.end_session(gridlint_session)
    response.delete_cookie("gridlint_session")
    return {"ok": True}


def _set_cookie(response: Response, token: str) -> None:
    response.set_cookie("gridlint_session", token, httponly=True, samesite="lax",
                        secure=os.environ.get("GRIDLINT_HTTPS") == "1",
                        max_age=60 * 60 * 24 * 14, path="/")


@app.get("/api/me")
def me(gridlint_session: str | None = Cookie(default=None)) -> dict[str, Any]:
    """Always 200. Signing in is optional here, so an anonymous visit is a
    normal state, not an error the browser should log."""
    member = store.member_for_token(gridlint_session)
    if member is None:
        return {"signed_in": False}
    return {"signed_in": True, "email": member["email"], "workspace_id": member["workspace_id"],
            "summary": store.workspace_summary(member["workspace_id"])}


# -------------------------------------------------------------------- workspace

@app.get("/api/workbooks")
def list_workbooks(member=Depends(current_member)) -> list[dict[str, Any]]:
    return [dict(r) for r in store.workbooks(member["workspace_id"])]


@app.post("/api/workbooks")
async def upload_workbook(file: UploadFile = File(...), member=Depends(current_member)) -> dict[str, Any]:
    data = await file.read()
    _check_upload(file, data)
    name = _safe_name(file.filename or "workbook.xlsx")
    wb_id = store.add_workbook(member["workspace_id"], name, data, member["id"])
    report = _audit_bytes(data, name)
    run_id = store.add_run(wb_id, member["workspace_id"], report)
    return {"workbook_id": wb_id, "run_id": run_id, "report": report}


@app.post("/api/workbooks/{workbook_id}/recheck")
def recheck(workbook_id: str, member=Depends(current_member)) -> dict[str, Any]:
    wb = store.workbook(workbook_id)
    if wb is None or wb["workspace_id"] != member["workspace_id"]:
        raise HTTPException(404, "No such workbook.")
    report = audit(wb["stored_path"]).to_dict()
    report["path"] = wb["name"]
    run_id = store.add_run(workbook_id, member["workspace_id"], report)
    return {"run_id": run_id, "report": report}


@app.get("/api/workbooks/{workbook_id}/runs")
def workbook_runs(workbook_id: str, member=Depends(current_member)) -> list[dict[str, Any]]:
    wb = store.workbook(workbook_id)
    if wb is None or wb["workspace_id"] != member["workspace_id"]:
        raise HTTPException(404, "No such workbook.")
    return [dict(r) for r in store.runs_for_workbook(workbook_id)]


@app.delete("/api/workbooks/{workbook_id}")
def remove_workbook(workbook_id: str, member=Depends(current_member)) -> dict[str, bool]:
    wb = store.workbook(workbook_id)
    if wb is None or wb["workspace_id"] != member["workspace_id"]:
        raise HTTPException(404, "No such workbook.")
    store.delete_workbook(workbook_id)
    return {"ok": True}


@app.get("/api/runs/{run_id}")
def get_run(run_id: str, member=Depends(current_member)) -> dict[str, Any]:
    run = store.run(run_id)
    if run is None or run["workspace_id"] != member["workspace_id"]:
        raise HTTPException(404, "No such report.")
    return run


@app.post("/api/runs/{run_id}/share")
def share_run(run_id: str, member=Depends(current_member)) -> dict[str, str]:
    run = store.run(run_id)
    if run is None or run["workspace_id"] != member["workspace_id"]:
        raise HTTPException(404, "No such report.")
    return {"token": store.share(run_id)}


@app.delete("/api/runs/{run_id}/share")
def unshare_run(run_id: str, member=Depends(current_member)) -> dict[str, bool]:
    run = store.run(run_id)
    if run is None or run["workspace_id"] != member["workspace_id"]:
        raise HTTPException(404, "No such report.")
    store.unshare(run_id)
    return {"ok": True}


@app.get("/api/shared/{token}")
def shared_report(token: str) -> dict[str, Any]:
    run = store.run_by_share_token(token)
    if run is None:
        raise HTTPException(404, "This link is no longer active.")
    return {"created_at": run["created_at"], "report": run["report"]}


@app.post("/api/workbooks/{workbook_id}/fixed")
def download_fixed(workbook_id: str, payload: dict = None, member=Depends(current_member)):
    """Return a copy of the workbook with the chosen verified fixes applied."""
    from .apply import apply_fixes

    wb = store.workbook(workbook_id)
    if wb is None or wb["workspace_id"] != member["workspace_id"]:
        raise HTTPException(404, "No such workbook.")
    wanted = set((payload or {}).get("finding_ids") or [])
    report = audit(wb["stored_path"])
    chosen = [f for f in report.findings
              if f.fix and f.fix_verified and (not wanted or f.id in wanted)]
    if not chosen:
        raise HTTPException(400, "None of the selected findings have a fix that passed verification.")
    out = Path(tempfile.mkdtemp(prefix="gridlint-fix-")) / f"{Path(wb['name']).stem}-fixed.xlsx"
    apply_fixes(wb["stored_path"], chosen, out)
    return FileResponse(out, filename=out.name,
                        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


# ------------------------------------------------------------------------- web

@app.get("/", response_class=HTMLResponse)
def index() -> HTMLResponse:
    return HTMLResponse((WEB / "index.html").read_text(encoding="utf-8"))


@app.get("/app", response_class=HTMLResponse)
@app.get("/report/{_rest:path}", response_class=HTMLResponse)
def spa(_rest: str = "") -> HTMLResponse:
    return HTMLResponse((WEB / "index.html").read_text(encoding="utf-8"))


app.mount("/static", StaticFiles(directory=str(WEB)), name="static")
