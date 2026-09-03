"""Storage: workspaces, members, uploaded workbooks, audit runs, shared reports.

SQLite with no ORM. Every query is written out, so the whole data model fits on
one screen and a reviewer can see exactly what is stored and what is not:
workbook bytes stay on disk under the workspace directory, and nothing is ever
sent anywhere unless the user turns on plain-English notes, which sends only
the finding metadata (never the file, never cell values outside the evidence).
"""
from __future__ import annotations

import json
import secrets
import sqlite3
import time
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Iterable

SCHEMA = """
CREATE TABLE IF NOT EXISTS workspace (
    id          TEXT PRIMARY KEY,
    name        TEXT NOT NULL,
    created_at  REAL NOT NULL,
    plan        TEXT NOT NULL DEFAULT 'free'
);
CREATE TABLE IF NOT EXISTS member (
    id            TEXT PRIMARY KEY,
    workspace_id  TEXT NOT NULL REFERENCES workspace(id) ON DELETE CASCADE,
    email         TEXT NOT NULL,
    password_hash TEXT NOT NULL,
    role          TEXT NOT NULL DEFAULT 'owner',
    created_at    REAL NOT NULL,
    UNIQUE (email)
);
CREATE TABLE IF NOT EXISTS session (
    token       TEXT PRIMARY KEY,
    member_id   TEXT NOT NULL REFERENCES member(id) ON DELETE CASCADE,
    created_at  REAL NOT NULL,
    expires_at  REAL NOT NULL
);
CREATE TABLE IF NOT EXISTS workbook (
    id            TEXT PRIMARY KEY,
    workspace_id  TEXT NOT NULL REFERENCES workspace(id) ON DELETE CASCADE,
    name          TEXT NOT NULL,
    stored_path   TEXT NOT NULL,
    size_bytes    INTEGER NOT NULL,
    uploaded_by   TEXT,
    created_at    REAL NOT NULL
);
CREATE TABLE IF NOT EXISTS run (
    id            TEXT PRIMARY KEY,
    workbook_id   TEXT NOT NULL REFERENCES workbook(id) ON DELETE CASCADE,
    workspace_id  TEXT NOT NULL,
    created_at    REAL NOT NULL,
    duration_ms   INTEGER NOT NULL,
    critical      INTEGER NOT NULL,
    warning       INTEGER NOT NULL,
    info          INTEGER NOT NULL,
    money_at_risk REAL NOT NULL DEFAULT 0,
    agreement     REAL NOT NULL DEFAULT 0,
    report_json   TEXT NOT NULL,
    share_token   TEXT
);
CREATE INDEX IF NOT EXISTS run_by_workbook ON run(workbook_id, created_at DESC);
CREATE INDEX IF NOT EXISTS wb_by_workspace ON workbook(workspace_id, created_at DESC);
CREATE UNIQUE INDEX IF NOT EXISTS run_share ON run(share_token) WHERE share_token IS NOT NULL;
"""


def _now() -> float:
    return time.time()


def new_id(prefix: str) -> str:
    return f"{prefix}_{secrets.token_hex(8)}"


class Store:
    def __init__(self, path: str | Path, data_dir: str | Path | None = None):
        self.path = str(path)
        self.data_dir = Path(data_dir or Path(path).parent / "workbooks")
        self.data_dir.mkdir(parents=True, exist_ok=True)
        self.conn = sqlite3.connect(self.path, check_same_thread=False)
        self.conn.row_factory = sqlite3.Row
        self.conn.execute("PRAGMA foreign_keys = ON")
        self.conn.execute("PRAGMA journal_mode = WAL")
        self.conn.executescript(SCHEMA)
        self.conn.commit()

    # ---------------------------------------------------------------- accounts
    def create_workspace(self, name: str, email: str, password_hash: str) -> tuple[str, str]:
        ws_id, m_id = new_id("ws"), new_id("mem")
        with self.conn:
            self.conn.execute("INSERT INTO workspace (id,name,created_at) VALUES (?,?,?)",
                              (ws_id, name, _now()))
            self.conn.execute(
                "INSERT INTO member (id,workspace_id,email,password_hash,role,created_at) "
                "VALUES (?,?,?,?,'owner',?)", (m_id, ws_id, email.lower(), password_hash, _now()))
        return ws_id, m_id

    def member_by_email(self, email: str) -> sqlite3.Row | None:
        return self.conn.execute("SELECT * FROM member WHERE email = ?", (email.lower(),)).fetchone()

    def start_session(self, member_id: str, ttl_seconds: int = 60 * 60 * 24 * 14) -> str:
        token = secrets.token_urlsafe(32)
        with self.conn:
            self.conn.execute("INSERT INTO session (token,member_id,created_at,expires_at) VALUES (?,?,?,?)",
                              (token, member_id, _now(), _now() + ttl_seconds))
        return token

    def member_for_token(self, token: str | None) -> sqlite3.Row | None:
        if not token:
            return None
        row = self.conn.execute(
            "SELECT m.* FROM session s JOIN member m ON m.id = s.member_id "
            "WHERE s.token = ? AND s.expires_at > ?", (token, _now())).fetchone()
        return row

    def end_session(self, token: str) -> None:
        with self.conn:
            self.conn.execute("DELETE FROM session WHERE token = ?", (token,))

    # --------------------------------------------------------------- workbooks
    def add_workbook(self, workspace_id: str, name: str, data: bytes, uploaded_by: str | None) -> str:
        wb_id = new_id("wb")
        folder = self.data_dir / workspace_id
        folder.mkdir(parents=True, exist_ok=True)
        dest = folder / f"{wb_id}.xlsx"
        dest.write_bytes(data)
        with self.conn:
            self.conn.execute(
                "INSERT INTO workbook (id,workspace_id,name,stored_path,size_bytes,uploaded_by,created_at) "
                "VALUES (?,?,?,?,?,?,?)",
                (wb_id, workspace_id, name, str(dest), len(data), uploaded_by, _now()))
        return wb_id

    def workbook(self, workbook_id: str) -> sqlite3.Row | None:
        return self.conn.execute("SELECT * FROM workbook WHERE id = ?", (workbook_id,)).fetchone()

    def workbooks(self, workspace_id: str, limit: int = 100) -> list[sqlite3.Row]:
        return self.conn.execute(
            "SELECT w.*, "
            " (SELECT COUNT(*) FROM run r WHERE r.workbook_id = w.id) AS runs, "
            " (SELECT r.critical FROM run r WHERE r.workbook_id = w.id ORDER BY r.created_at DESC LIMIT 1) AS critical, "
            " (SELECT r.id FROM run r WHERE r.workbook_id = w.id ORDER BY r.created_at DESC LIMIT 1) AS last_run "
            "FROM workbook w WHERE w.workspace_id = ? ORDER BY w.created_at DESC LIMIT ?",
            (workspace_id, limit)).fetchall()

    def delete_workbook(self, workbook_id: str) -> None:
        row = self.workbook(workbook_id)
        if row is None:
            return
        Path(row["stored_path"]).unlink(missing_ok=True)
        with self.conn:
            self.conn.execute("DELETE FROM workbook WHERE id = ?", (workbook_id,))

    # -------------------------------------------------------------------- runs
    def add_run(self, workbook_id: str, workspace_id: str, report: dict[str, Any]) -> str:
        run_id = new_id("run")
        counts = report["counts"]
        with self.conn:
            self.conn.execute(
                "INSERT INTO run (id,workbook_id,workspace_id,created_at,duration_ms,critical,warning,info,"
                "money_at_risk,agreement,report_json) VALUES (?,?,?,?,?,?,?,?,?,?,?)",
                (run_id, workbook_id, workspace_id, _now(), report["duration_ms"],
                 counts["critical"], counts["warning"], counts["info"],
                 report.get("money_at_risk", 0.0), report["engine"]["agreement"],
                 json.dumps(report, ensure_ascii=False, default=str)))
        return run_id

    def run(self, run_id: str) -> dict[str, Any] | None:
        row = self.conn.execute("SELECT * FROM run WHERE id = ?", (run_id,)).fetchone()
        return _run_dict(row) if row else None

    def run_by_share_token(self, token: str) -> dict[str, Any] | None:
        row = self.conn.execute("SELECT * FROM run WHERE share_token = ?", (token,)).fetchone()
        return _run_dict(row) if row else None

    def share(self, run_id: str) -> str:
        token = secrets.token_urlsafe(12)
        with self.conn:
            self.conn.execute("UPDATE run SET share_token = ? WHERE id = ?", (token, run_id))
        return token

    def unshare(self, run_id: str) -> None:
        with self.conn:
            self.conn.execute("UPDATE run SET share_token = NULL WHERE id = ?", (run_id,))

    def runs_for_workbook(self, workbook_id: str, limit: int = 30) -> list[sqlite3.Row]:
        return self.conn.execute(
            "SELECT id,created_at,critical,warning,info,money_at_risk,agreement,duration_ms,share_token "
            "FROM run WHERE workbook_id = ? ORDER BY created_at DESC LIMIT ?",
            (workbook_id, limit)).fetchall()

    def workspace_summary(self, workspace_id: str) -> dict[str, Any]:
        row = self.conn.execute(
            "SELECT COUNT(DISTINCT w.id) AS workbooks, COUNT(r.id) AS runs, "
            "COALESCE(SUM(r.critical),0) AS criticals "
            "FROM workbook w LEFT JOIN run r ON r.workbook_id = w.id WHERE w.workspace_id = ?",
            (workspace_id,)).fetchone()
        return dict(row) if row else {"workbooks": 0, "runs": 0, "criticals": 0}


def _run_dict(row: sqlite3.Row) -> dict[str, Any]:
    d = dict(row)
    d["report"] = json.loads(d.pop("report_json"))
    return d
