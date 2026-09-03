"""Tests for the fence around the language model.

The product's claim is that a hallucinated number cannot reach the report as a
fact. These tests are that claim, written down.
"""
from __future__ import annotations

import json
from pathlib import Path

import pytest

from gridlint import explain
from gridlint.audit import audit
from gridlint.explain import NumberGuard, allowed_numbers, numbers_in
from gridlint.rules.base import Edit, Finding, Fix

SAMPLES = Path(__file__).resolve().parent.parent / "samples"


@pytest.fixture(scope="module")
def report():
    return audit(SAMPLES / "board-model.xlsx")


# --------------------------------------------------------------------- guard
def test_guard_accepts_figures_that_are_in_the_evidence():
    g = NumberGuard([185000.0, 5.2, 38.6])
    assert g.check("Contractors is 185,000, so runway is 5.2 months not 38.6.").ok


def test_guard_rejects_a_number_the_model_invented():
    g = NumberGuard([185000.0])
    r = g.check("This costs about 240,000 a month.")
    assert not r.ok and "240,000" in r.offending


def test_guard_allows_small_counting_words():
    g = NumberGuard([])
    assert g.check("The same mistake appears in 12 cells across 3 sheets.").ok


def test_guard_handles_percentages_and_thousands_separators():
    g = NumberGuard([0.063, 1234567.0])
    assert g.check("Margin is 6.3% on 1,234,567.").ok
    assert not g.check("Margin is 9.9%.").ok


def test_guard_tolerates_rounding_of_an_allowed_figure():
    g = NumberGuard([2779774.0108])
    assert g.check("It moves 2,779,774.").ok


def test_numbers_in_text_are_extracted():
    assert 2026.0 in numbers_in("the twelve months of FY26 ending 2026")


def test_allowed_numbers_covers_every_measured_figure(report):
    f = report.findings[0]
    allowed = set(allowed_numbers(f))
    for h in f.headline_changes:
        assert float(h["before"]) in allowed and float(h["after"]) in allowed
    assert float(f.impact_value) in allowed


def test_context_numbers_are_allowed_because_we_supplied_them(report, monkeypatch):
    """A figure Gridlint put in the prompt is not a hallucination when repeated."""
    f = report.findings[0]
    ctx = explain.sheet_context("board-model.xlsx", f.sheet)
    guard = NumberGuard(allowed_numbers(f) + numbers_in(ctx))
    assert guard.check(f"In {f.sheet}, the total is short.").ok


# ------------------------------------------------------------------- replay
def test_explanations_replay_from_fixtures_with_no_api_key(report, monkeypatch):
    """A judge with no key must still see the written notes."""
    for var in ("ANTHROPIC_API_KEY", "OPENAI_API_KEY", "GROQ_API_KEY", "GEMINI_API_KEY", "LLM_PROVIDER"):
        monkeypatch.delenv(var, raising=False)
    written = 0
    for f in report.findings:
        text, guard = explain.explain_finding(
            f, sheet_context=explain.sheet_context("board-model.xlsx", f.sheet))
        assert guard.ok, f"a recorded note quoted an unverified number: {guard.offending}"
        if text:
            written += 1
            assert len(text) > 40
    assert written >= 3, "expected recorded explanations for the demo findings"


def test_missing_fixture_degrades_to_the_built_in_sentence(monkeypatch, tmp_path):
    """No key and no fixture must not raise: the deterministic detail still stands."""
    import importlib

    from gridlint import llm

    monkeypatch.setenv("LLM_FIXTURE_DIR", str(tmp_path))
    for var in ("ANTHROPIC_API_KEY", "OPENAI_API_KEY", "GROQ_API_KEY", "GEMINI_API_KEY", "LLM_PROVIDER"):
        monkeypatch.delenv(var, raising=False)
    importlib.reload(llm)
    importlib.reload(explain)
    try:
        f = Finding(rule="R001", title="t", severity="critical", cell="S!A1", sheet="S",
                    col=1, row=1, detail="something")
        text, guard = explain.explain_finding(f)
        assert text is None and guard.ok
    finally:
        monkeypatch.delenv("LLM_FIXTURE_DIR", raising=False)
        importlib.reload(llm)
        importlib.reload(explain)


# ------------------------------------------------------- proposed repairs
class _FakeResult:
    def __init__(self, text):
        self.text = text

    def json(self):
        from gridlint.llm import parse_json
        return parse_json(self.text)


def _patch_model(monkeypatch, payload):
    monkeypatch.setattr(explain, "complete", lambda *a, **k: _FakeResult(json.dumps(payload)))


def test_a_proposed_repair_that_does_not_parse_is_rejected(report, monkeypatch):
    from gridlint import engine

    wb_path = SAMPLES / "board-model.xlsx"
    from gridlint import workbook
    wb = workbook.load(wb_path)
    graph = engine.build_graph(wb)
    base = engine.recalc(wb, graph)

    _patch_model(monkeypatch, {"edits": [{"cell": "Operating Model!C28", "formula": "=MAX(0,-C18*"}],
                               "rationale": "nope"})
    out = explain.propose_repair(wb, graph, base, report.findings[0])
    assert not out.accepted and "does not parse" in out.rejected_reason


def test_a_proposed_repair_pointing_at_a_missing_sheet_is_rejected(report, monkeypatch):
    from gridlint import engine, workbook

    wb = workbook.load(SAMPLES / "board-model.xlsx")
    graph = engine.build_graph(wb)
    base = engine.recalc(wb, graph)
    _patch_model(monkeypatch, {"edits": [{"cell": "Ghost!A1", "formula": "=1"}], "rationale": "x"})
    out = explain.propose_repair(wb, graph, base, report.findings[0])
    assert not out.accepted and "not in this file" in out.rejected_reason


def test_a_proposed_repair_that_introduces_an_error_is_rejected(report, monkeypatch):
    from gridlint import engine, workbook

    wb = workbook.load(SAMPLES / "board-model.xlsx")
    graph = engine.build_graph(wb)
    base = engine.recalc(wb, graph)
    _patch_model(monkeypatch, {"edits": [{"cell": "Operating Model!C16", "formula": "=1/0"}],
                               "rationale": "x"})
    out = explain.propose_repair(wb, graph, base, report.findings[0])
    assert not out.accepted and "new error" in out.rejected_reason


def test_a_valid_proposed_repair_is_accepted_and_measured(report, monkeypatch):
    from gridlint import engine, workbook

    wb = workbook.load(SAMPLES / "board-model.xlsx")
    graph = engine.build_graph(wb)
    base = engine.recalc(wb, graph)
    _patch_model(monkeypatch, {"edits": [{"cell": "Operating Model!C16", "formula": "=SUM(C11:C15)"}],
                               "rationale": "include the Contractors row"})
    out = explain.propose_repair(wb, graph, base, report.findings[0])
    assert out.accepted
    assert out.changed_cells > 1
    assert out.fix.edits[0].new_formula == "=SUM(C11:C15)"


def test_a_declined_repair_is_reported_not_faked(report, monkeypatch):
    from gridlint import engine, workbook

    wb = workbook.load(SAMPLES / "board-model.xlsx")
    graph = engine.build_graph(wb)
    base = engine.recalc(wb, graph)
    _patch_model(monkeypatch, {"edits": [], "rationale": "I cannot repair this safely"})
    out = explain.propose_repair(wb, graph, base, report.findings[0])
    assert not out.accepted and out.rejected_reason == "model declined to repair"
    assert "cannot repair" in out.rationale
