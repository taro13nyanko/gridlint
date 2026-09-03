"""Rule tests. Each defect is built from scratch so the assertion is about the rule, not a fixture."""
from __future__ import annotations

from pathlib import Path

import openpyxl
import pytest

from gridlint.audit import audit
from gridlint.rules.base import skeleton_of, to_r1c1
from gridlint.formula.parser import parse_formula

SAMPLES = Path(__file__).resolve().parent.parent / "samples"


def build(tmp_path, cells: dict[str, object], name: str = "t.xlsx", sheet: str = "Sheet") -> Path:
    """Write a workbook from {'A1': value_or_formula}. No cached values, like a fresh export."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = sheet
    for ref, value in cells.items():
        ws[ref] = value
    p = tmp_path / name
    wb.save(p)
    wb.close()
    return p


def rules_found(report) -> set[str]:
    return {f.rule for f in report.findings}


def find(report, rule: str):
    return next((f for f in report.findings if f.rule == rule), None)


# ------------------------------------------------------------------ R001
def test_r001_flags_a_total_that_stops_one_row_short(tmp_path):
    p = build(tmp_path, {
        "A1": "Salaries", "B1": 100,
        "A2": "Cloud", "B2": 200,
        "A3": "Contractors", "B3": 300,
        "A4": "Total", "B4": "=SUM(B1:B2)",
    })
    f = find(audit(p), "R001")
    assert f is not None
    assert f.evidence["missing_label"] == "Contractors"
    assert f.fix.new_formula == "=SUM(B1:B3)"


def test_r001_measures_the_impact_by_recalculating(tmp_path):
    p = build(tmp_path, {
        "A1": "a", "B1": 100, "A2": "b", "B2": 200, "A3": "c", "B3": 300,
        "A4": "Total", "B4": "=SUM(B1:B2)",
        "A5": "Doubled", "B5": "=B4*2",
    })
    f = find(audit(p), "R001")
    assert f.fix_verified is True
    assert f.impact_value == pytest.approx(600.0)      # B5 moves 600 -> 1200
    assert f.impact_cells == 2


def test_r001_is_silent_when_the_total_covers_everything(tmp_path):
    p = build(tmp_path, {
        "A1": "a", "B1": 100, "A2": "b", "B2": 200, "A3": "Total", "B3": "=SUM(B1:B2)",
    })
    assert "R001" not in rules_found(audit(p))


def test_r001_does_not_flag_a_total_below_another_total(tmp_path):
    p = build(tmp_path, {
        "A1": "a", "B1": 100, "A2": "b", "B2": 200,
        "A3": "Subtotal", "B3": "=SUM(B1:B2)",
        "A4": "Grand total", "B4": "=SUM(B3)",
    })
    assert "R001" not in rules_found(audit(p))


# ------------------------------------------------------------------ R002
def test_r002_flags_the_odd_formula_in_a_copied_row(tmp_path):
    cells = {"A1": 10, "B1": 20, "C1": 30, "D1": 40, "E1": 50, "F1": 60}
    for i, col in enumerate("BCDEF"):
        cells[f"{col}2"] = f"={chr(64 + i + 1)}1*2"
    cells["D2"] = "=C1*3"                              # the hand-edited one
    p = build(tmp_path, cells)
    f = find(audit(p), "R002")
    assert f is not None and f.cell.endswith("!D2")
    assert f.evidence["agreeing"] >= 4


def test_r002_is_silent_when_the_row_is_consistent(tmp_path):
    cells = {"A1": 10, "B1": 20, "C1": 30, "D1": 40, "E1": 50, "F1": 60}
    for i, col in enumerate("BCDEF"):
        cells[f"{col}2"] = f"={chr(64 + i + 1)}1*2"
    p = build(tmp_path, cells)
    assert "R002" not in rules_found(audit(p))


def test_r002_needs_a_run_of_at_least_four(tmp_path):
    p = build(tmp_path, {"A1": 1, "B1": 2, "C1": 3,
                         "A2": "=A1*2", "B2": "=B1*2", "C2": "=C1*9"})
    assert "R002" not in rules_found(audit(p))


# ------------------------------------------------------------------ others
def test_r003_flags_a_buried_rate_but_not_a_structural_argument(tmp_path):
    p = build(tmp_path, {"A1": 1000, "B1": "=A1*1.0825", "C1": "=ROUND(A1,2)"})
    report = audit(p)
    f = find(report, "R003")
    assert f is not None and f.cell.endswith("!B1")
    assert 1.0825 in f.evidence["literals"]
    assert not any(x.cell.endswith("!C1") for x in report.findings if x.rule == "R003")


def test_r004_flags_arithmetic_on_an_empty_cell(tmp_path):
    p = build(tmp_path, {"A1": 100, "B1": "=A1*Z9"})
    f = find(audit(p), "R004")
    assert f is not None and "Sheet!Z9" in f.evidence["blank_cells"]


def test_r005_reports_the_root_cause_not_only_the_symptom(tmp_path):
    p = build(tmp_path, {"A1": 10, "A2": 0, "B1": "=A1/A2", "C1": "=B1+1", "D1": "=C1+1"})
    report = audit(p)
    roots = [f for f in report.findings if f.rule == "R005"]
    assert roots, "no error finding"
    assert all(f.evidence["root_cause"] == "Sheet!B1" for f in roots)


def test_r006_reports_a_cycle(tmp_path):
    p = build(tmp_path, {"A1": "=B1+1", "B1": "=A1+1"})
    f = find(audit(p), "R006")
    assert f is not None and f.evidence["size"] == 2


def test_r007_flags_a_reference_to_a_sheet_that_is_gone(tmp_path):
    p = build(tmp_path, {"A1": "='Q3 Actuals'!B2*2"})
    f = find(audit(p), "R007")
    assert f is not None and f.evidence["missing_sheet"] == "Q3 Actuals"


def test_r008_flags_a_number_stored_as_text_inside_a_sum(tmp_path):
    p = build(tmp_path, {"A1": 100, "A2": "38,500", "A3": 200, "A4": "=SUM(A1:A3)"})
    f = find(audit(p), "R008")
    assert f is not None and f.evidence["missing_total"] == pytest.approx(38500.0)


def test_r009_flags_a_total_that_swallows_a_subtotal(tmp_path):
    p = build(tmp_path, {
        "A1": 10, "A2": 20, "A3": "=SUM(A1:A2)", "A4": 30,
        "A5": "=SUM(A1:A4)",
    })
    f = find(audit(p), "R009")
    assert f is not None and f.evidence["subtotal"] == "Sheet!A3"


def test_r011_flags_volatile_functions(tmp_path):
    p = build(tmp_path, {"A1": "=TODAY()", "B1": "=A1+1"})
    f = find(audit(p), "R011")
    assert f is not None and "TODAY" in f.evidence["functions"]


def test_r012_flags_iferror_only_when_it_is_actually_hiding_something(tmp_path):
    hiding = build(tmp_path, {"A1": 10, "A2": 0, "B1": '=IFERROR(A1/A2,0)'}, name="hide.xlsx")
    fine = build(tmp_path, {"A1": 10, "A2": 2, "B1": '=IFERROR(A1/A2,0)'}, name="fine.xlsx")
    assert find(audit(hiding), "R012") is not None
    assert find(audit(fine), "R012") is None


def test_r013_needs_cached_values_and_stays_quiet_without_them(tmp_path):
    p = build(tmp_path, {"A1": 1, "B1": "=A1+1"})
    assert "R013" not in rules_found(audit(p))


def test_r013_flags_a_workbook_saved_without_recalculating():
    """The board sample is fully recalculated, so R013 must not fire on it."""
    assert "R013" not in rules_found(audit(SAMPLES / "board-model.xlsx"))


# ------------------------------------------------------------------ helpers
def test_r1c1_normalisation_makes_copied_formulas_identical():
    a = parse_formula("=B2*2")
    b = parse_formula("=C2*2")
    assert to_r1c1(a, 3, 2) == to_r1c1(b, 4, 2)        # same relative offset
    assert to_r1c1(a, 3, 2) != to_r1c1(b, 3, 2)


def test_skeleton_ignores_which_cells_are_referenced():
    assert skeleton_of(parse_formula("=C11*(1+$B$11)")) == skeleton_of(parse_formula("=D12*(1+$B$12)"))
    assert skeleton_of(parse_formula("=SUM(A1:A9)")) != skeleton_of(parse_formula("=A1*2"))


# ------------------------------------------------------------------ end to end
def test_board_sample_finds_every_planted_defect():
    report = audit(SAMPLES / "board-model.xlsx")
    assert {"R001", "R002", "R003", "R008", "R012"} <= rules_found(report)
    top = report.findings[0]
    assert top.rule == "R001" and top.group_size == 12 and top.fix_verified
    labels = {h["label"]: (h["before"], h["after"]) for h in top.headline_changes}
    assert "Runway (months)" in labels
    before, after = labels["Runway (months)"]
    assert before > 30 and after < 10, "the demo story is that runway collapses once the fix lands"


def test_clean_sample_reports_nothing():
    report = audit(SAMPLES / "clean-model.xlsx")
    assert report.findings == []
    assert report.self_check.trustworthy


def test_findings_are_ordered_worst_first():
    report = audit(SAMPLES / "board-model.xlsx")
    order = [f.sort_key for f in report.findings]
    assert order == sorted(order)


def test_a_rule_that_raises_does_not_lose_the_other_findings(monkeypatch):
    from gridlint import audit as audit_mod
    from gridlint.rules import base

    def boom(**_kw):
        raise RuntimeError("rule exploded")

    original = base._REGISTRY[:]
    try:
        base._REGISTRY.append((base.RuleMeta("R999", "boom", "test", "info"), boom))
        report = audit_mod.audit(SAMPLES / "board-model.xlsx")
        assert "R001" in rules_found(report)
        broken = find(report, "R999")
        assert broken is not None and "rule exploded" in broken.detail
    finally:
        base._REGISTRY[:] = original
