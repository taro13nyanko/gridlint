"""The language model's job, and the fence around it.

Gridlint gives the model exactly three tasks, none of which involve deciding
what is wrong or what a number is:

  1. explain a defect in one sentence to somebody who is not an Excel expert
  2. propose a repair for defects with no mechanical fix
  3. write a short review note for the whole file

Both outputs are checked by code before anyone sees them:

  * every number in an explanation must already appear in the evidence the
    detector produced (`NumberGuard`), otherwise the sentence is thrown away
  * every proposed formula is parsed, executed by the engine, and compared
    against the original workbook; if it fails to parse, changes cells it was
    not supposed to touch, or introduces a new error, it is rejected

So a hallucination cannot reach the report as a fact, only as a discarded
suggestion. The deterministic sentence is always there to fall back to.
"""
from __future__ import annotations

import json
import re
from dataclasses import dataclass, field
from typing import Any

from . import engine
from .llm import FixtureMissing, LLMError, complete, provider_info
from .rules.base import Edit, Finding, Fix
from .workbook import Workbook

SYSTEM = (
    "You explain spreadsheet defects to finance and operations people who are not "
    "Excel experts. You never compute or invent numbers: every figure you mention "
    "must already appear in the evidence you were given. You never claim something "
    "is wrong that the evidence does not state. Write plainly, in the second person, "
    "with no jargon and no exclamation marks."
)

_NUM_RE = re.compile(r"-?\d[\d,]*(?:\.\d+)?%?")
#: Small integers are ordinary English ("one row", "12 cells"), not claims about data.
_ORDINALS = {float(i) for i in range(0, 101)}


@dataclass
class GuardResult:
    ok: bool
    offending: list[str] = field(default_factory=list)

    def __bool__(self) -> bool:
        return self.ok


class NumberGuard:
    """Rejects any sentence containing a figure that is not in the evidence."""

    def __init__(self, allowed: list[float]):
        self.allowed = {round(float(a), 4) for a in allowed} | _ORDINALS

    def check(self, text: str) -> GuardResult:
        bad: list[str] = []
        for raw in _NUM_RE.findall(text or ""):
            v = _parse(raw)
            if v is None:
                continue
            if not any(abs(v - a) <= max(0.01, abs(a) * 1e-6) for a in self.allowed):
                bad.append(raw)
        return GuardResult(not bad, bad)


def _parse(raw: str) -> float | None:
    t = raw.replace(",", "")
    pct = t.endswith("%")
    t = t.rstrip("%")
    try:
        v = float(t)
    except ValueError:
        return None
    return v / 100 if pct else v


def allowed_numbers(f: Finding) -> list[float]:
    """Every figure the detector actually measured, which is all the model may cite."""
    out: list[float] = []

    def add(v: Any) -> None:
        if isinstance(v, bool) or v is None:
            return
        if isinstance(v, (int, float)):
            out.append(float(v))
            out.append(round(float(v)))
            out.append(round(float(v), 1))
            out.append(round(float(v), 2))
        elif isinstance(v, str):
            p = _parse(v)
            if p is not None:
                out.append(p)
        elif isinstance(v, (list, tuple)):
            for x in v:
                add(x)
        elif isinstance(v, dict):
            for x in v.values():
                add(x)

    add(f.evidence)
    add(f.impact_value)
    add(f.impact_before)
    add(f.impact_after)
    add(f.impact_cells)
    add(f.group_size)
    for h in f.headline_changes:
        add(h.get("before"))
        add(h.get("after"))
    for d in f.fix_diff:
        add(d.get("before"))
        add(d.get("after"))
    # cell addresses contain row numbers the model will naturally repeat
    for ref in [f.cell, *f.related, *f.group_cells]:
        for m in re.findall(r"\d+", ref or ""):
            out.append(float(m))
    return out


def _context(f: Finding) -> dict[str, Any]:
    return {
        "cell": f.cell,
        "formula": f.formula,
        "what_the_detector_found": f.detail,
        "evidence": f.evidence,
        "cells_affected": f.impact_cells,
        "same_mistake_repeated_in": f.group_size,
        "numbers_that_change_if_fixed": [
            {"label": h["label"], "from": h["before"], "to": h["after"]} for h in f.headline_changes
        ],
        "suggested_fix": f.fix.new_formula if f.fix else None,
    }


def numbers_in(text: str) -> list[float]:
    """Figures that appear in text handed to the model, which it may therefore repeat."""
    out = []
    for raw in _NUM_RE.findall(text or ""):
        v = _parse(raw)
        if v is not None:
            out.append(v)
    return out


def explain_finding(f: Finding, *, sheet_context: str = "") -> tuple[str | None, GuardResult]:
    """One short paragraph for a non-expert. Returns (text or None, guard result).

    The invariant the guard enforces: the model may only state figures that were
    given to it. Anything else is a number it made up, and the sentence is dropped.
    """
    guard = NumberGuard(allowed_numbers(f) + numbers_in(sheet_context))
    prompt = (
        "A spreadsheet checker found this defect. Write ONE short paragraph (at most "
        "three sentences) telling the file's owner what is wrong and why it matters to "
        "them, in plain English. Mention the business consequence, not the Excel "
        "mechanics. Do not repeat the cell address more than once. Use only figures "
        "that appear in the JSON below.\n\n"
        f"Sheet context:\n{sheet_context}\n\n"
        f"Defect:\n{json.dumps(_context(f), ensure_ascii=False, indent=1, default=str)}"
    )
    try:
        r = complete([{"role": "user", "content": prompt}], system=SYSTEM,
                     max_tokens=300, tag="explain_v1")
    except (FixtureMissing, LLMError):
        return None, GuardResult(True)
    text = r.text.strip()
    check = guard.check(text)
    if not check.ok:
        return None, check
    return text, check


def review_note(report_dict: dict[str, Any]) -> str | None:
    """Three sentences the reviewer can paste into an email."""
    top = report_dict.get("findings", [])[:5]
    # Every field is named explicitly rather than passing report structures through.
    # A prompt built from whole dicts changes the moment a new key is added to the
    # report, which silently invalidates every recorded fixture.
    payload = {
        "file": report_dict.get("path"),
        "critical": report_dict["counts"]["critical"],
        "warnings": report_dict["counts"]["warning"],
        "engine_agreement": report_dict["engine"]["summary"],
        "defects": [
            {"what": f["title"], "where": f["cell"], "detail": f["detail"],
             "changes_if_fixed": [{"label": h["label"], "from": h["before"], "to": h["after"]}
                                  for h in f.get("headline_changes", [])[:3]]}
            for f in top
        ],
    }
    allowed: list[float] = []
    for f in top:
        allowed += allowed_numbers(_finding_from_dict(f))
    allowed += [report_dict["counts"]["critical"], report_dict["counts"]["warning"],
                report_dict["counts"]["total"], report_dict["engine"]["checked"],
                report_dict["engine"]["matched"]]
    guard = NumberGuard(allowed)
    prompt = (
        "Write a three-sentence note the reviewer can send to whoever owns this "
        "spreadsheet. Say what was checked, what the most serious problem is, and what "
        "to do next. No greeting, no sign-off, no bullet points. Use only figures that "
        "appear below.\n\n" + json.dumps(payload, ensure_ascii=False, indent=1, default=str)
    )
    try:
        r = complete([{"role": "user", "content": prompt}], system=SYSTEM,
                     max_tokens=300, tag="review_v1")
    except (FixtureMissing, LLMError):
        return None
    text = r.text.strip()
    return text if guard.check(text).ok else None


def sheet_context(report_path: str, sheet: str) -> str:
    """Built server-side and only from the report, so the same finding always
    produces the same prompt -- which is what makes the recorded fixtures replay."""
    return f"Workbook {report_path}, sheet {sheet}. Figures are exactly as stored in the file."


def _finding_from_dict(d: dict[str, Any]) -> Finding:
    f = Finding(rule=d.get("rule", ""), title=d.get("title", ""), severity=d.get("severity", "info"),
                cell=d.get("cell", ""), sheet=d.get("sheet", ""), col=d.get("col", 0),
                row=d.get("row", 0), detail=d.get("detail", ""), formula=d.get("formula"))
    fix = d.get("fix")
    if fix:
        f.fix = Fix(edits=[Edit(e["cell"], e["new_formula"], e.get("old_formula"))
                           for e in fix.get("edits", [])],
                    label=fix.get("label", ""))
    f.group_cells = d.get("group_cells", [])
    f.related = d.get("related", [])
    f.confidence = d.get("confidence", 1.0)
    f.evidence = d.get("evidence", {})
    f.impact_value = d.get("impact_value")
    f.impact_before = d.get("impact_before")
    f.impact_after = d.get("impact_after")
    f.impact_cells = d.get("impact_cells", 0)
    f.group_size = d.get("group_size", 1)
    f.headline_changes = d.get("headline_changes", [])
    f.fix_diff = d.get("fix_diff", [])
    return f


# ---------------------------------------------------------------------------
# Proposed repairs: written by the model, then compiled and executed before use.
# ---------------------------------------------------------------------------

REPAIR_SYSTEM = (
    "You repair spreadsheet formulas. You reply with JSON only, in the form "
    '{"edits":[{"cell":"Sheet!A1","formula":"=..."}],"rationale":"one sentence"}. '
    "Every formula must be valid Excel starting with '='. Change as few cells as "
    "possible. Never invent a number that is not already in the workbook; if a "
    "constant must move somewhere, put it in the empty cell you are told to use. "
    'If you cannot repair it safely, reply {"edits":[],"rationale":"why not"}.'
)


@dataclass
class ProposedFix:
    fix: Fix | None
    rationale: str = ""
    rejected_reason: str | None = None
    changed_cells: int = 0
    proposed_raw: dict[str, Any] | None = None

    @property
    def accepted(self) -> bool:
        return self.fix is not None and self.rejected_reason is None


def propose_repair(wb: Workbook, graph: engine.Graph, baseline: dict, f: Finding,
                   *, spare_cell: str | None = None) -> ProposedFix:
    """Ask for a repair, then prove it before offering it.

    A proposal is accepted only if every formula parses, the workbook still
    computes, no new error appears anywhere, and only cells that are supposed to
    move actually move.
    """
    payload = {
        "defect": _context(f),
        "rule": f.rule,
        "cell_to_repair": f.cell,
        "empty_cell_you_may_use": spare_cell,
        "instruction": {
            "R003": "Move the hard-coded number into the empty cell provided, label it, "
                    "and reference that cell from the formula.",
            "R012": "Replace the IFERROR with a check that handles the real cause, so a "
                    "genuine failure is still visible.",
            "R004": "Make the missing input explicit rather than silently treating it as zero.",
        }.get(f.rule, "Repair the formula so the defect described above is gone."),
    }
    try:
        r = complete([{"role": "user", "content": json.dumps(payload, ensure_ascii=False,
                                                             indent=1, default=str)}],
                     system=REPAIR_SYSTEM, json_mode=True, max_tokens=600, tag="repair_v1")
        data = r.json()
    except (FixtureMissing, LLMError) as e:
        return ProposedFix(None, rejected_reason=f"model unavailable: {e}")
    except Exception as e:
        return ProposedFix(None, rejected_reason=f"model returned unusable output: {e}")

    edits_in = data.get("edits") or []
    rationale = str(data.get("rationale", ""))[:400]
    if not edits_in:
        return ProposedFix(None, rationale=rationale, rejected_reason="model declined to repair",
                           proposed_raw=data)

    overrides: dict[tuple[str, int, int], str] = {}
    edits: list[Edit] = []
    for e in edits_in[:8]:
        ref, formula = str(e.get("cell", "")), str(e.get("formula", ""))
        if not formula.startswith("="):
            return ProposedFix(None, rationale=rationale, proposed_raw=data,
                               rejected_reason=f"proposed value for {ref} is not a formula")
        key = _key(ref)
        if key is None or key[0] not in wb.sheets:
            return ProposedFix(None, rationale=rationale, proposed_raw=data,
                               rejected_reason=f"proposed edit points at {ref}, which is not in this file")
        try:
            from .formula.parser import parse_formula
            parse_formula(formula)
        except Exception as pe:
            return ProposedFix(None, rationale=rationale, proposed_raw=data,
                               rejected_reason=f"proposed formula does not parse: {pe}")
        overrides[key] = formula
        old = wb.get(*key)
        edits.append(Edit(ref, formula, old.formula if old else None))

    try:
        after = engine.recalc(wb, graph, formula_overrides=overrides)
    except Exception as ex:
        return ProposedFix(None, rationale=rationale, proposed_raw=data,
                           rejected_reason=f"workbook failed to recompute with the repair: {ex}")

    new_errors = [k for k, v in after.items()
                  if _is_err(v) and not _is_err(baseline.get(k))]
    if new_errors:
        return ProposedFix(None, rationale=rationale, proposed_raw=data,
                           rejected_reason=f"repair introduces {len(new_errors)} new error cell(s)")
    changed = [k for k, v in after.items() if not engine.values_equal(baseline.get(k), v)]
    return ProposedFix(Fix(edits=edits, label=rationale or "Proposed repair"),
                       rationale=rationale, changed_cells=len(changed), proposed_raw=data)


def _is_err(v: Any) -> bool:
    from .formula.values import ExcelError
    return isinstance(v, ExcelError)


def _key(ref: str):
    if "!" not in ref:
        return None
    sheet, _, a1 = ref.rpartition("!")
    col, i = 0, 0
    while i < len(a1) and a1[i].isalpha():
        col = col * 26 + (ord(a1[i].upper()) - 64)
        i += 1
    if not col or not a1[i:].isdigit():
        return None
    return (sheet, col, int(a1[i:]))


def status() -> dict[str, Any]:
    info = provider_info()
    return {**info, "explanations": info["provider"] != "replay" or True}
