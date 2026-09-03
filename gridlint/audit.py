"""Run every rule, price each finding by recalculation, and rank by money at risk.

The ordering matters more than the detection. A reviewer given 200 warnings
reads none of them; a reviewer told "this one moves 2.1 million yen" reads that
one. Pricing a finding means applying its fix to a copy of the workbook,
recomputing, and diffing -- so the number is measured, never estimated.
"""
from __future__ import annotations

import time
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

from . import engine
from .formula.parser import index_to_col
from .formula.values import BLANK, ExcelError
from .rules import Finding, registry
from .rules.base import CRITICAL, WARNING, format_kind, is_currency_format
from .workbook import Workbook, load

#: A fix that changes more than this many cells is reported but not auto-applied.
LARGE_BLAST = 500


@dataclass
class Report:
    path: str
    findings: list[Finding] = field(default_factory=list)
    stats: dict[str, Any] = field(default_factory=dict)
    self_check: engine.SelfCheck | None = None
    duration_ms: int = 0
    sheet_previews: dict[str, Any] = field(default_factory=dict)

    @property
    def critical(self) -> list[Finding]:
        return [f for f in self.findings if f.severity == CRITICAL]

    @property
    def money_at_risk(self) -> float:
        """Largest single measured impact, not a sum: impacts overlap and must not be added."""
        vals = [f.impact_value for f in self.findings if f.impact_value]
        return max(vals) if vals else 0.0

    def headline(self) -> str:
        if not self.findings:
            return "No defects found."
        top = self.findings[0]
        if top.impact_value:
            return f"{top.title} in {top.cell}, moving {top.impact_value:,.0f}."
        return f"{top.title} in {top.cell}."

    def to_dict(self) -> dict[str, Any]:
        sc = self.self_check
        return {
            "path": Path(self.path).name,
            "stats": self.stats,
            "duration_ms": self.duration_ms,
            "engine": {
                "checked": sc.checked if sc else 0,
                "matched": sc.matched if sc else 0,
                "agreement": round(sc.agreement, 4) if sc else 0.0,
                "trustworthy": bool(sc.trustworthy) if sc else False,
                "unsupported": (sc.unsupported[:20] if sc else []),
                "summary": sc.summary() if sc else "not run",
            },
            "counts": {
                "total": len(self.findings),
                "critical": sum(1 for f in self.findings if f.severity == CRITICAL),
                "warning": sum(1 for f in self.findings if f.severity == WARNING),
                "info": sum(1 for f in self.findings if f.severity not in (CRITICAL, WARNING)),
            },
            "money_at_risk": self.money_at_risk,
            "headline": self.headline(),
            "findings": [f.to_dict() for f in self.findings],
            "sheets": self.sheet_previews,
        }


def audit(path: str | Path, *, price: bool = True, max_priced: int = 40,
          preview_rows: int = 60, preview_cols: int = 14) -> Report:
    t0 = time.time()
    wb = load(path)
    graph = engine.build_graph(wb)
    computed = engine.recalc(wb, graph)
    sc = engine.self_check(wb, computed, graph)

    findings: list[Finding] = []
    for meta, fn in registry():
        try:
            for f in fn(wb=wb, computed=computed, graph=graph, self_check=sc):
                findings.append(f)
        except Exception as e:                         # a broken rule must not lose the others
            findings.append(Finding(
                rule=meta.code, title=f"Rule {meta.code} could not run", severity="info",
                cell="", sheet="", col=0, row=0,
                detail=f"{type(e).__name__}: {e}", confidence=0.0,
            ))

    findings = _group(_dedupe(findings))
    if price:
        _price(wb, graph, computed, findings, max_priced)
    findings.sort(key=lambda f: f.sort_key)

    report = Report(
        path=str(path), findings=findings, self_check=sc,
        stats={**wb.stats(), "cycles": len(graph.cycles),
               "has_cached_values": wb.has_cached_values},
        duration_ms=int((time.time() - t0) * 1000),
    )
    report.sheet_previews = _previews(wb, computed, findings, preview_rows, preview_cols)
    return report


def _dedupe(findings: list[Finding]) -> list[Finding]:
    seen: dict[tuple[str, str], Finding] = {}
    for f in findings:
        key = (f.rule, f.cell)
        prev = seen.get(key)
        if prev is None or f.confidence > prev.confidence:
            seen[key] = f
    return list(seen.values())


#: Rules where the same mistake is normally copied across a row or column of totals.
GROUPABLE = {"R001", "R002", "R004", "R008"}


def _group(findings: list[Finding]) -> list[Finding]:
    """Collapse one mistake copied across a row or column into a single finding.

    Six monthly totals that all forget the same line item are one defect a
    person has to think about once, not six items in a list.
    """
    from collections import defaultdict

    buckets: dict[tuple, list[Finding]] = defaultdict(list)
    passthrough: list[Finding] = []
    for f in findings:
        if f.rule not in GROUPABLE or f.fix is None:
            passthrough.append(f)
            continue
        sig = (f.rule, f.sheet, f.evidence.get("missing_label"), f.evidence.get("side"),
               f.evidence.get("pattern"), f.title)
        buckets[(sig, "row", f.row)].append(f)

    out = list(passthrough)
    used: set[str] = set()
    for (_sig, _axis, _line), group in sorted(buckets.items(), key=lambda kv: str(kv[0])):
        group.sort(key=lambda f: f.col)
        for run in _consecutive(group, lambda f: f.col):
            if len(run) == 1:
                out.append(run[0])
                used.add(run[0].id)
                continue
            head = run[0]
            head.group_size = len(run)
            head.group_cells = [f.cell for f in run]
            head.confidence = max(head.confidence, min(0.97, head.confidence + 0.05 * (len(run) - 1)))
            edits = [e for f in run if f.fix for e in f.fix.edits]
            span = f"{run[0].cell} to {run[-1].cell.split('!')[-1]}"
            head.fix = type(head.fix)(edits=edits,
                                      label=f"{head.fix.label} in all {len(run)} cells ({span})")
            head.detail = head.detail.rstrip(".") + f". The same mistake is repeated in {len(run)} cells ({span})."
            out.append(head)
    return out


def _consecutive(items: list[Finding], key) -> list[list[Finding]]:
    runs: list[list[Finding]] = []
    cur: list[Finding] = []
    prev = None
    for it in items:
        k = key(it)
        if prev is not None and k != prev + 1:
            runs.append(cur)
            cur = []
        cur.append(it)
        prev = k
    if cur:
        runs.append(cur)
    return runs


def _price(wb: Workbook, graph: engine.Graph, baseline: dict, findings: list[Finding],
           max_priced: int) -> None:
    """Measure each fixable finding by recomputing the workbook with the fix applied."""
    candidates = [f for f in findings if f.fix is not None]
    candidates.sort(key=lambda f: (0 if f.severity == CRITICAL else 1, -f.confidence))
    for f in candidates[:max_priced]:
        key = _key(f.cell)
        if key is None:
            continue
        overrides = {}
        for e in f.fix.edits:
            ek = _key(e.cell)
            if ek is not None:
                overrides[ek] = e.new_formula
        if not overrides:
            continue
        try:
            after = engine.recalc(wb, graph, formula_overrides=overrides)
        except Exception:
            f.fix_verified = False
            continue
        diff = []
        biggest = 0.0
        biggest_cell = None
        before_v = after_v = None
        for k, new in after.items():
            old = baseline.get(k)
            if engine.values_equal(old, new):
                continue
            cell = wb.get(*k)
            entry = {
                "cell": f"{k[0]}!{index_to_col(k[1])}{k[2]}",
                "before": _plain(old), "after": _plain(new),
                "currency": is_currency_format(cell.number_format) if cell else False,
                "fmt": format_kind(cell.number_format) if cell else "plain",
            }
            diff.append(entry)
            if isinstance(old, (int, float)) and isinstance(new, (int, float)) \
                    and not isinstance(old, bool) and not isinstance(new, bool):
                delta = abs(float(new) - float(old))
                weight = delta * (3.0 if entry["currency"] else 1.0)
                if weight > biggest:
                    biggest, biggest_cell = weight, entry["cell"]
                    before_v, after_v = float(old), float(new)
        f.impact_cells = len(diff)
        f.fix_diff = sorted(diff, key=lambda d: -abs(_num(d["after"]) - _num(d["before"])))[:30]
        f.headline_changes = _headline_changes(wb, graph, diff)
        f.trace = _trace_for(wb, graph, baseline, f, diff)
        f.fix_verified = _fix_is_sane(f, baseline, after)
        if biggest_cell is not None:
            f.impact_value = abs(after_v - before_v)
            f.impact_cell = biggest_cell
            f.impact_before = before_v
            f.impact_after = after_v
            f.impact_currency = any(d["currency"] for d in f.fix_diff)


def _trace_for(wb: Workbook, graph: engine.Graph, baseline: dict, f: Finding,
               diff: list[dict]) -> list[dict]:
    """Where the headline number comes from, so a reviewer can follow it back.

    The walk starts at the summary line a person reads -- runway, margin -- not
    at the cell holding the defect, because that is the number they are asking
    about.
    """
    if not f.headline_changes:
        return []
    start = _key(f.headline_changes[0]["cell"])
    if start is None:
        return []
    changed = {k for k in (_key(d["cell"]) for d in diff) if k is not None}
    steps = engine.trace(wb, graph, baseline, start, changed=changed)
    return [{"cell": s.cell, "label": s.label, "formula": s.formula, "value": s.value,
             "depth": s.depth, "is_input": s.is_input, "changed": s.changed}
            for s in steps]


def _headline_changes(wb: Workbook, graph: engine.Graph, diff: list[dict]) -> list[dict]:
    """The changed cells a person actually reads: labelled summary cells that
    nothing else is computed from. These are the numbers on the slide."""
    out = []
    for d in diff:
        key = _key(d["cell"])
        if key is None:
            continue
        if graph.dependents.get(key):
            continue                                   # something is computed from it: not a headline
        label = _row_label(wb, key[0], key[2], key[1])
        if not label:
            continue
        out.append({"cell": d["cell"], "label": label, "before": d["before"],
                    "after": d["after"], "currency": d["currency"], "fmt": d.get("fmt", "plain")})
    out.sort(key=lambda d: -abs(_num(d["after"]) - _num(d["before"])))
    return out[:6]


def _row_label(wb: Workbook, sheet: str, row: int, before_col: int) -> str | None:
    for c in range(1, min(before_col, 8)):
        cell = wb.get(sheet, c, row)
        if cell is not None and isinstance(cell.static, str) and cell.static.strip():
            return cell.static.strip()
    return None


def _fix_is_sane(f: Finding, baseline: dict, after: dict) -> bool:
    """A fix is only offered if every edited cell computes, and the fix adds no new error."""
    for e in f.fix.edits if f.fix else []:
        k = _key(e.cell)
        if k is None:
            return False
        if isinstance(after.get(k), ExcelError):
            return False
    new_errors = sum(1 for k, v in after.items()
                     if isinstance(v, ExcelError) and not isinstance(baseline.get(k), ExcelError))
    return new_errors == 0


def _key(ref: str):
    if "!" not in ref:
        return None
    sheet, _, a1 = ref.rpartition("!")
    col = 0
    i = 0
    while i < len(a1) and a1[i].isalpha():
        col = col * 26 + (ord(a1[i].upper()) - 64)
        i += 1
    if not col or not a1[i:].isdigit():
        return None
    return (sheet, col, int(a1[i:]))


def _num(v: Any) -> float:
    return float(v) if isinstance(v, (int, float)) and not isinstance(v, bool) else 0.0


def _plain(v: Any) -> Any:
    if isinstance(v, ExcelError):
        return v.code
    if v is BLANK or v is None:
        return None
    if isinstance(v, float) and v == int(v) and abs(v) < 1e15:
        return int(v)
    if isinstance(v, float):
        return round(v, 6)
    return v


def _previews(wb: Workbook, computed: dict, findings: list[Finding],
              max_rows: int, max_cols: int) -> dict[str, Any]:
    """A small window of each sheet so the web view can draw the grid."""
    flagged = {f.cell for f in findings}
    out: dict[str, Any] = {}
    for sheet in wb.sheet_list:
        rows = []
        n_rows = min(sheet.max_row, max_rows)
        n_cols = min(sheet.max_col, max_cols)
        for r in range(1, n_rows + 1):
            row = []
            for c in range(1, n_cols + 1):
                cell = sheet.cell(c, r)
                if cell is None:
                    row.append(None)
                    continue
                key = (sheet.name, c, r)
                value = computed.get(key, cell.static if cell.static is not None else cell.cached)
                row.append({
                    "v": _plain(value),
                    "f": cell.formula,
                    "flag": f"{sheet.name}!{index_to_col(c)}{r}" in flagged or None,
                })
            rows.append(row)
        out[sheet.name] = {"rows": rows, "max_row": sheet.max_row, "max_col": sheet.max_col,
                           "shown_rows": n_rows, "shown_cols": n_cols}
    return out
