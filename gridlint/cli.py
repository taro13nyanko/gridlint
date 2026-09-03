"""Command line interface.

    gridlint check FILE [--json] [--explain] [--fail-on critical]
    gridlint fix FILE [--out FILE] [--only R001] [--dry-run]
    gridlint verify FILE            # engine self-check only
    gridlint rules                  # what it looks for
    gridlint serve [--port 8000]

`check` exits 1 when a defect at or above --fail-on is found, so it can run in
CI over a workbook committed to a repository.
"""
from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

from . import __version__
from .audit import audit
from .rules import registry
from .rules.base import CRITICAL, INFO, WARNING

SEV_RANK = {CRITICAL: 0, WARNING: 1, INFO: 2}
COLOR = {CRITICAL: "\033[91m", WARNING: "\033[93m", INFO: "\033[94m"}
RESET = "\033[0m"
BOLD = "\033[1m"


def _supports_colour(stream) -> bool:
    return hasattr(stream, "isatty") and stream.isatty() and sys.platform != "emscripten"


def _money(v: float | None, currency: bool) -> str:
    if v is None:
        return ""

    return f"{'$' if currency else ''}{v:,.0f}"


def cmd_check(args: argparse.Namespace) -> int:
    report = audit(args.file, price=not args.no_price)
    data = report.to_dict()

    if args.explain:
        _add_explanations(report, data)

    if args.json:
        print(json.dumps(data, ensure_ascii=False, indent=2, default=str))
    else:
        _print_human(data, colour=_supports_colour(sys.stdout))

    threshold = SEV_RANK.get(args.fail_on, 0)
    worst = min((SEV_RANK.get(f["severity"], 3) for f in data["findings"]), default=9)
    return 1 if worst <= threshold else 0


def _add_explanations(report, data) -> None:
    from . import explain

    for finding, out in zip(report.findings, data["findings"]):
        # The same deterministic context the web app uses, so a recorded fixture
        # replays identically from either entry point.
        ctx = explain.sheet_context(data["path"], finding.sheet)
        text, guard = explain.explain_finding(finding, sheet_context=ctx)
        out["explanation"] = text
        out["explanation_rejected"] = None if guard.ok else guard.offending
    data["review_note"] = explain.review_note(data)


def _print_human(data: dict, *, colour: bool) -> None:
    def c(code: str, s: str) -> str:
        return f"{code}{s}{RESET}" if colour else s

    eng = data["engine"]
    print(c(BOLD, f"\n{data['path']}"))
    print(f"  {data['stats']['formulas']} formulas across {data['stats']['sheets']} sheet(s), "
          f"checked in {data['duration_ms']} ms")
    mark = "OK" if eng["trustworthy"] else "!!"
    print(f"  engine self-check [{mark}]: {eng['summary']}")
    if eng["unsupported"]:
        print(f"  {len(eng['unsupported'])} formula(s) use features the engine does not model")

    if not data["findings"]:
        print(c("\033[92m", "\n  No defects found.\n"))
        return

    counts = data["counts"]
    print(f"\n  {counts['critical']} critical, {counts['warning']} warning, {counts['info']} info\n")
    for i, f in enumerate(data["findings"], 1):
        sev = f["severity"]
        head = f"{i}. [{f['rule']}] {f['title']}"
        print(c(COLOR.get(sev, ""), c(BOLD, head)))
        where = f["cell"] + (f" and {f['group_size'] - 1} more" if f["group_size"] > 1 else "")
        print(f"     where: {where}")
        if f["formula"]:
            print(f"     formula: {f['formula']}")
        print(f"     {f['detail']}")
        if f.get("explanation"):
            print(f"     note: {f['explanation']}")
        for h in f.get("headline_changes", [])[:3]:
            print(c("\033[96m", f"     -> {h['label']}: {h['before']} becomes {h['after']}"))
        if f.get("fix"):
            status = {True: "verified by recalculation", False: "REJECTED", None: "not tested"}[f["fix_verified"]]
            print(f"     fix: {f['fix']['label']}  ({status}, {f['impact_cells']} cells change)")
            print(f"          {f['fix']['new_formula']}")
        print()
    if data.get("review_note"):
        print(c(BOLD, "  Review note"))
        print("   ", data["review_note"], "\n")


def cmd_fix(args: argparse.Namespace) -> int:
    from .apply import apply_fixes

    report = audit(args.file)
    chosen = [f for f in report.findings
              if f.fix and f.fix_verified and (not args.only or f.rule in args.only)]
    if not chosen:
        print("Nothing to apply: no verified fixes matched.")
        return 0
    out = Path(args.out) if args.out else Path(args.file).with_name(Path(args.file).stem + "-fixed.xlsx")
    if args.dry_run:
        for f in chosen:
            for e in f.fix.edits:
                print(f"{e.cell}: {e.old_formula}  ->  {e.new_formula}")
        return 0
    written = apply_fixes(args.file, chosen, out)
    print(f"Applied {len(chosen)} fix(es) across {written} cell(s).")
    print(f"Wrote {out}")
    print("Open it in Excel or Sheets to recalculate, then compare against the original.")
    return 0


def cmd_verify(args: argparse.Namespace) -> int:
    from . import engine, workbook

    wb = workbook.load(args.file)
    sc = engine.self_check(wb)
    print(f"{Path(args.file).name}: {sc.summary()}")
    for m in sc.mismatches[:20]:
        print(f"  {m['cell']}: saved {m['expected']!r}, recomputed {m['got']!r}   {m['formula']}")
    for u in sc.unsupported[:20]:
        print(f"  {u['cell']}: not modelled ({u['reason']})   {u['formula']}")
    return 0 if not sc.mismatches else 1


def cmd_rules(args: argparse.Namespace) -> int:
    for meta, _fn in registry():
        print(f"{meta.code}  {meta.default_severity:8}  {meta.name}")
        print(f"          {meta.why}")
    return 0


def cmd_serve(args: argparse.Namespace) -> int:
    import uvicorn

    uvicorn.run("gridlint.server:app", host=args.host, port=args.port, reload=args.reload)
    return 0


def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(prog="gridlint", description="Find spreadsheet defects and prove the fix.")
    p.add_argument("--version", action="version", version=f"gridlint {__version__}")
    sub = p.add_subparsers(dest="command", required=True)

    c = sub.add_parser("check", help="audit a workbook")
    c.add_argument("file")
    c.add_argument("--json", action="store_true", help="machine-readable output")
    c.add_argument("--explain", action="store_true", help="add plain-English notes (needs a model)")
    c.add_argument("--no-price", action="store_true", help="skip impact measurement (faster)")
    c.add_argument("--fail-on", choices=[CRITICAL, WARNING, INFO], default=CRITICAL,
                   help="exit 1 when a defect at this level or worse is found")
    c.set_defaults(func=cmd_check)

    f = sub.add_parser("fix", help="write a copy with verified fixes applied")
    f.add_argument("file")
    f.add_argument("--out")
    f.add_argument("--only", nargs="*", help="limit to these rule codes")
    f.add_argument("--dry-run", action="store_true")
    f.set_defaults(func=cmd_fix)

    v = sub.add_parser("verify", help="check the engine against the values the file already holds")
    v.add_argument("file")
    v.set_defaults(func=cmd_verify)

    r = sub.add_parser("rules", help="list the defects it looks for")
    r.set_defaults(func=cmd_rules)

    s = sub.add_parser("serve", help="run the web app")
    s.add_argument("--host", default="127.0.0.1")
    s.add_argument("--port", type=int, default=8000)
    s.add_argument("--reload", action="store_true")
    s.set_defaults(func=cmd_serve)
    return p


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    try:
        return args.func(args)
    except FileNotFoundError as e:
        print(f"gridlint: {e}", file=sys.stderr)
        return 2
    except KeyboardInterrupt:
        return 130


if __name__ == "__main__":
    sys.exit(main())
