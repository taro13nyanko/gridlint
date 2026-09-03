"""Generate a corpus of structurally varied, defect-free workbooks.

The benchmark needs both halves of the question. Mutants measure recall; this
corpus measures the other half: on a healthy file, the right number of findings
is zero. The generator deliberately produces the shapes that trip naive linters
-- blank spacer rows, subtotals summed into a grand total, cross-sheet
assumptions, lookup tables, percentage rows, text columns -- because a tool that
cries wolf on those is one nobody keeps installed.

    python bench/make_corpus.py --n 12
"""
from __future__ import annotations

import argparse
import random
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "bench" / "corpus"

MONEY = "$#,##0"
PCT = "0.0%"
SECTION_NAMES = ["REVENUE", "COST OF SALES", "OPERATING COSTS", "HEADCOUNT COSTS", "MARKETING"]
LINE_ITEMS = [
    "Subscriptions", "Services", "Licences", "Support", "Training",
    "Salaries", "Benefits", "Cloud hosting", "Data vendors", "Travel",
    "Office rent", "Software tools", "Contractors", "Recruiting", "Insurance",
    "Events", "Advertising", "Content", "Agency fees", "Sponsorships",
]
REGIONS = ["North", "South", "East", "West"]


def col(n: int) -> str:
    from openpyxl.utils import get_column_letter
    return get_column_letter(n)


def build_one(app, path: Path, rng: random.Random) -> dict:
    n_sections = rng.randint(1, 3)
    n_months = rng.choice([3, 6, 12])
    spacer = rng.random() < 0.7
    use_lookup = rng.random() < 0.4
    use_pct_row = rng.random() < 0.5
    assumptions_sheet = rng.random() < 0.7

    wb = app.Workbooks.Add()
    while wb.Worksheets.Count < (2 if assumptions_sheet else 1):
        wb.Worksheets.Add()
    ws = wb.Worksheets(1)
    ws.Name = "Model"
    inputs = wb.Worksheets(2) if assumptions_sheet else None
    if inputs is not None:
        inputs.Name = "Assumptions"
        inputs.Range("A1").Value = "Assumption"
        inputs.Range("B1").Value = "Value"
        inputs.Range("A2").Value = "Headcount"
        inputs.Range("B2").Value = rng.randint(8, 120)
        inputs.Range("A3").Value = "Cash on hand"
        inputs.Range("B3").Value = rng.randint(5, 60) * 100_000
        inputs.Range("B3").NumberFormat = MONEY

    months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun",
              "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"][:n_months]
    ws.Range("A1").Value = "Plan"
    ws.Range("B3").Value = "Growth"
    for i, m in enumerate(months):
        ws.Cells(3, 3 + i).Value = m
    last = 2 + n_months

    r = 3
    total_rows: list[int] = []
    items = LINE_ITEMS[:]
    rng.shuffle(items)
    for s in range(n_sections):
        r += 2 if spacer else 1
        ws.Cells(r, 1).Value = SECTION_NAMES[s % len(SECTION_NAMES)]
        start = r + 1
        for _ in range(rng.randint(3, 6)):
            r += 1
            ws.Cells(r, 1).Value = items.pop() if items else f"Line {r}"
            ws.Cells(r, 2).Value = round(rng.uniform(-0.02, 0.09), 3)
            ws.Cells(r, 2).NumberFormat = PCT
            ws.Cells(r, 3).Value = rng.randint(10, 400) * 1000
            for c in range(4, last + 1):
                ws.Cells(r, c).Formula = f"={col(c-1)}{r}*(1+$B${r})"
            ws.Range(f"C{r}:{col(last)}{r}").NumberFormat = MONEY
        end = r
        r += 1
        ws.Cells(r, 1).Value = f"Total {SECTION_NAMES[s % len(SECTION_NAMES)].lower()}"
        for c in range(3, last + 1):
            ws.Cells(r, c).Formula = f"=SUM({col(c)}{start}:{col(c)}{end})"
        ws.Range(f"C{r}:{col(last)}{r}").NumberFormat = MONEY
        total_rows.append(r)

    if len(total_rows) > 1:
        # A grand total that adds the section totals by reference, which is correct
        # and must not be mistaken for double counting.
        r += 2
        ws.Cells(r, 1).Value = "Grand total"
        for c in range(3, last + 1):
            refs = ",".join(f"{col(c)}{tr}" for tr in total_rows)
            ws.Cells(r, c).Formula = f"=SUM({refs})"
        ws.Range(f"C{r}:{col(last)}{r}").NumberFormat = MONEY
        grand = r
    else:
        grand = total_rows[0]

    if use_pct_row:
        r += 1
        ws.Cells(r, 1).Value = "Share of first month"
        for c in range(3, last + 1):
            ws.Cells(r, c).Formula = f"={col(c)}{grand}/$C${grand}"
        ws.Range(f"C{r}:{col(last)}{r}").NumberFormat = PCT

    if use_lookup:
        r += 2
        ws.Cells(r, 1).Value = "Region"
        ws.Cells(r, 2).Value = "Weight"
        table_start = r + 1
        for i, region in enumerate(REGIONS):
            ws.Cells(table_start + i, 1).Value = region
            ws.Cells(table_start + i, 2).Value = round(0.1 + i * 0.2, 2)
        r = table_start + len(REGIONS)
        ws.Cells(r, 1).Value = "Weighted first month"
        ws.Cells(r, 3).Formula = (
            f'=ROUND(C{grand}*VLOOKUP("{rng.choice(REGIONS)}",'
            f'A{table_start}:B{table_start + len(REGIONS) - 1},2,FALSE),0)'
        )
        ws.Cells(r, 3).NumberFormat = MONEY

    r += 2
    ws.Cells(r, 1).Value = "Full year"
    ws.Cells(r, 3).Formula = f"=SUM(C{grand}:{col(last)}{grand})"
    ws.Cells(r, 3).NumberFormat = MONEY
    fy = r
    if inputs is not None:
        r += 1
        ws.Cells(r, 1).Value = "Per head"
        ws.Cells(r, 3).Formula = f"=ROUND(C{fy}/Assumptions!B2,0)"
        ws.Cells(r, 3).NumberFormat = MONEY
        r += 1
        ws.Cells(r, 1).Value = "Months of cash"
        ws.Cells(r, 3).Formula = f"=ROUND(Assumptions!B3/AVERAGE(C{grand}:{col(last)}{grand}),1)"

    ws.Columns("A:A").ColumnWidth = 28
    app.Calculate()
    if path.exists():
        path.unlink()
    wb.SaveAs(str(path), FileFormat=51)
    wb.Close(SaveChanges=False)
    return {"sections": n_sections, "months": n_months, "spacer": spacer,
            "lookup": use_lookup, "pct_row": use_pct_row, "assumptions": assumptions_sheet}


def main() -> int:
    import win32com.client as win32

    ap = argparse.ArgumentParser()
    ap.add_argument("--n", type=int, default=12)
    ap.add_argument("--seed", type=int, default=20260915)
    args = ap.parse_args()

    OUT.mkdir(parents=True, exist_ok=True)
    for old in OUT.glob("*.xlsx"):
        old.unlink()

    rng = random.Random(args.seed)
    app = win32.Dispatch("Excel.Application")
    app.Visible = False
    app.DisplayAlerts = False
    try:
        for i in range(args.n):
            p = OUT / f"clean-{i:02d}.xlsx"
            shape = build_one(app, p, rng)
            print(f"  {p.name}: {shape}")
    finally:
        app.Quit()
    print(f"wrote {args.n} clean workbooks to {OUT}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
