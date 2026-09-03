"""Build the sample workbooks with Excel, so their cached values are real.

    python bench/make_samples.py

Writes:
  samples/board-model.xlsx   the demo file: a board deck's operating model with
                             five planted defects, including a SUM that stops one
                             row short and turns a cash-burning company profitable
  samples/clean-model.xlsx   the same model with nothing wrong, to show the tool
                             stays quiet on a healthy file
  samples/runway.xlsx        a smaller model used by the tests

Excel is only needed to create these fixtures. The committed .xlsx files make
every test runnable on any machine.
"""
from __future__ import annotations

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
SAMPLES = ROOT / "samples"

MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
MONEY = "$#,##0"
PCT = "0.0%"

REVENUE = [
    ("Subscription revenue", 420_000, 0.020),
    ("Services revenue", 90_000, 0.010),
]
COSTS = [
    ("Salaries & benefits", 310_000, 0.010),
    ("Cloud & infrastructure", 64_000, 0.060),
    ("Marketing", 81_000, 0.100),
    ("Office & software", 23_000, 0.005),
    ("Contractors", 185_000, 0.040),      # the line item the broken SUM leaves out
]
CASH_ON_HAND = 1_400_000
HEADCOUNT = 41


def build(app, out: Path, *, planted: bool) -> None:
    wb = app.Workbooks.Add()
    while wb.Worksheets.Count < 2:
        wb.Worksheets.Add()
    ws = wb.Worksheets(1)
    ws.Name = "Operating Model"
    inputs = wb.Worksheets(2)
    inputs.Name = "Inputs"

    inputs.Range("A1").Value = "Assumption"
    inputs.Range("B1").Value = "Value"
    inputs.Range("A2").Value = "Cash on hand"
    inputs.Range("B2").Value = CASH_ON_HAND
    inputs.Range("B2").NumberFormat = MONEY
    inputs.Range("A3").Value = "Headcount"
    inputs.Range("B3").Value = HEADCOUNT
    inputs.Range("A4").Value = "Target gross margin"
    inputs.Range("B4").Value = 0.72
    inputs.Range("B4").NumberFormat = PCT

    ws.Range("A1").Value = "FY26 Operating Model"
    ws.Range("A2").Value = "All figures USD, prepared for the September board meeting"
    ws.Range("B4").Value = "Monthly growth"
    for i, m in enumerate(MONTHS):
        c = ws.Cells(4, 3 + i)
        c.Value = m
        c.Font.Bold = True
    last_col = 2 + len(MONTHS)

    def col(n: int) -> str:
        return chr(64 + n) if n <= 26 else chr(64 + (n - 1) // 26) + chr(65 + (n - 1) % 26)

    r = 5
    ws.Cells(r, 1).Value = "REVENUE"
    ws.Cells(r, 1).Font.Bold = True
    rev_start = r + 1
    for name, first, growth in REVENUE:
        r += 1
        ws.Cells(r, 1).Value = name
        ws.Cells(r, 2).Value = growth
        ws.Cells(r, 2).NumberFormat = PCT
        ws.Cells(r, 3).Value = first
        for c in range(4, last_col + 1):
            ws.Cells(r, c).Formula = f"={col(c-1)}{r}*(1+$B${r})"
        ws.Range(f"C{r}:{col(last_col)}{r}").NumberFormat = MONEY
    rev_end = r
    r += 1
    rev_total = r
    ws.Cells(r, 1).Value = "Total revenue"
    ws.Cells(r, 1).Font.Bold = True
    for c in range(3, last_col + 1):
        ws.Cells(r, c).Formula = f"=SUM({col(c)}{rev_start}:{col(c)}{rev_end})"
    ws.Range(f"C{r}:{col(last_col)}{r}").NumberFormat = MONEY

    r += 2
    ws.Cells(r, 1).Value = "OPERATING COSTS"
    ws.Cells(r, 1).Font.Bold = True
    cost_start = r + 1
    for name, first, growth in COSTS:
        r += 1
        ws.Cells(r, 1).Value = name
        ws.Cells(r, 2).Value = growth
        ws.Cells(r, 2).NumberFormat = PCT
        ws.Cells(r, 3).Value = first
        for c in range(4, last_col + 1):
            if planted and name == "Marketing" and c == 11:
                # R002: someone typed a one-off number into September instead of
                # using the growth assumption like every other month.
                ws.Cells(r, c).Formula = f"={col(c-1)}{r}*1.35"
            else:
                ws.Cells(r, c).Formula = f"={col(c-1)}{r}*(1+$B${r})"
        ws.Range(f"C{r}:{col(last_col)}{r}").NumberFormat = MONEY
    cost_end = r
    r += 1
    cost_total = r
    ws.Cells(r, 1).Value = "Total operating costs"
    ws.Cells(r, 1).Font.Bold = True
    # R001: the range stops one row short and silently drops Contractors.
    end_row = cost_end - 1 if planted else cost_end
    for c in range(3, last_col + 1):
        ws.Cells(r, c).Formula = f"=SUM({col(c)}{cost_start}:{col(c)}{end_row})"
    ws.Range(f"C{r}:{col(last_col)}{r}").NumberFormat = MONEY

    r += 2
    burn = r
    ws.Cells(r, 1).Value = "Net burn (costs - revenue)"
    ws.Cells(r, 1).Font.Bold = True
    for c in range(3, last_col + 1):
        ws.Cells(r, c).Formula = f"={col(c)}{cost_total}-{col(c)}{rev_total}"
    ws.Range(f"C{r}:{col(last_col)}{r}").NumberFormat = MONEY

    r += 2
    ws.Cells(r, 1).Value = "SUMMARY"
    ws.Cells(r, 1).Font.Bold = True
    r += 1
    ws.Cells(r, 1).Value = "Cash on hand"
    ws.Cells(r, 3).Formula = "=Inputs!B2"
    ws.Cells(r, 3).NumberFormat = MONEY
    cash = r
    r += 1
    ws.Cells(r, 1).Value = "Average monthly burn"
    ws.Cells(r, 3).Formula = f"=AVERAGE(C{burn}:{col(last_col)}{burn})"
    ws.Cells(r, 3).NumberFormat = MONEY
    avg_burn = r
    r += 1
    ws.Cells(r, 1).Value = "Runway (months)"
    ws.Cells(r, 3).Formula = f'=IF(C{avg_burn}>0,ROUND(C{cash}/C{avg_burn},1),"Profitable")'
    runway = r
    r += 1
    ws.Cells(r, 1).Value = "Full-year revenue"
    ws.Cells(r, 3).Formula = f"=SUM(C{rev_total}:{col(last_col)}{rev_total})"
    ws.Cells(r, 3).NumberFormat = MONEY
    r += 1
    ws.Cells(r, 1).Value = "Full-year costs"
    ws.Cells(r, 3).Formula = f"=SUM(C{cost_total}:{col(last_col)}{cost_total})"
    ws.Cells(r, 3).NumberFormat = MONEY
    fy_costs = r
    r += 1
    ws.Cells(r, 1).Value = "Cost per head"
    ws.Cells(r, 3).Formula = f"=ROUND(C{fy_costs}/Inputs!B3,0)"
    ws.Cells(r, 3).NumberFormat = MONEY
    r += 1
    ws.Cells(r, 1).Value = "Gross margin"
    ws.Cells(r, 3).Formula = f"=(C{rev_total}-C{cost_total})/C{rev_total}"
    ws.Cells(r, 3).NumberFormat = PCT
    r += 1
    ws.Cells(r, 1).Value = "Estimated tax provision"
    if planted:
        # R003: the tax rate is typed into the formula instead of living in Inputs.
        ws.Cells(r, 3).Formula = f"=MAX(0,-C{burn}*0.0825)"
    else:
        inputs.Range("A5").Value = "Tax rate"
        inputs.Range("B5").Value = 0.0825
        inputs.Range("B5").NumberFormat = PCT
        ws.Cells(r, 3).Formula = f"=MAX(0,-C{burn}*Inputs!B5)"
    ws.Cells(r, 3).NumberFormat = MONEY
    r += 1
    ws.Cells(r, 1).Value = "Revenue per head"
    if planted:
        # R012: IFERROR hides a divide-by-zero caused by an empty cell.
        ws.Cells(r, 3).Formula = f"=IFERROR(C{rev_total}/Inputs!B7,0)"
    else:
        ws.Cells(r, 3).Formula = f"=ROUND(C{rev_total}/Inputs!B3,0)"
    ws.Cells(r, 3).NumberFormat = MONEY

    if planted:
        # R008: a number pasted as text is skipped by the SUM below it.
        r += 2
        ws.Cells(r, 1).Value = "One-off costs"
        ws.Cells(r, 1).Font.Bold = True
        oneoff_start = r + 1
        for label, value in (("Legal fees", 24_000), ("Recruiting", "38,500"), ("Relocation", 12_000)):
            r += 1
            ws.Cells(r, 1).Value = label
            if isinstance(value, str):
                ws.Cells(r, 3).NumberFormat = "@"
                ws.Cells(r, 3).Value = value
            else:
                ws.Cells(r, 3).Value = value
                ws.Cells(r, 3).NumberFormat = MONEY
        r += 1
        ws.Cells(r, 1).Value = "Total one-off"
        ws.Cells(r, 3).Formula = f"=SUM(C{oneoff_start}:C{r-1})"
        ws.Cells(r, 3).NumberFormat = MONEY

    ws.Columns("A:A").ColumnWidth = 30
    ws.Columns("B:B").ColumnWidth = 14
    ws.Range(f"C4:{col(last_col)}4").HorizontalAlignment = -4152
    app.Calculate()
    SAMPLES.mkdir(parents=True, exist_ok=True)
    if out.exists():
        out.unlink()
    wb.SaveAs(str(out), FileFormat=51)
    wb.Close(SaveChanges=False)
    print(f"wrote {out.name}  (runway row {runway}, total costs row {cost_total})")


def main() -> int:
    import win32com.client as win32

    app = win32.Dispatch("Excel.Application")
    app.Visible = False
    app.DisplayAlerts = False
    try:
        build(app, SAMPLES / "board-model.xlsx", planted=True)
        build(app, SAMPLES / "clean-model.xlsx", planted=False)
        return 0
    finally:
        app.Quit()


if __name__ == "__main__":
    sys.exit(main())
