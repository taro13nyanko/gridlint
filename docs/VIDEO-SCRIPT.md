# Demo video script — 4 minutes 55 seconds

Hard limit is 5 minutes. This runs to about 4:55, which still leaves a little room.

**Before you record**

- `python -m gridlint serve --port 8000`, browser at `http://127.0.0.1:8000`, window 1440×900.
- Have `samples/board-model.xlsx` **open in Excel on a second desktop** (Alt+Tab away, not
  side by side — switching is cleaner than shrinking).
- Zoom the browser to 110% so the numbers are readable after compression.
- Clear the terminal, and have these two commands ready to paste:
  - `python -m gridlint check samples/board-model.xlsx`
  - `python -m pytest tests/test_engine.py::test_conformance_workbook_matches_excel -q`
- Record at 1080p. Speak slowly; you are a non-native speaker and the judges are watching
  a hundred of these. Clarity beats speed every time.

---

## 0:00 – 0:35  ·  The hook (screen: Excel, the board model)

> *(Excel is open on the operating model, scrolled so rows 11 to 23 are visible.)*

"This is a startup's operating model. It goes to the board on Friday.

Down here it says: **runway, thirty-eight point six months**. Three years of cash. On that
number you hire, you raise, you plan.

*(click cell C16, so the formula bar shows `=SUM(C11:C14)`)*

And this is the total operating costs. `SUM`, C11 to C14.

*(click cell A15, "Contractors", then C15, `185,000`)*

Row fifteen is Contractors. A hundred and eighty-five thousand dollars, every month. The
total stops at row fourteen. It has never included it.

Nothing is broken. No red cell, no error. The number is just wrong, in the direction that
makes you feel safe."

---

## 0:35 – 1:20  ·  What it does (screen: browser, landing → report)

> *(Switch to the browser on the Gridlint landing page.)*

"This is Gridlint. Drop in a workbook and it tells you which mistake costs the most.

*(drag `board-model.xlsx` onto the drop zone; the report appears)*

One second. A hundred and twenty-three formulas.

*(point at the red headline box)*

Top of the report: that `SUM` leaves out the Contractors row, and the same mistake is
repeated in all twelve monthly columns.

And this is the part that matters — **what the numbers become once it is fixed**:

- Cost per head, one seventy-five thousand, becomes two forty-three thousand.
- Gross margin, plus six point three percent, becomes **minus thirty percent**.
- Runway, thirty-eight point six months, becomes **five point two**.

The company is not three years from running out of money. It is five months."

---

## 1:20 – 2:05  ·  Why you can believe the number (screen: report detail)

> *(Scroll to the detail panel on the right.)*

"That number is not an estimate. Gridlint applied the fix to a copy of the workbook,
**recalculated the whole thing**, and diffed it. Thirty cells changed. No new errors
appeared. That is what this badge means: *verified by recalculation*.

*(click 'Every cell that changes')*

Here is every one of them, before and after.

*(point at the badge in the header: 'engine matches the file on 123/123 values')*

And this is why you can trust the recalculation. Every workbook stores the values Excel
last computed. Gridlint recomputes all of them and compares. A hundred and twenty-three
out of a hundred and twenty-three.

If that agreement dropped, Gridlint would say so and **withhold the money figures** rather
than guess. It knows when to distrust itself."

---

## 2:05 – 2:50  ·  Where the AI is (screen: report, click Explain)

> *(Click 'Explain in plain English'.)*

"Now, the AI.

*(the note appears)*

A model wrote this paragraph for whoever owns the file — no jargon, business consequence
first.

But notice what the model did **not** do. It did not decide anything was wrong. It did not
calculate a single number. Every figure in that sentence had to already exist in the
evidence the detector measured. If the model writes a number Gridlint cannot source — the
whole sentence is thrown away, and you see the plain built-in description instead.

Same rule for repairs: if the model proposes a formula, Gridlint parses it, runs it through
the engine, and diffs the workbook. One new error cell anywhere and the suggestion is
rejected. It never reaches you.

Delete the model entirely and Gridlint still finds every defect and still measures every
impact. The model makes the report readable. It does not make it true."

---

## 2:50 – 3:10  ·  Where the number comes from (screen: report detail, trace)

> *(Scroll to the trace block.)*

"One more thing the graph makes possible. This is the chain behind that runway figure —
cost per head, from the full-year total, from the twelve monthly totals, down to the line
items somebody typed. The highlighted rows are the ones that move when the fix lands.

“Where does this number come from” is the question every reviewer asks, and this is the
answer, drawn from the dependency graph rather than from a guess."

---

## 3:10 – 3:25  ·  The rest of the findings (screen: findings list)

> *(Click through findings 2, 3 and 4 in the left list.)*

"The rest of the file.

*(finding 2)* One month's marketing formula was typed by hand instead of using the growth
rate — ten of the eleven other months agree, September does not.

*(finding 3)* An `IFERROR` that is hiding a live divide-by-zero. The cell shows a confident
zero. There is no number behind it.

*(finding 4)* Thirty-eight thousand five hundred pasted as text, so the `SUM` skips it.

*(tick 'show formulas' on the grid)*

And the grid shows it in place. Row sixteen, red, twelve times — `SUM` C11 to C14, with
Contractors sitting right above it."

---

## 3:25 – 4:10  ·  Proof (screen: terminal)

> *(Switch to the terminal.)*

"Four numbers, and you can run all of them from the repository.

*(paste: `python -m pytest tests/test_engine.py::test_conformance_workbook_matches_excel -q`)*

One. The engine is not a guess at Excel. This test compares it against a workbook Excel
itself generated — a hundred and fifty-six formulas, including the ones people get wrong.
In Excel, minus two squared is **positive four**. Two to the three to the two is
**sixty-four**, not five hundred and twelve. `ROUND` two point five is **three**, not two.
A hundred and fifty-six out of a hundred and fifty-six.

*(show `bench/results.json` or the README table)*

Two. A mutation benchmark: eight kinds of defect planted into twelve clean workbooks, every
mutant recalculated by Excel. **A hundred and eighty-two out of a hundred and eighty-two**
found by the right rule.

Three, and this one matters more: on those twelve clean workbooks — **zero findings**. A
thousand and fifty-nine formulas of silence. A checker that cries wolf is one nobody keeps
installed.

And it is fast enough to be useful: a twenty-three thousand formula model, audited end to
end, in seven seconds.

That benchmark also caught a real bug in Gridlint's own rounding while I was building it.
That is what a benchmark is for."

---

## 4:10 – 4:40  ·  The product, and why me (screen: landing / pricing, then face or logo)

> *(Back to the browser, scroll to pricing. Or cut to yourself on camera.)*

"Gridlint is a workspace. Your team's workbooks live in it, the check runs again next
month, and a report can be shared as a read-only link with whoever owns the file. There is
a command-line version and a GitHub Action, so a pull request that breaks a total fails the
build — the way it would for code.

I pay for university by fixing other people's spreadsheets. Every single one had a silent
formula bug in it. This is the tool I wished existed.

It is open source, it runs with no API key, and the whole thing is at
**github.com/taro13nyanko/gridlint**. Thank you."

---

## Shot list (if you record clips separately and cut them together)

| # | Length | Screen | Must be visible |
|---|---|---|---|
| 1 | 35 s | Excel, board model | `38.6` runway, formula bar `=SUM(C11:C14)`, row 15 Contractors 185,000 |
| 2 | 45 s | Gridlint landing → report | drag and drop, the red headline box with all four changes |
| 3 | 45 s | Report detail | verified badge, cell diff, `123/123` badge |
| 4 | 45 s | Report + Explain | the written note appearing |
| 5 | 35 s | Findings list + grid | findings 2–4, then formulas toggled on, row 16 red |
| 6 | 20 s | Report detail | the dependency trace, highlighted rows |
| 7 | 45 s | Terminal | conformance test passing, benchmark table |
| 8 | 30 s | Pricing / face | pricing tiers, repo URL on screen at the end |

## Things to avoid

- Do not read the rubric back to the judges ("this shows innovation…"). Show the thing.
- Do not apologise for anything. If something is missing it is on the roadmap slide.
- Do not let the terminal font be small. Bump it to 18pt before recording.
- Do not say "AI-powered". Say what the model does and what it is not allowed to do.
