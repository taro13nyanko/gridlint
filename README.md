# Gridlint

**Find the spreadsheet mistake that costs the most, and prove the fix before you accept it.**

Gridlint reads the formulas in an `.xlsx`, builds a dependency graph, recalculates every
cell with its own engine, and reports the defects that quietly change the answer — ranked
by how much money each one moves. Every proposed fix is applied to a copy, the workbook is
recomputed, and the difference is measured, so the number you are shown was **calculated,
not estimated**.

```
$ gridlint check samples/board-model.xlsx

board-model.xlsx
  123 formulas across 2 sheet(s), checked in 85 ms
  engine self-check [OK]: reproduced 123/123 cached values (100.0%)

  3 critical, 3 warning, 0 info

1. [R001] Total leaves out a row of data
     where: Operating Model!C16 and 11 more
     formula: =SUM(C11:C14)
     SUM covers C11:C14 but "Contractors" sits right below it and is left out.
     The same mistake is repeated in 12 cells (Operating Model!C16 to N16).
     -> Cost per head: 175859 becomes 243659
     -> Runway (months): 38.6 becomes 5.2
     fix: Extend C11:C14 to C11:C15 in all 12 cells  (verified by recalculation, 30 cells change)
          =SUM(C11:C15)
```

That board deck says the company has **38.6 months of runway**. One `SUM` stops one row
short. It has **5.2**.

---

## Try it in two minutes

No account, no API key, no upload.

```bash
git clone https://github.com/taro13nyanko/gridlint && cd gridlint
pip install -r requirements.txt

python -m gridlint check samples/board-model.xlsx   # the demo file
python -m gridlint check samples/clean-model.xlsx   # the healthy one: silence
python -m gridlint serve                            # the web app on :8000
```

On Windows, double-click **`run.cmd`**. On macOS or Linux, run **`./run.sh`**.
Then open <http://127.0.0.1:8000> and press **Check the example model**.

Docker, if you prefer: `docker build -t gridlint . && docker run -p 7860:7860 gridlint`

---

## Why this exists

Field audits keep finding errors in **24% to 94%** of real business spreadsheets
([EuSpRIG](https://eusprig.org/research-info/research-and-best-practice/); Panko found at
least one error in 94% of 88 audited workbooks). The reason those errors survive is not
that they are hard to see — it is that **a broken spreadsheet does not look broken**. A
`SUM` that stops one row short still returns a number, still formats as currency, still
lands in the board deck.

The existing tools for this are desktop add-ins costing **$249 to $2,000 a year**
(PerfectXL, Operis Analysis Kit, Spreadsheet Detective). They hand an analyst a long list
of warnings. None of them tells you which warning is worth reading, and none of them can
show you what the number becomes once it is fixed.

Gridlint is built around that gap: **rank by money, and prove the fix.**

---

## What makes this different from an AI wrapper

The organisers asked for products, not "AI wrappers with minimal differentiation". Here is
the honest division of labour inside Gridlint:

| Job | Done by |
|---|---|
| Parse formulas, build the dependency graph | **Code** (`gridlint/formula/`, `gridlint/engine.py`) |
| Recalculate the workbook | **Code** — a from-scratch Excel evaluation engine |
| Decide whether something is a defect | **Code** — 13 deterministic rules |
| Measure what a fix changes | **Code** — recalculate and diff |
| Explain the defect to a non-expert | **A language model**, fenced (below) |
| Draft a repair where no mechanical fix exists | **A language model**, then compiled and executed before it is offered |

**No model output is ever trusted as a fact.** Two fences enforce that, and both are tested:

1. **Number guard.** Every figure in a generated sentence must already appear in the
   evidence the detector measured, or in the context Gridlint itself supplied. Anything
   else is a number the model made up, and the whole sentence is discarded — the
   deterministic sentence is always there to fall back to.
   (`gridlint/explain.py::NumberGuard`, `tests/test_explain.py`)
2. **Executed repairs.** A model-proposed formula is parsed, run through the engine, and
   diffed against the original workbook. If it does not parse, points at a sheet that is
   not in the file, or introduces a single new error cell anywhere, it is rejected and
   never shown as a fix. (`gridlint/explain.py::propose_repair`)

Delete the model entirely and Gridlint still finds every defect and still measures every
impact. The model makes the report readable; it does not make it true.

---

## Proof

Three numbers, all reproducible from this repository.

### 1. The engine agrees with Excel: 93 / 93

`samples/conformance.xlsx` is generated **by Excel itself**, so the values it carries are
ground truth. It covers the cases a re-implementation usually gets wrong:

| Case | Excel | Naive implementation |
|---|---|---|
| `=-2^2` | `4` | `-4` |
| `=2^3^2` | `64` (left-associative) | `512` |
| `=ROUND(2.5,0)` | `3` (half away from zero) | `2` (banker's rounding) |
| `=ROUND(2.675,2)` | `2.68` | `2.67` (binary representation) |
| `=SUM(A1:A3)` with text in `A2` | text ignored | `#VALUE!` |
| `=SUM("3",4)` | `7` (direct argument coerced) | `4` |
| `="a">1` | `TRUE` (text sorts after numbers) | error |
| `="abc"="ABC"` | `TRUE` (case-insensitive) | `FALSE` |

```bash
python -m pytest tests/test_engine.py::test_conformance_workbook_matches_excel -q
```

### 2. Mutation benchmark: 141 / 141 planted defects found

`bench/benchmark.py` takes twelve structurally different clean workbooks, plants a known
defect of each kind, has **Excel recalculate every mutant** so the fixture is realistic,
and then checks whether the right rule reported the right cell.

| Rule | Defect planted | Planted | Found | Recall |
|---|---|---:|---:|---:|
| R001 | `SUM` range shortened by one row | 23 | 23 | 100% |
| R002 | one formula in a copied row hand-edited | 22 | 22 | 100% |
| R006 | two cells referring to each other | 24 | 24 | 100% |
| R007 | reference to a sheet that is gone | 24 | 24 | 100% |
| R008 | a number pasted as text inside a total | 24 | 24 | 100% |
| R012 | a divide-by-zero hidden behind `IFERROR` | 24 | 24 | 100% |
| | **total** | **141** | **141** | **100%** |

### 3. Silence on healthy files: 0 findings across 1,368 formulas

The same twelve clean workbooks — with blank spacer rows, subtotals rolled into a grand
total, cross-sheet assumptions, `VLOOKUP` tables, percentage rows — produce **zero**
findings. A checker that cries wolf is one nobody keeps installed.

```bash
python bench/make_corpus.py --n 12     # needs Excel; the corpus is committed
python bench/benchmark.py --per-kind 2 # writes bench/results.json
```

> The benchmark found a real bug in Gridlint's own arithmetic while it was being written:
> `ROUND` used a relative epsilon to correct binary representation error, which rounded
> `3,806,241.4967` up to `3,806,242`. It now rounds in decimal. That is the point of
> having a benchmark.

---

## What it looks for

| Rule | | What it catches |
|---|---|---|
| **R001** | critical | A `SUM` or `AVERAGE` whose range stops short of adjacent data |
| **R002** | critical | One formula in a copied row or column that breaks the pattern |
| **R005** | critical | A cell producing an error, traced back to the cell where it starts |
| **R006** | critical | A circular reference |
| **R007** | critical | A reference to a deleted row, column or sheet |
| **R009** | critical | A total whose range swallows a subtotal, double-counting it |
| **R012** | critical | `IFERROR` hiding a failure that is happening right now |
| **R013** | critical | A saved value that no longer matches its own formula |
| **R003** | warning | A rate or assumption typed inside a formula instead of a labelled cell |
| **R004** | warning | Arithmetic on an empty cell, which silently counts as zero |
| **R008** | warning | A number stored as text inside a range being totalled |
| **R011** | warning | `NOW`, `TODAY`, `RAND`, `INDIRECT` or `OFFSET`, which make a result nobody can reproduce |
| **R010** | info | The same formula duplicated across the workbook |

`gridlint rules` prints this list with the reasoning for each.

**R013 is the one only a tool with its own engine can run.** Every workbook saved by Excel
or Sheets stores the values that app last computed. Gridlint recomputes them and compares.
If they disagree — someone saved with calculation turned off — the number people are
reading is not the number the formula produces. And because that same comparison validates
the engine, Gridlint knows when to distrust *itself*: if agreement drops below 99.5% it
says so and withholds the money figures rather than guessing.

---

## Using it

### Command line

```bash
gridlint check model.xlsx                  # ranked findings
gridlint check model.xlsx --json           # machine-readable
gridlint check model.xlsx --explain        # add plain-English notes (needs a model)
gridlint check model.xlsx --fail-on critical   # exit 1 for CI
gridlint fix model.xlsx --dry-run          # show the edits
gridlint fix model.xlsx --out fixed.xlsx   # write a corrected copy
gridlint verify model.xlsx                 # engine self-check only
gridlint rules                             # what it looks for
gridlint serve --port 8000                 # the web app
```

`gridlint fix` never edits your file. It writes a new one with the corrected cells
highlighted and the original formula kept in a cell comment.

### In CI

`.github/workflows/spreadsheet-check.yml` in this repository is a drop-in workflow that
fails a pull request when a committed workbook gains a critical defect.

### As a library

```python
from gridlint.audit import audit

report = audit("model.xlsx")
print(report.headline())
for f in report.findings:
    print(f.rule, f.cell, f.impact_value, f.fix_verified)
```

### As a web app

Sign-in is optional. Anonymous visitors can check a file and get a full report; a
workspace additionally keeps workbooks, their history, re-checks, shareable read-only
report links, and the corrected-file download.

---

## Bring your own model (or none)

Gridlint is not tied to one vendor. Set **one** of these and the explanations turn on:

```bash
export ANTHROPIC_API_KEY=...     # or OPENAI_API_KEY, GROQ_API_KEY, GEMINI_API_KEY
export LLM_PROVIDER=ollama       # or run a local model, no key at all
```

With **no key set**, Gridlint runs in replay mode: the notes for the bundled samples come
from `fixtures/llm/`, recorded once and committed, so a reviewer sees the finished product
without spending anything. Everything else — detection, recalculation, impact, fixes —
never touches a model at all.

Only finding metadata is ever sent to a model. **Your workbook is not uploaded anywhere.**

---

## How it is put together

```
gridlint/
  formula/          tokenizer -> parser -> AST -> evaluator (an Excel subset, from scratch)
    tokenizer.py      single-pass lexer; every formula tokenizes or raises
    parser.py         precedence climbing, including Excel's two surprises
    values.py         the coercion rules: number/text/bool ordering, decimal rounding
    functions.py      42 worksheet functions with Excel's range-vs-argument semantics
    evaluator.py      pure: (node, sheet, resolver) -> value
  workbook.py       loads an .xlsx twice: formulas, and the values the app cached
  engine.py         dependency graph, topological recalculation, cycles, self-check
  rules/            13 detectors, each yielding Findings with evidence and a Fix
  audit.py          runs the rules, prices each fix by recalculation, ranks and groups
  explain.py        the model's three jobs and the fences around them
  apply.py          writes a corrected copy with comments and highlights
  server.py         FastAPI: workspaces, uploads, runs, shared reports
  web/              one HTML file, one stylesheet, one script. No build step.
bench/              corpus generator, mutation benchmark, results
tests/              125 tests
```

Design decisions worth knowing:

- **Evaluation never recurses into other cells.** The engine sorts the dependency graph
  topologically and evaluates in order, so a 3,000-cell reference chain cannot hit
  Python's recursion limit. There is a test for exactly that.
- **Ranges are clamped to the used area**, so `=SUM(A:A)` does not expand to a million
  cells in the graph.
- **Findings are grouped.** Twelve monthly totals that all forget the same line item are
  one thing a person has to think about, not twelve rows in a list.
- **Impacts are never summed.** They overlap. The headline is the largest single measured
  impact, not a total that would be double counting — which would be an odd mistake for
  this particular product to make.

Longer notes: [docs/ARCHITECTURE.md](docs/ARCHITECTURE.md).

---

## Limits, stated plainly

- **`.xlsx` and `.xlsm` only.** No `.xls`, no `.csv`, no Google Sheets API yet — export to
  `.xlsx` first.
- **The engine models a subset of Excel**: 42 functions, no array formulas, no pivot
  tables, no macros, no `INDIRECT`/`OFFSET` resolution (they are flagged as unauditable
  instead). Unsupported formulas are listed in the report rather than silently skipped,
  and they lower the self-check score, which in turn makes Gridlint withhold money figures.
- **A workbook with no cached values** (one written by a script rather than saved by a
  spreadsheet app) cannot be cross-checked. Gridlint says so and still reports structural
  defects.
- **R003 and R004 are judgement calls** and are warnings, never critical. R003 in
  particular fires on any unusual literal; that is deliberate, and it is why it is not
  ranked with the defects that move money.
- Everything runs in one process against SQLite. That is the right size for the teams this
  is aimed at, and the wrong size for a hundred thousand of them.

---

## Tests

```bash
python -m pytest -q          # 125 tests, about 40 seconds
```

They cover Excel's coercion and rounding rules directly, the graph and recalculation, each
rule against a workbook built for it, the fences around the model, and the API including
one workspace being unable to read another's files.

Excel is needed only to *regenerate* the fixtures (`bench/make_samples.py`,
`bench/make_conformance.py`, `bench/make_corpus.py`). The generated `.xlsx` files are
committed, so the whole suite runs on Linux in CI.

---

## Built for the AI Builders Hackathon 2026

Written from scratch during the submission window by one undergraduate at the University
of Tokyo, who pays for university by fixing other people's spreadsheets. Every one of them
had a silent formula bug in it. This is the tool that would have found them.

MIT licensed.
