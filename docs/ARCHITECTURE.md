# Architecture

The short version: **code finds and prices the defects, a model only writes about them.**
Everything below is the detail behind that sentence.

```
   .xlsx
     │  loaded twice: once for formulas, once for the values the app cached
     ▼
  workbook.py ──► Cell{formula, ast, cached, static, number_format}
     │
     ▼
  formula/  tokenizer ──► parser ──► AST
     │
     ▼
  engine.py  precedents ──► dependency graph ──► topological order
     │                                   │
     │                                   └─► Tarjan SCC ──► cycles (R006)
     ▼
  recalc()  evaluate every formula cell in order
     │
     ├─► self_check()  compare against the cached values
     │        ├─ validates the engine
     │        └─ detects a file saved without recalculating (R013)
     ▼
  rules/    15 detectors ──► Finding{evidence, Fix}
     │
     ▼
  audit.py  group repeats ──► for each Fix: recalc a copy, diff ──► impact in money
     │                                                    └─► rank
     ├─► cli.py      terminal / CI
     ├─► server.py   FastAPI + web/
     └─► explain.py  the model, fenced
```

---

## The formula engine (`gridlint/formula/`)

### `tokenizer.py`

One left-to-right pass, no backtracking. Every formula either tokenizes or raises
`FormulaSyntaxError` — there is no "best effort" path that silently drops a token, because
a dropped token would change a number.

Handles: quoted sheet names with escaped quotes (`'My ''Sheet'''!A1`), error literals,
whole-row and whole-column ranges, `$` anchors, and the distinction between a function name
and a defined name (a name followed by `(` is a call).

### `parser.py`

Precedence climbing. Excel has two rules that catch people out and both are tested:

| Expression | Excel | Why |
|---|---|---|
| `-2^2` | `4` | unary minus binds **tighter** than `^` |
| `2^3^2` | `64` | `^` is **left**-associative |

`%` is postfix. Empty arguments are legal (`IF(a,,b)`). Array literals are rejected rather
than mis-parsed.

### `values.py`

Excel's coercion rules, isolated so they can be tested alone:

- **Ordering across types:** numbers < text < `FALSE` < `TRUE`. So `="a">1` is `TRUE`.
- **Text comparison is case-insensitive:** `="abc"="ABC"` is `TRUE`.
- **Blank is zero in arithmetic and `""` in text**, which is exactly why R004 exists.
- **`ROUND` is half away from zero**, and is done in decimal on `repr(float)` — the
  shortest string that round-trips — because that is the number the person believes they
  typed. `ROUND(2.675, 2)` is `2.68` even though 2.675 is stored as
  2.67499999999999982…

### `functions.py`

90 functions. The rule that matters most is Excel's asymmetry: **inside a range, text and
blanks are ignored; as a direct argument, text is coerced.**

```
=SUM(A1:A3)  with A2 = "text"   →  text ignored
=SUM("3", 4)                     →  7
```

Getting that backwards changes totals silently, which is the whole class of bug this
product is about.

### `evaluator.py`

A pure function of `(node, context sheet, resolver)`. It never looks up another cell's
*formula* — only its already-computed value. That is what lets the engine handle arbitrarily
deep reference chains without recursion, and it makes the evaluator trivial to test with a
fake resolver.

`IF`, `IFERROR`, `IFNA` and `ISERROR` are evaluated lazily so an error in an untaken branch
does not propagate.

---

## The graph and recalculation (`gridlint/engine.py`)

`build_graph` walks each AST and records precedents. Ranges are expanded, but **clamped to
the sheet's used area** and capped at 50,000 cells, so `=SUM(A:A)` does not put a million
nodes in the graph.

Kahn's algorithm gives the evaluation order. Whatever never becomes ready is, by definition,
in a cycle; Tarjan's SCC on that remainder gives the exact cycles, which become R006
findings, and those cells evaluate to `#CIRC!`.

`recalc(wb, graph, overrides=..., formula_overrides=...)` is the workhorse:

- `overrides` replaces **values** — used for what-if.
- `formula_overrides` replaces **formulas** on a shallow copy of the workbook — this is how
  a candidate fix is tested without touching the original. There is a test asserting the
  original is unmodified afterwards.

### The self-check

```python
sc = engine.self_check(wb)      # compares recalculated vs cached, cell by cell
sc.agreement                    # 0.0 … 1.0
sc.trustworthy                  # agreement >= 99.5% on a non-empty sample
```

Comparison is date-aware: a date-formatted cell comes back from the file as a `datetime`
while the engine computes the serial number underneath it, and those are the same value.
Getting this wrong made every date formula look stale.

Numeric comparison uses a **relative tolerance of 1e-9**. That was 1e-6 at first, which
meant a $1 discrepancy on a $1,000,000 value counted as a match — the exact thing this tool
exists to catch. A test now pins it.

---

## Rules (`gridlint/rules/`)

Each rule is a generator taking `(wb, computed, graph, self_check)` and yielding `Finding`s.
A rule that raises is caught and reported as an info finding, so one broken detector cannot
lose the others (there is a test for that too).

Two shared helpers do most of the work:

- **`to_r1c1(ast, col, row)`** — a signature where relative references become offsets, so
  two cells doing "the same thing one column across" produce the identical string. This is
  how R002 finds the one formula in a row that was typed by hand.
- **`skeleton_of(ast)`** — the same idea with every reference *and* literal erased. Sibling
  rows legitimately differ in which absolute cell they anchor to (`=C11*(1+$B$11)` versus
  `=C12*(1+$B$12)`), so R001 uses skeletons to decide whether a neighbouring cell belongs to
  the same block. Using exact shapes here caused Gridlint to find only 1 of 6 identical
  defects during development.

R013 deserves its own note. It fires when the engine's recomputed value disagrees with the
cached one — but only when agreement *excluding those cells* is still above 99.5% and the
mismatches are a small minority. Otherwise the right conclusion is "the engine does not
understand this workbook", not "this workbook is stale", and Gridlint says that instead.

---

## Pricing and ranking (`gridlint/audit.py`)

1. **Group.** Findings of the same rule, same evidence signature, in consecutive cells of
   one row are collapsed into a single finding with a multi-cell fix. Twelve monthly totals
   with the same mistake are one thing a person has to think about.
2. **Price.** For each grouped finding with a fix: apply every edit at once, recalculate,
   diff against the baseline. Record the number of changed cells, the largest single
   currency delta, and — separately — the **headline changes**: cells that changed, carry a
   row label, and have **no dependents**. Those are the summary lines a human actually
   reads. "Runway (months) 38.6 → 5.2" comes from there.
3. **Verify.** A fix is only offered if every edited cell computes and **no new error cell
   appears anywhere in the workbook**.
4. **Rank.** Severity, then measured impact, then blast radius, then confidence.

`Report.money_at_risk` is the **largest** impact, never the sum. Impacts overlap, and
double-counting would be a strange mistake for this product to make.

---

## The model, and the fences (`gridlint/explain.py`)

Three jobs, none of which is deciding anything:

| Function | What it produces | What checks it |
|---|---|---|
| `explain_finding` | one paragraph for a non-expert | `NumberGuard` |
| `review_note` | three sentences to send to the file's owner | `NumberGuard` |
| `propose_repair` | a candidate formula | parse → recalc → diff |

**`NumberGuard`** extracts every numeric token from the generated text and requires each to
match something in `allowed_numbers(finding)` — the evidence dict, the measured impact, the
before/after values, the cell addresses — or in the context string Gridlint itself supplied.
Integers 0–100 are allowed as ordinary English ("12 cells"). Anything else means the model
produced a figure from nowhere, and the whole sentence is discarded.

**`propose_repair`** rejects a proposal that does not parse, points at a sheet not in the
file, fails to recompute, or introduces a new error cell. All four rejections are tested.

**Prompt stability matters.** Prompts are built from explicitly named fields, never by
dumping a report structure. Adding one key to the report would otherwise change every
prompt and silently invalidate every recorded fixture — which happened once during
development and is why the rule is written down here.

---

## Storage and the web app

`db.py` is SQLite with hand-written SQL: workspaces, members, sessions, workbooks, runs,
share tokens. Workbook bytes live on disk under the workspace directory; the report JSON
lives in the `run` row so history is cheap to render.

`server.py` is FastAPI. Notable choices:

- `/api/demo` runs a bundled sample — no account, no key, no upload. That is the path a
  judge takes.
- `/api/me` returns `{"signed_in": false}` with a 200 for anonymous visitors, because being
  signed out is a normal state and not an error the browser should log.
- Uploads are checked for size, extension, **and the ZIP magic bytes**, then audited in a
  temporary directory. `/api/check` never stores anything.
- Every workspace-scoped endpoint verifies `workspace_id` against the session. There is a
  test that one workspace cannot read, re-check or delete another's workbook.
- Passwords are PBKDF2-SHA256, 240,000 iterations, per-user salt.

`web/` is one HTML file, one stylesheet, one script. No framework, no build step, no
`node_modules` — which also means the repository a judge clones is the repository that runs.

---

## Performance

Two passes in the original code were quadratic, and profiling a 23,000-formula model is
what surfaced them:

- `build_graph` rebuilt `set(order)` inside a list comprehension, once per formula.
- `double_counting` (R009) scanned every cell of every summed range looking for a subtotal.
  It now indexes the cells that contain a `SUM` by `(sheet, row)` and looks up only the
  range's row window.

Together those took the same workbook from 37 seconds to 7. The tokenizer also stopped
trying the external-reference pattern at every character position, which alone cost more
than the rest of lexing put together.

The evaluator runs about **88,000 formulas per second**; roughly half of what remains is
openpyxl reading the file twice, which is the price of having Excel's own values to compare
against. `bench/perf.py` prints the stage-by-stage breakdown after a warm-up pass, so the
figures are steady state rather than a measurement of Python's import cost.

## The dependency trace

`engine.trace()` walks backwards from a cell through the graph, breadth-first, returning
each step with its row label, formula, value and depth. Given the set of cells a fix moves,
it prunes the branches that do not move, so the chain shown is the chain that matters.

The walk starts at the **headline change** rather than at the defect, because the number a
reviewer is asking about is "runway", not "C16". Cycles cannot loop it: each cell is
visited once, and both depth and step count are bounded.

## Testing

151 tests, about 15 seconds, green on Linux and Windows across Python 3.10 and 3.12.

- `test_formula.py` — coercion, precedence, rounding, the function semantics, in isolation.
- `test_engine.py` — the conformance workbook against Excel, graph ordering, cycles, a
  3,000-cell chain, tolerance.
- `test_rules.py` — each rule against a workbook built for it, plus the silence cases.
- `test_explain.py` — the guard accepts real figures and rejects invented ones; every
  rejection path of `propose_repair`.
- `test_llm.py` — replay, cache keys, and fixture hygiene.
- `test_server.py` — the API, including cross-workspace isolation and the corrected download.
- `test_cli.py` — exit codes, `--fail-on` levels, and that `fix` never touches the original.
- `test_real_world.py` — the things a real workbook contains: `_xlfn.` prefixes, array
  idioms, dates as serials, table references, external links, hidden rows.
- `test_trace.py` — the dependency walk, including that a cycle cannot make it loop.

Excel is needed only to *regenerate* fixtures. The generated `.xlsx` files are committed, so
CI runs without it.
