# Benchmark

Two questions, one corpus: does it find planted defects, and does it stay quiet otherwise?

Run it yourself (Excel required to build the fixtures; the corpus is committed so you can
skip straight to scoring):

```bash
python bench/make_corpus.py --n 12       # 12 clean, structurally varied workbooks
python bench/benchmark.py --per-kind 2   # plant, recalculate with Excel, score
```

Results are written to `bench/results.json`. Seed 20260915, so it repeats.

## Recall — 182 / 182

Eight kinds of defect, planted into twelve base workbooks, **every mutant re-opened and
recalculated by Excel** so the fixture carries the same cached values a real upload would.
A mutant counts as found only when the *right rule* names the *right cell*.

| Rule | Defect planted | Planted | Found | Recall | Ranked first |
|---|---|---:|---:|---:|---:|
| R001 | `SUM` range shortened by one row | 22 | 22 | 100% | 13 |
| R002 | one formula in a copied row hand-edited | 16 | 16 | 100% | 16 |
| R006 | two cells referring to each other | 24 | 24 | 100% | 24 |
| R007 | reference to a sheet that is gone | 24 | 24 | 100% | 0 |
| R008 | a number pasted as text inside a total | 24 | 24 | 100% | 24 |
| R012 | divide-by-zero hidden behind `IFERROR` | 24 | 24 | 100% | 24 |
| R014 | a link into a workbook that is not attached | 24 | 24 | 100% | 24 |
| R015 | a row hidden inside a summed range | 24 | 24 | 100% | 24 |
| | **total** | **182** | **182** | **100%** | |

R007 is never ranked first because a reference to a missing sheet also produces `#REF!`,
which R005 reports as a critical error with a blast radius. Both findings are correct and
both point at the same cell; R005 simply sorts higher.

## False positives — 0 across 1,059 formulas

The twelve unmutated workbooks produce **zero** findings. They deliberately contain the
shapes that trip naive linters:

- blank spacer rows between sections,
- section subtotals rolled into a grand total by reference (correct, and not double
  counting),
- cross-sheet assumption cells referenced absolutely,
- `VLOOKUP` against a lookup table,
- percentage rows dividing by an anchored cell,
- text label columns beside numeric blocks,
- a hidden helper column that is not inside any total, because hiding something is not by
  itself a defect and R015 has to stay quiet about it.

| Workbook | Formulas | Engine agreement | Findings |
|---|---:|---:|---:|
| clean-00 … clean-11 | 1,059 total | 1,059 / 1,059 | **0** |

## What the benchmark caught

While it was being written, the clean corpus flagged `Model!C30` in `clean-00` as a stale
value. It was not stale — the engine was wrong. `excel_round` corrected binary
representation error with a *relative* epsilon:

```python
nudged = abs(scaled) + 1e-9 * max(1.0, abs(scaled))
```

At 3,806,241.4967 that epsilon is 0.0038, which pushed the value over the .5 boundary and
returned 3,806,242. Excel returns 3,806,241. Rounding now happens in decimal, on the
shortest string that round-trips to the same float, and the case is pinned by a test.

A benchmark that never fails its own author is not measuring anything.

## What this does not measure

- **Real-world recall.** Planted defects are synthetic. They are drawn from the mistakes
  that show up in field audits, but a corpus of genuinely messy files that somebody's
  finance team actually shipped would be better evidence. That corpus does not exist
  publicly, which is itself part of the roadmap.
- **Precision on messy files.** Zero false positives on clean workbooks is necessary, not
  sufficient. A file with merged cells, hidden sheets and ten years of history is the real
  test.
- **Speed on a real file.** `bench/perf.py` measures a synthetic model, which is uniform in
  a way real workbooks are not. A 23,000-formula file audits in about 7 seconds; the same
  formula count spread over forty sheets with merged cells will not.
- **R003, R004, R010, R011.** These are judgement calls rather than defects with a ground
  truth, so they are excluded from the recall figures and are never ranked as critical.
