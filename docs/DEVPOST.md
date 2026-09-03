# Devpost submission — copy and paste

Submission page: <https://ai-builders-hackathon-2026.devpost.com/>
Deadline: **15 September 2026, 11:00 pm EDT** (= 16 September, 12:00 noon JST)

---

## Project name

```
Gridlint
```

## Elevator pitch (200 characters max)

```
Your board deck says 38.6 months of runway. One SUM stops a row short: it is 5.2. Gridlint recalculates your workbook, ranks defects by the money they move, and proves each fix before you accept it.
```

*(199 characters.)*

---

## Built with (tags)

```
python, fastapi, openpyxl, sqlite, javascript, html, css, anthropic, openai, docker, pytest, github-actions
```

## Try it out (links)

```
https://github.com/taro13nyanko/gridlint
```

Add the live URL here once it is deployed (see `docs/DEPLOY.md`).

---

## About the project

### Inspiration

I pay for university by fixing other people's spreadsheets — small automation jobs for
tutors, clubs, tiny businesses. Every single file I have been handed had a silent formula
bug in it. Not a crash. Not a red `#REF!`. A `SUM` that stopped one row short, or a
hand-typed number in the middle of a copied row.

Nobody ever noticed, because **a broken spreadsheet does not look broken**. It returns a
number, formats it as currency, and the number goes into a decision.

Field audits agree: studies collected by EuSpRIG find errors in 24% to 94% of real business
spreadsheets, and Panko found at least one error in 94% of 88 audited workbooks. The tools
that exist for this are desktop add-ins costing $249 to $2,000 a year, and what they
produce is a list of four hundred warnings with no order to it.

So I asked a different question: not *what is wrong*, but **what is it costing you** — and
can I prove it?

### What it does

Drop an `.xlsx` in. Gridlint:

1. Parses every formula into a syntax tree and builds a dependency graph.
2. Recalculates the whole workbook with its own engine, then **checks itself** against the
   values Excel or Sheets saved in the file.
3. Runs 13 deterministic rules to find defects.
4. For each fixable defect, applies the fix to a copy, recomputes, and **measures** the
   difference — down to which labelled lines change.
5. Ranks by money moved, groups repeats, and writes a plain-English note.

On the bundled example, the top finding reads:

> A `SUM` covers C11:C14, but "Contractors" sits right below it and is left out. The same
> mistake is repeated in 12 cells.
> **Cost per head** $175,859 → $243,659 · **Gross margin** +6.3% → −30% ·
> **Runway** 38.6 months → **5.2**

That company is not three years from running out of money. It is five months.

### How I built it

Everything is Python and vanilla JavaScript. No framework on the front end, no build step.

The hard part was the **Excel evaluation engine**, written from scratch:

- a single-pass tokenizer,
- a precedence-climbing parser that gets Excel's two surprises right (`-2^2` is `+4`, and
  `^` is left-associative so `2^3^2` is `64`),
- a value model with Excel's coercion rules (text ignored inside a range but coerced as a
  direct argument; numbers sort before text which sorts before booleans; text comparison is
  case-insensitive),
- decimal rounding, because Excel rounds half away from zero and Python does not,
- 42 worksheet functions,
- and an evaluator that **never recurses into other cells** — the dependency graph is
  sorted topologically and evaluated in order, so a 3,000-cell reference chain cannot blow
  the stack.

The self-check falls out of that for free, and it turned out to be the most useful feature
in the product. Every workbook already stores the values its app last computed. Recomputing
and comparing does two jobs at once: it validates the engine, and it detects a file saved
with calculation switched off — where the numbers on screen no longer match the formulas
underneath. If agreement drops below 99.5%, Gridlint says so and withholds the money
figures rather than guessing.

### Where the AI is, and where it is not

The organisers said they did not want AI wrappers. Here is the honest split:

| Job | Done by |
|---|---|
| Parse formulas, build the graph | Code |
| Recalculate the workbook | Code |
| Decide whether something is a defect | Code — 13 rules |
| Measure what a fix changes | Code — recalculate and diff |
| Explain the defect to a non-expert | A model, fenced |
| Draft a repair where no mechanical fix exists | A model, then executed |

Two fences, both covered by tests:

1. **Number guard.** Every figure in a generated sentence must already appear in the
   evidence the detector measured, or in the context Gridlint itself supplied. A number the
   model invented gets the whole sentence discarded, and the deterministic description is
   shown instead.
2. **Executed repairs.** A proposed formula is parsed, run through the engine, and diffed
   against the original workbook. If it does not parse, points at a sheet that is not in the
   file, or introduces one new error cell anywhere, it is rejected and never shown.

Delete the model entirely and Gridlint still finds every defect and still measures every
impact. The model makes the report readable; it does not make it true.

It is also provider-agnostic — Anthropic, OpenAI, Groq, Gemini or a local Ollama model —
and it ships with recorded fixtures so it runs with **no API key at all**.

### Challenges I ran into

**Excel's arithmetic is not Python's arithmetic.** `ROUND(2.5, 0)` is 3 in Excel and 2 in
Python. My first fix was a relative epsilon nudge, and it worked on the textbook case —
until the clean-workbook benchmark caught it rounding `3,806,241.4967` up to `3,806,242`.
It now rounds in decimal, on the shortest string that round-trips to the same float. That
bug is the reason I trust the benchmark.

**Not crying wolf.** Detecting a shortened `SUM` is easy; staying quiet on a legitimate
grand total that adds three subtotals is the hard half. The fix was to compare formula
*skeletons* (references and literals erased) rather than exact text, so a row of siblings
that differ only in which absolute cell they point at still counts as one block.

**A hackathon submission has to be believable in five minutes.** That is why the benchmark
exists at all, and why the report leads with the numbers a person recognises from their own
sheet — "Runway (months) 38.6 → 5.2" — rather than a cell address.

### Accomplishments I am proud of

- **93 / 93** — the engine reproduces every value in a conformance workbook generated by
  Excel itself, including the cases re-implementations usually get wrong.
- **141 / 141** — every planted defect found by the right rule in the mutation benchmark,
  across six defect kinds and twelve base workbooks, each mutant recalculated by Excel.
- **0** — findings across twelve clean workbooks totalling 1,368 formulas.
- **113 tests**, green on Linux and Windows, Python 3.10 and 3.12, with no Excel installed.

### What I learned

That the interesting engineering in an "AI product" is usually the part that lets you check
the AI. The number guard and the executed-repair check took an afternoon each, and they are
the reason I am willing to put a dollar figure on screen.

Also: a tool that reports 400 problems is a tool that reports none. Ranking is the product.

### What is next

- Google Sheets, read directly rather than via export.
- More of Excel: array formulas, `XLOOKUP`, `SUMIFS`, date arithmetic.
- Rule packs per domain — grant reporting, payroll, VAT returns.
- Watch mode: re-check on save, and diff this month's model against last month's.
- A public benchmark corpus, so other spreadsheet checkers can be measured against the same
  planted defects instead of everyone claiming accuracy.

---

## Submission checklist

- [x] Working product — `python -m gridlint serve`, or `run.cmd` on Windows
- [x] Public repository — <https://github.com/taro13nyanko/gridlint> (MIT)
- [x] Documentation — README with install, proof, architecture and stated limits
- [x] Presentation deck, 10 slides — `deck/Gridlint.pdf`
- [ ] Demo video, max 5 minutes — script in `docs/VIDEO-SCRIPT.md`
- [ ] Live demo URL — see `docs/DEPLOY.md`
- [ ] Devpost form submitted before 15 Sep, 11:00 pm EDT
- [ ] Tin Computer credits claimed after submitting ($299, first 100 teams)
