# Formula recalculation: the library is not a formula engine

Cluster: formulas

## Scenario

A user loads a workbook with a formula cell that depends on others (e.g. `A1+B1` with a cached
result), programmatically changes an input cell (sets `A1` to a new number), then reads the formula
cell expecting its result to reflect the change. It does not: the library stores the formula's cached
result and never recomputes it, so the dependent cell holds a now-stale value. This is correct by
design but a frequent source of confusion.

> Spec note, not a corpus case: this is a contract clarification and API-ergonomics question, not a
> defect — there is no correct "recomputed" output to assert (the library has no evaluator). The
> durable value is making the staleness explicit and safe.

## Desired behavior / contract

- The library is a **spreadsheet document reader/writer, not a formula engine.** A formula cell is
  stored as its formula text plus an **optional cached result**; the library never computes or
  recomputes formula values.
- Reading a formula cell returns the **formula text and the cached result as loaded from the file.**
  Mutating a dependency cell in memory does **not** update any dependent formula cell — the cached
  result is retained and is now stale.
- On write, the cached result is preserved as-is unless the caller overwrites it; the library does
  not blank or fabricate results.
- The contract is **documented explicitly**, and the surface makes staleness safe: a caller can tell
  a cached result from a computed one and can choose to clear/overwrite results they know are stale.

## Open questions

- Should the writer set the workbook's **`fullCalcOnLoad`** flag (so the consuming application
  recalculates on open) when the caller signals results may be stale — the standard OOXML way to
  delegate recalculation to Excel?
- A helper to **clear cached results** (write formulas with no `<v>`) so the consumer recomputes, vs
  leaving that to the caller?
- Is an optional pluggable evaluator ever in scope, or firmly out (keeping the library a document
  tool)? Firmly out is the least-surprise, smallest-surface answer.
- How is a formula cell's cached result surfaced distinctly from a plain value so callers are not
  misled into thinking it was recomputed?

Related: `formula-cell-preserves-falsy-result-values`, `formula-string-result-under-date-format-roundtrip`,
`shared-formula-master-survives-roundtrip-and-splice`, `modern-function-xlfn-prefix-roundtrip`.
