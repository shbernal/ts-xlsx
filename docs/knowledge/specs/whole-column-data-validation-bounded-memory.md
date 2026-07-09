# Whole-column data-validation ranges must not explode into per-cell state

Cluster: validations / security

## Scenario

A workbook applies a data validation (e.g. a dropdown list, or the empty catch-all rule) over a
whole-column or otherwise huge `sqref` — a range covering hundreds of thousands to over a million
cells. A reader that materializes one validation object per covered cell expands the rule into an
enormous per-cell structure: reading the file hangs or exhausts memory. This is an unbounded
allocation driven by attacker- or generator-controlled input, so it is a hostile-input concern,
not merely a performance nit.

> This behavior is captured as a spec note rather than a corpus case on purpose: the current
> reader **hangs** on the fixture, and a hanging read must never enter the corpus suite (it would
> stall CI). The requirement is asserted here; a bounded-time regression check belongs in a
> dedicated perf/security harness with a hard timeout, not in the behavior corpus.

## Desired behavior

- Reading a workbook whose validations cover whole columns (or any large/multi-part `sqref`)
  completes in **bounded time and memory**, independent of how many cells the ranges nominally
  cover. Memory scales with the number of *rules*, not the number of *covered cells*.
- A data validation is retained as a **rule associated with its range(s)** — the `sqref` is kept
  as ranges, not expanded to an entry per cell. Lookups ("does this cell have a validation?")
  resolve against the ranges.
- The list-type validation's allowed-values formula and the type/flags survive the read for the
  covered range.
- A `dataValidation` with no `type` (the empty catch-all frequently applied over many ranges) is
  handled without error and without generating per-cell objects.
- The model must round-trip the range-based representation back out as a compact `sqref`, not a
  cell-exploded one.

## Prior art / root cause

The blow-up is a whole-column `sqref` expanded into a per-cell map on read. Excel and other
generators routinely attach validations to entire columns, so this is common, not exotic. The
same range-based representation that fixes the memory problem also makes round-trip serialization
compact and correct.

## Open questions

- Internal representation: an interval/range set keyed per sheet, with membership tests over
  ranges rather than a per-cell hash.
- Caps for the hostile-input path: maximum number of rules, maximum ranges per rule, and refusal
  (or documented truncation with a `log`-style signal) beyond a threshold — consistent with the
  fork's zip-bomb / unbounded-allocation stance.
- How the public API exposes "the validation on cell X" without ever materializing the full
  covered set.

Related: `bounded-memory-large-workbook-read`, `excel-repair-on-open-structural-constraints`.
