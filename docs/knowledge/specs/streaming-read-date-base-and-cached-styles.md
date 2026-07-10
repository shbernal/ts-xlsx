# Streaming read of date cells: 1904 date base and cached/compact styles

Cluster: streaming

## Scenario

A large workbook is read row-by-row with the streaming reader. Two failure modes around date cells
appear, both rooted in the streaming path resolving less context than the buffered read:

1. **1904 date base ignored.** The workbook declares the 1904 date system
   (`workbookPr/@date1904`), under which serial `0` is 1904-01-01 rather than 1900-01-01 — a
   constant ~1462-day offset. The buffered read applies this base when converting a date-formatted
   serial to a `Date`; the streaming read resolves the date without the workbook-level base, so every
   date comes out ~4 years off (or defaults to 1900 semantics) — silently wrong, not an error.
2. **Cached / compact style mode.** When style resolution runs in a cache/compact mode (styles
   interned and looked up by index rather than fully materialized per cell), the streaming reader's
   date-typing path can fail to resolve a cell's number format at all — degrading date cells to raw
   serials or crashing, where the buffered read succeeds.

> Spec note, not a corpus case here: the *core* streaming-date defect — a date-formatted numeric cell
> streaming as a raw serial because styles are not applied when typing — is already locked as a
> known-open corpus case (`streaming-read-applies-date-format`). What this note carries is the two
> distinct *root causes* that the existing case does not exercise: the workbook-level 1904 base and
> the cached/compact style mode. Both need a fixture pair (a 1904-based workbook, and one exercised
> under both buffered and cached-style streaming reads) plus a streaming-read harness that surfaces
> per-cell resolved values under a chosen style mode — new adapter surface. Captured now so the
> rewrite covers the whole surface, not just the default-style 1900-base path.

## Desired behavior

- **Date base is workbook-level and applies to every read path.** The 1900-vs-1904 base is resolved
  once from `workbookPr/@date1904` and applied identically by the streaming and buffered readers, so a
  date-formatted cell yields the same `Date` under both. Absent the flag, both default to the 1900
  base rather than crashing on missing date-base context.
- **Streaming date typing does not depend on full style materialization.** Resolving a cell's number
  format (to decide date-vs-number) works whether styles are fully materialized or held in a
  cache/compact form. A date cell reads as a date in every style mode; no mode crashes or silently
  downgrades a date to a raw serial.
- **Parity is the invariant.** For the same workbook, streaming and buffered reads agree on both the
  *type* (date) and the *value* (same epoch offset applied) of every date-formatted cell.

## Open questions

- Where is the date base threaded into the streaming pipeline — is `workbookPr` guaranteed to be read
  before the first worksheet's cells stream, or can a worksheet entry precede the workbook part in the
  zip (the same ordering hazard as shared strings and styles)? If it can, the base must be buffered or
  the affected cells deferred.
- Does the compact/cache style mode intern the number-format *code* (recoverable) or only a resolved
  style *object* (which may omit the numFmt needed for date detection)? The fix differs accordingly.
- The 1904 base also affects date *validation* operands and formula date results, not just plain
  cells — confirm those paths share the same base resolution.

Related: `streaming-read-applies-date-format` (the default-style/1900 face of the same defect),
`streaming-reader-styles-option-and-defaults`, `streaming-read-styles-before-cells`,
`date-value-timezone-conversion`, `xlsx-date-detection-control`.
