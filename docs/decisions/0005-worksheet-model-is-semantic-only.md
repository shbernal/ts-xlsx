# ADR 0005 — `WorksheetModel` stays a semantic value; attached parts are out of scope

**Status:** Accepted (2026-07-18) · Phase 3 core slice

## Context

`Worksheet.model` is the sheet-to-sheet transfer contract: `dst.model = src.model`
snapshots a sheet and reproduces it elsewhere. Its documentation claimed the snapshot
was "everything that defines the sheet" and that the round-trip "drops nothing" — a
promise that exists to prevent a historical silent-data-loss bug (a getter field the
setter ignored).

That promise was false. The getter carried the sheet's value and overlay content
(cells, styles, column/row/page metadata, merges, data validations, conditional
formattings, tables, protection) but silently omitted every *attached part*: anchored
and background images, the autofilter, authored pivots, loaded pivots, and the
byte-preserved references (charts, vector drawings, slicers) captured for
round-tripping. A `dst.model = src.model` on a loaded sheet quietly dropped all of them.

The parked question was whether to make the model *faithful* by carrying those parts —
concretely, whether to put opaque preserved-part **byte closures** (and loaded-pivot
records) onto the public model so the "drops nothing" claim became true.

## Decision

1. **`WorksheetModel` remains a semantic, serialisable value type.** It carries pure
   sheet state; it does not carry attached parts that hold *workbook-level identity*.
   Those parts cannot be made correct inside a value snapshot: image bytes live on the
   `Workbook`, a pivot's source references a live `Worksheet` instance, and preserved
   parts are opaque package bytes with their own relationship graph — none survives a
   copy that only sees one sheet's state. Cramming byte blobs into the model would
   break its character (the types are the docs) without being correct cross-workbook.
   **Byte-closures-on-the-model is declined, not merely deferred.**

2. **The autofilter is in scope and was added.** Unlike the attached parts, an
   autofilter is pure sheet-level state — a range plus criteria, workbook-independent
   (its `_FilterDatabase` defined name is derived by the writer at emit time, not
   stored), in the exact category the model already carries (`dataValidations`,
   `conditionalFormattings`). Its omission was an unprincipled gap, so `autoFilter` now
   rides the getter and setter like the other overlays.

3. **The docs are made honest.** The `WorksheetModel` and `model`-getter comments now
   scope the "drops nothing" claim to value + overlay content and name the boundary
   explicitly: attached parts carrying workbook-level identity stay with their source
   sheet, and a model assignment *neither copies nor clears* them. The silent trap
   becomes a documented boundary.

4. **A faithful whole-sheet/package copy primitive is deferred pending a consumer.**
   Transferring attached parts (a chart, a pivot, an image) between sheets would be a
   *new* public API, not a change to `model`. Nothing in the codebase needs it yet, and
   building it now would be speculative abstraction (CLAUDE.md §4). It waits for a real
   use-case to force its shape.

## Consequences

- **Positive:** the model stays a clean, serialisable value with no opaque bytes; the
  autofilter is no longer silently dropped; the transfer boundary is documented rather
  than discovered by data loss.
- **Negative / deferred:** `dst.model = src.model` still does not transfer images,
  pivots, or byte-preserved charts/drawings between sheets — now stated in the contract,
  but a consumer who expected a deep clone must copy those parts by other means (none
  yet exists).
- **Revisit when:** a concrete consumer needs to copy a sheet's attached parts (chart,
  pivot, image) to another sheet or workbook. At that point design the dedicated copy
  primitive around that use-case — do not retrofit byte closures onto `WorksheetModel`.
