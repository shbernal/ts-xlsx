# Serialize a single worksheet part, and swap a sheet into an existing package

Cluster: streaming

## Scenario

A user has a rich template with features the library does not fully preserve on round-trip (charts,
named ranges, pivots). They want to populate just one worksheet with data, then splice that
populated sheet's XML into the untouched original package — replacing one sheet part while leaving
everything else byte-untouched — so the unpreserved features survive. Today they hand-unzip both
packages, copy the raw `xl/worksheets/sheetN.xml` text across, and re-zip. That is fragile because a
worksheet part is **not self-contained**: it references shared strings by index, styles via `cellXfs`
`s=` indices, and drawing/comment/table parts by relationship id — all of which differ between the
two packages.

> Spec note, not a corpus case: this is a new I/O capability. The durable value is the two-level API
> and the cross-reference reconciliation that makes a raw sheet-XML copy unsafe.

## Desired behavior

- **(1) Emit one worksheet's part.** Serialize an individual worksheet to its OOXML worksheet-part
  XML fragment (`xl/worksheets/sheetN.xml`), decoupled from a full-workbook write — given its cell
  contents and the workbook context it needs (shared strings, styles).
- **(2) Sheet swap.** A higher-level operation that replaces a single sheet part inside an existing
  package while leaving all other parts (charts, named ranges, defined names, drawings, VML, …)
  intact — reconciling the incoming sheet's **shared-string indices, style (`cellXfs`) indices, and
  relationship ids** into the target package rather than assuming they match. This is the safe
  version of the manual unzip-copy-rezip workaround.
- Both are bounded/streaming-friendly so a large sheet can be emitted without buffering the whole
  workbook (ties to the streaming-write family).

## Open questions

- Does the emit-one-sheet API take a live worksheet from a loaded workbook (so it can resolve shared
  strings/styles), or a standalone sheet + an explicit shared-string/style context?
- Sheet swap conflict policy: merge the incoming sheet's shared strings/styles into the target's
  tables (remapping indices), or require the two packages to already share them?
- How are relationship-backed sheet features (drawings, comments, tables) on the incoming sheet
  carried across — remapped rel ids + copied parts, or rejected if present?
- Is this preferable to the more general `roundtrip-preserves-unmodeled-package-parts` approach
  (preserve everything the model doesn't touch), or complementary?

Related: `roundtrip-preserves-unmodeled-package-parts`, `preserve-drawing-shapes-on-roundtrip`,
`streaming-write-per-sheet-memory-release`, `streaming-read-modify-write-template`.
