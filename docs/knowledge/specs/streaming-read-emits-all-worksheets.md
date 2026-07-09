# Streaming read must emit every worksheet, including large sheet counts

Cluster: streaming

## Scenario

A workbook with a large number of worksheets (well over 100 — one report cites ~188) is read with
the streaming reader. The streaming reader should emit every worksheet the workbook declares, in
order, and complete without stalling. In practice the many-sheet path is unreliable: worksheets
are dropped, or enumeration crashes.

> A synthesized 150-sheet workbook fed through the streaming reader crashes in the worksheet
> parser (`this.model.sheets` is undefined for a worksheet), indicating the workbook model is not
> reliably available before worksheets are parsed when there are many sheets. The originally
> attached repro file is a broken HTML download (not a zip), so the exact reported symptom
> ("drops worksheets") could not be reproduced from it — captured here as desired behavior rather
> than a corpus case with a fragile assertion.

## Desired behavior

- Streaming-reading a workbook with any number of worksheets emits a worksheet for **every** sheet
  part the workbook declares — the count surfaced equals the count of declared sheets.
- Worksheets are emitted in workbook order, each discoverable by its declared name.
- Enumeration completes (the stream never stalls) regardless of sheet count.
- The workbook-level model (sheet list, relationships, shared strings) needed to resolve a
  worksheet must be available *before* that worksheet is emitted — parsing order must guarantee it,
  so a worksheet handler never dereferences an unpopulated model.

## Open questions

- Ordering guarantee: does the streaming reader parse `workbook.xml` + `_rels` fully before
  emitting any worksheet, or lazily? The crash above suggests the model must be resolved first.
- Memory: many sheets must not force the whole workbook into memory — only the workbook-level
  index (small) plus the current worksheet's bounded window. Ties into
  `bounded-memory-large-workbook-read`.
- Error semantics if a declared sheet part is missing or unreadable — skip-with-signal vs. throw.
- A regression check for this belongs in a harness that can synthesize an N-sheet workbook and
  assert `emitted === declared`, since a durable fixture would be large.

Related: `bounded-memory-large-workbook-read`, `streaming-read-resolves-shared-strings`,
`streaming-delete-worksheet-before-write`.
