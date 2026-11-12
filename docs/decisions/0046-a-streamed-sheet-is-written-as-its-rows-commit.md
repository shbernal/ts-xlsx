# ADR 0046: A streamed sheet is written into its zip entry as its rows commit

**Status:** Accepted (2026-11-12) · I/O API · breaking, shipped in 4.0.0 · fixes #33

## Context

Issue #33 reported two defects in `WorkbookStreamWriter`. PR #47 fixed the second: the package is
fed to the deflater in 64 KiB slices and the writer waits for a slow sink's `drain`. The first was
left open. A committed row dropped its cell objects but kept its serialised `<row>` XML in
`#flushedRows` until the workbook committed, and `buildPackageParts` then built every sheet's XML as
one string. Peak memory grew with the row count, smaller per row than the cell graph but unbounded,
and the README said it stayed flat.

Bounding it means writing a committed row straight into the sheet's zip entry, which has to be open
from the first row on. The entry is deflated as it goes, so nothing written can be revised. Five
things the old writer decided at workbook commit, after every row was known, cannot be decided that
way:

- `<dimension>` precedes `<sheetData>` and is the bounding box of every used cell.
- `<cols>`, `<sheetFormatPr>`, `<sheetViews>` and `<sheetPr>` precede `<sheetData>` too, and
  `<sheetFormatPr outlineLevelRow>` is the deepest row outline level on the sheet.
- `collapsed="1"` rides a summary row when its whole detail group is hidden. With summary rows
  above their detail (`summaryBelow: false`), that is a fact about rows written after it.
- `getCell` could reach a row number below one already committed, which the old writer sorted
  into place at the end.
- A sheet's part went through `buildPackageParts` like every other part.

Option B in the doubts record was to deflate each sheet's rows into a per-sheet compressed buffer.
It cuts memory roughly five to ten times and changes no output, but the memory still grows with the
sheet, so it does not fix the issue.

## Decision

On a streaming writer with inline strings (the default), each sheet's part is written into its own
zip entry while the rows are being produced.

1. **The entry opens at the first flush.** That is the first `row.commit()` or the sheet's
   `commit()`, whichever comes first. The writer writes the part's head (everything before the
   first `<row>`) from the model as it stands at that moment. Committed rows are rendered, deflated
   in batches of about 64 KiB of XML, and handed to the sink. At sheet commit the writer renders the
   rows still open, writes the tail (everything after `</sheetData>`), and closes the entry.
   Memory now depends on the rows not yet committed, not on the sheet's length.

2. **`<dimension>` is omitted from a streamed sheet.** It is optional in CT_Worksheet, and Excel and
   this library's reader compute the used range when it is absent. A placeholder such as `A1` would
   be wrong for every sheet that has more than one cell. The buffered writer, and a streaming writer
   with `useSharedStrings`, still write it.

3. **The head is frozen at the first flush.** Column definitions, the sheet view, sheet properties
   and outline placement are read once. At sheet commit the writer renders the head again. If the
   model changed anything in it after the first flush, the sheet commit throws an `AuthoringError`,
   so the change is not silently dropped from the file. Set them before the first row commits.

4. **The row outline depth is declared or seen at the first flush.** `outlineLevelRow` is the
   larger of the new `outlineLevelRow` option of the streaming `addWorksheet` and the deepest level
   among the rows in the first flush. A row committed later with a deeper level throws an
   `AuthoringError` naming the option. An outline bar sized too small is still a schema-valid file,
   so no validator would catch it.

5. **`collapsed` is derived from the rows written with or before the summary row.** Rows are written
   in ascending order. With summary rows below their detail (Excel's default), the detail group has
   always been written by the time its summary row is, so the writer derives the flag from a
   per-level running state of the trailing rows. That state is constant-size, not a per-row map.
   With summary rows above, the writer derives the flag only from detail rows in the same flush
   batch. To get it derived, commit the group's last detail row; committing a row also commits every
   uncommitted row above it. Otherwise set `collapsed` on the summary row yourself.

6. **A write behind the flush point throws.** Committing row _n_ first commits every uncommitted
   row above it, so rows are written in ascending order. `getCell` on a row at or above the last
   written row throws an `AuthoringError`. It previously threw only for a row that was itself
   committed.

7. **Checks that span rows run at flush time.** An array formula's range conflicts only with
   formulas from its own row down to its last row. The writer keeps the array ranges that are still
   open and checks each batch against them, so the conflict is reported at `row.commit()`, not at the
   workbook commit. Notes are still collected per row, because they belong to the comments part.

8. **Streamed sheet parts leave `buildPackageParts`.** The sheet's tail needs the relationship ids
   its sheet-local ledger assigns: the drawing, the notes' VML, hyperlinks, tables and printer
   settings. At sheet commit the writer plans the whole package and takes this sheet's plan. The
   ids depend only on the sheet's own content, and a committed sheet cannot change, so planning the
   package again at workbook commit gives the same ids. `buildPackageParts` takes the streamed
   sheets, emits their rels, comments and drawings but not the sheet parts, and throws an
   `InternalError` if the ids it plans differ from the ones already written.

9. **Commits return promises for backpressure.** `row.commit()` and the sheet's `commit()` return
   a `Promise<void>` that resolves once the caller's sink can take more. A producer that awaits it
   is paced by the sink; one that does not still gets a correct file, with the sink's buffer
   absorbing the difference. The promise never rejects. A sink error is recorded, the writer stops
   writing to the sink, and the workbook's `commit()` rejects with that error.

## Consequences

- **Breaking, so 4.0.0.** The changes are the `Promise` return types, the `getCell` rule, the
  frozen head, the outline-depth rule, `collapsed` with summary rows above, the missing
  `<dimension>`, and array-formula conflicts reported at row commit.
- **Sheets written one after another stay bounded; interleaved sheets do not.** A zip stream holds
  one entry at a time. If sheet 2 is written while sheet 1's entry is still open, fflate holds sheet
  2's compressed bytes until sheet 1 closes. To keep memory bounded, commit each sheet before
  writing the next.
- **Retained state that is still proportional:** the notes collected for the comments part, the
  hyperlinks, merges, validations and conditional formats on the sheet, the style registry, the
  `stream` getter's copy when no sink is given, and the array formula ranges still open. Each grows
  with authored content, not with the number of rows.
- **`useSharedStrings` is unchanged.** It keeps every row live until the workbook commits and goes
  through `buildPackageParts`, because the pool's indices come from the whole workbook. A streamed
  shared-strings pool could also be bounded by unique strings, not rows, but that is a separate
  change.
- **Archive order changes.** Streamed sheet entries come first in the archive, ahead of
  `[Content_Types].xml`. OPC does not order entries, and the reader resolves parts through the
  central directory.
- **`writer.stream` with a sink must be taken before output starts.** Output now starts at the
  first flush. A `stream` getter reached after that would miss the bytes already written, so with a
  sink it throws. Without a sink every chunk is kept, as before, and a late `stream` replays them.
