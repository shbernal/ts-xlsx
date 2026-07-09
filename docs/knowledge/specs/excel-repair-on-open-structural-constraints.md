# Microsoft Excel needs to recover XLSX before opening

## Producing .xlsx files Excel opens without a "repair" prompt

### Problem
A file written by the library can be well-formed XML, open cleanly in LibreOffice and Google Sheets, yet still make Microsoft Excel display *"Excel found unreadable content... Do you want to recover the contents of this workbook?"* On recovery Excel logs entries like `Repaired Records: Cell information from /xl/worksheets/sheet1.xml`, `Worksheet properties from /xl/workbook.xml`, `String properties from /xl/sharedStrings.xml`, or `Table from /xl/tables/table1.xml`, then strips or rewrites the offending part. Excel's recovery log names only the *part*, never the *cause*, so users debug blindly.

This is not one bug. It is the absence of a validation/sanitization layer that enforces Excel's structural rules at write time, plus a couple of concrete correctness defects. The durable goal: **if a workbook can be represented and Excel accepts it, we write it faithfully; if it cannot, we fail loudly at write time (or sanitize with an explicit, documented policy) — never silently emit a file Excel must repair.**

### Excel constraints a writer must respect (each independently observed causing the repair prompt)
- **Worksheet name**: max 31 characters; may not contain `\ / ? * [ ] :`; may not be blank; must be unique (case-insensitive). Violations → `Worksheet properties` repair.
- **Table name**: must be unique across the *entire workbook* (not just the worksheet); must begin with a letter or underscore; may not contain spaces, dashes, or other non-identifier characters; may not collide with a cell reference like `C1`/`P1`. Violations → `Table` repair.
- **Table column headers**: must be unique within a single table. Duplicate header text (e.g. two `Fri/24` columns) → `Table` repair. Note: real Excel silently de-duplicates by appending a suffix; a faithful writer should do the same or reject.
- **Cell text length**: a single cell value may not exceed 32767 characters → `Cell information` / `String properties` repair.
- **Data validation dropdown list**: the combined length of an inline list exceeds Excel's limit (~255 chars) → repair.
- **Numeric/date values**: `NaN` (from `parseFloat(null)`, `Number(null)`, a division producing NaN, or a null in a date column) must never be serialized as a cell value or a coordinate. See defects below.
- **Formulas**: separator is always `,` regardless of locale; the library must not localize `,`→`;`. (User confusion, not a library defect, but worth documenting.)

### Concrete correctness defects worth regression coverage
1. **Whole-column print area corrupts to NaN.** A defined print-area range with column-only bounds — e.g. `$A:$F` — round-trips (read then write) to `$ANaN:$FNaN`. The empty row component is being coerced to a number, yielding `NaN`. Correct behavior: whole-column (and by symmetry whole-row) ranges must survive a round-trip unchanged and never produce `NaN` in an address. This is a real, reproducible parser/serializer bug and belongs in the address-decoding corpus.
2. **Empty rich-text run with a font produces an invalid shared string.** A cell whose value is `{ richText: [{ text: '', font: { bold: true } }, { text: 'test' }] }` emits a `sharedStrings.xml` entry Excel repairs (`String properties`). Correct behavior: an empty run must be dropped or serialized in a form Excel accepts; a written rich-text string must always be a valid, Excel-openable shared string.

### Prior art
- Excel specifications and limits (Microsoft support): the authoritative list of the numeric/length/naming ceilings above.
- Real Excel silently sanitizes many of these (de-duplicating table columns, truncating over-long strings) rather than rejecting — a reasonable model for our writer's default policy, provided the sanitization is documented and optionally strict.

### Open questions
- Policy: for each constraint, do we (a) throw at write time, (b) sanitize-and-warn, or (c) sanitize silently to match Excel? Leaning: sanitize-and-warn by default, with a strict mode that throws. Naming/uniqueness violations that would silently drop user data (duplicate table names) should probably throw.
- Where does validation live — at mutation time (e.g. `addTable`/`addWorksheet` rejects a bad name immediately, giving a good stack trace) or only at serialization time? Mutation-time gives the best developer experience; serialization-time is the last-resort safety net. Likely both.
- Do we surface a machine-readable diagnostics list (part + cause) so callers can inspect what was sanitized, instead of only console warnings?
