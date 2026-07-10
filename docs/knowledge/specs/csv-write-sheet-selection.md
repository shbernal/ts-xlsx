# CSV write must make multi-sheet selection explicit, not silently drop sheets

Cluster: csv

## Scenario

A user builds a workbook with several worksheets and asks the library to serialize it to CSV. CSV is
a flat, single-table format with no concept of multiple sheets, so the writer can represent only one.
Today it silently picks the first worksheet and drops the rest — no error, no warning, no way to
influence the choice. Users are surprised to find only one sheet's data in the output and cannot
target a specific sheet. The recurring confusion over many years shows the silent drop, not the
single-sheet limitation itself, is the actual pain.

> Spec note, not a corpus case: the current output is well-formed (it just contains the wrong sheet
> for a multi-sheet workbook), so there is no malformed serialization to assert — the defect is a
> silent-data-loss usability choice plus an undecided default. The durable value is the selection
> contract.

## Desired behavior

- CSV write accepts an option naming the worksheet to serialize (by name or index). When provided,
  that sheet is written; if the named/indexed sheet does not exist, **fail with a clear error naming
  the available sheets** rather than silently falling back.
- When the option is omitted and the workbook has exactly one worksheet, write it with no ceremony
  (the current happy path).
- When the option is omitted and the workbook has **more than one** worksheet, surface the ambiguity
  rather than silently taking the first. Erroring is safer (no silent data loss) and fits the fork's
  correctness-first stance, though it is a harder break for callers relying on first-sheet behavior;
  the alternative is default-to-first with a diagnostic.

## Open questions

- Omitted selection with multiple sheets: error, or warn-and-default? (correctness vs convenience)
- A batch mode that writes one CSV per worksheet (multiple named buffers, or a directory) for users
  who genuinely want every sheet exported?
- How is "active sheet" vs "first sheet" resolved when both concepts exist?
- The selection option lives under the existing CSV write `options` surface.

Related: `csv-write-date-format-honored`, `path-reader-is-node-only-clear-error`.
