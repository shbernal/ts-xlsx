# Tolerate the row-number quirks foreign generators emit that Excel accepts

Cluster: streaming

## Scenario

A user loads an `.xlsx` produced by a third-party server-side spreadsheet engine, not Excel, directly
from a download. The parser throws "Invalid row number" and refuses to open the file. Opening the
same file in Excel and re-saving it unchanged makes it load, the canonical signal that the input is
within the real-world tolerable envelope even though it trips a strict check.

> Spec note, not a corpus case: probing shows this library's current reader already tolerates
> injected out-of-range and zero row numbers, which load without throwing, but a synthesized fixture
> may not reproduce the exact `<row r>` shape the foreign engine emits. Rather than lock a possibly
> unrepresentative fixture, the durable requirement, to tolerate any row-number quirk Excel tolerates,
> is captured here. If a real foreign fixture that reproduces the throw surfaces, promote it to a
> tolerant-read corpus case.

## Desired behavior

- Reading a worksheet **tolerates the row-index quirks Excel itself tolerates** from foreign
  generators rather than throwing a hard "Invalid row number" that makes an otherwise-valid workbook
  unopenable.
- When a `<row>` element carries a **missing, empty, zero, or out-of-sequence `r` attribute**, the
  reader either infers the row number from document order and the row's position in the `sheetData`
  stream, or rejects only that individual malformed row while still surfacing the rest of the
  sheet. It never aborts the whole load.
- If a hard rejection is chosen for a truly invalid value, the error is **clear and actionable**,
  naming the offending value and its cell or row context, not a bare "Invalid row number".

## Open questions

- Infer-from-position or skip-the-bad-row: which default? Inference preserves more data but can
  misplace cells if the stream is genuinely disordered.
- What counts as "out of range": clamp to the sheet's 1,048,576-row maximum, or accept and let
  downstream bounds checks handle it?
- Does the same tolerance apply to column (`r` or cell address) quirks, or only row numbers?

Related: `read-workbook-missing-app-properties`, `foreign-generator-boolean-and-mixed-sharedstring`,
`streaming-reader-preserves-blank-row-numbers`, `unsupported-input-format-typed-error`.
