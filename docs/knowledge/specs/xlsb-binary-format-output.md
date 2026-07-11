# Reading and writing the binary XLSB format

Cluster: io / formats

## Scenario

A user generating very large spreadsheets wants an alternative to the XML-based `.xlsx` format:
the binary `.xlsb` (SpreadsheetML Binary, BIFF12) form. The motivation is size and speed — a binary
workbook stores cell data as packed records rather than verbose XML, so it is typically much smaller
on disk and faster to parse and emit for large datasets. The request is for an additional
serialization format (write, and by extension read), not a bug: "produce a valid `.xlsb` that
Excel opens, and read one back."

> Spec note, not a corpus case: this is a new capability with no failing behavior to assert yet. It
> becomes assertable — round-trip a binary workbook and inspect its records/values — once a codec
> exists. The durable value is the format's shape, the architectural constraint it places on the
> rebuild, and the scoping decisions it forces.

## Desired behavior

- **XLSB is a second serializer/parser over the same logical model, not a bolt-on.** The workbook
  model (sheets, cells, styles, defined names, tables…) is independent of on-disk encoding; XLSX
  (XML) and XLSB (binary) are two codecs over it. Designing this cleanly means a pluggable
  container/part-codec layer where a part is `{contentType, encode(model), decode(bytes)}` — the
  same seam that lets the streaming writer emit either form.
- **A produced `.xlsb` opens in Excel unmodified**, and a foreign-generated `.xlsb` reads back into
  the same model an `.xlsx` would, so a caller can convert between the two without loss of the
  supported feature set.

## Format notes / prior art

- XLSB uses the **same OPC/ZIP package** and relationship graph as XLSX, but the sheet / workbook /
  sharedStrings / styles parts are **binary BIFF12 record streams** (`BrtRow`, `BrtCellRk`,
  `BrtCellIsst`, …) rather than XML. Part names and content types differ (worksheet parts are
  `.bin`, the workbook is `xl/workbook.bin`).
- The canonical reference is the open **[MS-XLSB]** specification. SheetJS/`xlsx` implements XLSB
  read+write and is the strongest cross-check for record encodings and edge cases (RK number
  encoding, inline vs. shared strings, formula token streams).
- **Numeric encoding quirks:** `RK` values pack a 30-bit integer or a truncated float with a
  ×100-divisor flag; formulas are stored as **Ptg token streams**, not text — significantly more
  work than XLSX's plain formula strings.

## Scope decisions

- **Read before write.** Reading foreign `.xlsb` files is higher-value and lower-risk than writing
  them and could ship first; the motivating large-workbook use case then pairs write support with
  the streaming writer (binary record streams stream well).
- **Feature subset first:** values + shared strings + basic styles for read, deferring
  formulas / tables / pivots / rich formatting.

## Open questions

- Is XLSB in scope for the fork's first stable surface, or a later add-on? Full BIFF12 record
  coverage plus Ptg encode/decode is a large, self-contained sub-project.
- Formulas: decode Ptg tokens to text (and re-encode on write), store opaquely, or recompute? Full
  Ptg round-tripping is the hard part.
- **Security:** binary record parsing of untrusted input needs the same bounded-allocation and
  zip-bomb defenses as the XML path (see `bounded-memory-large-workbook-read`,
  `lean-zip-container-strategy`), plus per-record length sanity checks specific to BIFF12 — a
  malformed record length must never drive an unbounded allocation.

Related: `bounded-memory-large-workbook-read`, `lean-zip-container-strategy`,
`unified-streaming-and-buffered-io`, `unsupported-input-format-typed-error`.
