# Reject non-.xlsx input with a clear, typed error, not raw zip internals

Cluster: xlsx-io

## Scenario

A user hands the reader a legacy binary spreadsheet (the old BIFF `.xls` format) or any other file
that is not an OOXML `.xlsx` package. Because `.xlsx` is a ZIP container and `.xls` is not, the zip
layer fails deep in its own internals with an opaque message like *"Can't find end of central
directory : is this a zip file?"*. The user cannot tell from that message that the real problem is a
wrong/unsupported format — and the raw error may even leak absolute local filesystem paths from the
zip layer.

> Spec note, not a corpus case: the desired behavior is an error *contract* (message clarity, error
> type, no path leakage) rather than a serialization property the corpus asserts on a package. The
> durable value is the classification of input-format failures and their typed errors.

## Desired behavior

- Given input that is not a valid OOXML `.xlsx` ZIP package, the reader rejects with a **clear,
  format-aware, typed error** rather than leaking a raw zip-parsing failure.
- If the input is a **legacy BIFF `.xls`** file — detectable by the OLE Compound File magic bytes
  `D0 CF 11 E0 A1 B1 1A E1` — the error says the `.xls` binary format is not supported and only
  `.xlsx`/OOXML is handled.
- If the input is **not a ZIP at all** (no `PK\x03\x04`, no end-of-central-directory record), the
  error says the file is not a valid `.xlsx` package rather than surfacing zip-internals text.
- The error is a **distinct, catchable type/category** (e.g. `UnsupportedFormatError` vs
  `CorruptPackageError`) so callers can branch programmatically, and it **does not expose absolute
  local filesystem paths** from the zip layer.

## Open questions

- How much format sniffing is in scope beyond `.xls` — detect `.xlsb` (a ZIP but binary-parts),
  `.ods`, CSV-handed-to-the-xlsx-reader, and give each a tailored message?
- Where does detection live — a small magic-byte probe in front of the zip layer, so a non-ZIP fails
  fast before the zip library runs?
- Error taxonomy: one `UnsupportedFormatError` with a `format` field, or distinct subclasses?

Related: `path-reader-is-node-only-clear-error`, `load-accepts-arraybuffer-and-typed-arrays`,
`tolerant-parse-unclosed-vml-tags`.
