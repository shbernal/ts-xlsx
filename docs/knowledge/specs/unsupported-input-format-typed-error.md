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

## Known gap: a malformed ZIP is misreported

A `PK`-headed blob that the zip layer then rejects as malformed is classified
`UnsupportedFormatError('unknown')`, whose default message reads *"not a valid .xlsx package: no
OOXML workbook part was found"* (`io/opc/sniff-format.ts`). No part search ever ran — the archive
did not inflate at all — so the message describes a check that did not happen, and points an
investigation at the wrong layer.

The classification looks wrong too, by the taxonomy's own words: `PackageReadError` is documented as
"the input is the right kind of thing and we will not, or cannot, unpack it", which is exactly this.
Left unchanged deliberately, because it moves every truncated or corrupt package from code
`unsupported-format` to `malformed-input` — a behaviour change that wants its own corpus case and
its own commit, not a drive-by. What must not change with it: the underlying zip text stays
discarded rather than wrapped, since it can name an absolute filesystem path.

## Open questions

- How much format sniffing is in scope beyond `.xls` — detect `.xlsb` (a ZIP but binary-parts),
  `.ods`, CSV-handed-to-the-xlsx-reader, and give each a tailored message?
- Where does detection live — a small magic-byte probe in front of the zip layer, so a non-ZIP fails
  fast before the zip library runs?
- Error taxonomy: one `UnsupportedFormatError` with a `format` field, or distinct subclasses?

Related: `path-reader-is-node-only-clear-error`, `load-accepts-arraybuffer-and-typed-arrays`,
`tolerant-parse-unclosed-vml-tags`.
