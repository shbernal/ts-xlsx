# Xlsx Errors

<!-- Generated from the public types by `pnpm run docs`. Do not edit by hand. -->

### `UnsupportedFormat`

<sub>type</sub>

Which unsupported input the reader recognised:
- `'xls'` — a legacy BIFF `.xls` (an OLE2/CFB compound file), detected by its magic bytes.
- `'xlsb'` — a binary BIFF12 `.xlsb`: the same OPC/ZIP container as `.xlsx`, but its office document
  is `xl/workbook.bin` rather than `xl/workbook.xml`. Reading it is a future capability; for now it is
  classified rather than mis-parsed.
- `'unknown'` — not a recognised spreadsheet at all: not a ZIP, or a ZIP carrying no OOXML workbook
  part (nor a `.xlsb` binary one).

```ts
type UnsupportedFormat = 'xls' | 'xlsb' | 'unknown';
```

---

### `UnsupportedFormatError`

<sub>class</sub>

Thrown when input is not a readable `.xlsx` package. The single `format` field is the branch a
caller keys on (rather than a subclass per format), so a `catch` can distinguish a legacy `.xls`, a
binary `.xlsb`, and an unrecognised blob without string-matching the message.

The message never carries a filesystem path or the underlying zip library's internals — the whole
point of the type is that the classification, not a leaked lower-layer string, is what the caller sees.

```ts
class UnsupportedFormatError extends Error {
  override readonly name = 'UnsupportedFormatError';
  readonly format: UnsupportedFormat;
}
```
