# Opc Errors

<!-- Generated from the public types by `pnpm run docs`. Do not edit by hand. -->

### `PackageReadError`

<sub>class</sub>

Thrown when the input *is* a ZIP container but it cannot be unpacked: the archive is corrupt or
truncated, or inflating it would push total uncompressed output past the caller's bound (which is
how a zip bomb presents).

The neighbouring [`UnsupportedFormatError`](./opc-errors.md#unsupportedformaterror) says the input is a different *kind* of thing; this
one says it is the right kind and we will not (or cannot) unpack it. Keeping them apart is what
lets a caller answer "should I try another reader, or reject this file?" — and it is what replaced
the message-prefix match the bomb refusal used to be recognised by.

The zip library's own failure text never survives into either the message or `cause`: it can name
internals — or an absolute filesystem path — from the layer below, and this type carries the
classification precisely so no lower-layer string has to.

```ts
class PackageReadError extends XlsxError {
  override readonly name = 'PackageReadError';
  override readonly code = 'malformed-input';
}
```

---

### `UnsupportedFormat`

<sub>type</sub>

Which unsupported input the reader recognised:
- `'xls'` — a legacy BIFF `.xls` (an OLE2/CFB compound file), detected by its magic bytes.
- `'xlsb'` — a binary BIFF12 `.xlsb`: the same OPC/ZIP container as `.xlsx`, but its office document
  is `xl/workbook.bin` rather than `xl/workbook.xml`. `readXlsx`/`readXlsb` read one; the entry points
  that cannot yet (the row streamer) report it under this format with their own message.
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

`format` stays the branch for *which* unsupported input this was; the inherited
[`XlsxError.code`](./errors.md#xlsxerror) answers the coarser question of what kind of failure it is.

```ts
class UnsupportedFormatError extends XlsxError {
  override readonly name = 'UnsupportedFormatError';
  override readonly code = 'unsupported-format';
  readonly format: UnsupportedFormat;
}
```
