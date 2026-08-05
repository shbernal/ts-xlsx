# Xlsb Read

<!-- Generated from the public types by `pnpm run docs`. Do not edit by hand. -->

### `readXlsb`

<sub>function</sub>

Read an `.xlsb` (binary BIFF12) package into a `Workbook`.

```ts
function readXlsb(data: Uint8Array, options: ReadXlsxOptions = {}): Workbook;
```

**Throws** — `UnsupportedFormatError` if the input is not an `.xlsb` package — a legacy `.xls`
(`.format === 'xls'`), an XML `.xlsx` or unrecognised blob (`'unknown'`).
**Throws** — `XlsbParseError` if a binary part is malformed.
**Throws** — `PackageReadError` if the input is a ZIP that cannot be unpacked — a corrupt or
truncated archive, or one exceeding the inflate bound (a probable zip bomb).
