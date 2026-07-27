# Xlsb Read

<!-- Generated from the public types by `pnpm run docs`. Do not edit by hand. -->

### `readXlsb`

<sub>function</sub>

Read an `.xlsb` (binary BIFF12) package into a `Workbook`.

```ts
function readXlsb(data: Uint8Array, options: ReadXlsxOptions = {}): Workbook;
```

**Throws** — if the input is not an `.xlsb` package — a legacy `.xls`
(`.format === 'xls'`), an XML `.xlsx` or unrecognised blob (`'unknown'`).
**Throws** — if a binary part is malformed.
**Throws** — if the archive exceeds the inflate bound (a probable zip bomb).
