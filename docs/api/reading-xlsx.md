# Reading .xlsx

<!-- Generated from the public types by `pnpm run docs`. Do not edit by hand. -->

### `readXlsx`

<sub>function</sub>

Read a spreadsheet package into a `Workbook`.

Both OOXML serialisations are accepted: an XML `.xlsx`, and a binary `.xlsb` (BIFF12), which is the
same OPC container with binary office-document parts. The two are auto-detected from the package
itself rather than from a file extension, so a caller never branches on which form it holds — and
the model produced is the same either way. See `../xlsb/read.ts` for what the binary path does not
yet decode.

```ts
function readXlsx(data: Uint8Array, options: ReadXlsxOptions = {}): Workbook;
```

**Throws** — if the input is neither — a legacy `.xls` (`.format === 'xls'`) or
an unrecognised/non-ZIP blob (`'unknown'`).
**Throws** — if a binary `.xlsb` part is malformed.
**Throws** — {
