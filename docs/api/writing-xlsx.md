# Writing .xlsx

<!-- Generated from the public types by `npm run docs`. Do not edit by hand. -->

### `writeXlsx`

<sub>function</sub>

Serialise a workbook into an `.xlsx` package.

```ts
function writeXlsx(workbook: Workbook, options: WriteOptions = {}): Uint8Array;
```

**Throws** — if the workbook has no worksheets (a zero-sheet package is corrupt),
or holds a value the writer cannot yet represent.
