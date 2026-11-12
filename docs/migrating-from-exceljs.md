# Migrating from ExcelJS

`ts-xlsx` started as a fork of ExcelJS, but **it is not a drop-in replacement**. The API
changed on purpose wherever a better shape existed (see [`AGENTS.md`](../AGENTS.md) §1).
This page maps common ExcelJS patterns to their `ts-xlsx` equivalents so a port is
mechanical, and lists what has not been rebuilt.

Treat this as a translation guide, not a compatibility promise. Pin a version; the API
still breaks between major versions.

## The three shifts that cover most code

### 1. I/O is synchronous and byte-native: free functions, not `workbook.xlsx.*`

ExcelJS did I/O through async methods on the workbook that assumed Node `Buffer`s and
streams. `ts-xlsx` reads and writes plain `Uint8Array` synchronously, so the same call
works in Node and the browser, with nothing to `await` on the buffered path.

```ts
// ExcelJS
const wb = new ExcelJS.Workbook();
await wb.xlsx.readFile('in.xlsx');
await wb.xlsx.writeFile('out.xlsx');
const buf = await wb.xlsx.writeBuffer();

// ts-xlsx: I/O is separate from the model, synchronous, and Uint8Array in/out
import {Workbook, readXlsx, writeXlsx} from '@shbernal/ts-xlsx';
import {readFileSync, writeFileSync} from 'node:fs';

const wb = readXlsx(readFileSync('in.xlsx'));
writeFileSync('out.xlsx', writeXlsx(wb));
const bytes = writeXlsx(wb); // Uint8Array
```

Reading and writing files is the caller's job. Outside the `/node` streaming writer,
`ts-xlsx` never touches the filesystem, which keeps it browser-safe.

### 2. Cell values are one precisely-typed union

ExcelJS typed `cell.value` loosely and modelled formulas, hyperlinks and rich text as
ad-hoc object shapes. In `ts-xlsx`, `cell.value` is the single
[`CellValue`](api/cell-values.md) union: `number | string | boolean | Date | null`
plus typed formula, rich-text and error shapes. `null` is the empty cell.

A hyperlink is not a cell value any more. It sits on the sheet, over a cell or a range:
`sheet.addHyperlink({ref: 'B2', target: 'https://example.com'})`.

```ts
sheet.getCell('A1').value = 42;
sheet.getCell('A2').value = new Date();
sheet.getCell('A3').value = {formula: 'SUM(A1:A2)', result: 42};
sheet.getCell('A4').value = null; // empty
```

### 3. Absent axes are `undefined`, never sentinels

A whole-row reference (`$1`) has no column; a whole-column reference (`$A:$A`) has no
row. ExcelJS let those decay into `NaN` or `"undefined"`, which leaked into serialized addresses.
`ts-xlsx` models an omitted axis as `undefined` on
[`CellAddress`](api/addresses-ranges.md) and never fabricates a sentinel.

## Quick reference

| ExcelJS                         | ts-xlsx                                                       |
| ------------------------------- | ------------------------------------------------------------- |
| `await wb.xlsx.readFile(path)`  | `readXlsx(readFileSync(path))`                                |
| `await wb.xlsx.load(buffer)`    | `readXlsx(bytes)`                                             |
| `await wb.xlsx.writeFile(path)` | `writeFileSync(path, writeXlsx(wb))`                          |
| `await wb.xlsx.writeBuffer()`   | `writeXlsx(wb)` → `Uint8Array`                                |
| `await wb.csv.readFile(path)`   | `readCsv(readFileSync(path, 'utf8'))`                         |
| `await wb.csv.writeBuffer()`    | `writeCsv(wb)` / `writeCsvText(wb)`                           |
| `wb.addWorksheet('S')`          | `wb.addWorksheet('S')` _(unchanged)_                          |
| `wb.getWorksheet('S')`          | `wb.getWorksheet('S')` → `Worksheet \| undefined`             |
| `sheet.getCell('A1').value = …` | `sheet.getCell('A1').value = …` _(unchanged)_                 |
| `sheet.addRow([…])`             | `sheet.addRow([…])` _(unchanged)_                             |
| streaming `WorkbookReader`      | `readSheetRows(bytes, {sheet})` / `readWorkbookStream(bytes)` |
| streaming `WorkbookWriter`      | `WorkbookStreamWriter` from `@shbernal/ts-xlsx/node`          |

Where a method name is unchanged, its types are still stricter. `getWorksheet` returns
`Worksheet | undefined`, so handle the miss, and the arguments are precisely typed.

## What is not here (yet)

An API lands only once it is strictly typed and pinned by tests, so some ExcelJS features
are not here yet. The buffered writer refuses a value it cannot represent faithfully instead
of writing a lossy package. If you need a feature missing from the
[API reference](api/README.md), [open an issue](https://github.com/shbernal/ts-xlsx/issues).
Each one becomes a regression case (see [`architecture.md`](architecture.md)).

**Charts, vector shapes, slicers and legacy form controls are round-trip-only, by
decision.** `ts-xlsx` keeps them byte for byte through a load, edit and save, but has no
authoring API for any of the four. See [`docs/api/preserved.md`](api/preserved.md) and
[ADR-0014](decisions/0014-charts-shapes-slicers-are-round-trip-only-for-1-0.md) for the
reasoning and what it would take to pick one up.

## Why break compatibility at all?

Keeping the old API would have been easier and made the library worse
([`AGENTS.md`](../AGENTS.md) §5). What you get for the port:

- `.d.ts` types precise enough that the API reference is generated from them.
- An I/O path with no Node built-ins, so the same code runs in a browser.
- Behavior pinned by a regression corpus built from ExcelJS's issue backlog.
