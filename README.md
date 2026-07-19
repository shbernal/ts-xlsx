# ts-xlsx

A TypeScript-first library for reading and writing spreadsheet documents
(`.xlsx` / OOXML, and CSV) — synchronous, `Uint8Array`-native, and dependency-lean.

> **Status — an independent rebuild in progress.**
> `ts-xlsx` began as a hard fork of [ExcelJS](https://github.com/exceljs/exceljs)
> and is being rebuilt from the ground up into a modern, strict-TypeScript library.
> **It carries no backwards-compatibility guarantee with ExcelJS** — the API below is
> its own, not a drop-in.
> See [`CLAUDE.md`](CLAUDE.md) for the goals and [`docs/architecture.md`](docs/architecture.md)
> for the design, and [migrating from ExcelJS](docs/migrating-from-exceljs.md) if you are coming across.

## Why it exists

Upstream ExcelJS is effectively unmaintained — no release since 2023, a backlog of
hundreds of issues and PRs — while still serving tens of millions of downloads a month.
`ts-xlsx` extracts the accumulated value from that backlog (as a permanent regression
corpus) and discards the debt: no untyped surfaces, no callback APIs, a single small
runtime dependency, and every behavior pinned by a test. This is not a compatibility
shim. It is a different, better library that happens to share ancestry.

What that buys you today:

- **Strict, precise types are the contract.** `strict` + `noUncheckedIndexedAccess` +
  `exactOptionalPropertyTypes`. The published `.d.ts` *is* the documentation, and the
  [API reference](docs/api/README.md) is generated straight from it.
- **Synchronous, buffer-native I/O.** `readXlsx`/`writeXlsx` take and return a
  `Uint8Array` — no streams to await for the buffered path, no Node `Buffer` assumption,
  so the same code runs in Node and the browser.
- **Streaming when you need it.** Bounded-memory row streaming for reads, so a large
  workbook never has to be fully materialized.
- **One runtime dependency** ([`fflate`](https://github.com/101arrowz/fflate) for zip),
  a hand-written SAX reader with bounded allocation on every parser path, and a
  build-free, strict-typed source tree.

## Install

```shell
npm install @shbernal/ts-xlsx
```

Requires Node ≥ 18 (or any modern browser bundler). ESM only.

## Quick start

Everything is synchronous. `writeXlsx` returns the file bytes; `readXlsx` takes them back.

```ts
import {Workbook, writeXlsx, readXlsx} from '@shbernal/ts-xlsx';

// --- write ---
const wb = new Workbook();
const sheet = wb.addWorksheet('People');

sheet.getCell('A1').value = 'Name';
sheet.getCell('B1').value = 'Joined';
sheet.addRow(['Ada', new Date('2026-01-01')]);
sheet.addRow(['Grace', new Date('2026-02-14')]);
sheet.getCell('C1').value = {formula: 'COUNTA(A:A)', result: 3};

const bytes: Uint8Array = writeXlsx(wb);

// --- read it back ---
const reopened = readXlsx(bytes);
const people = reopened.getWorksheet('People');
console.log(people?.getCell('A2').value); // 'Ada'
```

Persisting to disk is your platform's job, not the library's — `writeXlsx` hands you the
bytes:

```ts
import {writeFileSync, readFileSync} from 'node:fs';

writeFileSync('people.xlsx', writeXlsx(wb));
const wb2 = readXlsx(readFileSync('people.xlsx')); // a Buffer is a Uint8Array
```

## Core model

- **`Workbook`** — the document. `addWorksheet(name)`, `getWorksheet(nameOrId)`,
  `worksheets`, defined names, images, and workbook-level properties.
- **`Worksheet`** — a sheet. `getCell('B3')`, `addRow(values)`, `addTable(...)`,
  `mergeCells('A1:B2')`, column/row properties, page setup, and print options.
- **`Cell`** — one cell. `cell.value` is the whole story: a `number`, `string`,
  `boolean`, `Date`, `null` (empty), a formula (`{formula, result}`), rich text, a
  hyperlink, or an error — all precisely typed as [`CellValue`](docs/api/cell-values.md).

Addresses are honest: an axis a reference doesn't mention is `undefined`, never a
sentinel — see [`decodeAddress`](docs/api/addresses-ranges.md).

## Reading, writing, streaming, CSV

```ts
import {
  readXlsx, writeXlsx,        // buffered .xlsx  (Uint8Array ⇄ Workbook)
  readSheetRows,              // stream one sheet's rows, bounded memory
  readWorkbookStream,         // stream every sheet, rows one at a time
  readCsv, writeCsv,          // CSV as Uint8Array
  writeCsvText,               // CSV as a string
} from '@shbernal/ts-xlsx';

// Bounded-memory extraction — the whole workbook is never materialized:
for (const row of readSheetRows(bytes, {sheet: 'People'})) {
  console.log(row.number, row.cells.map((c) => c.value));
}
```

The reader decodes untrusted input defensively — entities are decoded but never
expanded, and inflation is bounded by a running output counter rather than any declared
size, so a malformed or hostile package can't exhaust memory.

## API reference

The full reference is generated from the public types — it cannot drift from what the
compiler accepts — and lives in **[`docs/api/`](docs/api/README.md)**. Regenerate it with:

```shell
pnpm run docs
```

## Coming from ExcelJS?

Read **[migrating from ExcelJS](docs/migrating-from-exceljs.md)**. The short version: the
shapes are deliberately different (synchronous `Uint8Array` I/O instead of async
`Buffer`/stream methods, `readXlsx`/`writeXlsx` free functions instead of
`workbook.xlsx.*`), because the goal was the *right* API, not the familiar one.

## Design principles

This project is optimized to be built and maintained largely by autonomous agents, with
a machine-checkable safety net as the primary guarantor of correctness. The rules that
govern every change are in [`CLAUDE.md`](CLAUDE.md); the design and working agreements are
in [`docs/architecture.md`](docs/architecture.md); notable decisions are recorded under
[`docs/decisions/`](docs/decisions/).

## License

[MIT](LICENSE) — inherited from ExcelJS and retained.
