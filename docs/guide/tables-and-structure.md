# Structure: merges, widths, panes and tables

This page covers the layout around the values: column widths, merged cells, frozen panes,
and tables that are real table objects in the file.

## Columns and rows

`getColumn` and `getRow` return handles. Reading one creates nothing, so asking about
column 40 of a two-column sheet costs nothing and does not widen the sheet.

```ts
import {Workbook} from '@shbernal/ts-xlsx';

const sheet = new Workbook().addWorksheet('Sheet1');
sheet.addRow(['Region', 'Revenue']);

sheet.getColumn(1).width = 18;
sheet.getColumn(2).width = 14;
sheet.getRow(1).height = 22;
sheet.getRow(1).hidden = false;

console.log(sheet.getColumn(1).letter); // 'A'
console.log(sheet.getColumn(40).width); // undefined
console.log(sheet.columnCount); // 2
```

Width is in characters of the workbook's default font, the unit OOXML uses, not pixels. A
column with no width set reads `undefined`, not a made-up number, because "whatever the
consumer defaults to" and "18 characters" are different facts.

## Merges

```ts
import {readXlsx, Workbook, writeXlsx} from '@shbernal/ts-xlsx';

const source = new Workbook();
const built = source.addWorksheet('Report');
built.getCell('A1').value = 'Regional summary';
built.mergeCells('A1:C1');
built.getCell('A1').alignment = {horizontal: 'center'};

const sheet = readXlsx(writeXlsx(source)).requireWorksheet('Report');
console.log(sheet.merges); // ['A1:C1']
console.log(sheet.getCell('A1').value); // 'Regional summary'
```

The value lives on the top-left cell; the cells a merge covers hold nothing. `unmergeCells`
undoes one and tells you whether there was a merge there to undo. Merges may not overlap.
An overlapping one throws an `AuthoringError` now instead of producing a file that opens wrong.

## Freezing headers

```ts
import {readXlsx, Workbook, writeXlsx} from '@shbernal/ts-xlsx';

const source = new Workbook();
const built = source.addWorksheet('Ledger');
built.addRow(['Date', 'Amount']);
built.addRow([new Date('2026-03-04'), 12]);
// Two rows and one column stay put while the rest scrolls.
built.freeze(2, 1);

const sheet = readXlsx(writeXlsx(source)).requireWorksheet('Ledger');
console.log(sheet.view.ySplit); // 2
console.log(sheet.view.xSplit); // 1
console.log(sheet.view.state); // 'frozen'
```

`freeze()` with no arguments freezes the first row, which is the case you almost always
want. `unfreeze()` clears it.

## Tables

A table is an object in the file with a name, named columns, a header row, an optional
totals row, and a style from Excel's gallery. Formulas elsewhere
can refer to it structurally, as `Regions[Revenue]`.

```ts
import {readXlsx, Workbook, writeXlsx} from '@shbernal/ts-xlsx';

const source = new Workbook();
const built = source.addWorksheet('Regions');
built.getCell('A1').value = 'Region';
built.getCell('B1').value = 'Revenue';
built.addRow(['North', 84200]);
built.addRow(['South', 61050]);

built.addTable({
  name: 'Regions',
  ref: 'A1',
  columns: [{name: 'Region'}, {name: 'Revenue'}],
  rowCount: 2,
  style: {name: 'TableStyleMedium2', showRowStripes: true},
});

const sheet = readXlsx(writeXlsx(source)).requireWorksheet('Regions');
const table = sheet.getTable('Regions');
console.log(table?.range); // 'A1:B3'
console.log(table?.columnCount); // 2
```

`ref` is the table's top-left corner and `rowCount` is its number of data rows. The library
derives the range from those two. It does not validate the style name, on purpose: a file
from a newer Excel can name a style this library has never seen, and rejecting it would make
that file impossible to round-trip.

## Autofilter and data validation

```ts
import {readXlsx, Workbook, writeXlsx} from '@shbernal/ts-xlsx';

const source = new Workbook();
const built = source.addWorksheet('Filtered');
built.addRow(['Region', 'Revenue']);
built.addRow(['North', 84200]);
built.autoFilter = 'A1:B2';
built.addDataValidation('B2', {
  type: 'decimal',
  operator: 'greaterThan',
  formulae: ['0'],
  allowBlank: false,
});

const sheet = readXlsx(writeXlsx(source)).requireWorksheet('Filtered');
console.log(sheet.autoFilter?.ref); // 'A1:B2'
console.log(sheet.dataValidationAt('B2')?.type); // 'decimal'
```

## Inserting and removing rows

`spliceRows` is the array method's shape, applied to a sheet: a start, a count to remove, and
rows to insert. Everything below moves, and so do the things anchored to it.

```ts
import {Workbook} from '@shbernal/ts-xlsx';

const sheet = new Workbook().addWorksheet('Sheet1');
sheet.addRows([['a'], ['b'], ['c']]);

sheet.spliceRows(2, 1); // drop row 2
console.log(sheet.getCell('A2').value); // 'c'

sheet.insertRow(2, ['b again']);
console.log(sheet.getCell('A2').value); // 'b again'
```

A handle's position is fixed, exactly as a cell's is: after a splice, `getRow(3)` still means
row 3, now holding whatever moved there.

## Page setup, for a sheet that will be printed

```ts
import {readXlsx, Workbook, writeXlsx} from '@shbernal/ts-xlsx';

const source = new Workbook();
const built = source.addWorksheet('Print');
built.addRow(['Region', 'Revenue']);
built.pageSetup.orientation = 'landscape';
built.pageSetup.fitToWidth = 1;
built.printOptions.horizontalCentered = true;
built.headerFooter.oddFooter = '&CPage &P of &N';

const sheet = readXlsx(writeXlsx(source)).requireWorksheet('Print');
console.log(sheet.pageSetup.orientation); // 'landscape'
console.log(sheet.headerFooter.oddFooter); // '&CPage &P of &N'
```

Next: [streaming](./streaming.md), when the sheet is bigger than the memory you want to
spend on it.
