# Reading a workbook

`readXlsx` takes bytes and gives you a `Workbook`. It is synchronous, and it accepts both
OOXML serialisations: an XML `.xlsx` and a binary `.xlsb`. The reader tells them apart by
looking inside the package, not at a file extension. You never branch on which one you hold,
and you get the same model either way.

```ts
import {readXlsx, Workbook, writeXlsx} from '@shbernal/ts-xlsx';

const original = new Workbook();
const built = original.addWorksheet('People');
built.addRow(['Name', 'Joined']);
built.addRow(['Ada', new Date('2026-01-01')]);

const workbook = readXlsx(writeXlsx(original));
const sheet = workbook.requireWorksheet('People');
console.log(sheet.getCell('A2').value); // 'Ada'
console.log(sheet.usedRange?.address); // 'A1:B2'
```

`getWorksheet` returns `undefined` for a name that is not there. `requireWorksheet` throws
instead. Use it when a missing sheet means the input is wrong, not a case to handle.

## Walking a sheet

```ts
import {readXlsx, Workbook, writeXlsx} from '@shbernal/ts-xlsx';

const source = new Workbook();
const built = source.addWorksheet('Data');
built.addRow(['a', 1]);
built.addRow(['b', 2]);
built.addRow(['c', 3]);

const sheet = readXlsx(writeXlsx(source)).requireWorksheet('Data');

for (const row of sheet.rows()) {
  console.log(row.number, row.values);
}
console.log(sheet.actualRowCount); // 3
```

`rows()` yields only the rows the sheet declares, so a sparse sheet costs what it holds,
not what its bounds span. `usedRange` is the rectangle those cells occupy, and it
is `undefined` for a sheet with nothing in it.

To read a rectangle rather than a sheet, ask for a `Range`:

```ts
import {readXlsx, Workbook, writeXlsx} from '@shbernal/ts-xlsx';

const source = new Workbook();
const built = source.addWorksheet('Grid');
built.addRow([1, 2, 3]);
built.addRow([4, 5, 6]);

const sheet = readXlsx(writeXlsx(source)).requireWorksheet('Grid');
const block = sheet.getRange('A1:C2');
console.log(block.cellCount); // 6
console.log(block.cells.map((cell) => cell.value)); // [1, 2, 3, 4, 5, 6]
```

## What a cell reads back as

The same union you wrote. A cell you never touched is `null`, not `undefined` and not an
empty string. When you want text, `cellValueToText` handles every kind in the union so you
don't have to switch on them yourself.

```ts
import {cellValueToText, readXlsx, Workbook, writeXlsx} from '@shbernal/ts-xlsx';

const source = new Workbook();
const built = source.addWorksheet('Kinds');
built.getCell('A1').value = 1234.5;
built.getCell('A2').value = {formula: 'A1*2', result: 2469};
built.getCell('A3').value = {error: '#DIV/0!'};

const sheet = readXlsx(writeXlsx(source)).requireWorksheet('Kinds');
console.log(sheet.getCell('A1').type); // 'number'
console.log(cellValueToText(sheet.getCell('A2').value)); // '2469'
console.log(sheet.getCell('B9').value); // null
```

For a formula, `cellValueToText` returns the cached _result_, not the formula source,
because the sheet never displayed the source. It applies no number format either. The format
lives on the style and this function only sees the value, so a currency cell comes back
without its currency sign. This library stores number formats but does not render them.
[Styles and number formats](./styles-and-formats.md) says what that means for you.

## What the model does not interpret

A real workbook contains parts this library does not model: a pivot cache, a slicer, a
chart, a linked-workbook reference. The reader keeps those bytes and the writer puts them
back untouched, so a load and save does not delete them. They show up as _preserved_
references, each listing the package parts it reaches.

```ts
import {readXlsx, Workbook, writeXlsx} from '@shbernal/ts-xlsx';

const source = new Workbook();
source.addWorksheet('Plain').getCell('A1').value = 'nothing exotic';

const workbook = readXlsx(writeXlsx(source));
// A workbook this library wrote has nothing it had to preserve.
console.log(workbook.preservedReferences.length); // 0
```

A preserved part is expected, not a bug. Keeping what it does not model means the library
can be safe to load and save before it models everything. The
[preserved-content reference](../api/preserved.md) lists what is preserved today, and
[ADR-0014](../decisions/0014-charts-shapes-slicers-are-round-trip-only-for-1-0.md) records
the decision.

## Reading untrusted input

Treat every file you did not write as hostile. A zip archive of XML is cheap to weaponise.
Two defences are always on.

The reader bounds inflation by counting the bytes it produces, never by trusting the
archive's declared sizes, so a zip bomb that lies about its uncompressed size still gets
refused. The ceiling defaults to 512 MiB and you can lower it:

```ts
import {PackageReadError, readXlsx, Workbook, writeXlsx} from '@shbernal/ts-xlsx';

const source = new Workbook();
const built = source.addWorksheet('Big');
for (let row = 1; row <= 200; row++) built.addRow([row, 'x'.repeat(200)]);
const bytes = writeXlsx(source);

try {
  readXlsx(bytes, {maxUncompressedBytes: 1024});
  console.log('no bound was enforced');
} catch (error) {
  console.log(error instanceof PackageReadError); // true
}
```

XML entities are decoded but never expanded, so the billion-laughs family of attacks has
nothing to expand into. Neither defence needs configuring, and neither can be switched off.

## Why there is no `readXlsxAsync`

Writing has an async form because DEFLATE dominates it and workers can take that off your
thread. Reading spends most of its time parsing XML and building the model, which no worker
can move. An async read would promise a non-blocking call and then block for most of it.
The reader also enforces the zip-bomb ceiling by counting output between synchronous input
slices, and that guarantee weakens once inflation moves to a worker.

If you need a read that does not block, run the whole read in a worker.
[ADR-0024](../decisions/0024-async-is-one-writer-not-a-mirrored-pair.md) has the argument in
full.

For a workbook too large to hold in memory at all, see [streaming](./streaming.md).

Next: [styles and number formats](./styles-and-formats.md), or [errors](./errors.md) for
what the failures above mean.
