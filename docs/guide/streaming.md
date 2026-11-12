# Streaming

The buffered path builds the whole workbook in memory. That is the right default, and it
stops being the right default at the point where the workbook is larger than the memory you
are willing to spend on it. Streaming trades the model for bounded memory: you see rows go
past, once, in order, and nothing accumulates.

Reach for it when the file is large and your job is a pass over the data. Stay with
`readXlsx` and `writeXlsx` when you need to look at a cell twice, when you need styles, or
when the file is small enough that none of this matters.

## Reading rows without building a workbook

```ts
import {readSheetRows, Workbook, writeXlsx} from '@shbernal/ts-xlsx';

const source = new Workbook();
const built = source.addWorksheet('People');
built.addRow(['Name', 'Age']);
built.addRow(['Ada', 36]);
built.addRow(['Grace', 45]);
const bytes = writeXlsx(source);

let total = 0;
for (const row of readSheetRows(bytes, {sheet: 'People'})) {
  const age = row.cells[1]?.value;
  if (typeof age === 'number') total += age;
}
console.log(total); // 81
```

`readSheetRows` is a generator, so nothing is read until you ask for the next row and
nothing is kept once you have moved past it. Two things about what it yields are worth
knowing, because both differ from the buffered model:

- only rows the sheet actually declares appear, so a gap in the numbering is a gap in the
  file;
- within a row, only the cells that carry something appear, because a blank or style-only
  cell is not data.

`sheet` takes a name or a 1-based position, and defaults to the first sheet.

## Reading every sheet

```ts
import {readWorkbookStream, Workbook, writeXlsx} from '@shbernal/ts-xlsx';

const source = new Workbook();
source.addWorksheet('First').addRow(['a']);
source.addWorksheet('Second').addRow(['b']);
const bytes = writeXlsx(source);

const names: string[] = [];
for (const sheet of readWorkbookStream(bytes)) {
  names.push(sheet.name);
}
console.log(names); // ['First', 'Second']
```

This is the streaming analogue of walking `readXlsx(bytes).worksheets`. Each yielded sheet
lets you stream its rows and read its merge and hidden-column summaries, without the model
behind it ever existing.

## Writing a workbook incrementally

The streaming writer is the mirror: add a sheet, append rows, and commit each one to write
it out and let it go. It comes from `@shbernal/ts-xlsx/node` rather than from the root
specifier, because it is the one part of the library that opens files and pipes Node streams;
[in the browser](./browser.md) says what that boundary buys everything else.

```ts
import {WorkbookStreamWriter} from '@shbernal/ts-xlsx/node';

const writer = new WorkbookStreamWriter();
const sheet = writer.addWorksheet('Big');
for (let n = 1; n <= 1000; n++) await sheet.addRow([n, n * n]).commit();
await sheet.commit();

const bytes = await writer.commit();
console.log(bytes !== undefined && bytes.length > 0); // true
```

`commit()` on a row writes it into the sheet's part in the archive and frees its cells. The
first committed row opens the part; each later one is deflated in batches of about 64 KiB and
handed to the destination. `commit()` on the sheet writes the rows you never committed and the
rest of the part, and closes it. `commit()` on the writer writes the remaining parts (styles,
the workbook, rels) and ends the destination.

So memory depends on the rows you have not committed yet, not on the sheet's length. In a test
with 45,000 rows of about 300 bytes, the heap after a collection was the same at row 5,000 and at
row 45,000; the 3.x writer kept about 30 MB more by the end. Three things still grow with what
you author rather than with the row count: notes, hyperlinks, merges, validations and
conditional formats (they are written after the rows); the style table; and the whole archive
when there is no sink to send it to.

The promise a row's `commit()` returns resolves once the sink can take more. Await it and the
loop runs at the sink's pace; skip the `await` and the file is still correct, but a slow sink
buffers the difference. A sink error never rejects that promise. It rejects the writer's
`commit()`.

Sheets share one archive, and an archive holds one open entry at a time. Commit each sheet
before you start the next: rows committed to a second sheet while the first is still open wait
in memory, compressed, until the first one closes.

### Rows are written in order

Writing as you go fixes the order of the file, and the writer refuses what it can no longer
put in the right place:

- Committing a row first commits every uncommitted row above it.
- `getCell` on a row at or above the last one written throws an `AuthoringError`. Fill a row
  before you commit it or any row below it.
- What comes before the rows in the file is written with the first committed row: column
  definitions, the frozen pane, sheet properties, and where outline summary rows sit. Set them
  through `sheet.model` before then. If one changes after that, the sheet's `commit()` throws
  rather than leave the change out of the file.
- The outline bar's depth (`outlineLevelRow`) is written with the first row too. It is the
  deepest level among the rows written then, or the `outlineLevelRow` you pass to
  `addWorksheet`, whichever is larger. A deeper row after that throws, so pass the option when
  your groups start further down.
- A summary row is marked `collapsed` when its detail rows are all hidden. With summary rows
  below their detail (Excel's default) the writer always knows. With summary rows above, it
  knows only if the detail rows go out in the same commit as the summary, so commit the group's
  last detail row rather than the summary row, or set `collapsed` on the row yourself.
- A streamed sheet has no `<dimension>` element. Excel and this library's reader compute the
  used range without it.

```ts
import {AuthoringError} from '@shbernal/ts-xlsx';
import {WorkbookStreamWriter} from '@shbernal/ts-xlsx/node';

const writer = new WorkbookStreamWriter();
const sheet = writer.addWorksheet('Report');
const header = sheet.addRow(['Region', 'Revenue']);
sheet.getCell('B1').value = 'Revenue (EUR)'; // not written yet, so still editable
await header.commit();

let refused = false;
try {
  sheet.getCell('B1').value = 'too late';
} catch (error) {
  refused = error instanceof AuthoringError;
}
console.log(refused); // true

await sheet.commit();
const written = await writer.commit();
console.log(written !== undefined && written.length > 0); // true
```

All of this applies with strings inline, the default. With `useSharedStrings: true` the rows
stay in memory until the writer commits, because the shared-strings table numbers strings across
the whole workbook, and the sheet is written the way the buffered writer writes it.

This half of the API is asynchronous where the buffered path is synchronous, and the bytes
arrive as the resolved value above, through `writer.stream`, a Node `Readable`, or into a sink
you hand the writer. The sink is the one that matters on a server, because it means a workbook
can go straight to an HTTP response without ever being one object in memory:

<!-- sample: illustrative. It needs a live HTTP response object, which a guide has none of. -->

```ts
// Inside a request handler:
const writer = new WorkbookStreamWriter({stream: response});
const sheet = writer.addWorksheet('Export');
for (const record of records) await sheet.addRow([record.id, record.total]).commit();
await sheet.commit();
await writer.commit();
```

`commit()` resolves with the package _only when something is going to want it_: when you
supplied no sink, or when you touched `writer.stream`. Hand the writer a `stream` or a
`filename` and never reach for `writer.stream`, as the handler above does not, and it resolves
with `undefined` and the archive is never assembled as a whole object. That is the point of
passing a sink, and the return type says so rather than handing back a copy of everything you
just streamed away. With a sink, take `writer.stream` before the first row is committed or not
at all: after that the sink already has bytes nobody kept, and the getter throws.

## What streaming costs you

The streaming API does less than the buffered one. Here is what it gives up.

- The streaming **writer is Node-only**, which is why it is published from `/node` and not
  from the root specifier: it writes to a Node stream. A browser build that imports it links
  a module that throws by name rather than one a bundler cannot resolve.
  [In the browser](./browser.md) says what runs in a tab.
- The streaming **reader does not read `.xlsb`**. A binary package cannot be row-streamed,
  and asking raises `UnsupportedFormatError` with `format` saying which kind it was, rather
  than a vague refusal.
- You do not get a `Worksheet`. `getCell` is there, but only below the last row written, and
  the whole-sheet reads (`usedRange`, `getRange`, `hasCell`, `actualRowCount`) and the
  structural edits (`spliceRows`, `insertRow`, the column verbs) are absent altogether: a
  streamed sheet could only answer them over whichever rows happened to still be in memory,
  which is a plausible answer rather than a true one.

If you are streaming because writing is slow rather than because it is large, look at
`writeXlsxAsync` in [writing a workbook](./writing.md) first: it keeps the whole buffered
API and just moves compression off your thread.

Next: [errors](./errors.md), which is where `UnsupportedFormatError` above is explained.
