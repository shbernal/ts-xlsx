# Styles and number formats

A cell's appearance is six independent facets: `font`, `fill`, `border`, `alignment`,
`numFmt` and `protection`. Each is set on its own, each is absent until you set it, and an
absent facet is genuinely absent rather than an empty placeholder.

```ts
import {Workbook} from '@shbernal/ts-xlsx';

const sheet = new Workbook().addWorksheet('Report');
const cell = sheet.getCell('A1');

cell.value = 'Total';
cell.font = {bold: true, size: 12, color: {argb: 'FFFFFFFF'}};
cell.fill = {type: 'pattern', pattern: 'solid', fgColor: {argb: 'FF146C43'}};
cell.alignment = {horizontal: 'center', vertical: 'center'};

console.log(cell.font?.bold); // true
console.log(cell.border); // undefined
```

Colour is an ARGB hex string, uppercase, alpha first, no leading `#`. `FF` is opaque, so
opaque white is `FFFFFFFF` and opaque black is `FF000000`.

## Styling more than one cell

Setting a facet on a `Range` sets it on every cell in the range, and setting one on a
`Column` or `Row` sets the default for that line. They are separate mechanisms: a column
style applies to cells the column does not otherwise have, while a range style is written
onto each cell.

```ts
import {Workbook} from '@shbernal/ts-xlsx';

const sheet = new Workbook().addWorksheet('Report');
sheet.addRow(['Region', 'Revenue']);
sheet.addRow(['North', 84200]);
sheet.addRow(['South', 61050]);

sheet.getRange('A1:B1').font = {bold: true};
sheet.getColumn(2).numFmt = '#,##0';

console.log(sheet.getCell('B1').font?.bold); // true
console.log(sheet.getColumn(2).numFmt); // '#,##0'
```

Facets do not leak between cells. Setting a border on `B2` leaves `B3` alone, even though
the two may share a style record underneath; the sharing is a storage detail the model does
not let you observe.

## Number formats

A number format is a string, and the library carries it verbatim. It is not parsed, not
normalised, and not validated, which is the only behaviour that lets an author use a format
this library has never heard of.

```ts
import {readXlsx, Workbook, writeXlsx} from '@shbernal/ts-xlsx';

const accounting = '_("$"* #,##0.00_);_("$"* (#,##0.00);_("$"* "-"??_);_(@_)';

const source = new Workbook();
const built = source.addWorksheet('Money');
built.getCell('A1').value = 1234.5;
built.getCell('A1').numFmt = accounting;
built.getCell('A2').value = new Date('2026-03-04');
built.getCell('A2').numFmt = 'yyyy-mm-dd';
built.getCell('A3').value = 0.62;
built.getCell('A3').numFmt = '0.0%';

const sheet = readXlsx(writeXlsx(source)).requireWorksheet('Money');
console.log(sheet.getCell('A1').numFmt === accounting); // true
```

That round trip is not a coincidence. It is a regression case
(`custom-numfmt-string-roundtrips-verbatim`), because a format code silently losing its group
separators renders every cell in a column wrong and nothing throws.

**The library does not apply formats.** It stores them, round-trips them, and hands them to
whatever opens the file. `cell.text` and `cellValueToText` give you the value's text with no
format applied, so a currency cell has no currency sign and `0.1 + 0.2` reads as
`0.30000000000000004`. If you need the string Excel would paint, you need a formatting
engine, and this is not one. The [playground](https://shbernal.github.io/ts-xlsx/playground)
shows the same distinction on screen.

## Rich text, when one cell holds two fonts

```ts
import {readXlsx, richTextToPlain, Workbook, writeXlsx} from '@shbernal/ts-xlsx';

const source = new Workbook();
source.addWorksheet('Note').getCell('A1').value = {
  richText: [
    {text: 'Warning: ', font: {bold: true, color: {argb: 'FF8A5300'}}},
    {text: 'the total is provisional.'},
  ],
};

const sheet = readXlsx(writeXlsx(source)).requireWorksheet('Note');
const value = sheet.getCell('A1').value;
console.log(typeof value === 'object' && value !== null && 'richText' in value); // true
console.log(richTextToPlain({richText: [{text: 'a'}, {text: 'b'}]})); // 'ab'
```

`richTextToPlain` is what a consumer that cannot render per-run formatting sees, which is
also what a CSV field or a pivot cache entry gets.

## Themes and the default font

A workbook carries a theme: a colour scheme and a font scheme that styles refer to by role
rather than by value. Reading a file keeps the theme it came with, and `resolveColor` turns
a theme reference into the ARGB it currently means.

```ts
import {Workbook} from '@shbernal/ts-xlsx';

const workbook = new Workbook();
console.log(workbook.defaultFont.name !== undefined); // true
console.log(workbook.themeColors.accent1 !== undefined); // true
```

The default font is declared rather than assumed. That matters more than it sounds: column
width in OOXML is measured in characters of the workbook's default font, so a library that
guessed the font would compute the wrong widths for any workbook that set one.
[ADR-0025](../decisions/0025-the-default-font-is-declared-not-assumed.md) has the case.

Next: [structure](./tables-and-structure.md) for the geometry around the values.
