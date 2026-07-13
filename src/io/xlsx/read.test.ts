import assert from 'node:assert/strict';
import {test} from 'node:test';

import {zipSync, strToU8, strFromU8, unzipSync} from 'fflate';

import {Workbook} from '../../core/workbook.ts';
import {isFormulaValue} from '../../core/value.ts';
import {readXlsx} from './read.ts';
import {writeXlsx} from './write.ts';

/** Write a workbook and read it straight back — the round-trip under test. */
function roundtrip(workbook: Workbook): Workbook {
  return readXlsx(writeXlsx(workbook));
}

test('scalar cell values survive the round-trip with their types', () => {
  const wb = new Workbook();
  const sheet = wb.addWorksheet('S');
  sheet.getCell('A1').value = 42;
  sheet.getCell('A2').value = 'hello';
  sheet.getCell('A3').value = true;
  sheet.getCell('A4').value = false;

  const back = roundtrip(wb).getWorksheet('S');
  assert.equal(back?.getCell('A1').value, 42);
  assert.equal(back?.getCell('A2').value, 'hello');
  assert.equal(back?.getCell('A3').value, true);
  assert.equal(back?.getCell('A4').value, false);
});

test('a string with markup-significant and leading/trailing space round-trips exactly', () => {
  const wb = new Workbook();
  wb.addWorksheet('S').getCell('A1').value = '  <a> & "b" \t end  ';
  const back = roundtrip(wb).getWorksheet('S');
  assert.equal(back?.getCell('A1').value, '  <a> & "b" \t end  ');
});

test('a formula with a numeric result round-trips as {formula, result}', () => {
  const wb = new Workbook();
  wb.addWorksheet('S').getCell('A1').value = {formula: 'SUM(B1:B2)', result: 7};
  const value = roundtrip(wb).getWorksheet('S')?.getCell('A1').value;
  assert.ok(value && isFormulaValue(value));
  assert.equal(value.formula, 'SUM(B1:B2)');
  assert.equal(value.result, 7);
});

test('a formula with a string result carries its t="str" result back', () => {
  const wb = new Workbook();
  wb.addWorksheet('S').getCell('A1').value = {formula: 'CONCAT("a","b")', result: 'ab'};
  const value = roundtrip(wb).getWorksheet('S')?.getCell('A1').value;
  assert.ok(value && isFormulaValue(value));
  assert.equal(value.result, 'ab');
});

test('a formula with no cached result round-trips without inventing one', () => {
  const wb = new Workbook();
  wb.addWorksheet('S').getCell('A1').value = {formula: 'NOW()'};
  const value = roundtrip(wb).getWorksheet('S')?.getCell('A1').value;
  assert.ok(value && isFormulaValue(value));
  assert.equal(value.result, undefined);
});

test('a modern function is stored with _xlfn. on disk but round-trips as its plain name', () => {
  const wb = new Workbook();
  wb.addWorksheet('S').getCell('A1').value = {formula: 'FILTER(B1:D1,B2:D2=1)', result: 0};
  const xml = strFromU8(unzipSync(writeXlsx(wb))['xl/worksheets/sheet1.xml'] as Uint8Array);
  assert.match(xml, /<f>_xlfn\.FILTER\(B1:D1,B2:D2=1\)<\/f>/);

  const value = roundtrip(wb).getWorksheet('S')?.getCell('A1').value;
  assert.ok(value && isFormulaValue(value));
  assert.equal(value.formula, 'FILTER(B1:D1,B2:D2=1)');
});

test('a formula that already carries _xlfn. is not double-prefixed on write', () => {
  const wb = new Workbook();
  wb.addWorksheet('S').getCell('A1').value = {formula: '_xlfn.XLOOKUP(1,B:B,C:C)', result: 0};
  const xml = strFromU8(unzipSync(writeXlsx(wb))['xl/worksheets/sheet1.xml'] as Uint8Array);
  assert.match(xml, /_xlfn\.XLOOKUP/);
  assert.doesNotMatch(xml, /_xlfn\._xlfn/);
});

test('a dotted statistical function is stored whole with _xlfn. and round-trips as its plain name', () => {
  const wb = new Workbook();
  wb.addWorksheet('S').getCell('A1').value = {formula: 'NORM.DIST(A2,0,1,TRUE)', result: 0.5};
  const xml = strFromU8(unzipSync(writeXlsx(wb))['xl/worksheets/sheet1.xml'] as Uint8Array);
  assert.match(xml, /<f>_xlfn\.NORM\.DIST\(A2,0,1,TRUE\)<\/f>/);
  assert.doesNotMatch(xml, /_xlfn\.DIST/);

  const value = roundtrip(wb).getWorksheet('S')?.getCell('A1').value;
  assert.ok(value && isFormulaValue(value));
  assert.equal(value.formula, 'NORM.DIST(A2,0,1,TRUE)');
});

test('column width and visibility round-trip', () => {
  const wb = new Workbook();
  const sheet = wb.addWorksheet('S');
  sheet.getColumn(2).width = 24;
  sheet.getColumn(4).hidden = true;
  const back = roundtrip(wb).getWorksheet('S');
  assert.equal(back?.getColumn(2).width, 24);
  assert.equal(back?.getColumn(4).hidden, true);
});

test('column outline grouping round-trips', () => {
  const wb = new Workbook();
  const sheet = wb.addWorksheet('S');
  sheet.getColumn(2).outlineLevel = 2;
  sheet.getColumn(3).collapsed = true;
  const back = roundtrip(wb).getWorksheet('S');
  assert.equal(back?.getColumn(2).outlineLevel, 2);
  assert.equal(back?.getColumn(3).collapsed, true);
});

test('row height and visibility round-trip', () => {
  const wb = new Workbook();
  const sheet = wb.addWorksheet('S');
  sheet.getCell('A1').value = 'x';
  sheet.getRow(1).height = 33;
  sheet.getRow(2).hidden = true;
  const back = roundtrip(wb).getWorksheet('S');
  assert.equal(back?.getRow(1).height, 33);
  assert.equal(back?.getRow(2).hidden, true);
});

test('merged ranges round-trip', () => {
  const wb = new Workbook();
  const sheet = wb.addWorksheet('S');
  sheet.getCell('A1').value = 'x';
  sheet.mergeCells('A1:B2');
  const back = roundtrip(wb).getWorksheet('S');
  assert.deepEqual([...(back?.merges ?? [])], ['A1:B2']);
});

test('page margins round-trip', () => {
  const wb = new Workbook();
  const sheet = wb.addWorksheet('S');
  sheet.getCell('A1').value = 'x';
  sheet.pageMargins.left = 0.5;
  sheet.pageMargins.top = 1.25;
  const back = roundtrip(wb).getWorksheet('S');
  assert.equal(back?.pageMargins.left, 0.5);
  assert.equal(back?.pageMargins.top, 1.25);
});

test('workbook core properties round-trip', () => {
  const wb = new Workbook();
  wb.properties.creator = 'Ada Lovelace';
  wb.properties.lastModifiedBy = 'Grace Hopper';
  wb.properties.created = new Date('2020-01-02T03:04:05.000Z');
  wb.properties.modified = new Date('2021-06-07T08:09:10.000Z');
  wb.addWorksheet('S').getCell('A1').value = 1;

  const back = roundtrip(wb);
  assert.equal(back.properties.creator, 'Ada Lovelace');
  assert.equal(back.properties.lastModifiedBy, 'Grace Hopper');
  assert.equal(back.properties.created?.toISOString(), '2020-01-02T03:04:05.000Z');
  assert.equal(back.properties.modified?.toISOString(), '2021-06-07T08:09:10.000Z');
});

test('multiple sheets round-trip in order and are addressable by name', () => {
  const wb = new Workbook();
  for (const name of ['First', 'Second', 'Third']) wb.addWorksheet(name).getCell('A1').value = name;
  const back = roundtrip(wb);
  assert.deepEqual(
    back.worksheets.map(s => s.name),
    ['First', 'Second', 'Third']
  );
  assert.equal(back.getWorksheet('Second')?.getCell('A1').value, 'Second');
});

test('rowCount spans a gap; actualRowCount counts only populated rows', () => {
  const wb = new Workbook();
  const sheet = wb.addWorksheet('S');
  sheet.getCell('A1').value = 'a';
  sheet.getCell('A3').value = 'c';
  const back = roundtrip(wb).getWorksheet('S');
  assert.equal(back?.rowCount, 3);
  assert.equal(back?.actualRowCount, 2);
});

test('a solid pattern fill round-trips with its foreground colour on a single cell', () => {
  const wb = new Workbook();
  const sheet = wb.addWorksheet('S');
  sheet.getCell('A1').value = 'painted';
  sheet.getCell('A1').fill = {type: 'pattern', pattern: 'solid', fgColor: {argb: 'FFFF0000'}};
  sheet.getCell('B2').value = 'plain';

  const back = roundtrip(wb).getWorksheet('S');
  const fill = back?.getCell('A1').fill;
  assert.equal(fill?.type, 'pattern');
  assert.equal(fill?.pattern, 'solid');
  assert.equal(fill?.fgColor?.argb, 'FFFF0000');
  assert.equal(fill?.bgColor?.indexed, 64, 'solid fill keeps the automatic indexed background');
  assert.equal(back?.getCell('B2').fill, undefined, 'an unfilled cell reads back with no fill');
});

test('a formatted-but-empty cell round-trips with its fill and a null value', () => {
  const wb = new Workbook();
  const sheet = wb.addWorksheet('S');
  sheet.getCell('A1').value = 'anchor';
  // B2 is styled but never given a value — the fill must survive without a value being invented.
  sheet.getCell('B2').fill = {type: 'pattern', pattern: 'solid', fgColor: {argb: 'FF00FF00'}};

  const back = roundtrip(wb).getWorksheet('S');
  assert.equal(back?.getCell('B2').value, null, 'the empty cell stays empty, not fabricated a value');
  assert.equal(back?.getCell('B2').fill?.fgColor?.argb, 'FF00FF00', 'the fill survives on the empty cell');
});

test('two cells with different fills stay distinct across the round-trip', () => {
  const wb = new Workbook();
  const sheet = wb.addWorksheet('S');
  sheet.getCell('A1').value = 1;
  sheet.getCell('A1').fill = {type: 'pattern', pattern: 'solid', fgColor: {argb: 'FFFF0000'}};
  sheet.getCell('A2').value = 2;
  sheet.getCell('A2').fill = {type: 'pattern', pattern: 'solid', fgColor: {argb: 'FF0000FF'}};

  const back = roundtrip(wb).getWorksheet('S');
  assert.equal(back?.getCell('A1').fill?.fgColor?.argb, 'FFFF0000');
  assert.equal(back?.getCell('A2').fill?.fgColor?.argb, 'FF0000FF');
});

test('a row-level fill is inherited by the row cells that carry no fill of their own', () => {
  const wb = new Workbook();
  const sheet = wb.addWorksheet('S');
  for (let r = 1; r <= 4; r++) sheet.getCell(`A${r}`).value = `r${r}`;
  sheet.getRow(3).fill = {type: 'pattern', pattern: 'solid', fgColor: {argb: 'FFFF4500'}};

  const back = roundtrip(wb).getWorksheet('S');
  assert.equal(back?.getCell('A3').fill?.fgColor?.argb, 'FFFF4500', 'the formatted row paints its cell');
  for (const ref of ['A1', 'A2', 'A4']) {
    assert.equal(back?.getCell(ref).fill, undefined, `${ref} does not inherit the row-3 fill`);
  }
});

test('many identically-filled cells collapse to one shared style entry in the package', () => {
  const wb = new Workbook();
  const sheet = wb.addWorksheet('S');
  for (let r = 1; r <= 40; r++) {
    sheet.getCell(`A${r}`).value = r;
    sheet.getCell(`A${r}`).fill = {type: 'pattern', pattern: 'solid', fgColor: {argb: 'FFDDEEFF'}};
  }
  const files = unzipSync(writeXlsx(wb));
  const stylesXml = strFromU8(files['xl/styles.xml'] as Uint8Array);
  // Default xf + the single shared fill = two entries, never ~40.
  assert.match(stylesXml, /<cellXfs count="2">/);
});

test('a custom cell number format round-trips byte-for-byte', () => {
  const wb = new Workbook();
  const sheet = wb.addWorksheet('S');
  const fmt = '_("$"* #,##0.00_);_("$"* (#,##0.00);_("$"* "-"??_);_(@_)';
  sheet.getCell('A1').value = 1234.5;
  sheet.getCell('A1').numFmt = fmt;
  sheet.getCell('B1').value = 'plain';

  const back = roundtrip(wb).getWorksheet('S');
  assert.equal(back?.getCell('A1').numFmt, fmt, 'the custom format survives character-for-character');
  assert.equal(back?.getCell('B1').numFmt, undefined, 'an unformatted cell reads back with no numFmt');
});

test('a column number format round-trips and is inherited by the column cells that carry no format', () => {
  const wb = new Workbook();
  const sheet = wb.addWorksheet('S');
  sheet.getColumn(1).numFmt = '0.00%';
  sheet.getColumn(2).numFmt = '"$"#,##0.00';
  sheet.getCell('A1').value = 0.1;
  sheet.getCell('B1').value = 3;

  const back = roundtrip(wb).getWorksheet('S');
  assert.equal(back?.getColumn(1).numFmt, '0.00%', 'column 1 keeps its own format');
  assert.equal(back?.getColumn(2).numFmt, '"$"#,##0.00', 'column 2 keeps its own format, not the last-assigned one');
  assert.equal(back?.getCell('A1').numFmt, '0.00%', 'a bare cell inherits its column format');
  assert.equal(back?.getCell('B1').numFmt, '"$"#,##0.00', 'a bare cell inherits its own column, not a sibling’s');
});

test('a cell fill and its column number format both survive — overriding one facet keeps the other', () => {
  const wb = new Workbook();
  const sheet = wb.addWorksheet('S');
  sheet.getColumn(1).numFmt = '0.00';
  sheet.getCell('A1').value = 1;
  sheet.getCell('A2').value = 2;
  sheet.getCell('A2').fill = {type: 'pattern', pattern: 'solid', fgColor: {argb: 'FFFFFF00'}};
  sheet.getCell('A3').value = 3;

  const back = roundtrip(wb).getWorksheet('S');
  assert.equal(back?.getCell('A2').fill?.fgColor?.argb, 'FFFFFF00', 'the filled cell keeps its fill');
  for (const ref of ['A1', 'A2', 'A3']) {
    assert.equal(back?.getCell(ref).numFmt, '0.00', `${ref} keeps the column number format`);
  }
  assert.equal(back?.getCell('A1').fill, undefined, 'a sibling does not pick up the fill');
  assert.equal(back?.getCell('A3').fill, undefined, 'a sibling does not pick up the fill');
});

test('a column alignment round-trips and is inherited by its own cells only', () => {
  const wb = new Workbook();
  const sheet = wb.addWorksheet('S');
  sheet.getColumn(1).alignment = {textRotation: 45};
  sheet.getCell('A1').value = 'a';
  sheet.getCell('A2').value = 'c';
  sheet.getCell('B1').value = 'b';

  const back = roundtrip(wb).getWorksheet('S');
  assert.equal(back?.getColumn(1).alignment?.textRotation, 45, 'the column keeps its own alignment');
  assert.equal(back?.getCell('A1').alignment?.textRotation, 45, 'a bare cell inherits its column alignment');
  assert.equal(back?.getCell('A2').alignment?.textRotation, 45, 'a second cell in the column inherits it too');
  assert.equal(back?.getCell('B1').alignment, undefined, 'a cell in another column does not inherit it');
});

test('a column border applies only to its declaring column, not to later width-only columns', () => {
  const wb = new Workbook();
  const sheet = wb.addWorksheet('S');
  sheet.getColumn(1).border = {right: {style: 'thin', color: {argb: 'FF000000'}}};
  sheet.getColumn(2).width = 10;
  sheet.getCell('A1').value = 'a';
  sheet.getCell('B1').value = 'b';
  sheet.getCell('C1').value = 'c';

  const back = roundtrip(wb).getWorksheet('S');
  assert.equal(back?.getCell('A1').border?.right?.style, 'thin', 'the declaring column’s cell has the border');
  assert.equal(back?.getCell('B1').border, undefined, 'a width-only column’s cell gets no border');
  assert.equal(back?.getCell('C1').border, undefined, 'an undeclared column’s cell gets no border');
});

test('a cell overriding one facet keeps the column’s other facet default', () => {
  // The column defaults both an alignment and a number format; a cell that overrides only its
  // alignment must still carry the column's number format — the generalized column composition
  // must not drop a non-overridden facet.
  const wb = new Workbook();
  const sheet = wb.addWorksheet('S');
  sheet.getColumn(1).alignment = {horizontal: 'center'};
  sheet.getColumn(1).numFmt = '0.00';
  sheet.getCell('A1').value = 1;
  sheet.getCell('A2').value = 2;
  sheet.getCell('A2').alignment = {horizontal: 'right'};

  const back = roundtrip(wb).getWorksheet('S');
  assert.equal(back?.getCell('A1').alignment?.horizontal, 'center', 'a bare cell inherits the column alignment');
  assert.equal(back?.getCell('A2').alignment?.horizontal, 'right', 'the overriding cell keeps its own alignment');
  assert.equal(back?.getCell('A2').numFmt, '0.00', 'the overriding cell still carries the column number format');
});

test('a built-in numFmt id on a foreign cell resolves to its standard format code', () => {
  // A foreign generator names a built-in format by id with no <numFmt> entry; the reader
  // resolves it from the standard table (id 10 = "0.00%").
  const files: Record<string, Uint8Array> = {
    '[Content_Types].xml': strToU8(
      '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
        '<Default Extension="xml" ContentType="application/xml"/></Types>'
    ),
    'xl/workbook.xml': strToU8(
      '<?xml version="1.0"?><workbook xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
        '<sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets></workbook>'
    ),
    'xl/_rels/workbook.xml.rels': strToU8(
      '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        '<Relationship Id="rId1" Type="x" Target="worksheets/sheet1.xml"/></Relationships>'
    ),
    'xl/styles.xml': strToU8(
      '<?xml version="1.0"?><styleSheet><fills count="1"><fill><patternFill patternType="none"/></fill></fills>' +
        '<cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>' +
        '<xf numFmtId="10" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/></cellXfs></styleSheet>'
    ),
    'xl/worksheets/sheet1.xml': strToU8(
      '<?xml version="1.0"?><worksheet><sheetData><row r="1"><c r="A1" s="1"><v>0.5</v></c></row></sheetData></worksheet>'
    ),
  };
  const back = readXlsx(zipSync(files)).getWorksheet('S');
  assert.equal(back?.getCell('A1').numFmt, '0.00%');
});

test('a Strict-mode t="d" cell parses to the ISO date it states, not a 1900 serial', () => {
  const files: Record<string, Uint8Array> = {
    'xl/workbook.xml': strToU8(
      '<?xml version="1.0"?><workbook xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
        '<sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets></workbook>'
    ),
    'xl/_rels/workbook.xml.rels': strToU8(
      '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        '<Relationship Id="rId1" Type="x" Target="worksheets/sheet1.xml"/></Relationships>'
    ),
    'xl/worksheets/sheet1.xml': strToU8(
      '<?xml version="1.0"?><worksheet><sheetData><row r="1"><c r="A1" t="d"><v>2024-02-09</v></c></row></sheetData></worksheet>'
    ),
  };
  const cell = readXlsx(zipSync(files)).getWorksheet('S')?.getCell('A1');
  assert.equal(cell?.type, 'date');
  assert.equal((cell?.value as Date).toISOString(), '2024-02-09T00:00:00.000Z');
});

test('a serial under a built-in locale date id (57) reads as a date, not a bare number', () => {
  const files: Record<string, Uint8Array> = {
    'xl/workbook.xml': strToU8(
      '<?xml version="1.0"?><workbook xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
        '<sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets></workbook>'
    ),
    'xl/_rels/workbook.xml.rels': strToU8(
      '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        '<Relationship Id="rId1" Type="x" Target="worksheets/sheet1.xml"/></Relationships>'
    ),
    // Built-in id 57 carries no <numFmt> entry — it is resolved from the built-in table.
    'xl/styles.xml': strToU8(
      '<?xml version="1.0"?><styleSheet><cellXfs count="2"><xf numFmtId="0"/>' +
        '<xf numFmtId="57" applyNumberFormat="1"/></cellXfs></styleSheet>'
    ),
    'xl/worksheets/sheet1.xml': strToU8(
      '<?xml version="1.0"?><worksheet><sheetData><row r="1"><c r="A1" s="1"><v>45809</v></c></row></sheetData></worksheet>'
    ),
  };
  const cell = readXlsx(zipSync(files)).getWorksheet('S')?.getCell('A1');
  assert.equal(cell?.type, 'date', 'a serial under a built-in date id reads as a date');
  assert.ok(cell?.numFmt, 'the built-in id resolves to a non-empty format code');
});

test('a cell font round-trips through the <fonts> table, and only the styled cell carries it', () => {
  const wb = new Workbook();
  const sheet = wb.addWorksheet('S');
  sheet.getCell('A1').value = 'heading';
  sheet.getCell('A1').font = {bold: true, italic: true, size: 14, name: 'Arial', color: {argb: 'FF3A80D5'}};
  sheet.getCell('B2').value = 'plain';

  const back = roundtrip(wb).getWorksheet('S');
  const font = back?.getCell('A1').font;
  assert.deepEqual(font, {bold: true, italic: true, size: 14, color: {argb: 'FF3A80D5'}, name: 'Arial'});
  // The unstyled sibling renders in the workbook default font, not A1's — so it must not pick up
  // any of A1's overrides (bold/italic/Arial/the blue), though it does resolve the default face.
  const b2 = back?.getCell('B2').font;
  assert.equal(b2?.bold, undefined, 'an unstyled sibling does not inherit the bold');
  assert.equal(b2?.italic, undefined, 'an unstyled sibling does not inherit the italic');
  assert.notEqual(b2?.name, 'Arial', 'an unstyled sibling keeps the default face, not A1 Arial');
});

test('an unstyled cell resolves to the workbook default font (a concrete face), not nothing', () => {
  const wb = new Workbook();
  wb.addWorksheet('S').getCell('A1').value = 'plain';

  const font = roundtrip(wb).getWorksheet('S')?.getCell('A1').font;
  assert.ok(font, 'an unstyled cell must resolve the default font, not undefined');
  assert.equal(font?.name, 'Calibri', 'the default face is Calibri');
  assert.equal(font?.size, 11, 'the default size is 11');
});

test('a cell carrying exactly the default font interns back to font id 0 — no redundant entry', () => {
  const wb = new Workbook();
  const sheet = wb.addWorksheet('S');
  sheet.getCell('A1').value = 'x';
  // Assign the default face explicitly; on write it must collapse to id 0, so the emitted <fonts>
  // table stays a single default entry rather than gaining a duplicate.
  sheet.getCell('A1').font = {size: 11, color: {theme: 1}, name: 'Calibri', family: 2, scheme: 'minor'};

  const styles = strFromU8(unzipSync(writeXlsx(wb))['xl/styles.xml'] ?? new Uint8Array());
  const fontsBlock = styles.match(/<fonts\b[^>]*>[\s\S]*?<\/fonts>/)?.[0] ?? '';
  assert.equal((fontsBlock.match(/<font\b/g) ?? []).length, 1, 'only the single default font entry');
});

test('an underline font round-trips: single stays single, a named variant keeps its value', () => {
  const wb = new Workbook();
  const sheet = wb.addWorksheet('S');
  sheet.getCell('A1').value = 'x';
  sheet.getCell('A1').font = {underline: true};
  sheet.getCell('A2').value = 'y';
  sheet.getCell('A2').font = {underline: 'double'};

  const back = roundtrip(wb).getWorksheet('S');
  assert.equal(back?.getCell('A1').font?.underline, true, 'a bare underline reads back as true');
  assert.equal(back?.getCell('A2').font?.underline, 'double', 'a named underline keeps its variant');
});

test('a foreign font’s <u val="none"/> reads back as not underlined, not the truthy string "none"', () => {
  // Real producers write <u val="none"/> for the explicit ABSENCE of an underline. Surfacing the
  // literal "none" would be truthy — a consumer’s `if (font.underline)` would mistake it for an
  // underline — so the reader must read it back falsy.
  const files: Record<string, Uint8Array> = {
    'xl/workbook.xml': strToU8(
      '<?xml version="1.0"?><workbook xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
        '<sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets></workbook>'
    ),
    'xl/_rels/workbook.xml.rels': strToU8(
      '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        '<Relationship Id="rId1" Type="x" Target="worksheets/sheet1.xml"/></Relationships>'
    ),
    'xl/styles.xml': strToU8(
      '<?xml version="1.0"?><styleSheet>' +
        '<fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><u val="none"/></font></fonts>' +
        '<fills count="1"><fill><patternFill patternType="none"/></fill></fills>' +
        '<cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>' +
        '<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/></cellXfs></styleSheet>'
    ),
    'xl/worksheets/sheet1.xml': strToU8(
      '<?xml version="1.0"?><worksheet><sheetData><row r="1"><c r="A1" s="1" t="inlineStr"><is><t>x</t></is></c></row></sheetData></worksheet>'
    ),
  };
  const underline = readXlsx(zipSync(files)).getWorksheet('S')?.getCell('A1').font?.underline;
  assert.ok(!underline, `<u val="none"/> must read back falsy, not the truthy string "none"; got ${JSON.stringify(underline)}`);
});

test('a foreign font honours an explicit-false boolean flag rather than tag presence', () => {
  // A foreign generator writes <b/> (bold on) but <i val="0"/> (italic explicitly off). The
  // reader must honour the val — a present tag is not truthy on its own.
  const files: Record<string, Uint8Array> = {
    '[Content_Types].xml': strToU8(
      '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
        '<Default Extension="xml" ContentType="application/xml"/></Types>'
    ),
    'xl/workbook.xml': strToU8(
      '<?xml version="1.0"?><workbook xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
        '<sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets></workbook>'
    ),
    'xl/_rels/workbook.xml.rels': strToU8(
      '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        '<Relationship Id="rId1" Type="x" Target="worksheets/sheet1.xml"/></Relationships>'
    ),
    'xl/styles.xml': strToU8(
      '<?xml version="1.0"?><styleSheet>' +
        '<fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><b/><i val="0"/></font></fonts>' +
        '<fills count="1"><fill><patternFill patternType="none"/></fill></fills>' +
        '<cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>' +
        '<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/></cellXfs></styleSheet>'
    ),
    'xl/worksheets/sheet1.xml': strToU8(
      '<?xml version="1.0"?><worksheet><sheetData><row r="1"><c r="A1" s="1" t="inlineStr"><is><t>x</t></is></c></row></sheetData></worksheet>'
    ),
  };
  const font = readXlsx(zipSync(files)).getWorksheet('S')?.getCell('A1').font;
  assert.equal(font?.bold, true, 'a bare <b/> is bold');
  assert.equal(font?.italic, false, '<i val="0"/> is explicitly not italic — the val is honoured');
});

test('a cell border round-trips through the <borders> table, and only the styled cell carries it', () => {
  const wb = new Workbook();
  const sheet = wb.addWorksheet('S');
  sheet.getCell('A1').value = 'boxed';
  sheet.getCell('A1').border = {
    top: {style: 'thin'},
    bottom: {style: 'medium', color: {argb: 'FF3A80D5'}},
  };
  sheet.getCell('B2').value = 'plain';

  const back = roundtrip(wb).getWorksheet('S');
  assert.deepEqual(back?.getCell('A1').border, {
    top: {style: 'thin'},
    bottom: {style: 'medium', color: {argb: 'FF3A80D5'}},
  });
  assert.equal(back?.getCell('B2').border, undefined, 'an unbordered sibling stays borderless');
});

test('a cell bordered on one side does not fabricate the other three across a round-trip', () => {
  const wb = new Workbook();
  const sheet = wb.addWorksheet('S');
  sheet.getCell('A1').value = 'x';
  sheet.getCell('A1').border = {top: {style: 'thin'}};

  const border = roundtrip(wb).getWorksheet('S')?.getCell('A1').border;
  assert.equal(border?.top?.style, 'thin', 'the declared top edge survives');
  for (const side of ['left', 'right', 'bottom', 'diagonal'] as const) {
    assert.equal(border?.[side], undefined, `${side} is not fabricated`);
  }
});

test('a foreign diagonal border reads its edge and diagonal direction', () => {
  // A foreign generator declares a diagonal border with diagonalUp; the reader must carry the
  // edge and the direction flag rather than dropping either.
  const files: Record<string, Uint8Array> = {
    '[Content_Types].xml': strToU8(
      '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
        '<Default Extension="xml" ContentType="application/xml"/></Types>'
    ),
    'xl/workbook.xml': strToU8(
      '<?xml version="1.0"?><workbook xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
        '<sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets></workbook>'
    ),
    'xl/_rels/workbook.xml.rels': strToU8(
      '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        '<Relationship Id="rId1" Type="x" Target="worksheets/sheet1.xml"/></Relationships>'
    ),
    'xl/styles.xml': strToU8(
      '<?xml version="1.0"?><styleSheet>' +
        '<fills count="1"><fill><patternFill patternType="none"/></fill></fills>' +
        '<borders count="2"><border><left/><right/><top/><bottom/><diagonal/></border>' +
        '<border diagonalUp="1"><left/><right/><top/><bottom/><diagonal style="thin"/></border></borders>' +
        '<cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>' +
        '<xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1"/></cellXfs></styleSheet>'
    ),
    'xl/worksheets/sheet1.xml': strToU8(
      '<?xml version="1.0"?><worksheet><sheetData><row r="1"><c r="A1" s="1" t="inlineStr"><is><t>x</t></is></c></row></sheetData></worksheet>'
    ),
  };
  const border = readXlsx(zipSync(files)).getWorksheet('S')?.getCell('A1').border;
  assert.equal(border?.diagonal?.style, 'thin', 'the diagonal edge survives');
  assert.equal(border?.diagonalUp, true, 'the diagonalUp direction is honoured');
  assert.equal(border?.top, undefined, 'a styleless edge is not fabricated');
});

test('cell alignment round-trips through the xf, and only the aligned cell carries it', () => {
  const wb = new Workbook();
  const sheet = wb.addWorksheet('S');
  sheet.getCell('A1').value = 'x';
  sheet.getCell('A1').alignment = {horizontal: 'center', vertical: 'top', wrapText: true, indent: 2, textRotation: 45};
  sheet.getCell('B2').value = 'plain';

  const back = roundtrip(wb).getWorksheet('S');
  assert.deepEqual(back?.getCell('A1').alignment, {
    horizontal: 'center',
    vertical: 'top',
    wrapText: true,
    indent: 2,
    textRotation: 45,
  });
  assert.equal(back?.getCell('B2').alignment, undefined, 'an unaligned sibling stays alignment-free');
});

test('alignment boolean flags left off do not read back spuriously enabled', () => {
  const wb = new Workbook();
  const sheet = wb.addWorksheet('S');
  sheet.getCell('A1').value = 'x';
  sheet.getCell('A1').alignment = {wrapText: false, shrinkToFit: false};

  // An all-default alignment serialises to nothing, so it resolves to xf 0 and reads back absent.
  assert.equal(roundtrip(wb).getWorksheet('S')?.getCell('A1').alignment, undefined);
});

test('a foreign alignment carrying only wrapText="0" reads back with no alignment', () => {
  // Excel writes an all-false alignment as wrapText="0"; the raw "0" is a truthy JS string, so a
  // reader that guards on presence rather than the parsed boolean would mistake it for present.
  const files: Record<string, Uint8Array> = {
    '[Content_Types].xml': strToU8(
      '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
        '<Default Extension="xml" ContentType="application/xml"/></Types>'
    ),
    'xl/workbook.xml': strToU8(
      '<?xml version="1.0"?><workbook xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
        '<sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets></workbook>'
    ),
    'xl/_rels/workbook.xml.rels': strToU8(
      '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        '<Relationship Id="rId1" Type="x" Target="worksheets/sheet1.xml"/></Relationships>'
    ),
    'xl/styles.xml': strToU8(
      '<?xml version="1.0"?><styleSheet>' +
        '<fills count="1"><fill><patternFill patternType="none"/></fill></fills>' +
        '<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1">' +
        '<alignment wrapText="0"/></xf></cellXfs></styleSheet>'
    ),
    'xl/worksheets/sheet1.xml': strToU8(
      '<?xml version="1.0"?><worksheet><sheetData><row r="1"><c r="A1" s="0" t="inlineStr"><is><t>x</t></is></c></row></sheetData></worksheet>'
    ),
  };
  assert.equal(readXlsx(zipSync(files)).getWorksheet('S')?.getCell('A1').alignment, undefined);
});

test('cell protection round-trips the meaningful flags, and only the protected cell carries them', () => {
  const wb = new Workbook();
  const sheet = wb.addWorksheet('S');
  sheet.getCell('A1').value = 'x';
  sheet.getCell('A1').protection = {locked: false, hidden: true};
  sheet.getCell('B2').value = 'plain';

  const back = roundtrip(wb).getWorksheet('S');
  assert.deepEqual(back?.getCell('A1').protection, {locked: false, hidden: true});
  assert.equal(back?.getCell('B2').protection, undefined, 'a default-protection sibling stays protection-free');
});

test('a default-locked cell does not read back as explicitly protected', () => {
  const wb = new Workbook();
  const sheet = wb.addWorksheet('S');
  sheet.getCell('A1').value = 'x';
  // locked defaults to on, hidden to off, so this restates the default and serialises to nothing.
  sheet.getCell('A1').protection = {locked: true, hidden: false};

  assert.equal(roundtrip(wb).getWorksheet('S')?.getCell('A1').protection, undefined);
});

test('a foreign <protection locked="1"> — an explicit default — reads back with no protection', () => {
  // A foreign generator states the default explicitly (locked on). Since locked defaults true,
  // that carries no information; the reader must not fabricate a { locked: true } protection.
  const files: Record<string, Uint8Array> = {
    '[Content_Types].xml': strToU8(
      '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
        '<Default Extension="xml" ContentType="application/xml"/></Types>'
    ),
    'xl/workbook.xml': strToU8(
      '<?xml version="1.0"?><workbook xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
        '<sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets></workbook>'
    ),
    'xl/_rels/workbook.xml.rels': strToU8(
      '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        '<Relationship Id="rId1" Type="x" Target="worksheets/sheet1.xml"/></Relationships>'
    ),
    'xl/styles.xml': strToU8(
      '<?xml version="1.0"?><styleSheet>' +
        '<fills count="1"><fill><patternFill patternType="none"/></fill></fills>' +
        '<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyProtection="1">' +
        '<protection locked="1"/></xf></cellXfs></styleSheet>'
    ),
    'xl/worksheets/sheet1.xml': strToU8(
      '<?xml version="1.0"?><worksheet><sheetData><row r="1"><c r="A1" s="0" t="inlineStr"><is><t>x</t></is></c></row></sheetData></worksheet>'
    ),
  };
  assert.equal(readXlsx(zipSync(files)).getWorksheet('S')?.getCell('A1').protection, undefined);
});

test('the inflate bound rejects a part whose declared size is over the cap', () => {
  const wb = new Workbook();
  wb.addWorksheet('S').getCell('A1').value = 'x';
  const buffer = writeXlsx(wb);
  assert.throws(() => readXlsx(buffer, {maxUncompressedBytes: 16}), /possible zip bomb/);
});

test('a zip that is not an xlsx (no workbook part) is rejected, not misread', () => {
  const bogus = zipSync({'hello.txt': strToU8('not a spreadsheet')});
  assert.throws(() => readXlsx(bogus), /xl\/workbook\.xml is missing/);
});

test('a t="s" shared-string cell resolves against the shared table', () => {
  // Our writer emits inlineStr, but the reader must also resolve shared strings that
  // foreign generators use. Assemble a minimal package by hand to exercise that path.
  const files: Record<string, Uint8Array> = {
    '[Content_Types].xml': strToU8(
      '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
        '<Default Extension="xml" ContentType="application/xml"/></Types>'
    ),
    'xl/workbook.xml': strToU8(
      '<?xml version="1.0"?><workbook xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
        '<sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets></workbook>'
    ),
    'xl/_rels/workbook.xml.rels': strToU8(
      '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        '<Relationship Id="rId1" Type="x" Target="worksheets/sheet1.xml"/></Relationships>'
    ),
    'xl/sharedStrings.xml': strToU8('<?xml version="1.0"?><sst><si><t>shared</t></si></sst>'),
    'xl/worksheets/sheet1.xml': strToU8(
      '<?xml version="1.0"?><worksheet><sheetData><row r="1"><c r="A1" t="s"><v>0</v></c></row></sheetData></worksheet>'
    ),
  };
  const back = readXlsx(zipSync(files)).getWorksheet('S');
  assert.equal(back?.getCell('A1').value, 'shared');
});

test('a password-protected sheet round-trips its credential and permissive flags', () => {
  const wb = new Workbook();
  const ws = wb.addWorksheet('S');
  ws.getCell('A1').value = 'x';
  ws.protect('secret', {sort: true, autoFilter: true, selectLockedCells: false});

  const back = roundtrip(wb).getWorksheet('S');
  const protection = back?.protection;
  assert.ok(protection, 'the reloaded sheet reports protection');

  // The permissive flags the author set survive; a default-valued flag stays absent.
  assert.equal(protection.flags.sort, true);
  assert.equal(protection.flags.autoFilter, true);
  assert.equal(protection.flags.selectLockedCells, false);
  assert.equal(protection.flags.deleteRows, undefined);

  // The agile credential is preserved verbatim — the reader cannot (and must not) re-hash it.
  const original = ws.protection?.credential;
  assert.deepEqual(protection.credential, original);
  assert.equal(protection.credential?.algorithmName, 'SHA-512');
});

test('a passwordless protected sheet round-trips as protected with no credential', () => {
  const wb = new Workbook();
  const ws = wb.addWorksheet('S');
  ws.getCell('A1').value = 'x';
  ws.protect();

  const protection = roundtrip(wb).getWorksheet('S')?.protection;
  assert.ok(protection, 'a soft (passwordless) lock still round-trips as protection');
  assert.equal(protection.credential, undefined);
});

test('an unprotected sheet reports no protection after a round-trip', () => {
  const wb = new Workbook();
  wb.addWorksheet('S').getCell('A1').value = 'x';
  assert.equal(roundtrip(wb).getWorksheet('S')?.protection, undefined);
});

test('a re-written loaded protection preserves the credential byte-for-byte', () => {
  const wb = new Workbook();
  const ws = wb.addWorksheet('S');
  ws.getCell('A1').value = 'x';
  ws.protect('pw', {formatCells: true});

  const firstXml = strFromU8(unzipSync(writeXlsx(wb))['xl/worksheets/sheet1.xml']!);
  const secondXml = strFromU8(unzipSync(writeXlsx(roundtrip(wb)))['xl/worksheets/sheet1.xml']!);
  const prot = (xml: string): string => (xml.match(/<sheetProtection\b[^>]*\/>/) ?? [''])[0];

  assert.ok(prot(firstXml), 'the first write emits a sheetProtection element');
  assert.equal(prot(secondXml), prot(firstXml), 'a passthrough save re-emits the identical protection');
});

test('a Date value round-trips as a Date on the calendar day it was written', () => {
  const wb = new Workbook();
  wb.addWorksheet('S').getCell('A1').value = new Date('2020-03-04T00:00:00.000Z');

  const back = roundtrip(wb).getWorksheet('S')?.getCell('A1');
  assert.equal(back?.type, 'date');
  assert.ok(back?.value instanceof Date);
  assert.equal((back?.value as Date).toISOString(), '2020-03-04T00:00:00.000Z');
});

test('a bare Date is written under a date number format so it reads back as a date', () => {
  const wb = new Workbook();
  wb.addWorksheet('S').getCell('A1').value = new Date('2020-03-04T00:00:00.000Z');

  const xml = strFromU8(unzipSync(writeXlsx(wb))['xl/worksheets/sheet1.xml'] as Uint8Array);
  assert.ok(!/t="/.test(xml.match(/<c r="A1"[^>]*>/)?.[0] ?? ''), 'a date serial is a plain number cell, no t=');
});

test('an explicit date numFmt on a Date cell survives verbatim, not swapped for the default', () => {
  const wb = new Workbook();
  const cell = wb.addWorksheet('S').getCell('A1');
  cell.value = new Date('2020-03-04T00:00:00.000Z');
  cell.numFmt = 'DD/MM/YYYY';

  const back = roundtrip(wb).getWorksheet('S')?.getCell('A1');
  assert.equal(back?.numFmt, 'DD/MM/YYYY');
  assert.equal(back?.type, 'date');
});

test('a serial under a non-date format reads back as a plain number, not a date', () => {
  const wb = new Workbook();
  const cell = wb.addWorksheet('S').getCell('A1');
  cell.value = 43_894; // the serial for 2020-03-04
  cell.numFmt = '0.00';

  const back = roundtrip(wb).getWorksheet('S')?.getCell('A1');
  assert.equal(back?.type, 'number');
  assert.equal(back?.value, 43_894);
});

test('an Invalid Date does not throw on write and does not drop sibling cells', () => {
  const wb = new Workbook();
  const sheet = wb.addWorksheet('S');
  sheet.getCell('A1').value = new Date(NaN);
  sheet.getCell('B1').value = 'still here';
  sheet.getCell('C1').value = 42;

  const back = roundtrip(wb).getWorksheet('S');
  assert.equal(back?.getCell('B1').value, 'still here');
  assert.equal(back?.getCell('C1').value, 42);
  assert.equal(back?.getCell('A1').value, null, 'the invalid date carries no serial');
});

test('a worksheet tab colour round-trips as the exact 8-digit ARGB', () => {
  const wb = new Workbook();
  const sheet = wb.addWorksheet('S');
  sheet.tabColor = {argb: 'FFFF0000'};
  sheet.getCell('A1').value = 'x';

  const back = roundtrip(wb).getWorksheet('S');
  assert.deepEqual(back?.tabColor, {argb: 'FFFF0000'});
});

test('the written package carries the tab colour under <sheetPr>', () => {
  const wb = new Workbook();
  wb.addWorksheet('S').tabColor = {argb: 'FF00FF00'};

  const xml = strFromU8(unzipSync(writeXlsx(wb))['xl/worksheets/sheet1.xml']!);
  assert.match(xml, /<sheetPr><tabColor rgb="FF00FF00"\/><\/sheetPr>/);
  // <sheetPr> must lead the worksheet, before <dimension>.
  assert.ok(xml.indexOf('<sheetPr>') < xml.indexOf('<dimension'), 'sheetPr precedes dimension');
});

test('a sheet with no tab colour acquires none and emits no <sheetPr>', () => {
  const wb = new Workbook();
  wb.addWorksheet('S').getCell('A1').value = 'y';

  const xml = strFromU8(unzipSync(writeXlsx(wb))['xl/worksheets/sheet1.xml']!);
  assert.doesNotMatch(xml, /<sheetPr>/);
  assert.equal(roundtrip(wb).getWorksheet('S')?.tabColor, undefined);
});

test('a theme-based tab colour with a tint round-trips', () => {
  const wb = new Workbook();
  wb.addWorksheet('S').tabColor = {theme: 4, tint: -0.25};

  const back = roundtrip(wb).getWorksheet('S');
  assert.deepEqual(back?.tabColor, {theme: 4, tint: -0.25});
});

test('a tab colour survives a worksheet model export/import', () => {
  const wb = new Workbook();
  const src = wb.addWorksheet('Src');
  src.tabColor = {argb: 'FF0000FF'};
  const dst = wb.addWorksheet('Dst');
  dst.model = src.model;
  assert.deepEqual(dst.tabColor, {argb: 'FF0000FF'});
});

test('inverted outline summary positions round-trip as false', () => {
  const wb = new Workbook();
  const sheet = wb.addWorksheet('S');
  sheet.outline.summaryBelow = false;
  sheet.outline.summaryRight = false;
  sheet.getCell('A1').value = 'x';

  const back = roundtrip(wb).getWorksheet('S');
  assert.equal(back?.outline.summaryBelow, false);
  assert.equal(back?.outline.summaryRight, false);
});

test('the written package carries the outline flags under <sheetPr>', () => {
  const wb = new Workbook();
  const sheet = wb.addWorksheet('S');
  sheet.outline.summaryBelow = false;
  sheet.outline.summaryRight = false;

  const xml = strFromU8(unzipSync(writeXlsx(wb))['xl/worksheets/sheet1.xml']!);
  assert.match(xml, /<sheetPr><outlinePr summaryBelow="0" summaryRight="0"\/><\/sheetPr>/);
});

test('the tab colour and outline flags share one <sheetPr> in CT_SheetPr order', () => {
  const wb = new Workbook();
  const sheet = wb.addWorksheet('S');
  sheet.tabColor = {argb: 'FFFF0000'};
  sheet.outline.summaryBelow = false;

  const xml = strFromU8(unzipSync(writeXlsx(wb))['xl/worksheets/sheet1.xml']!);
  assert.match(xml, /<sheetPr><tabColor rgb="FFFF0000"\/><outlinePr summaryBelow="0"\/><\/sheetPr>/);
});

test('only the set outline flag serializes; the untouched one stays absent', () => {
  const wb = new Workbook();
  wb.addWorksheet('S').outline.summaryRight = false;

  const xml = strFromU8(unzipSync(writeXlsx(wb))['xl/worksheets/sheet1.xml']!);
  assert.match(xml, /<outlinePr summaryRight="0"\/>/);
  assert.doesNotMatch(xml, /summaryBelow/);
});

test('a sheet with default outline positions emits no <outlinePr>', () => {
  const wb = new Workbook();
  wb.addWorksheet('S').getCell('A1').value = 'y';

  const xml = strFromU8(unzipSync(writeXlsx(wb))['xl/worksheets/sheet1.xml']!);
  assert.doesNotMatch(xml, /<outlinePr/);
  const back = roundtrip(wb).getWorksheet('S');
  assert.equal(back?.outline.summaryBelow, undefined);
  assert.equal(back?.outline.summaryRight, undefined);
});

test('outline flags survive a worksheet model export/import', () => {
  const wb = new Workbook();
  const src = wb.addWorksheet('Src');
  src.outline.summaryBelow = false;
  const dst = wb.addWorksheet('Dst');
  dst.model = src.model;
  assert.equal(dst.outline.summaryBelow, false);
});

test('a fit-to-page setup round-trips its flag, counts, and scale', () => {
  const wb = new Workbook();
  const sheet = wb.addWorksheet('S');
  sheet.pageSetup.fitToPage = true;
  sheet.pageSetup.fitToWidth = 1;
  sheet.pageSetup.fitToHeight = 0;
  sheet.pageSetup.scale = 80;
  sheet.getCell('A1').value = 'x';

  const back = roundtrip(wb).getWorksheet('S');
  assert.equal(back?.pageSetup.fitToPage, true);
  assert.equal(back?.pageSetup.fitToWidth, 1);
  assert.equal(back?.pageSetup.fitToHeight, 0);
  assert.equal(back?.pageSetup.scale, 80);
});

test('the fit-to-page flag rides <pageSetUpPr> under <sheetPr>, the counts ride <pageSetup>', () => {
  const wb = new Workbook();
  const sheet = wb.addWorksheet('S');
  sheet.pageSetup.fitToPage = true;
  sheet.pageSetup.fitToWidth = 1;

  const xml = strFromU8(unzipSync(writeXlsx(wb))['xl/worksheets/sheet1.xml']!);
  assert.match(xml, /<sheetPr><pageSetUpPr fitToPage="1"\/><\/sheetPr>/);
  assert.match(xml, /<pageSetup fitToWidth="1"\/>/);
});

test('<pageSetUpPr> follows <outlinePr> under <sheetPr> in CT_SheetPr order', () => {
  const wb = new Workbook();
  const sheet = wb.addWorksheet('S');
  sheet.outline.summaryBelow = false;
  sheet.pageSetup.fitToPage = true;

  const xml = strFromU8(unzipSync(writeXlsx(wb))['xl/worksheets/sheet1.xml']!);
  assert.match(xml, /<sheetPr><outlinePr summaryBelow="0"\/><pageSetUpPr fitToPage="1"\/><\/sheetPr>/);
});

test('<pageSetup> sits between <pageMargins> and <headerFooter>', () => {
  const wb = new Workbook();
  const sheet = wb.addWorksheet('S');
  sheet.pageMargins.top = 1;
  sheet.pageSetup.scale = 90;
  sheet.headerFooter.oddHeader = 'H';

  const xml = strFromU8(unzipSync(writeXlsx(wb))['xl/worksheets/sheet1.xml']!);
  assert.ok(xml.indexOf('<pageMargins') < xml.indexOf('<pageSetup'), 'pageSetup follows pageMargins');
  assert.ok(xml.indexOf('<pageSetup') < xml.indexOf('<headerFooter'), 'pageSetup precedes headerFooter');
});

test('orientation and pageOrder round-trip and emit only when set', () => {
  const wb = new Workbook();
  const sheet = wb.addWorksheet('S');
  sheet.pageSetup.orientation = 'landscape';
  sheet.pageSetup.pageOrder = 'overThenDown';

  const xml = strFromU8(unzipSync(writeXlsx(wb))['xl/worksheets/sheet1.xml']!);
  assert.match(xml, /<pageSetup pageOrder="overThenDown" orientation="landscape"\/>/);
  assert.doesNotMatch(xml, /scale=|fitToWidth=|fitToHeight=/);

  const back = roundtrip(wb).getWorksheet('S');
  assert.equal(back?.pageSetup.orientation, 'landscape');
  assert.equal(back?.pageSetup.pageOrder, 'overThenDown');
});

test('paperSize round-trips and leads the <pageSetup> attributes', () => {
  const wb = new Workbook();
  const sheet = wb.addWorksheet('S');
  sheet.pageSetup.paperSize = 9;
  sheet.pageSetup.scale = 96;

  const xml = strFromU8(unzipSync(writeXlsx(wb))['xl/worksheets/sheet1.xml']!);
  assert.match(xml, /<pageSetup paperSize="9" scale="96"\/>/);

  const back = roundtrip(wb).getWorksheet('S');
  assert.equal(back?.pageSetup.paperSize, 9);
});

test('a non-numeric paperSize is dropped on read, not stored as NaN', () => {
  const wb = new Workbook();
  wb.addWorksheet('S').pageSetup.scale = 96;
  const files = unzipSync(writeXlsx(wb));
  files['xl/worksheets/sheet1.xml'] = strToU8(
    strFromU8(files['xl/worksheets/sheet1.xml']!).replace('<pageSetup ', '<pageSetup paperSize="A4" ')
  );
  const back = readXlsx(zipSync(files)).getWorksheet('S');
  assert.equal(back?.pageSetup.paperSize, undefined);
  assert.equal(back?.pageSetup.scale, 96);
});

test('a sheet with no page setup emits neither <pageSetUpPr> nor <pageSetup>', () => {
  const wb = new Workbook();
  wb.addWorksheet('S').getCell('A1').value = 'y';

  const xml = strFromU8(unzipSync(writeXlsx(wb))['xl/worksheets/sheet1.xml']!);
  assert.doesNotMatch(xml, /<pageSetUpPr/);
  assert.doesNotMatch(xml, /<pageSetup/);
  const back = roundtrip(wb).getWorksheet('S');
  assert.equal(back?.pageSetup.fitToPage, undefined);
  assert.equal(back?.pageSetup.scale, undefined);
});

test('a <pageSetUpPr> present only for other reasons leaves fitToPage unset', () => {
  const wb = new Workbook();
  wb.addWorksheet('S').getCell('A1').value = 'y';
  const files = unzipSync(writeXlsx(wb));
  files['xl/worksheets/sheet1.xml'] = strToU8(
    strFromU8(files['xl/worksheets/sheet1.xml']!).replace('<dimension', '<sheetPr><pageSetUpPr autoPageBreaks="0"/></sheetPr><dimension')
  );
  const back = readXlsx(zipSync(files)).getWorksheet('S');
  assert.equal(back?.pageSetup.fitToPage, undefined);
});

test('page setup survives a worksheet model export/import', () => {
  const wb = new Workbook();
  const src = wb.addWorksheet('Src');
  src.pageSetup.fitToPage = true;
  src.pageSetup.scale = 75;
  const dst = wb.addWorksheet('Dst');
  dst.model = src.model;
  assert.equal(dst.pageSetup.fitToPage, true);
  assert.equal(dst.pageSetup.scale, 75);
});
