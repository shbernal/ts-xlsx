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

test('column width and visibility round-trip', () => {
  const wb = new Workbook();
  const sheet = wb.addWorksheet('S');
  sheet.getColumn(2).width = 24;
  sheet.getColumn(4).hidden = true;
  const back = roundtrip(wb).getWorksheet('S');
  assert.equal(back?.getColumn(2).width, 24);
  assert.equal(back?.getColumn(4).hidden, true);
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
