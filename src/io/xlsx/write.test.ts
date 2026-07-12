import assert from 'node:assert/strict';
import {test} from 'node:test';

import {strFromU8, unzipSync} from 'fflate';

import {Workbook} from '../../core/workbook.ts';
import {writeXlsx} from './write.ts';

function partsOf(workbook: Workbook): Record<string, string> {
  const unzipped = unzipSync(writeXlsx(workbook));
  const out: Record<string, string> = {};
  for (const [name, bytes] of Object.entries(unzipped)) out[name] = strFromU8(bytes);
  return out;
}

test('a one-sheet workbook writes the full set of OPC parts', () => {
  const wb = new Workbook();
  wb.addWorksheet('S').getCell('A1').value = 'x';
  const names = Object.keys(partsOf(wb)).sort();
  assert.deepEqual(names, [
    '[Content_Types].xml',
    '_rels/.rels',
    'docProps/app.xml',
    'docProps/core.xml',
    'xl/_rels/workbook.xml.rels',
    'xl/styles.xml',
    'xl/theme/theme1.xml',
    'xl/workbook.xml',
    'xl/worksheets/sheet1.xml',
  ]);
});

test('a workbook with no worksheets is refused rather than written corrupt', () => {
  assert.throws(() => writeXlsx(new Workbook()), /no worksheets/);
});

test('the content types and rels declare each worksheet consistently', () => {
  const wb = new Workbook();
  wb.addWorksheet('One').getCell('A1').value = 1;
  wb.addWorksheet('Two').getCell('A1').value = 2;
  const parts = partsOf(wb);
  for (const i of [1, 2]) {
    assert.match(parts['[Content_Types].xml'] as string, new RegExp(`/xl/worksheets/sheet${i}\\.xml`));
    assert.match(parts['xl/_rels/workbook.xml.rels'] as string, new RegExp(`worksheets/sheet${i}\\.xml`));
    assert.ok(parts[`xl/worksheets/sheet${i}.xml`], `sheet${i}.xml part exists`);
  }
});

test('a default sheet is visible (no state attribute); explicit states are written', () => {
  const wb = new Workbook();
  wb.addWorksheet('Visible').getCell('A1').value = 'x';
  wb.addWorksheet('Hidden', {state: 'hidden'}).getCell('A1').value = 'x';
  const xml = partsOf(wb)['xl/workbook.xml'] as string;
  assert.match(xml, /<sheet name="Visible" sheetId="1" r:id="rId1"\/>/);
  assert.match(xml, /<sheet name="Hidden" sheetId="2" state="hidden" r:id="rId2"\/>/);
});

test('cell values serialise by type with a computed dimension', () => {
  const wb = new Workbook();
  const s = wb.addWorksheet('S');
  s.getCell('B2').value = 42;
  s.getCell('C2').value = true;
  s.getCell('B3').value = 'hi';
  const xml = partsOf(wb)['xl/worksheets/sheet1.xml'] as string;
  assert.match(xml, /<dimension ref="B2:C3"\/>/);
  assert.match(xml, /<c r="B2"><v>42<\/v><\/c>/);
  assert.match(xml, /<c r="C2" t="b"><v>1<\/v><\/c>/);
  assert.match(xml, /<c r="B3" t="inlineStr"><is><t>hi<\/t><\/is><\/c>/);
});

test('XML-special characters in text and formulas are escaped', () => {
  const wb = new Workbook();
  const s = wb.addWorksheet('S');
  s.getCell('A1').value = 'a < b & c > d';
  s.getCell('A2').value = {formula: 'IF(A1<B1,"x"&"y","")'};
  const xml = partsOf(wb)['xl/worksheets/sheet1.xml'] as string;
  assert.match(xml, /<t>a &lt; b &amp; c &gt; d<\/t>/);
  assert.match(xml, /<f>IF\(A1&lt;B1,"x"&amp;"y",""\)<\/f>/);
  // No raw ampersand survives except as the head of an entity — the check the corpus's
  // xmlWellFormed applies to reject unescaped specials.
  assert.doesNotMatch(xml, /&(?!(amp|lt|gt|quot|apos|#\d+|#x[0-9a-fA-F]+);)/);
});

test('a formula supplied with a leading = is stored without it', () => {
  const wb = new Workbook();
  wb.addWorksheet('S').getCell('A1').value = {formula: '=1+2', result: 3};
  const xml = partsOf(wb)['xl/worksheets/sheet1.xml'] as string;
  assert.match(xml, /<f>1\+2<\/f>/);
  assert.doesNotMatch(xml, /<f>=/);
});

test('a string with edge whitespace carries xml:space="preserve"', () => {
  const wb = new Workbook();
  wb.addWorksheet('S').getCell('A1').value = '  padded  ';
  const xml = partsOf(wb)['xl/worksheets/sheet1.xml'] as string;
  assert.match(xml, /<t xml:space="preserve">  padded  <\/t>/);
});

test('a non-finite number is refused, not written as NaN', () => {
  const wb = new Workbook();
  wb.addWorksheet('S').getCell('A1').value = Number.POSITIVE_INFINITY;
  assert.throws(() => writeXlsx(wb), /non-finite/);
});

test('a value kind the writer cannot represent yet is refused', () => {
  const wb = new Workbook();
  wb.addWorksheet('S').getCell('A1').value = new Date();
  assert.throws(() => writeXlsx(wb), /not implemented yet/);
});

test('a formula cell with a string result is typed t="str"', () => {
  const wb = new Workbook();
  wb.addWorksheet('S').getCell('A1').value = {formula: 'A2&A3', result: 'joined'};
  const xml = partsOf(wb)['xl/worksheets/sheet1.xml'] as string;
  assert.match(xml, /<c r="A1" t="str"><f>A2&amp;A3<\/f><v>joined<\/v><\/c>/);
});

test('every written package ships a default theme part', () => {
  const wb = new Workbook();
  wb.addWorksheet('S').getCell('A1').value = 'x';
  const parts = partsOf(wb);
  assert.ok(parts['xl/theme/theme1.xml'], 'theme part present');
  assert.match(parts['[Content_Types].xml'] as string, /theme\+xml/);
});

test('a column width emits a <col> with customWidth', () => {
  const wb = new Workbook();
  const s = wb.addWorksheet('S');
  s.getCell('A1').value = 'x';
  s.getColumn(2).width = 12;
  const xml = partsOf(wb)['xl/worksheets/sheet1.xml'] as string;
  assert.match(xml, /<cols><col min="2" max="2" width="12" customWidth="1"\/><\/cols>/);
});

test('a hidden column emits hidden="1" and needs no width', () => {
  const wb = new Workbook();
  const s = wb.addWorksheet('S');
  s.getCell('A1').value = 'x';
  s.getColumn(3).hidden = true;
  const xml = partsOf(wb)['xl/worksheets/sheet1.xml'] as string;
  assert.match(xml, /<col min="3" max="3" hidden="1"\/>/);
});

test('a column past the 16384 limit is dropped, never serialized', () => {
  const wb = new Workbook();
  const s = wb.addWorksheet('S');
  s.getCell('A1').value = 'x';
  s.getColumn(16384).width = 10;
  s.getColumn(16385).width = 10;
  const xml = partsOf(wb)['xl/worksheets/sheet1.xml'] as string;
  assert.match(xml, /min="16384" max="16384"/);
  assert.doesNotMatch(xml, /16385/);
});

test('a sheet with no column definitions emits no <cols>', () => {
  const wb = new Workbook();
  wb.addWorksheet('S').getCell('A1').value = 'x';
  const xml = partsOf(wb)['xl/worksheets/sheet1.xml'] as string;
  assert.doesNotMatch(xml, /<cols>/);
});

test('row height and outline flags serialize onto the <row>', () => {
  const wb = new Workbook();
  const s = wb.addWorksheet('S');
  s.getCell('A2').value = 'd';
  const r = s.getRow(2);
  r.height = 30;
  r.hidden = true;
  r.outlineLevel = 1;
  const xml = partsOf(wb)['xl/worksheets/sheet1.xml'] as string;
  assert.match(xml, /<row r="2" ht="30" customHeight="1" hidden="1" outlineLevel="1">/);
});

test('a row carrying only metadata is emitted with no cells', () => {
  const wb = new Workbook();
  const s = wb.addWorksheet('S');
  s.getCell('A1').value = 'x';
  s.getRow(5).hidden = true;
  const xml = partsOf(wb)['xl/worksheets/sheet1.xml'] as string;
  assert.match(xml, /<row r="5" hidden="1"><\/row>/);
});

test('a collapsed flag is emitted only where set, not on sibling rows', () => {
  const wb = new Workbook();
  const s = wb.addWorksheet('S');
  s.getRow(2).outlineLevel = 1;
  s.getRow(3).collapsed = true;
  const xml = partsOf(wb)['xl/worksheets/sheet1.xml'] as string;
  assert.match(xml, /<row r="2" outlineLevel="1">/);
  assert.match(xml, /<row r="3" collapsed="1">/);
});

test('sheet default row height and column width land on <sheetFormatPr>', () => {
  const wb = new Workbook();
  const s = wb.addWorksheet('S');
  s.getCell('A1').value = 'x';
  s.properties.defaultRowHeight = 30;
  s.properties.defaultColWidth = 20;
  const xml = partsOf(wb)['xl/worksheets/sheet1.xml'] as string;
  assert.match(xml, /<sheetFormatPr defaultRowHeight="30" defaultColWidth="20" customHeight="1"\/>/);
});

test('an unset default row height falls back to 15 with no customHeight', () => {
  const wb = new Workbook();
  wb.addWorksheet('S').getCell('A1').value = 'x';
  const xml = partsOf(wb)['xl/worksheets/sheet1.xml'] as string;
  assert.match(xml, /<sheetFormatPr defaultRowHeight="15"\/>/);
});

test('<cols> is placed after <sheetFormatPr> and before <sheetData>', () => {
  const wb = new Workbook();
  const s = wb.addWorksheet('S');
  s.getCell('A1').value = 'x';
  s.getColumn(1).width = 8;
  const xml = partsOf(wb)['xl/worksheets/sheet1.xml'] as string;
  const fmt = xml.indexOf('<sheetFormatPr');
  const cols = xml.indexOf('<cols>');
  const data = xml.indexOf('<sheetData>');
  assert.ok(fmt < cols && cols < data, `expected sheetFormatPr < cols < sheetData, got ${fmt},${cols},${data}`);
});
