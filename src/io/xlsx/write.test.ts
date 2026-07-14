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
  wb.addWorksheet('S').getCell('A1').value = {error: '#REF!'};
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

test('column outline grouping serializes onto the <col>', () => {
  const wb = new Workbook();
  const s = wb.addWorksheet('S');
  s.getCell('A1').value = 'x';
  const c = s.getColumn(2);
  c.outlineLevel = 1;
  c.collapsed = true;
  const xml = partsOf(wb)['xl/worksheets/sheet1.xml'] as string;
  assert.match(xml, /<col min="2" max="2" outlineLevel="1" collapsed="1"\/>/);
});

test('an ungrouped column emits no outline attributes', () => {
  const wb = new Workbook();
  const s = wb.addWorksheet('S');
  s.getCell('A1').value = 'x';
  s.getColumn(2).width = 10;
  const xml = partsOf(wb)['xl/worksheets/sheet1.xml'] as string;
  assert.doesNotMatch(xml, /outlineLevel|collapsed/);
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

test('a formatted-but-empty cell is emitted as a styled <c> with no value, not dropped', () => {
  const wb = new Workbook();
  const s = wb.addWorksheet('S');
  s.getCell('A1').value = 'x';
  // B2 carries a fill but no value — a real formatted blank Excel keeps, not a cell to discard.
  s.getCell('B2').fill = {type: 'pattern', pattern: 'solid', fgColor: {argb: 'FF00FF00'}};
  const xml = partsOf(wb)['xl/worksheets/sheet1.xml'] as string;
  assert.match(xml, /<c r="B2" s="\d+"\/>/, 'the styled empty cell is written with its style and no <v>');
  assert.doesNotMatch(xml, /<c r="B2"[^>]*>.*<\/c>/, 'a valueless styled cell carries no child content');
});

test('an empty cell with no style of its own is not serialised', () => {
  const wb = new Workbook();
  const s = wb.addWorksheet('S');
  s.getCell('A1').value = 'x';
  // Touching a cell without giving it a value or style must not fabricate a <c> for it.
  s.getCell('C3');
  const xml = partsOf(wb)['xl/worksheets/sheet1.xml'] as string;
  assert.doesNotMatch(xml, /r="C3"/, 'a value-less, style-less cell contributes nothing');
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

test('setting a subset of margins emits all six pageMargins attributes', () => {
  const wb = new Workbook();
  const s = wb.addWorksheet('S');
  s.getCell('A1').value = 'x';
  s.pageMargins.left = 0.1;
  s.pageMargins.right = 0.1;
  const xml = partsOf(wb)['xl/worksheets/sheet1.xml'] as string;
  const tag = /<pageMargins ([^/]*)\/>/.exec(xml)?.[1] ?? '';
  for (const side of ['left', 'right', 'top', 'bottom', 'header', 'footer']) {
    assert.match(tag, new RegExp(`\\b${side}="[0-9.]+"`), `missing ${side}`);
  }
  // The explicitly-set sides keep their values; the untouched ones fall back to defaults.
  assert.match(tag, /left="0.1"/);
  assert.match(tag, /right="0.1"/);
  assert.match(tag, /top="0.75"/);
});

test('a sheet with no margins set emits no <pageMargins>', () => {
  const wb = new Workbook();
  wb.addWorksheet('S').getCell('A1').value = 'x';
  const xml = partsOf(wb)['xl/worksheets/sheet1.xml'] as string;
  assert.doesNotMatch(xml, /<pageMargins/);
});

test('<pageMargins> is placed after <sheetData>', () => {
  const wb = new Workbook();
  const s = wb.addWorksheet('S');
  s.getCell('A1').value = 'x';
  s.pageMargins.top = 1;
  const xml = partsOf(wb)['xl/worksheets/sheet1.xml'] as string;
  assert.ok(xml.indexOf('<sheetData') < xml.indexOf('<pageMargins'), 'pageMargins must follow sheetData');
});

test('header/footer variants emit their children and gate them with different* flags', () => {
  const wb = new Workbook();
  const s = wb.addWorksheet('S');
  s.getCell('A1').value = 'x';
  Object.assign(s.headerFooter, {
    oddHeader: 'ODD-H',
    evenHeader: 'EVEN-H',
    firstFooter: 'FIRST-F',
  });
  const xml = partsOf(wb)['xl/worksheets/sheet1.xml'] as string;
  assert.match(xml, /<headerFooter[^>]* differentOddEven="1"/);
  assert.match(xml, /<headerFooter[^>]* differentFirst="1"/);
  assert.match(xml, /<oddHeader>ODD-H<\/oddHeader>/);
  assert.match(xml, /<evenHeader>EVEN-H<\/evenHeader>/);
  assert.match(xml, /<firstFooter>FIRST-F<\/firstFooter>/);
});

test('an odd-only header/footer sets no different* flags', () => {
  const wb = new Workbook();
  const s = wb.addWorksheet('S');
  s.getCell('A1').value = 'x';
  s.headerFooter.oddHeader = 'H';
  const xml = partsOf(wb)['xl/worksheets/sheet1.xml'] as string;
  assert.match(xml, /<headerFooter><oddHeader>H<\/oddHeader><\/headerFooter>/);
  assert.doesNotMatch(xml, /different/);
});

test('a sheet with no header/footer emits no <headerFooter>', () => {
  const wb = new Workbook();
  wb.addWorksheet('S').getCell('A1').value = 'x';
  const xml = partsOf(wb)['xl/worksheets/sheet1.xml'] as string;
  assert.doesNotMatch(xml, /<headerFooter/);
});

test('header/footer text is XML-escaped and placed after <pageMargins>', () => {
  const wb = new Workbook();
  const s = wb.addWorksheet('S');
  s.getCell('A1').value = 'x';
  s.pageMargins.top = 1;
  s.headerFooter.oddHeader = 'a & b < c';
  const xml = partsOf(wb)['xl/worksheets/sheet1.xml'] as string;
  assert.match(xml, /<oddHeader>a &amp; b &lt; c<\/oddHeader>/);
  assert.ok(xml.indexOf('<pageMargins') < xml.indexOf('<headerFooter'), 'headerFooter follows pageMargins');
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

test('an empty-body table refs the full header row and writes a table part', () => {
  const wb = new Workbook();
  wb.addWorksheet('S').addTable({
    name: 'T1',
    ref: 'A1',
    columns: [{name: 'Alpha'}, {name: 'Beta'}],
    rowCount: 0,
  });
  const parts = partsOf(wb);
  const table = parts['xl/tables/table1.xml'] as string;
  assert.match(table, /ref="A1:B1"/);
  assert.match(table, /<tableColumns count="2">/);
  assert.match(table, /<autoFilter ref="A1:B1"\/>/);
  assert.match(parts['[Content_Types].xml'] as string, /\/xl\/tables\/table1\.xml/);
  assert.match(parts['xl/worksheets/_rels/sheet1.xml.rels'] as string, /Target="\.\.\/tables\/table1\.xml"/);
  assert.match(parts['xl/worksheets/sheet1.xml'] as string, /<tableParts count="1"><tablePart r:id="rId1"\/><\/tableParts>/);
});

test('a data row extends the table ref by one row', () => {
  const wb = new Workbook();
  wb.addWorksheet('S').addTable({name: 'T', ref: 'A1', columns: [{name: 'A'}, {name: 'B'}], rowCount: 1});
  const table = partsOf(wb)['xl/tables/table1.xml'] as string;
  assert.match(table, /ref="A1:B2"/);
});

test('a headerless table sets headerRowCount="0" and emits no autoFilter', () => {
  const wb = new Workbook();
  wb.addWorksheet('S').addTable({
    name: 'H',
    ref: 'A1',
    columns: [{name: 'A'}, {name: 'B'}],
    rowCount: 2,
    headerRow: false,
  });
  const table = partsOf(wb)['xl/tables/table1.xml'] as string;
  assert.match(table, /headerRowCount="0"/);
  assert.doesNotMatch(table, /<autoFilter/);
});

test('a totals-row column serialises its function and keeps every column', () => {
  const wb = new Workbook();
  wb.addWorksheet('S').addTable({
    name: 'T',
    ref: 'A1',
    columns: [{name: 'Item', totalsRowLabel: 'Total'}, {name: 'Amount', totalsRowFunction: 'sum'}],
    rowCount: 2,
    totalsRow: true,
  });
  const table = partsOf(wb)['xl/tables/table1.xml'] as string;
  assert.match(table, /ref="A1:B4"/);
  assert.match(table, /totalsRowCount="1"/);
  assert.match(table, /<tableColumn id="1" name="Item" totalsRowLabel="Total"\/>/);
  assert.match(table, /<tableColumn id="2" name="Amount" totalsRowFunction="sum"\/>/);
});

test('an illegal table name is rejected at definition time', () => {
  const s = new Workbook().addWorksheet('S');
  assert.throws(() => s.addTable({name: "Bob's Accounts", ref: 'A1', columns: [{name: 'A'}], rowCount: 1}), /identifier/);
  assert.throws(() => s.addTable({name: '1Digit', ref: 'A1', columns: [{name: 'A'}], rowCount: 1}), /identifier/);
  assert.throws(() => s.addTable({name: 'test-name', ref: 'A1', columns: [{name: 'A'}], rowCount: 1}), /identifier/);
});

test('a valid identifier table name is written verbatim', () => {
  const wb = new Workbook();
  wb.addWorksheet('S').addTable({name: 'Valid_Name', ref: 'A1', columns: [{name: 'A'}], rowCount: 1});
  const table = partsOf(wb)['xl/tables/table1.xml'] as string;
  assert.match(table, /name="Valid_Name"/);
  assert.match(table, /displayName="Valid_Name"/);
});

test('tables are numbered globally across sheets with sheet-local rel ids', () => {
  const wb = new Workbook();
  wb.addWorksheet('One').addTable({name: 'Ta', ref: 'A1', columns: [{name: 'A'}], rowCount: 1});
  wb.addWorksheet('Two').addTable({name: 'Tb', ref: 'A1', columns: [{name: 'A'}], rowCount: 1});
  const parts = partsOf(wb);
  assert.ok(parts['xl/tables/table1.xml'], 'first table part');
  assert.ok(parts['xl/tables/table2.xml'], 'second table part (globally numbered)');
  assert.match(parts['xl/worksheets/_rels/sheet2.xml.rels'] as string, /Target="\.\.\/tables\/table2\.xml"/);
});

test('a merge overlapping a table is rejected; a disjoint merge is written', () => {
  const overlap = new Workbook();
  const s1 = overlap.addWorksheet('S');
  s1.addTable({name: 'T', ref: 'A1', columns: [{name: 'A'}, {name: 'B'}], rowCount: 2});
  s1.mergeCells('A2:B2');
  assert.throws(() => writeXlsx(overlap), /overlaps table/);

  const disjoint = new Workbook();
  const s2 = disjoint.addWorksheet('S');
  s2.addTable({name: 'T', ref: 'A1', columns: [{name: 'A'}, {name: 'B'}], rowCount: 2});
  s2.mergeCells('D5:E5');
  const xml = partsOf(disjoint)['xl/worksheets/sheet1.xml'] as string;
  assert.match(xml, /<mergeCells count="1"><mergeCell ref="D5:E5"\/><\/mergeCells>/);
});

test('<tableParts> follows <headerFooter> in the worksheet element order', () => {
  const wb = new Workbook();
  const s = wb.addWorksheet('S');
  s.getCell('A1').value = 'x';
  s.headerFooter.oddHeader = 'H';
  s.addTable({name: 'T', ref: 'A1', columns: [{name: 'A'}], rowCount: 1});
  const xml = partsOf(wb)['xl/worksheets/sheet1.xml'] as string;
  assert.ok(
    xml.indexOf('<headerFooter') < xml.indexOf('<tableParts'),
    'tableParts must follow headerFooter per CT_Worksheet'
  );
});

test('an unprotected sheet emits no <sheetProtection> element', () => {
  const wb = new Workbook();
  wb.addWorksheet('S').getCell('A1').value = 'x';
  const xml = partsOf(wb)['xl/worksheets/sheet1.xml'] as string;
  assert.doesNotMatch(xml, /<sheetProtection/);
});

test('protecting a sheet emits a self-closing <sheetProtection sheet="1"> after <sheetData>', () => {
  const wb = new Workbook();
  const s = wb.addWorksheet('S');
  s.getCell('A1').value = 'x';
  s.protect();
  const xml = partsOf(wb)['xl/worksheets/sheet1.xml'] as string;
  assert.match(xml, /<sheetProtection sheet="1"\/>/);
  assert.ok(
    xml.indexOf('</sheetData>') < xml.indexOf('<sheetProtection'),
    'sheetProtection must follow sheetData per CT_Worksheet'
  );
});

test('an unprotected password derives an OOXML-agile credential onto <sheetProtection>', () => {
  const wb = new Workbook();
  const s = wb.addWorksheet('S');
  s.getCell('A1').value = 'x';
  s.protect('secret');
  const xml = partsOf(wb)['xl/worksheets/sheet1.xml'] as string;
  assert.match(xml, /algorithmName="SHA-512"/);
  assert.match(xml, /hashValue="[^"]+"/);
  assert.match(xml, /saltValue="[^"]+"/);
  assert.match(xml, /spinCount="100000"/);
  assert.match(xml, /sheet="1"/);
});

test('protection flags invert to OOXML "forbidden" booleans, writing only non-default values', () => {
  const wb = new Workbook();
  const s = wb.addWorksheet('S');
  s.getCell('A1').value = 'x';
  // sort defaults to forbidden under protection, so allowing it must emit sort="0".
  // selectLockedCells defaults to permitted, so forbidding it must emit selectLockedCells="1".
  s.protect(undefined, {sort: true, autoFilter: true, selectLockedCells: false});
  const xml = partsOf(wb)['xl/worksheets/sheet1.xml'] as string;
  assert.match(xml, /sort="0"/);
  assert.match(xml, /autoFilter="0"/);
  assert.match(xml, /selectLockedCells="1"/);
});

test('a flag left at its OOXML default is omitted from <sheetProtection>', () => {
  const wb = new Workbook();
  const s = wb.addWorksheet('S');
  s.getCell('A1').value = 'x';
  // Allowing selection (its default) and forbidding sort (its default) are both no-ops.
  s.protect(undefined, {selectLockedCells: true, sort: false});
  const xml = partsOf(wb)['xl/worksheets/sheet1.xml'] as string;
  const tag = (xml.match(/<sheetProtection[^>]*\/>/) as RegExpMatchArray)[0];
  assert.doesNotMatch(tag, /selectLockedCells=/);
  assert.doesNotMatch(tag, /sort=/);
  assert.match(tag, /sheet="1"/);
});

test('unprotect() removes a sheet-protection element previously set', () => {
  const wb = new Workbook();
  const s = wb.addWorksheet('S');
  s.getCell('A1').value = 'x';
  s.protect('pw');
  s.unprotect();
  const xml = partsOf(wb)['xl/worksheets/sheet1.xml'] as string;
  assert.doesNotMatch(xml, /<sheetProtection/);
});
