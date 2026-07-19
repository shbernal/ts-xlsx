import assert from 'node:assert/strict';
import {test} from 'node:test';

import {strFromU8, strToU8, unzipSync, zipSync} from 'fflate';

import {Workbook} from '../../core/workbook.ts';
import {readXlsx} from './read.ts';
import {writeXlsx} from './write.ts';

function sheetXml(pkg: Uint8Array): string {
  return strFromU8(unzipSync(pkg)['xl/worksheets/sheet1.xml'] ?? new Uint8Array());
}

function stylesXml(pkg: Uint8Array): string {
  return strFromU8(unzipSync(pkg)['xl/styles.xml'] ?? new Uint8Array());
}

// Read an xlsx built from a hand-authored sheet1 (and optionally styles) part, so a case can feed the
// reader markup the writer itself only produces on round-trip — an Excel-authored x14 extLst block, a
// foreign dxf table.
function readParts(parts: {sheet1?: string; styles?: string}): Workbook {
  const base = new Workbook();
  base.addWorksheet('S');
  const files = unzipSync(writeXlsx(base));
  if (parts.sheet1 !== undefined) files['xl/worksheets/sheet1.xml'] = strToU8(parts.sheet1);
  if (parts.styles !== undefined) files['xl/styles.xml'] = strToU8(parts.styles);
  return readXlsx(zipSync(files));
}

test('a dataBar rule emits a well-formed block with two cfvo anchors and a bar colour', () => {
  const workbook = new Workbook();
  workbook.addWorksheet('S').addConditionalFormatting({
    ref: 'A1:A3',
    rules: [
      {
        type: 'dataBar',
        color: {argb: 'FF638EC6'},
        cfvo: [
          {type: 'num', value: 0},
          {type: 'num', value: 1},
        ],
      },
    ],
  });
  const xml = sheetXml(writeXlsx(workbook));
  const block = xml.match(/<conditionalFormatting[\s\S]*?<\/conditionalFormatting>/)?.[0] ?? '';

  assert.equal([...xml.matchAll(/<conditionalFormatting\b/g)].length, 1, 'one block');
  assert.match(block, /<dataBar>/);
  assert.equal([...block.matchAll(/<cfvo\b/g)].length, 2, 'both anchors present');
  assert.match(block, /<color rgb="FF638EC6"\/>/);
});

test('a dataBar rule reads back on the same range with its type, colour, and both anchors', () => {
  const workbook = new Workbook();
  workbook.addWorksheet('S').addConditionalFormatting({
    ref: 'A1:A3',
    rules: [
      {
        type: 'dataBar',
        color: {argb: 'FF638EC6'},
        cfvo: [
          {type: 'num', value: 0},
          {type: 'num', value: 1},
        ],
      },
    ],
  });
  const rule = readXlsx(writeXlsx(workbook)).getWorksheet('S')?.conditionalFormattings[0]?.rules[0];
  assert.equal(rule?.type, 'dataBar');
  assert.equal(rule?.color?.argb, 'FF638EC6');
  assert.deepEqual(
    rule?.cfvo?.map((v) => v.value),
    [0, 1],
    'both numeric anchors survive in order',
  );
});

test('a minimal dataBar (no cfvo, no colour) gains Excel default min/max anchors and a bar colour', () => {
  const workbook = new Workbook();
  workbook
    .addWorksheet('S')
    .addConditionalFormatting({ref: 'A1:A3', rules: [{type: 'dataBar', priority: 1}]});
  const block = sheetXml(writeXlsx(workbook)).match(/<dataBar[\s\S]*?<\/dataBar>/)?.[0] ?? '';

  assert.equal(
    [...block.matchAll(/<cfvo\b/g)].length,
    2,
    'a default data bar carries a min and a max cfvo',
  );
  assert.match(block, /<cfvo type="min"\/>/);
  assert.match(block, /<cfvo type="max"\/>/);
  assert.match(block, /<color\b/, 'a default bar colour is emitted');
});

test('a rule over a multi-area ref emits one block whose sqref lists every area', () => {
  const workbook = new Workbook();
  workbook.addWorksheet('S').addConditionalFormatting({
    ref: 'A1:C1 A3:C3 A5:C5',
    rules: [
      {
        type: 'colorScale',
        cfvo: [{type: 'min'}, {type: 'max'}],
        colors: [{argb: 'FFFF0000'}, {argb: 'FF00FF00'}],
      },
    ],
  });
  const xml = sheetXml(writeXlsx(workbook));

  assert.equal(
    [...xml.matchAll(/<conditionalFormatting\b/g)].length,
    1,
    'one block, not one per area',
  );
  assert.equal(xml.match(/<conditionalFormatting\b[^>]*sqref="([^"]*)"/)?.[1], 'A1:C1 A3:C3 A5:C5');
});

test('a stopIfTrue rule serialises the flag, references a dxf for its style, and both survive a read', () => {
  const workbook = new Workbook();
  workbook.addWorksheet('S').addConditionalFormatting({
    ref: 'A1:A10',
    rules: [
      {
        type: 'cellIs',
        operator: 'greaterThan',
        formulae: [3],
        stopIfTrue: true,
        style: {fill: {type: 'pattern', pattern: 'solid', bgColor: {argb: 'FFFF0000'}}},
      },
    ],
  });
  const pkg = writeXlsx(workbook);
  const xml = sheetXml(pkg);

  assert.match(
    xml,
    /<cfRule type="cellIs" dxfId="0" priority="1" stopIfTrue="1" operator="greaterThan">/,
  );
  assert.match(
    stylesXml(pkg),
    /<dxfs count="1"><dxf><fill><patternFill patternType="solid"><bgColor rgb="FFFF0000"\/>/,
  );

  const rule = readXlsx(pkg).getWorksheet('S')?.conditionalFormattings[0]?.rules[0];
  assert.equal(rule?.stopIfTrue, true, 'the flag reads back');
  assert.equal(rule?.dxfId, '0', 'the differential-style reference reads back');
});

test('an expression rule with no formula serialises without throwing', () => {
  const workbook = new Workbook();
  workbook
    .addWorksheet('S')
    .addConditionalFormatting({ref: 'A1:A10', rules: [{type: 'expression', style: {}}]});
  const xml = sheetXml(writeXlsx(workbook));
  assert.match(
    xml,
    /<cfRule type="expression"[^>]*\/>/,
    'a formula-less rule is a self-closing cfRule',
  );
});

test('a self-closing rule read from a file (a duplicateValues) round-trips its type, dxfId, and priority', () => {
  const part =
    '<?xml version="1.0"?>' +
    '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/>' +
    '<conditionalFormatting sqref="A1:A1048576">' +
    '<cfRule type="duplicateValues" dxfId="0" priority="1"/></conditionalFormatting></worksheet>';
  const xml = sheetXml(writeXlsx(readParts({sheet1: part})));

  assert.match(
    xml,
    /<cfRule type="duplicateValues" dxfId="0" priority="1"\/>/,
    'the rule survives, not dropped',
  );
  assert.equal([...xml.matchAll(/<conditionalFormatting\b/g)].length, 1, 'no empty block shell');
});

test('a foreign differential style with a custom number format round-trips verbatim, never "[object Object]"', () => {
  const code = '_(* #,##0_);_(* \\(#,##0\\);_(* &quot;-&quot;_);_(@_)';
  const styles =
    '<?xml version="1.0"?>' +
    '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
    '<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>' +
    '<fills count="1"><fill><patternFill patternType="none"/></fill></fills>' +
    '<borders count="1"><border/></borders>' +
    '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
    '<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>' +
    `<dxfs count="1"><dxf><numFmt numFmtId="164" formatCode="${code}"/></dxf></dxfs>` +
    '</styleSheet>';
  const sheet1 =
    '<?xml version="1.0"?>' +
    '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/>' +
    '<conditionalFormatting sqref="A1:A5">' +
    '<cfRule type="expression" dxfId="0" priority="1"><formula>A1&gt;2</formula></cfRule>' +
    '</conditionalFormatting></worksheet>';
  const out = stylesXml(writeXlsx(readParts({sheet1, styles})));

  assert.doesNotMatch(
    out,
    /\[object Object\]/,
    'the dxf numFmt is a real format code, not a coerced object',
  );
  assert.match(out, /formatCode="_\(\* #,##0_\)/, 'the exact custom format code is preserved');
});

test('an x14 extLst conditional formatting is left untouched and writing the sheet does not crash', () => {
  const sheet1 =
    '<?xml version="1.0"?>' +
    '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/>' +
    '<extLst><ext uri="{78C0D931-6437-407d-A8EE-F0AAD7539E65}" ' +
    'xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main">' +
    '<x14:conditionalFormattings><x14:conditionalFormatting ' +
    'xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">' +
    '<x14:cfRule type="expression" priority="1" id="{GUID}"><xm:f>A1&gt;2</xm:f></x14:cfRule>' +
    '<xm:sqref>A1:A5</xm:sqref></x14:conditionalFormatting></x14:conditionalFormattings></ext></extLst></worksheet>';
  const workbook = readParts({sheet1});

  // The x14 rule is namespace-prefixed and is not read into the classic model — so it is neither
  // half-parsed into a broken rule nor does it make the writer throw.
  assert.equal(
    workbook.getWorksheet('S')?.conditionalFormattings.length,
    0,
    'no classic rule was fabricated',
  );
  assert.doesNotThrow(
    () => writeXlsx(workbook),
    'writing a sheet whose CF lived only in x14 does not crash',
  );
});

function dataBarBook(rule: Record<string, unknown>): Workbook {
  const workbook = new Workbook();
  workbook.addWorksheet('S').addConditionalFormatting({
    ref: 'A1:A3',
    rules: [
      {
        type: 'dataBar',
        color: {argb: 'FF638EC6'},
        cfvo: [
          {type: 'num', value: 0},
          {type: 'num', value: 1},
        ],
        ...rule,
      },
    ],
  });
  return workbook;
}

test('a gradient dataBar writes the classic element plus a linked x14 extension carrying the flag', () => {
  const xml = sheetXml(writeXlsx(dataBarBook({gradient: true})));

  // The classic element is unchanged — every consumer still understands the bar's anchors and colour.
  assert.match(
    xml,
    /<dataBar><cfvo type="num" val="0"\/><cfvo type="num" val="1"\/><color rgb="FF638EC6"\/><\/dataBar>/,
  );
  // The cfRule links its extension by a shared id, which the worksheet x14 block echoes on its cfRule.
  const id = xml.match(/<x14:id>(\{[^}]+\})<\/x14:id>/)?.[1];
  assert.ok(id, 'the classic cfRule carries an x14:id link');
  assert.match(
    xml,
    new RegExp(`<x14:cfRule type="dataBar" id="${id.replace(/[{}]/g, '\\$&')}">`),
    'the extension echoes the same id',
  );
  assert.match(
    xml,
    /<x14:dataBar gradient="1"><x14:cfvo type="num"><xm:f>0<\/xm:f><\/x14:cfvo><x14:cfvo type="num"><xm:f>1<\/xm:f><\/x14:cfvo><\/x14:dataBar>/,
  );
  assert.match(xml, /<xm:sqref>A1:A3<\/xm:sqref>/);
});

test('the gradient flag survives a write→read round-trip through the x14 extension', () => {
  for (const gradient of [true, false]) {
    const rule = readXlsx(writeXlsx(dataBarBook({gradient}))).getWorksheet('S')
      ?.conditionalFormattings[0]?.rules[0];
    assert.equal(rule?.type, 'dataBar');
    assert.equal(rule?.gradient, gradient, `gradient=${gradient} reads back`);
    // The classic facets still survive alongside the enriched ones.
    assert.equal(rule?.color?.argb, 'FF638EC6');
    assert.deepEqual(
      rule?.cfvo?.map((v) => v.value),
      [0, 1],
    );
  }
});

test('a dataBar negative-fill and axis colour round-trip through the x14 extension', () => {
  const book = dataBarBook({negativeFillColor: {argb: 'FFFF0000'}, axisColor: {argb: 'FF000000'}});
  const xml = sheetXml(writeXlsx(book));
  assert.match(xml, /<x14:negativeFillColor rgb="FFFF0000"\/>/);
  assert.match(xml, /<x14:axisColor rgb="FF000000"\/>/);

  const rule = readXlsx(writeXlsx(book)).getWorksheet('S')?.conditionalFormattings[0]?.rules[0];
  assert.equal(rule?.negativeFillColor?.argb, 'FFFF0000');
  assert.equal(rule?.axisColor?.argb, 'FF000000');
});

test('a plain dataBar carrying no x14 facet stays classic-only, fabricating no extension', () => {
  const xml = sheetXml(writeXlsx(dataBarBook({})));
  assert.match(xml, /<dataBar>/, 'the classic element is present');
  assert.doesNotMatch(xml, /x14:dataBar/, 'no x14 data-bar extension is written');
  assert.doesNotMatch(xml, /<extLst>/, 'no extLst is fabricated for a plain data bar');
});

test('a sheet with both an extended validation and a gradient dataBar emits one worksheet extLst', () => {
  const workbook = dataBarBook({gradient: true});
  workbook
    .getWorksheet('S')
    ?.addDataValidation('B1', {type: 'list', formulae: ['Other!$A$1:$A$3']}, {extended: true});
  const xml = sheetXml(writeXlsx(workbook));

  // Both extensions ride as sibling <ext> blocks inside a single worksheet <extLst> — never two.
  assert.match(
    xml,
    /uri="\{78C0D931-6437-407d-A8EE-F0AAD7539E65\}"/,
    'the CF extension is present',
  );
  assert.match(
    xml,
    /uri="\{CCE6A557-97BC-4b89-ADB6-D9C93CAAB3DF\}"/,
    'the DV extension is present',
  );
  assert.doesNotMatch(
    xml,
    /<\/extLst><extLst>/,
    'the two extensions are not split across two worksheet extLst',
  );

  const sheet = readXlsx(writeXlsx(workbook)).getWorksheet('S');
  assert.equal(sheet?.conditionalFormattings[0]?.rules[0]?.gradient, true, 'the gradient survives');
  assert.equal(sheet?.dataValidations.length, 1, 'the extended validation survives');
});
