import assert from 'node:assert/strict';
import {test} from 'node:test';

import {Workbook} from '../../core/workbook.ts';
import {
  elementIn,
  foreignPackage,
  partIn,
  partsWritten,
  patchParts,
  SHEET1,
} from './package.test-support.ts';
import {readXlsx} from './read.ts';
import {writeXlsx} from './write.ts';

test('an external link produces exactly one External relationship of hyperlink type', () => {
  const wb = new Workbook();
  wb.addWorksheet('S').addHyperlink({ref: 'A1', target: 'https://example.com'});

  const parts = partsWritten(wb);
  const rels = partIn(parts, 'xl/worksheets/_rels/sheet1.xml.rels');
  const external = /<Relationship\b[^>]*TargetMode="External"[^>]*\/>/g;
  assert.equal([...rels.matchAll(external)].length, 1);
  const link = elementIn(rels, external);
  assert.match(link, /Type="[^"]*\/hyperlink"/);
  assert.match(link, /Target="https:\/\/example\.com"/);
});

test('the <hyperlinks> element sits after <mergeCells> and before <pageMargins>', () => {
  const wb = new Workbook();
  const sheet = wb.addWorksheet('S');
  sheet.addHyperlink({ref: 'A1', target: 'https://example.com'});
  sheet.mergeCells('B1:C1');
  sheet.pageMargins.left = 0.5;

  const sheetXml = partIn(partsWritten(wb), 'xl/worksheets/sheet1.xml');
  const merge = sheetXml.indexOf('<mergeCells');
  const links = sheetXml.indexOf('<hyperlinks>');
  const margins = sheetXml.indexOf('<pageMargins');
  assert.ok(merge >= 0 && links >= 0 && margins >= 0);
  assert.ok(
    merge < links && links < margins,
    `order was mergeCells@${merge} hyperlinks@${links} pageMargins@${margins}`,
  );
});

test('a hyperlink spanning a range is read with its whole range', () => {
  const archive = foreignPackage({
    [SHEET1]:
      '<?xml version="1.0"?><worksheet xmlns:r="x"><sheetData>' +
      '<row r="1"><c r="D1" t="inlineStr"><is><t>go</t></is></c></row>' +
      '</sheetData><hyperlinks><hyperlink ref="D1:H1" location="Sheet1!A1"/></hyperlinks></worksheet>',
  });

  const back = readXlsx(archive).getWorksheet('S');
  assert.deepEqual(back?.hyperlinks, [{ref: 'D1:H1', target: '#Sheet1!A1'}]);
  assert.equal(back?.hyperlinkAt('F1')?.target, '#Sheet1!A1', 'every covered cell opens it');
});

test('a link naming a whole column, or nothing at all, is dropped on read and the rest are kept', () => {
  const wb = new Workbook();
  wb.addWorksheet('S').getCell('A1').value = 'x';
  const links =
    '<hyperlink ref="A:A" location="S!H1"/><hyperlink ref="junk!!" location="S!H1"/>' +
    '<hyperlink ref="B2" location="S!H1"/>';
  const back = readXlsx(
    patchParts(writeXlsx(wb), {
      [SHEET1]: (xml) =>
        xml.replace('</sheetData>', `</sheetData><hyperlinks>${links}</hyperlinks>`),
    }),
  ).getWorksheet('S');

  assert.deepEqual(back?.hyperlinks, [{ref: 'B2', target: '#S!H1'}]);
});
