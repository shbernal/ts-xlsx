import assert from 'node:assert/strict';
import {test} from 'node:test';

import {write} from './lanes.ts';
import {listParts, prettyXml} from './package.ts';
import {findSample} from './samples.ts';

function partsOf(id: string): ReturnType<typeof listParts> {
  const sample = findSample(id);
  assert.ok(sample !== undefined, id);
  const written = write(sample.build());
  assert.ok(written.ok);
  return listParts(written.value.bytes);
}

test('the part tree holds what an xlsx package must contain', () => {
  const paths = partsOf('values').map((entry) => entry.path);
  for (const required of [
    '[Content_Types].xml',
    '_rels/.rels',
    'xl/workbook.xml',
    'xl/_rels/workbook.xml.rels',
    'xl/worksheets/sheet1.xml',
  ]) {
    assert.ok(paths.includes(required), `${required} is missing from ${paths.join(', ')}`);
  }
});

test('the part tree is sorted by path, not left in zip order', () => {
  const paths = partsOf('styles').map((entry) => entry.path);
  assert.deepEqual(paths, [...paths].sort());
});

test('a text part carries its text and a size that matches its bytes', () => {
  const workbookPart = partsOf('values').find((entry) => entry.path === 'xl/workbook.xml');
  assert.ok(workbookPart !== undefined);
  assert.ok(workbookPart.text !== undefined);
  assert.match(workbookPart.text, /<workbook/);
  assert.equal(workbookPart.byteLength, new TextEncoder().encode(workbookPart.text).length);
});

test('a part with no text extension is reported as a size and nothing else', () => {
  // Nothing this playground writes is binary today, so the classification is checked on the
  // rule rather than on an artefact: an unknown extension must not be decoded as text and
  // shown as mojibake.
  const parts = partsOf('values');
  for (const part of parts) {
    const isTextPath = /\.(?:xml|rels)$/.test(part.path);
    assert.equal(part.text !== undefined, isTextPath, part.path);
  }
});

test('prettyXml puts one element on each line and indents by depth', () => {
  assert.equal(
    prettyXml('<a><b><c/></b></a>'),
    ['<a>', '  <b>', '    <c/>', '  </b>', '</a>'].join('\n'),
  );
});

test('prettyXml keeps an element with only text on one line', () => {
  assert.equal(prettyXml('<si><t>hello</t></si>'), ['<si>', '  <t>hello</t>', '</si>'].join('\n'));
});

test('prettyXml does not touch the text it moves', () => {
  // `xml:space="preserve"` is how a cell holding spaces says so. A pretty-printer that
  // trimmed here would show a different workbook from the one in the file.
  const xml = '<t xml:space="preserve">  padded  </t>';
  assert.equal(prettyXml(xml), xml);
});

test('prettyXml leaves the declaration at the top level', () => {
  const printed = prettyXml('<?xml version="1.0"?><root><leaf/></root>');
  assert.equal(printed.split('\n')[0], '<?xml version="1.0"?>');
  assert.equal(printed.split('\n')[1], '<root>');
});

test('prettyXml survives a stray closing tag rather than indenting into the negative', () => {
  assert.equal(prettyXml('</lonely>'), '</lonely>');
  assert.equal(prettyXml(''), '');
});
