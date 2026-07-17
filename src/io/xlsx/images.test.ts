import assert from 'node:assert/strict';
import {test} from 'node:test';

import {strFromU8, unzipSync} from 'fflate';

import {isOneCellAnchor} from '../../core/image.ts';
import {Workbook} from '../../core/workbook.ts';
import {readXlsx} from './read.ts';
import {writeXlsx} from './write.ts';

// A 1×1 transparent PNG — enough bytes to prove the media round-trips verbatim.
const ONE_PX_PNG = Uint8Array.from(
  atob(
    'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg=='
  ),
  c => c.charCodeAt(0)
);

function anchored(): Workbook {
  const wb = new Workbook();
  const ws = wb.addWorksheet('S');
  const id = wb.addImage({buffer: ONE_PX_PNG, extension: 'png'});
  ws.addImage(id, {tl: {col: 0, row: 5}, br: {col: 2, row: 8}});
  return wb;
}

test('an anchored image survives the write/read round-trip with its bytes intact', () => {
  const back = readXlsx(writeXlsx(anchored())).getWorksheet('S');
  const images = back?.images ?? [];
  assert.strictEqual(images.length, 1);
  const anchor = images[0]?.anchor;
  assert.ok(anchor && !isOneCellAnchor(anchor));
  assert.deepStrictEqual(anchor.from, {col: 0, row: 5, colOff: 0, rowOff: 0});
  assert.deepStrictEqual(anchor.to, {col: 2, row: 8, colOff: 0, rowOff: 0});
  const media = readXlsx(writeXlsx(anchored())).getImage(images[0]?.imageId ?? -1);
  assert.strictEqual(media?.extension, 'png');
  assert.deepStrictEqual(media?.data, ONE_PX_PNG);
});

test('a noted-free workbook with an image emits a drawing, a media part, and a <drawing> reference', () => {
  const files = unzipSync(writeXlsx(anchored()));
  assert.ok(files['xl/drawings/drawing1.xml'], 'a drawing part is written');
  assert.ok(files['xl/media/image1.png'], 'the media bytes are written');
  const sheetXml = strFromU8(files['xl/worksheets/sheet1.xml'] as Uint8Array);
  assert.match(sheetXml, /<drawing r:id="[^"]+"\/>/);
  const drawingXml = strFromU8(files['xl/drawings/drawing1.xml'] as Uint8Array);
  assert.match(drawingXml, /<xdr:from>[\s\S]*?<xdr:row>5<\/xdr:row>/);
  const contentTypes = strFromU8(files['[Content_Types].xml'] as Uint8Array);
  assert.match(contentTypes, /Extension="png" ContentType="image\/png"/);
  assert.match(contentTypes, /PartName="\/xl\/drawings\/drawing1\.xml"/);
});

test('an image-free workbook writes no drawing or media parts', () => {
  const wb = new Workbook();
  wb.addWorksheet('S').getCell('A1').value = 'plain';
  const names = Object.keys(unzipSync(writeXlsx(wb)));
  assert.ok(!names.some(n => /drawing\d+\.xml$/.test(n)));
  assert.ok(!names.some(n => n.startsWith('xl/media/')));
});

test('inserting a row above an anchored image shifts its anchor down a row', () => {
  const wb = anchored();
  wb.getWorksheet('S')?.spliceRows(1, 0, ['inserted']);
  const back = readXlsx(writeXlsx(wb)).getWorksheet('S');
  const anchor = back?.images[0]?.anchor;
  assert.ok(anchor && !isOneCellAnchor(anchor));
  assert.strictEqual(anchor.from.row, 6, 'from-anchor row 5 shifts to 6');
  assert.strictEqual(anchor.to.row, 9, 'to-anchor row 8 shifts to 9');
  assert.strictEqual(anchor.from.col, 0, 'columns are untouched by a row splice');
});

test('a one-cell anchor emits a oneCellAnchor with a pixel extent converted to EMUs', () => {
  const wb = new Workbook();
  const ws = wb.addWorksheet('S');
  const id = wb.addImage({buffer: ONE_PX_PNG, extension: 'png'});
  ws.addImage(id, {tl: {col: 2, row: 3}, ext: {width: 191, height: 47}});
  const files = unzipSync(writeXlsx(wb));
  const drawing = strFromU8(files['xl/drawings/drawing1.xml'] as Uint8Array);
  assert.match(drawing, /<xdr:oneCellAnchor>/);
  assert.doesNotMatch(drawing, /<xdr:to>/, 'a one-cell anchor has no bottom-right point');
  assert.match(drawing, new RegExp(`<xdr:ext cx="${191 * 9525}" cy="${47 * 9525}"/>`));
});

test('a two-cell anchor honors its editAs mode and defaults to oneCell', () => {
  const wb = new Workbook();
  const id = wb.addImage({buffer: ONE_PX_PNG, extension: 'png'});
  wb.addWorksheet('A').addImage(id, {tl: {col: 1, row: 1}, br: {col: 4, row: 6}, editAs: 'absolute'});
  wb.addWorksheet('B').addImage(id, {tl: {col: 1, row: 1}, br: {col: 4, row: 6}});
  const files = unzipSync(writeXlsx(wb));
  assert.match(strFromU8(files['xl/drawings/drawing1.xml'] as Uint8Array), /<xdr:twoCellAnchor editAs="absolute">/);
  assert.match(strFromU8(files['xl/drawings/drawing2.xml'] as Uint8Array), /<xdr:twoCellAnchor editAs="oneCell">/);
});

test('an anchored picture carries no absolute spPr transform that would override the anchor', () => {
  const files = unzipSync(writeXlsx(anchored()));
  const drawing = strFromU8(files['xl/drawings/drawing1.xml'] as Uint8Array);
  assert.doesNotMatch(drawing, /<a:xfrm/, 'the geometry is the anchor, not a zeroed transform');
});

test('a one-cell anchor round-trips through the reader as an extent, not a to-point', () => {
  const wb = new Workbook();
  const ws = wb.addWorksheet('S');
  const id = wb.addImage({buffer: ONE_PX_PNG, extension: 'png'});
  ws.addImage(id, {tl: {col: 2, row: 3}, ext: {width: 191, height: 47}});
  const anchor = readXlsx(writeXlsx(wb)).getWorksheet('S')?.images[0]?.anchor;
  assert.ok(anchor && isOneCellAnchor(anchor));
  assert.deepStrictEqual(anchor.from, {col: 2, row: 3, colOff: 0, rowOff: 0});
  assert.deepStrictEqual(anchor.ext, {cx: 191 * 9525, cy: 47 * 9525});
});

test('one image anchored on two sheets is stored as a single media part', () => {
  const wb = new Workbook();
  const id = wb.addImage({buffer: ONE_PX_PNG, extension: 'png'});
  wb.addWorksheet('A').addImage(id, {tl: {col: 0, row: 0}, br: {col: 1, row: 1}});
  wb.addWorksheet('B').addImage(id, {tl: {col: 3, row: 3}, br: {col: 4, row: 4}});
  const files = unzipSync(writeXlsx(wb));
  const mediaParts = Object.keys(files).filter(n => n.startsWith('xl/media/'));
  assert.strictEqual(mediaParts.length, 1, 'the shared image is written once');
});
