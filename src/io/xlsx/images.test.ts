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
    'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==',
  ),
  (c) => c.charCodeAt(0),
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
  assert.ok(!names.some((n) => /drawing\d+\.xml$/.test(n)));
  assert.ok(!names.some((n) => n.startsWith('xl/media/')));
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
  wb.addWorksheet('A').addImage(id, {
    tl: {col: 1, row: 1},
    br: {col: 4, row: 6},
    editAs: 'absolute',
  });
  wb.addWorksheet('B').addImage(id, {tl: {col: 1, row: 1}, br: {col: 4, row: 6}});
  const files = unzipSync(writeXlsx(wb));
  assert.match(
    strFromU8(files['xl/drawings/drawing1.xml'] as Uint8Array),
    /<xdr:twoCellAnchor editAs="absolute">/,
  );
  assert.match(
    strFromU8(files['xl/drawings/drawing2.xml'] as Uint8Array),
    /<xdr:twoCellAnchor editAs="oneCell">/,
  );
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

test('a dirty or missing image extension is sanitised to a well-formed media name', () => {
  const wb = new Workbook();
  const ws = wb.addWorksheet('S');
  const dirty = wb.addImage({buffer: ONE_PX_PNG, extension: '.png?alt=media&token=abc'});
  const missing = wb.addImage({buffer: ONE_PX_PNG});
  ws.addImage(dirty, {tl: {col: 0, row: 0}, br: {col: 1, row: 1}});
  ws.addImage(missing, {tl: {col: 2, row: 2}, br: {col: 3, row: 3}});
  const files = unzipSync(writeXlsx(wb));
  const media = Object.keys(files).filter((n) => n.startsWith('xl/media/'));
  assert.deepStrictEqual(media.sort(), ['xl/media/image1.png', 'xl/media/image2.png']);
  const contentTypes = strFromU8(files['[Content_Types].xml'] as Uint8Array);
  const defaults = [...contentTypes.matchAll(/<Default Extension="([^"]*)"/g)].map((m) => m[1]);
  assert.ok(
    defaults.every((e) => /^[A-Za-z0-9]+$/.test(e as string)),
    'every Default extension is a bare token',
  );
  assert.doesNotMatch(contentTypes, /image\/undefined/, 'no bogus media type');
  assert.strictEqual(readXlsx(writeXlsx(wb)).getWorksheet('S')?.images.length, 2);
});

test('removeImage drops exactly the targeted anchor and omits its now-orphaned media', () => {
  const wb = new Workbook();
  const ws = wb.addWorksheet('S');
  const id1 = wb.addImage({buffer: ONE_PX_PNG, extension: 'png'});
  const id2 = wb.addImage({buffer: ONE_PX_PNG, extension: 'png'});
  ws.addImage(id1, {tl: {col: 0, row: 0}, br: {col: 1, row: 1}});
  ws.addImage(id2, {tl: {col: 2, row: 0}, br: {col: 3, row: 1}});
  ws.removeImage(id1);
  assert.deepStrictEqual(
    ws.images.map((i) => i.imageId),
    [id2],
  );
  const media = Object.keys(unzipSync(writeXlsx(wb))).filter((n) => n.startsWith('xl/media/'));
  assert.strictEqual(media.length, 1, 'the orphaned image is not written');
});

test('a fractional anchor floors to its cell with a sub-cell offset scaled by column width', () => {
  const fromCol = (width: number): {col: number; colOff: number} => {
    const wb = new Workbook();
    const ws = wb.addWorksheet('S');
    ws.getColumn(4).width = width;
    const id = wb.addImage({buffer: ONE_PX_PNG, extension: 'png'});
    ws.addImage(id, {tl: {col: 3.5, row: 0}, br: {col: 5, row: 2}});
    const anchor = ws.images[0]?.anchor;
    assert.ok(anchor && !isOneCellAnchor(anchor));
    return {col: anchor.from.col, colOff: anchor.from.colOff ?? 0};
  };
  const narrow = fromCol(5);
  const wide = fromCol(50);
  assert.strictEqual(narrow.col, 3, 'col 3.5 floors to cell column 3');
  assert.ok(
    narrow.colOff > 0 && wide.colOff > narrow.colOff,
    'a wider column yields a larger offset',
  );
});

test('a picture rotation survives the write/read round-trip on a rot-only transform', () => {
  const wb = new Workbook();
  const ws = wb.addWorksheet('S');
  const id = wb.addImage({buffer: ONE_PX_PNG, extension: 'png'});
  ws.addImageAnchor(id, {from: {col: 1, row: 1}, to: {col: 3, row: 3}, rotation: 2700000});
  const drawing = strFromU8(unzipSync(writeXlsx(wb))['xl/drawings/drawing1.xml'] as Uint8Array);
  assert.match(drawing, /<a:xfrm rot="2700000"\/>/);
  assert.doesNotMatch(drawing, /<a:off|<a:ext/, 'the rot rides alone, no zeroed offset/extent');
  assert.strictEqual(
    readXlsx(writeXlsx(wb)).getWorksheet('S')?.images[0]?.anchor.rotation,
    2700000,
  );
});

test('a sheet background image writes a <picture>, an image relationship, and its media, and round-trips', () => {
  const wb = new Workbook();
  const ws = wb.addWorksheet('S');
  const id = wb.addImage({buffer: ONE_PX_PNG, extension: 'png'});
  ws.addBackgroundImage(id);
  const files = unzipSync(writeXlsx(wb));
  const sheetXml = strFromU8(files['xl/worksheets/sheet1.xml'] as Uint8Array);
  const picture = sheetXml.match(/<picture r:id="([^"]+)"\/>/);
  assert.ok(picture, 'the sheet references a background picture');
  const relsXml = strFromU8(files['xl/worksheets/_rels/sheet1.xml.rels'] as Uint8Array);
  assert.match(
    relsXml,
    new RegExp(
      `<Relationship Id="${picture![1]}"[^>]*Type="[^"]*/image"[^>]*Target="\\.\\./media/image1\\.png"`,
    ),
  );
  assert.ok(files['xl/media/image1.png'], 'the background bytes are written once');
  const back = readXlsx(writeXlsx(wb)).getWorksheet('S');
  assert.strictEqual(
    back?.backgroundImageId !== undefined,
    true,
    'the background survives the round-trip',
  );
  assert.deepStrictEqual(
    back && readXlsx(writeXlsx(wb)).getImage(back.backgroundImageId ?? -1)?.data,
    ONE_PX_PNG,
  );
});

test('a background image, a note, and an anchored image on one sheet keep unique relationship ids', () => {
  const wb = new Workbook();
  const ws = wb.addWorksheet('S');
  ws.getCell('B2').note = 'a note';
  const anchoredId = wb.addImage({buffer: ONE_PX_PNG, extension: 'png'});
  ws.addImage(anchoredId, {tl: {col: 0, row: 0}, br: {col: 1, row: 1}});
  const bgId = wb.addImage({buffer: ONE_PX_PNG, extension: 'png'});
  ws.addBackgroundImage(bgId);
  const files = unzipSync(writeXlsx(wb));
  const relsXml = strFromU8(files['xl/worksheets/_rels/sheet1.xml.rels'] as Uint8Array);
  const ids = [...relsXml.matchAll(/<Relationship Id="([^"]+)"/g)].map((m) => m[1]);
  assert.strictEqual(new Set(ids).size, ids.length, 'no two worksheet relationships share an id');
  // The note (comments + VML), the drawing, and the background all resolve to distinct part classes.
  const types = [...relsXml.matchAll(/Type="[^"]*\/(\w+)"/g)].map((m) => m[1]);
  assert.ok(types.includes('image'), 'the background rides an image relationship');
  assert.ok(types.includes('drawing'), 'the anchored image rides a drawing relationship');
  assert.ok(
    types.includes('comments') && types.includes('vmlDrawing'),
    'the note rides comments + VML',
  );
});

test('removeBackgroundImage clears the background and omits its now-orphaned media', () => {
  const wb = new Workbook();
  const ws = wb.addWorksheet('S');
  const id = wb.addImage({buffer: ONE_PX_PNG, extension: 'png'});
  ws.addBackgroundImage(id);
  ws.removeBackgroundImage();
  assert.strictEqual(ws.backgroundImageId, undefined);
  const files = unzipSync(writeXlsx(wb));
  assert.ok(
    !Object.keys(files).some((n) => n.startsWith('xl/media/')),
    'the orphaned background is not written',
  );
  assert.doesNotMatch(strFromU8(files['xl/worksheets/sheet1.xml'] as Uint8Array), /<picture\b/);
});

test('one image anchored on two sheets is stored as a single media part', () => {
  const wb = new Workbook();
  const id = wb.addImage({buffer: ONE_PX_PNG, extension: 'png'});
  wb.addWorksheet('A').addImage(id, {tl: {col: 0, row: 0}, br: {col: 1, row: 1}});
  wb.addWorksheet('B').addImage(id, {tl: {col: 3, row: 3}, br: {col: 4, row: 4}});
  const files = unzipSync(writeXlsx(wb));
  const mediaParts = Object.keys(files).filter((n) => n.startsWith('xl/media/'));
  assert.strictEqual(mediaParts.length, 1, 'the shared image is written once');
});
