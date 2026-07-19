import assert from 'node:assert/strict';
import {readFileSync, rmSync} from 'node:fs';
import {tmpdir} from 'node:os';
import {join} from 'node:path';
import {Duplex, PassThrough} from 'node:stream';
import {test} from 'node:test';

import {strFromU8, unzipSync} from 'fflate';

import {isOneCellAnchor} from '../../core/image.ts';
import {readXlsx} from './read.ts';
import {WorkbookStreamWriter} from './write-stream.ts';

// Decode one package part back to its XML text — the streamed archive is a real zip, so read it the
// same way a consumer would.
function partText(pkg: Uint8Array, name: string): string {
  const bytes = unzipSync(pkg)[name];
  assert.ok(bytes, `expected part ${name}`);
  return strFromU8(bytes);
}

// A 1×1 transparent PNG — enough bytes to prove the streamed media round-trips verbatim.
const ONE_PX_PNG = Uint8Array.from(
  atob(
    'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg=='
  ),
  c => c.charCodeAt(0)
);

// Drain a Node readable into one buffer, resolving once it ends.
function drain(stream: NodeJS.ReadableStream): Promise<Uint8Array> {
  return new Promise((resolve, reject) => {
    const chunks: Buffer[] = [];
    stream.on('data', (chunk: Buffer) => chunks.push(chunk));
    stream.on('end', () => resolve(Buffer.concat(chunks)));
    stream.on('error', reject);
  });
}

test('a streamed workbook reloads to the same sheet names and cell values as a whole-file write', async () => {
  const writer = new WorkbookStreamWriter();
  const sheet = writer.addWorksheet('S');
  for (let i = 1; i <= 20; i++) sheet.addRow([`r${i}`, i]).commit();
  sheet.commit();
  const bytes = await writer.commit();

  const workbook = readXlsx(bytes);
  assert.deepEqual(
    workbook.worksheets.map(s => s.name),
    ['S']
  );
  const reread = workbook.getWorksheet('S');
  assert.ok(reread);
  assert.equal(reread.getCell('A1').value, 'r1');
  assert.equal(reread.getCell('B1').value, 1);
  assert.equal(reread.getCell('A20').value, 'r20');
});

test('addRows appends a batch identically to adding rows one at a time', async () => {
  const writer = new WorkbookStreamWriter();
  const sheet = writer.addWorksheet('S');
  sheet.addRows([
    ['a', 1],
    ['b', 2],
  ]);
  sheet.commit();
  const workbook = readXlsx(await writer.commit());
  const reread = workbook.getWorksheet('S');
  assert.ok(reread);
  assert.equal(reread.rowCount, 2);
  assert.equal(reread.getCell('A1').value, 'a');
  assert.equal(reread.getCell('B1').value, 1);
  assert.equal(reread.getCell('A2').value, 'b');
  assert.equal(reread.getCell('B2').value, 2);
});

test('a row added to a committed sheet is rejected with a legible "already committed" error', () => {
  const writer = new WorkbookStreamWriter();
  const sheet = writer.addWorksheet('S');
  sheet.addRow(['a']).commit();
  sheet.commit();
  assert.throws(() => sheet.addRow(['b']), /already committed/i);
  assert.throws(() => sheet.getCell('A2'), /already committed/i);
});

test('writer.stream.pipe(dest) returns dest and delivers the whole package', async () => {
  const writer = new WorkbookStreamWriter();
  const sink = new PassThrough();
  const drained = drain(sink);
  const returned = writer.stream.pipe(sink);
  assert.equal(returned, sink, 'pipe(dest) must return dest so .pipe(out).on(...) composes');

  const sheet = writer.addWorksheet('S');
  sheet.addRow(['a', 'b']).commit();
  sheet.commit();
  const committed = await writer.commit();

  const piped = await drained;
  assert.deepEqual(Uint8Array.from(piped), committed, 'the piped bytes match the committed package');
  const reread = readXlsx(piped).getWorksheet('S');
  assert.ok(reread);
  assert.equal(reread.getCell('A1').value, 'a');
});

test('fullCalcOnLoad set through calcProperties is emitted; unset it is absent', async () => {
  const withFlag = new WorkbookStreamWriter();
  withFlag.calcProperties.fullCalcOnLoad = true;
  withFlag.addWorksheet('S').getCell('A1').value = 1;
  const flagged = await withFlag.commit();

  const without = new WorkbookStreamWriter();
  without.addWorksheet('S').getCell('A1').value = 1;
  const plain = await without.commit();

  assert.match(partText(flagged, 'xl/workbook.xml'), /fullCalcOnLoad="1"/);
  assert.doesNotMatch(partText(plain, 'xl/workbook.xml'), /fullCalcOnLoad/);
});

test('useSharedStrings pools streamed string cells into a shared table that reads back intact', async () => {
  const writer = new WorkbookStreamWriter({useSharedStrings: true});
  const sheet = writer.addWorksheet('S');
  sheet.addRow(['dup']).commit();
  sheet.addRow(['dup']).commit();
  sheet.addRow(['other']).commit();
  sheet.commit();
  const bytes = await writer.commit();

  const sst = partText(bytes, 'xl/sharedStrings.xml');
  assert.match(sst, /uniqueCount="2"/);
  assert.match(partText(bytes, 'xl/worksheets/sheet1.xml'), /t="s"><v>0<\/v>/);

  const reread = readXlsx(bytes).getWorksheet('S');
  assert.ok(reread);
  assert.equal(reread.getCell('A1').value, 'dup');
  assert.equal(reread.getCell('A3').value, 'other');
});

test('a streamed workbook without the option keeps strings inline and writes no shared table', async () => {
  const writer = new WorkbookStreamWriter();
  writer.addWorksheet('S').addRow(['inline']).commit();
  const bytes = await writer.commit();

  assert.throws(() => partText(bytes, 'xl/sharedStrings.xml'), /expected part/);
  assert.match(partText(bytes, 'xl/worksheets/sheet1.xml'), /t="inlineStr"/);
});

test('commit over a caller-supplied PassThrough sink resolves and delivers a valid package', async () => {
  const sink = new PassThrough();
  const chunks: Buffer[] = [];
  sink.on('data', (chunk: Buffer) => chunks.push(chunk));

  const writer = new WorkbookStreamWriter({stream: sink});
  writer.addWorksheet('S').addRow(['a', 'b']).commit();
  const bytes = await writer.commit();

  const delivered = Buffer.concat(chunks);
  assert.ok(delivered.length > 0, 'the sink received bytes');
  assert.deepEqual(Uint8Array.from(delivered), bytes, 'the sink got exactly the committed package');
  const reread = readXlsx(delivered).getWorksheet('S');
  assert.ok(reread);
  assert.equal(reread.getCell('A1').value, 'a');
});

test('commit over a Duplex sink resolves — completion does not depend on the writer owning the stream', async () => {
  const chunks: Buffer[] = [];
  const sink = new Duplex({
    read() {},
    write(chunk: Buffer, _enc, cb) {
      chunks.push(chunk);
      cb();
    },
  });

  const writer = new WorkbookStreamWriter({stream: sink});
  writer.addWorksheet('S').addRow(['a']).commit();
  await writer.commit();

  const reread = readXlsx(Buffer.concat(chunks)).getWorksheet('S');
  assert.ok(reread);
  assert.equal(reread.getCell('A1').value, 'a');
});

test('commit to a valid filename writes a re-openable package to disk', async () => {
  const target = join(tmpdir(), `ts-xlsx-stream-${process.pid}.xlsx`);
  try {
    const writer = new WorkbookStreamWriter({filename: target});
    writer.addWorksheet('S').addRow(['a']).commit();
    await writer.commit();

    const reread = readXlsx(readFileSync(target)).getWorksheet('S');
    assert.ok(reread);
    assert.equal(reread.getCell('A1').value, 'a');
  } finally {
    rmSync(target, {force: true});
  }
});

test('commit to an unopenable filename rejects with the underlying I/O error rather than hanging', async () => {
  // A path whose parent directory does not exist cannot be opened for writing; the write stream errors
  // on a later tick and commit must surface it.
  const badPath = join(tmpdir(), 'ts-xlsx-no-such-dir', `${'x'.repeat(300)}.xlsx`);
  const writer = new WorkbookStreamWriter({filename: badPath});
  writer.addWorksheet('S').addRow(['a']).commit();

  await assert.rejects(writer.commit(), (err: NodeJS.ErrnoException) => {
    assert.match(String(err.code ?? err.message), /ENOENT|ENAMETOOLONG/);
    return true;
  });
});

test('supplying both a stream and a filename is rejected at construction', () => {
  assert.throws(
    () => new WorkbookStreamWriter({stream: new PassThrough(), filename: 'out.xlsx'}),
    /either a stream or a filename/i
  );
});

test('shared-formula slave cells authored on the stream reload populated, not empty', async () => {
  const writer = new WorkbookStreamWriter();
  const sheet = writer.addWorksheet('yua');
  for (let i = 1; i <= 10; i++) sheet.getCell(`A${i}`).value = i * 10;
  sheet.getCell('B1').value = {formula: 'A1*2', result: 20};
  for (let j = 2; j <= 10; j++) sheet.getCell(`B${j}`).value = {sharedFormula: 'B1'};
  sheet.commit();

  const reread = readXlsx(await writer.commit()).getWorksheet('yua');
  assert.ok(reread);
  const master = reread.getCell('B1').value;
  assert.ok(master && typeof master === 'object' && 'formula' in master, 'master is a formula cell');
  const slave = reread.getCell('B3').value;
  assert.ok(slave && typeof slave === 'object', 'slave reloads as a value object, not empty');
  assert.ok('sharedFormula' in slave || 'formula' in slave, 'slave carries a (shared) formula');
});

test('a streamed sheet emits <conditionalFormatting> before <hyperlinks>, per the CT_Worksheet sequence', async () => {
  const writer = new WorkbookStreamWriter();
  const sheet = writer.addWorksheet('S');
  sheet.getCell('A1').value = {text: 'link', hyperlink: 'https://example.com'};
  sheet.addConditionalFormatting({
    ref: 'A1:A10',
    rules: [{type: 'expression', formulae: ['MOD(ROW(),2)=0'], style: {fill: {type: 'pattern', pattern: 'solid', bgColor: {argb: 'FFEEEEEE'}}}}],
  });
  sheet.addRow(['x']).commit();
  sheet.commit();

  const xml = partText(await writer.commit(), 'xl/worksheets/sheet1.xml');
  const posCf = xml.indexOf('<conditionalFormatting');
  const posHl = xml.indexOf('<hyperlinks');
  assert.ok(posCf >= 0 && posHl >= 0, 'both blocks are present');
  assert.ok(posCf < posHl, 'conditionalFormatting precedes hyperlinks');
});

test('a streamed sheet emits <dataValidations> before <hyperlinks>, per the CT_Worksheet sequence', async () => {
  const writer = new WorkbookStreamWriter();
  const sheet = writer.addWorksheet('S');
  sheet.getCell('A1').value = {text: 'link', hyperlink: 'https://example.com'};
  sheet.addDataValidation('B1', {type: 'list', allowBlank: true, formulae: ['"x,y,z"']});
  sheet.addRow(['r']).commit();
  sheet.commit();

  const xml = partText(await writer.commit(), 'xl/worksheets/sheet1.xml');
  const posDv = xml.indexOf('<dataValidations');
  const posHl = xml.indexOf('<hyperlinks');
  assert.ok(posDv >= 0 && posHl >= 0, 'both blocks are present');
  assert.ok(posDv < posHl, 'dataValidations precedes hyperlinks');
});

test('streamed conditional formatting and data validations reload through the tolerant reader', async () => {
  const writer = new WorkbookStreamWriter();
  const sheet = writer.addWorksheet('S');
  sheet.getCell('A1').value = {text: 'link', hyperlink: 'https://example.com'};
  sheet.addDataValidation('B1', {type: 'list', allowBlank: true, formulae: ['"x,y,z"']});
  sheet.addConditionalFormatting({ref: 'A1:A10', rules: [{type: 'cellIs', operator: 'greaterThan', formulae: [3], priority: 1}]});
  sheet.addRow(['r']).commit();
  sheet.commit();

  const reread = readXlsx(await writer.commit()).getWorksheet('S');
  assert.ok(reread);
  assert.equal(reread.dataValidations.length, 1, 'the data validation survives');
  assert.equal(reread.conditionalFormattings.length, 1, 'the conditional formatting survives');
  assert.equal(reread.conditionalFormattings[0]?.rules[0]?.type, 'cellIs');
});

test('authoring conditional formatting or a data validation on a committed streamed sheet is rejected legibly', async () => {
  const writer = new WorkbookStreamWriter();
  const sheet = writer.addWorksheet('S');
  sheet.commit();
  assert.throws(() => sheet.addConditionalFormatting({ref: 'A1', rules: []}), /already committed/);
  assert.throws(() => sheet.addDataValidation('A1', {type: 'list', formulae: ['"a"']}), /already committed/);
  await writer.commit();
});

test('an image anchored on the stream reloads with its anchor and bytes intact', async () => {
  const writer = new WorkbookStreamWriter();
  const sheet = writer.addWorksheet('S');
  const id = writer.addImage({buffer: ONE_PX_PNG, extension: 'png'});
  sheet.addImage(id, {tl: {col: 0, row: 5}, br: {col: 2, row: 8}});
  sheet.commit();

  const bytes = await writer.commit();
  const reloaded = readXlsx(bytes);
  const [image] = reloaded.getWorksheet('S')?.images ?? [];
  const anchor = image?.anchor;
  assert.ok(anchor && !isOneCellAnchor(anchor));
  assert.deepStrictEqual(anchor.from, {col: 0, row: 5, colOff: 0, rowOff: 0});
  assert.deepStrictEqual(anchor.to, {col: 2, row: 8, colOff: 0, rowOff: 0});
  const media = reloaded.getImage(image?.imageId ?? -1);
  assert.strictEqual(media?.extension, 'png');
  assert.deepStrictEqual(media?.data, ONE_PX_PNG, 'the streamed media bytes survive verbatim');
});

test('a streamed image emits the drawing, media, and <drawing> reference like a buffered write', async () => {
  const writer = new WorkbookStreamWriter();
  const sheet = writer.addWorksheet('S');
  const id = writer.addImage({buffer: ONE_PX_PNG, extension: 'png'});
  sheet.addImage(id, {tl: {col: 0, row: 0}, br: {col: 1, row: 1}});
  sheet.commit();

  const bytes = await writer.commit();
  const files = unzipSync(bytes);
  assert.ok(files['xl/drawings/drawing1.xml'], 'a drawing part is streamed');
  assert.ok(files['xl/media/image1.png'], 'the media bytes are streamed');
  assert.match(partText(bytes, 'xl/worksheets/sheet1.xml'), /<drawing r:id="[^"]+"\/>/);
  const contentTypes = partText(bytes, '[Content_Types].xml');
  assert.match(contentTypes, /Extension="png" ContentType="image\/png"/);
  assert.match(contentTypes, /PartName="\/xl\/drawings\/drawing1\.xml"/);
});

test('one streamed image anchored on two sheets is stored as a single media part', async () => {
  const writer = new WorkbookStreamWriter();
  const id = writer.addImage({buffer: ONE_PX_PNG, extension: 'png'});
  writer.addWorksheet('A').addImage(id, {tl: {col: 0, row: 0}, br: {col: 1, row: 1}});
  writer.addWorksheet('B').addImage(id, {tl: {col: 3, row: 3}, br: {col: 4, row: 4}});

  const files = unzipSync(await writer.commit());
  const mediaParts = Object.keys(files).filter(n => n.startsWith('xl/media/'));
  assert.strictEqual(mediaParts.length, 1, 'the shared image is streamed once');
});

test('registering an image on a committed streamed workbook is rejected legibly', async () => {
  const writer = new WorkbookStreamWriter();
  writer.addWorksheet('S').commit();
  await writer.commit();
  assert.throws(() => writer.addImage({buffer: ONE_PX_PNG, extension: 'png'}), /already committed/);
});

test('a streamed sheet carrying both protection and an autofilter emits them in CT_Worksheet order', async () => {
  const writer = new WorkbookStreamWriter();
  const sheet = writer.addWorksheet('S');
  sheet.addRow(['H1', 'H2']).commit();
  sheet.addRow(['a', 'b']).commit();
  sheet.autoFilter = 'A1:B1';
  sheet.protect('pw', {});
  sheet.commit();

  const bytes = await writer.commit();
  const xml = partText(bytes, 'xl/worksheets/sheet1.xml');
  const posProtection = xml.indexOf('<sheetProtection');
  const posAutoFilter = xml.indexOf('<autoFilter');
  assert.ok(posProtection >= 0, 'the streamed sheet carries sheetProtection');
  assert.ok(posAutoFilter >= 0, 'the streamed sheet carries autoFilter');
  assert.ok(posProtection < posAutoFilter, '<sheetProtection> precedes <autoFilter> per the schema');
  // The whole package still reloads, and the autofilter survives.
  assert.deepEqual(readXlsx(bytes).getWorksheet('S')?.autoFilter, {ref: 'A1:B1', columns: []});
});

test('the streamed sheet autoFilter getter reflects what was set', () => {
  const sheet = new WorkbookStreamWriter().addWorksheet('S');
  assert.strictEqual(sheet.autoFilter, undefined);
  sheet.autoFilter = 'A1:C3';
  assert.deepEqual(sheet.autoFilter, {ref: 'A1:C3', columns: []});
});

test('setting protection or an autofilter on a committed streamed sheet is rejected legibly', () => {
  const sheet = new WorkbookStreamWriter().addWorksheet('S');
  sheet.commit();
  assert.throws(() => sheet.protect('pw', {}), /already committed/);
  assert.throws(() => {
    sheet.autoFilter = 'A1:B2';
  }, /already committed/);
});

test('committing a streamed row evicts its cells from the model, bounding peak memory', async () => {
  const writer = new WorkbookStreamWriter();
  const sheet = writer.addWorksheet('S');
  const row = sheet.addRow(['a', 'b']);
  assert.ok(sheet.model.hasCell(1, 1), 'the cell is live while the row is being styled');
  row.commit();
  assert.equal(sheet.model.hasCell(1, 1), false, 'commit released the cell graph');
  assert.equal(sheet.model.hasCell(1, 2), false);
  // rowCount survives the eviction, so the next append lands below the evicted row rather than reusing
  // its number — the correctness hazard that makes owning the row counter necessary.
  assert.equal(sheet.rowCount, 1);
  sheet.addRow(['c']).commit();
  const reread = readXlsx(await writer.commit()).getWorksheet('S');
  assert.ok(reread);
  assert.equal(reread.getCell('A1').value, 'a');
  assert.equal(reread.getCell('A2').value, 'c', 'the second append did not collide with the evicted first');
});

test('a fully committed streamed sheet holds no live cells at commit — every row was evicted', async () => {
  const writer = new WorkbookStreamWriter();
  const sheet = writer.addWorksheet('S');
  for (let i = 1; i <= 50; i++) sheet.addRow([i, `v${i}`]).commit();
  for (let i = 1; i <= 50; i++) assert.equal(sheet.model.hasCell(i, 1), false, `row ${i} was freed`);
  assert.equal(sheet.rowCount, 50, 'rowCount still reflects every appended row');
  const reread = readXlsx(await writer.commit()).getWorksheet('S');
  assert.ok(reread);
  assert.equal(reread.getCell('A50').value, 50);
  assert.equal(reread.getCell('B1').value, 'v1');
});

test('with useSharedStrings a streamed row stays live until commit — the shared pool defeats bounding', async () => {
  const writer = new WorkbookStreamWriter({useSharedStrings: true});
  const sheet = writer.addWorksheet('S');
  sheet.addRow(['x']).commit();
  assert.ok(sheet.model.hasCell(1, 1), 'shared-strings mode keeps the row live rather than evicting it');
  const reread = readXlsx(await writer.commit()).getWorksheet('S');
  assert.equal(reread?.getCell('A1').value, 'x');
});

test('committing a row that carries a shared-formula cell is rejected legibly', () => {
  const sheet = new WorkbookStreamWriter().addWorksheet('S');
  const row = sheet.addRow([{sharedFormula: 'A1'}]);
  assert.throws(() => row.commit(), /shared-formula/);
});

test('committing a streamed row twice does not duplicate it in the sheet', async () => {
  const writer = new WorkbookStreamWriter();
  const sheet = writer.addWorksheet('S');
  const row = sheet.addRow(['only']);
  row.commit();
  row.commit();
  const xml = partText(await writer.commit(), 'xl/worksheets/sheet1.xml');
  assert.equal((xml.match(/<row /g) ?? []).length, 1, 'the row appears exactly once despite the double commit');
});

test('a styled row committed eagerly round-trips its fill through the shared style registry', async () => {
  const writer = new WorkbookStreamWriter();
  const sheet = writer.addWorksheet('S');
  const row = sheet.addRow(['tinted']);
  const cell = row.cells[0];
  assert.ok(cell);
  cell.fill = {type: 'pattern', pattern: 'solid', fgColor: {argb: 'FFFF0000'}};
  row.commit();

  const fill = readXlsx(await writer.commit()).getWorksheet('S')?.getCell('A1').fill;
  assert.ok(fill && fill.type === 'pattern' && fill.pattern === 'solid', 'the eagerly-flushed row kept its fill');
});

test('a live getCell row and a committed appended row serialise in ascending order and both reload', async () => {
  const writer = new WorkbookStreamWriter();
  const sheet = writer.addWorksheet('S');
  sheet.getCell('A1').value = 'live'; // row 1, never committed → stays in the model
  sheet.addRow(['flushed']).commit(); // row 2, serialised and evicted
  const bytes = await writer.commit();

  const xml = partText(bytes, 'xl/worksheets/sheet1.xml');
  assert.ok(xml.indexOf('r="1"') < xml.indexOf('r="2"'), 'the live row 1 precedes the flushed row 2');
  const reread = readXlsx(bytes).getWorksheet('S');
  assert.ok(reread);
  assert.equal(reread.getCell('A1').value, 'live');
  assert.equal(reread.getCell('A2').value, 'flushed');
});
