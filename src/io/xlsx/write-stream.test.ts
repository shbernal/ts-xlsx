import assert from 'node:assert/strict';
import {PassThrough} from 'node:stream';
import {test} from 'node:test';

import {strFromU8, unzipSync} from 'fflate';

import {readXlsx} from './read.ts';
import {WorkbookStreamWriter} from './write-stream.ts';

// Decode one package part back to its XML text — the streamed archive is a real zip, so read it the
// same way a consumer would.
function partText(pkg: Uint8Array, name: string): string {
  const bytes = unzipSync(pkg)[name];
  assert.ok(bytes, `expected part ${name}`);
  return strFromU8(bytes);
}

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
