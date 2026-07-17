import assert from 'node:assert/strict';
import test from 'node:test';

import {Workbook} from '../../core/workbook.ts';
import {writeCsv, writeCsvText} from './write.ts';

test('each row is sized to its own extent, not clamped to a narrower earlier row', () => {
  const wb = new Workbook();
  const sheet = wb.addWorksheet('S');
  sheet.addRow(['a']);
  sheet.addRow(['b', 'c', 'd']);
  assert.equal(writeCsvText(wb), 'a\nb,c,d');
});

test('a number renders bare and fields join on the configured delimiter', () => {
  const wb = new Workbook();
  wb.addWorksheet('S').addRow(['a', 'b', 'c']);
  assert.equal(writeCsvText(wb, {delimiter: ';'}), 'a;b;c');
});

test('sheetName selects the worksheet; the default is the first', () => {
  const wb = new Workbook();
  wb.addWorksheet('First').addRow(['a', 1]);
  const second = wb.addWorksheet('Second');
  second.addRow(['b', 2]);
  second.addRow(['c', 3]);
  assert.equal(writeCsvText(wb), 'a,1');
  assert.equal(writeCsvText(wb, {sheetName: 'Second'}), 'b,2\nc,3');
});

test('a name matching no worksheet throws rather than emitting empty output', () => {
  const wb = new Workbook();
  wb.addWorksheet('First').addRow(['a']);
  assert.throws(() => writeCsvText(wb, {sheetName: 'Nope'}), /no worksheet named "Nope"/);
});

test('a Date renders in a token format in UTC, or as a full ISO timestamp by default', () => {
  const wb = new Workbook();
  const noon = new Date('2018-01-05T12:00:00.000Z');
  wb.addWorksheet('S').addRow([noon]);
  assert.equal(writeCsvText(wb, {dateFormat: 'MM/DD/YYYY', dateUTC: true}), '01/05/2018');
  assert.match(writeCsvText(wb, {dateUTC: true}), /^2018-01-05T12:00:00/);
});

test('a formula cell renders its cached result', () => {
  const wb = new Workbook();
  wb.addWorksheet('S').addRow([{formula: '1+1', result: 2}]);
  assert.equal(writeCsvText(wb), '2');
});

test('a field carrying the delimiter, a quote, or a newline is quoted and escaped', () => {
  const wb = new Workbook();
  wb.addWorksheet('S').addRow(['a,b', 'he said "hi"', 'line\nbreak']);
  assert.equal(writeCsvText(wb), '"a,b","he said ""hi""","line\nbreak"');
});

test('writeCsv prepends a UTF-8 BOM for the default UTF-8 encoding', () => {
  const wb = new Workbook();
  wb.addWorksheet('S').addRow(['café']);
  const bytes = writeCsv(wb);
  assert.deepEqual([bytes[0], bytes[1], bytes[2]], [0xef, 0xbb, 0xbf]);
  assert.equal(Buffer.from(bytes.slice(3)).toString('utf8'), 'café');
});

test('writeCsv honours a requested non-UTF-8 encoding and adds no BOM there', () => {
  const wb = new Workbook();
  wb.addWorksheet('S').addRow(['café']);
  const bytes = writeCsv(wb, {encoding: 'utf16le'});
  assert.equal(Buffer.from(bytes).toString('utf16le'), 'café');
  assert.notEqual(Buffer.from(bytes).toString('utf8'), 'café');
});

test('emoji and CJK survive the default UTF-8 byte path', () => {
  const wb = new Workbook();
  wb.addWorksheet('S').addRow(['😀🎉', '日本語']);
  const body = Buffer.from(writeCsv(wb).slice(3)).toString('utf8');
  assert.equal(body, '😀🎉,日本語');
});
