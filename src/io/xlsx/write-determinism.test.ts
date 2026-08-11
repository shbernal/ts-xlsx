// Writing a workbook is a function of the workbook and nothing else.
//
// The three writers reach fflate through three different calls — `zipSync`, `zip`, and the streamed
// `Zip`/`ZipDeflate` container — and each of them defaults an entry's timestamp to `Date.now()`. One
// call site left unpinned is enough to make a regenerated deliverable churn, and the failure is
// silent: the package stays valid, reloads identically, and only shows up as a diff in someone
// else's repository. So each writer is asked here, in the same terms.

import assert from 'node:assert/strict';
import {test} from 'node:test';

import {Workbook} from '../../core/workbook.ts';
import {writeXlsx, writeXlsxAsync} from './write.ts';
import {WorkbookStreamWriter} from './write-stream.ts';

// The instant every entry must carry, spelled out rather than imported from `FIXED_ENTRY_MTIME`, so
// these tests can disagree with the constant instead of restating it. Read from the clock instead,
// this would be today — which is exactly the churn the pin exists to stop.
const EXPECTED_STAMP = {year: 2001, month: 1, day: 1, hour: 12, minute: 0, second: 0};

// A workbook with enough shape to spread across several package parts, so a stamp assertion covers
// more than one entry and a byte comparison has something to disagree about.
function severalPartWorkbook(): Workbook {
  const wb = new Workbook();
  const sheet = wb.addWorksheet('S');
  for (let row = 1; row <= 50; row++) {
    sheet.getCell(`A${row}`).value = `row ${row}`;
    sheet.getCell(`B${row}`).value = row * 3;
  }
  sheet.getCell('A1').font = {bold: true};
  sheet.getColumn(1).width = 18;
  wb.addWorksheet('T').getCell('A1').value = new Date(Date.UTC(2020, 4, 17));
  return wb;
}

test('the same workbook written twice produces byte-identical packages', () => {
  assert.deepEqual(writeXlsx(severalPartWorkbook()), writeXlsx(severalPartWorkbook()));
});

test('the async writer is deterministic too, and agrees with the buffered one', async () => {
  const [first, second] = await Promise.all([
    writeXlsxAsync(severalPartWorkbook()),
    writeXlsxAsync(severalPartWorkbook()),
  ]);
  assert.deepEqual(first, second);
  assert.deepEqual(first, writeXlsx(severalPartWorkbook()));
});

test('every entry of a buffered package carries the fixed timestamp', () => {
  const stamps = entryStamps(writeXlsx(severalPartWorkbook()));
  assert.ok(stamps.length >= 8, `expected the package to have entries, got ${stamps.length}`);
  for (const stamp of stamps) assert.deepEqual(stamp, EXPECTED_STAMP);
});

test('the streamed writer stamps its entries the same way', async () => {
  const writer = new WorkbookStreamWriter();
  const sheet = writer.addWorksheet('S');
  for (let row = 1; row <= 50; row++) sheet.addRow([`row ${row}`, row * 3]).commit();
  sheet.commit();
  const streamed = await writer.commit();

  const stamps = entryStamps(streamed);
  assert.ok(stamps.length >= 8, `expected the package to have entries, got ${stamps.length}`);
  for (const stamp of stamps) assert.deepEqual(stamp, EXPECTED_STAMP);
});

// Decode the DOS date/time dword every archive entry carries, read from the central directory. The
// directory is walked record by record rather than scanned for its signature, which deflated bytes
// can imitate by chance. The dword packs year-since-1980 / month / day / hour / minute / two-second
// tick, from the high bits down.
function entryStamps(pkg: Uint8Array): Array<Record<string, number>> {
  const view = new DataView(pkg.buffer, pkg.byteOffset, pkg.byteLength);
  // These writers emit no archive comment, so the end-of-central-directory record sits exactly 22
  // bytes from the end; its offset field says where the directory itself starts.
  const eocd = pkg.length - 22;
  assert.equal(view.getUint32(eocd, true), 0x0605_4b50, 'the archive ends with an EOCD record');
  const count = view.getUint16(eocd + 10, true);

  const stamps: Array<Record<string, number>> = [];
  let at = view.getUint32(eocd + 16, true);
  for (let i = 0; i < count; i++) {
    assert.equal(view.getUint32(at, true), 0x0201_4b50, `central directory record ${i}`);
    const dos = view.getUint32(at + 12, true);
    stamps.push({
      year: ((dos >>> 25) & 0x7f) + 1980,
      month: (dos >>> 21) & 0x0f,
      day: (dos >>> 16) & 0x1f,
      hour: (dos >>> 11) & 0x1f,
      minute: (dos >>> 5) & 0x3f,
      second: (dos & 0x1f) * 2,
    });
    const nameLen = view.getUint16(at + 28, true);
    const extraLen = view.getUint16(at + 30, true);
    const commentLen = view.getUint16(at + 32, true);
    at += 46 + nameLen + extraLen + commentLen;
  }
  return stamps;
}
