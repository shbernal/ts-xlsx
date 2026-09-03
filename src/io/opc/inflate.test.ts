import assert from 'node:assert/strict';
import {test} from 'node:test';

import {strFromU8, strToU8, unzipSync, zipSync} from 'fflate';

import {PackageReadError} from './errors.ts';
import {inflatePackage} from './inflate.ts';

/** A deterministic, near-incompressible byte pattern, large enough to span several input
 *  slices once zipped, so the multi-chunk reassembly path is exercised. */
function noise(length: number): Uint8Array {
  const out = new Uint8Array(length);
  let state = 0x2545f491;
  for (let i = 0; i < length; i += 1) {
    state ^= state << 13;
    state ^= state >>> 17;
    state ^= state << 5;
    out[i] = state & 0xff;
  }
  return out;
}

const GENEROUS_CAP = 64 * 1024 * 1024;

test('a deflated package inflates to its parts', () => {
  const archive = zipSync({
    'a.txt': strToU8('hello'),
    'nested/b.xml': strToU8('<x/>'),
  });
  const files = inflatePackage(archive, GENEROUS_CAP);
  assert.equal(strFromU8(files['a.txt'] as Uint8Array), 'hello');
  assert.equal(strFromU8(files['nested/b.xml'] as Uint8Array), '<x/>');
});

test('a stored (uncompressed) entry passes through unchanged', () => {
  const bytes = noise(4096);
  const archive = zipSync({'raw.bin': bytes}, {level: 0});
  const files = inflatePackage(archive, GENEROUS_CAP);
  assert.deepEqual(files['raw.bin'], bytes);
});

test('a part whose data spans several input slices is reassembled byte-for-byte', () => {
  const bytes = noise(200 * 1024);
  const files = inflatePackage(zipSync({'big.bin': bytes}), GENEROUS_CAP);
  assert.deepEqual(files['big.bin'], bytes);
});

test('the running counter rejects output that exceeds the cap', () => {
  // 1 MiB of zeros compresses to a fraction of a kilobyte but inflates well past a 4 KiB cap.
  const archive = zipSync({'bomb.bin': new Uint8Array(1024 * 1024)});
  assert.throws(() => inflatePackage(archive, 4096), PackageReadError);
  assert.throws(() => inflatePackage(archive, 4096), /possible zip bomb/);
});

test('a package under the cap inflates without complaint', () => {
  const archive = zipSync({'ok.bin': new Uint8Array(1024 * 1024)});
  const files = inflatePackage(archive, 2 * 1024 * 1024);
  assert.equal((files['ok.bin'] as Uint8Array).length, 1024 * 1024);
});

test('a header that lies small about its uncompressed size is still bounded by real output', () => {
  // The bound must consult *produced* bytes, never the archive's declared size. Forge the
  // local header's uncompressed-size field down to one byte: a declared-size filter would
  // wave this through, but the true 1 MiB of output must still trip the counter.
  const archive = zipSync({'liar.bin': new Uint8Array(1024 * 1024)});
  assert.deepEqual(
    [...archive.subarray(0, 4)],
    [0x50, 0x4b, 0x03, 0x04],
    'local file header at offset 0',
  );
  const view = new DataView(archive.buffer, archive.byteOffset, archive.byteLength);
  view.setUint32(22, 1, true); // uncompressed size → 1
  assert.equal(view.getUint32(22, true), 1, 'the forged header now declares a single byte');

  assert.throws(() => inflatePackage(archive, 64 * 1024), /possible zip bomb/);
});

test('an unsupported compression method is rejected, not silently dropped', () => {
  const archive = zipSync({'weird.bin': strToU8('data')});
  const view = new DataView(archive.buffer, archive.byteOffset, archive.byteLength);
  view.setUint16(8, 99, true); // compression method → an unknown value
  assert.throws(() => inflatePackage(archive, GENEROUS_CAP), /unknown compression/);
});

// ── Hand-assembled archives ─────────────────────────────────────────────────────────────────────────
// `zipSync` takes an object, so it cannot express two entries under one name or a name no OPC package
// may carry. These are written out byte by byte instead: stored (method 0) entries, a central
// directory, and an end-of-central-directory record.

const CRC_TABLE = Uint32Array.from({length: 256}, (_, byte) => {
  let value = byte;
  for (let bit = 0; bit < 8; bit++) value = value & 1 ? 0xedb8_8320 ^ (value >>> 1) : value >>> 1;
  return value >>> 0;
});

function crc32(bytes: Uint8Array): number {
  let value = 0xffff_ffff;
  for (const byte of bytes) value = (CRC_TABLE[(value ^ byte) & 0xff] as number) ^ (value >>> 8);
  return (value ^ 0xffff_ffff) >>> 0;
}

function storedZip(entries: readonly {name: string; data: Uint8Array}[]): Uint8Array {
  const local: number[] = [];
  const central: number[] = [];
  const u16 = (n: number): number[] => [n & 0xff, (n >> 8) & 0xff];
  const u32 = (n: number): number[] => [...u16(n & 0xffff), ...u16((n >>> 16) & 0xffff)];

  for (const {name, data} of entries) {
    const nameBytes = [...strToU8(name)];
    const shared = [
      ...u16(0), // flags
      ...u16(0), // method: stored
      ...u16(0), // modified time
      ...u16(0), // modified date
      ...u32(crc32(data)),
      ...u32(data.length), // compressed size
      ...u32(data.length), // uncompressed size
      ...u16(nameBytes.length),
    ];
    const localHeaderAt = local.length;
    local.push(...u32(0x0403_4b50), ...u16(20), ...shared, ...u16(0), ...nameBytes, ...data);
    central.push(
      ...u32(0x0201_4b50),
      ...u16(20), // version made by
      ...u16(20), // version needed
      ...shared,
      ...u16(0), // extra length
      ...u16(0), // comment length
      ...u16(0), // disk number
      ...u16(0), // internal attributes
      ...u32(0), // external attributes
      ...u32(localHeaderAt),
      ...nameBytes,
    );
  }

  return Uint8Array.from([
    ...local,
    ...central,
    ...u32(0x0605_4b50),
    ...u16(0),
    ...u16(0),
    ...u16(entries.length),
    ...u16(entries.length),
    ...u32(central.length),
    ...u32(local.length),
    ...u16(0), // comment length
  ]);
}

test('the hand-assembled archive builder produces something this reader accepts', () => {
  const files = inflatePackage(
    storedZip([
      {name: 'one.xml', data: strToU8('<a/>')},
      {name: 'sub/two.xml', data: strToU8('<b/>')},
    ]),
    GENEROUS_CAP,
  );
  assert.deepEqual(Object.keys(files).sort(), ['one.xml', 'sub/two.xml']);
  assert.equal(strFromU8(files['sub/two.xml'] as Uint8Array), '<b/>');
});

test('two parts under one name are refused rather than one of them being chosen', () => {
  // Walking the central directory takes the first, streaming the local headers takes the last, and a
  // package that carries both is a package that means different things to different readers.
  const archive = storedZip([
    {name: 'xl/worksheets/sheet1.xml', data: strToU8('<first/>')},
    {name: 'xl/worksheets/sheet1.xml', data: strToU8('<second/>')},
  ]);
  assert.throws(() => inflatePackage(archive, GENEROUS_CAP), {
    name: 'PackageReadError',
    message: /duplicate part "xl\/worksheets\/sheet1.xml"/,
  });
});

test('an entry name that cannot be a part name is refused, not normalised', () => {
  for (const name of ['/xl/workbook.xml', 'xl\\workbook.xml', '../outside.xml', 'C:/outside.xml']) {
    assert.throws(
      () => inflatePackage(storedZip([{name, data: strToU8('<x/>')}]), GENEROUS_CAP),
      {name: 'PackageReadError', message: /illegal part name/},
      name,
    );
  }
});

test('inflatePackage agrees with fflate on a well-formed archive', () => {
  const archive = zipSync({a: strToU8('one'), b: noise(3000)});
  const ours = inflatePackage(archive, GENEROUS_CAP);
  const theirs = unzipSync(archive);
  assert.deepEqual(Object.keys(ours).sort(), Object.keys(theirs).sort());
  for (const name of Object.keys(theirs)) {
    assert.deepEqual(ours[name], theirs[name], name);
  }
});
