import assert from 'node:assert/strict';
import {test} from 'node:test';

import {strFromU8, strToU8, unzipSync, zipSync} from 'fflate';

import {inflatePackage} from './inflate.ts';

/** A deterministic, near-incompressible byte pattern — large enough to span several input
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

test('inflatePackage agrees with fflate on a well-formed archive', () => {
  const archive = zipSync({a: strToU8('one'), b: noise(3000)});
  const ours = inflatePackage(archive, GENEROUS_CAP);
  const theirs = unzipSync(archive);
  assert.deepEqual(Object.keys(ours).sort(), Object.keys(theirs).sort());
  for (const name of Object.keys(theirs)) {
    assert.deepEqual(ours[name], theirs[name], name);
  }
});
