// SHA-512, written out here rather than imported from `node:crypto`.
//
// This library hashes exactly one thing: a sheet-protection password (`src/core/protection.ts`),
// through OOXML's agile scheme. That is one hash on an authoring call, and it was the only reason
// `node:crypto` appeared in the module graph reachable from the package entry, which is a
// bundler's problem long before it is a runtime one, because a bundler resolves imports rather than
// call graphs and so pulled Node's crypto into every browser build whether or not anything ever
// called it (`docs/knowledge/specs/browser-safe-io-boundary.md`).
//
// Web Crypto would have been the obvious replacement and is not usable here: `crypto.subtle.digest`
// is asynchronous, and the agile scheme feeds each digest into the next for a hundred thousand
// rounds, so an async primitive would make `Worksheet.protect()` async for every caller: an API
// cost far larger than the one this module pays. Salt generation does use the platform
// (`crypto.getRandomValues`), which is synchronous and present in browsers and Node alike.
//
// The 64-bit words are carried as hi/lo `Uint32` pairs because JavaScript has no fast 64-bit
// integer: BigInt is the honest spelling and is several times slower over this loop, which is the
// entire cost of the operation. `sha512.test.ts` checks every digest here against Node's, so the
// two cannot drift, and states what the default 100000-spin derivation measures at.

const BLOCK_BYTES = 128;
const DIGEST_BYTES = 64;

// The first 64 bits of the fractional parts of the cube roots of the first 80 primes (FIPS 180-4
// §4.2.3), hi word then lo word.
const K = new Uint32Array([
  0x428a2f98, 0xd728ae22, 0x71374491, 0x23ef65cd, 0xb5c0fbcf, 0xec4d3b2f, 0xe9b5dba5, 0x8189dbbc,
  0x3956c25b, 0xf348b538, 0x59f111f1, 0xb605d019, 0x923f82a4, 0xaf194f9b, 0xab1c5ed5, 0xda6d8118,
  0xd807aa98, 0xa3030242, 0x12835b01, 0x45706fbe, 0x243185be, 0x4ee4b28c, 0x550c7dc3, 0xd5ffb4e2,
  0x72be5d74, 0xf27b896f, 0x80deb1fe, 0x3b1696b1, 0x9bdc06a7, 0x25c71235, 0xc19bf174, 0xcf692694,
  0xe49b69c1, 0x9ef14ad2, 0xefbe4786, 0x384f25e3, 0x0fc19dc6, 0x8b8cd5b5, 0x240ca1cc, 0x77ac9c65,
  0x2de92c6f, 0x592b0275, 0x4a7484aa, 0x6ea6e483, 0x5cb0a9dc, 0xbd41fbd4, 0x76f988da, 0x831153b5,
  0x983e5152, 0xee66dfab, 0xa831c66d, 0x2db43210, 0xb00327c8, 0x98fb213f, 0xbf597fc7, 0xbeef0ee4,
  0xc6e00bf3, 0x3da88fc2, 0xd5a79147, 0x930aa725, 0x06ca6351, 0xe003826f, 0x14292967, 0x0a0e6e70,
  0x27b70a85, 0x46d22ffc, 0x2e1b2138, 0x5c26c926, 0x4d2c6dfc, 0x5ac42aed, 0x53380d13, 0x9d95b3df,
  0x650a7354, 0x8baf63de, 0x766a0abb, 0x3c77b2a8, 0x81c2c92e, 0x47edaee6, 0x92722c85, 0x1482353b,
  0xa2bfe8a1, 0x4cf10364, 0xa81a664b, 0xbc423001, 0xc24b8b70, 0xd0f89791, 0xc76c51a3, 0x0654be30,
  0xd192e819, 0xd6ef5218, 0xd6990624, 0x5565a910, 0xf40e3585, 0x5771202a, 0x106aa070, 0x32bbd1b8,
  0x19a4c116, 0xb8d2d0c8, 0x1e376c08, 0x5141ab53, 0x2748774c, 0xdf8eeb99, 0x34b0bcb5, 0xe19b48a8,
  0x391c0cb3, 0xc5c95a63, 0x4ed8aa4a, 0xe3418acb, 0x5b9cca4f, 0x7763e373, 0x682e6ff3, 0xd6b2b8a3,
  0x748f82ee, 0x5defb2fc, 0x78a5636f, 0x43172f60, 0x84c87814, 0xa1f0ab72, 0x8cc70208, 0x1a6439ec,
  0x90befffa, 0x23631e28, 0xa4506ceb, 0xde82bde9, 0xbef9a3f7, 0xb2c67915, 0xc67178f2, 0xe372532b,
  0xca273ece, 0xea26619c, 0xd186b8c7, 0x21c0c207, 0xeada7dd6, 0xcde0eb1e, 0xf57d4f7f, 0xee6ed178,
  0x06f067aa, 0x72176fba, 0x0a637dc5, 0xa2c898a6, 0x113f9804, 0xbef90dae, 0x1b710b35, 0x131c471b,
  0x28db77f5, 0x23047d84, 0x32caab7b, 0x40c72493, 0x3c9ebe0a, 0x15c9bebc, 0x431d67c4, 0x9c100d4c,
  0x4cc5d4be, 0xcb3e42b6, 0x597f299c, 0xfc657e2a, 0x5fcb6fab, 0x3ad6faec, 0x6c44198c, 0x4a475817,
]);

// The first 64 bits of the fractional parts of the square roots of the first 8 primes (§5.3.5).
const INITIAL = new Uint32Array([
  0x6a09e667, 0xf3bcc908, 0xbb67ae85, 0x84caa73b, 0x3c6ef372, 0xfe94f82b, 0xa54ff53a, 0x5f1d36f1,
  0x510e527f, 0xade682d1, 0x9b05688c, 0x2b3e6c1f, 0x1f83d9ab, 0xfb41bd6b, 0x5be0cd19, 0x137e2179,
]);

// `noUncheckedIndexedAccess` types every typed-array read as `number | undefined`, and the non-null
// assertion that would answer it is banned in this tree. Every index below is a loop counter
// bounded by the array's own length, so the fallback is unreachable; naming it once keeps the
// arithmetic legible instead of scattering `?? 0` through eighty rounds.
function word(words: Uint32Array, index: number): number {
  return words[index] ?? 0;
}

/** The SHA-512 digest of `message`, as its 64 bytes. */
export function sha512(message: Uint8Array): Uint8Array {
  // One 0x80 byte, then zeroes, then the 16-byte big-endian bit length on the block boundary.
  const blocks = Math.ceil((message.length + 17) / BLOCK_BYTES);
  const padded = new Uint8Array(blocks * BLOCK_BYTES);
  padded.set(message);
  padded[message.length] = 0x80;
  const view = new DataView(padded.buffer);
  // A `Uint8Array` cannot be long enough for the length field's high half to be non-zero, so only
  // its low 64 bits are ever written.
  const bits = message.length * 8;
  view.setUint32(padded.length - 8, Math.floor(bits / 0x1_0000_0000), false);
  view.setUint32(padded.length - 4, bits >>> 0, false);

  const h = Uint32Array.from(INITIAL);
  const schedule = new Uint32Array(160);
  for (let block = 0; block < blocks; block++) compress(h, view, block * BLOCK_BYTES, schedule);

  const digest = new Uint8Array(DIGEST_BYTES);
  const out = new DataView(digest.buffer);
  for (let i = 0; i < 16; i++) out.setUint32(i * 4, word(h, i), false);
  return digest;
}

/** Mix one 128-byte block of `input` into the running state `h`, using `schedule` as scratch. */
function compress(h: Uint32Array, input: DataView, offset: number, schedule: Uint32Array): void {
  for (let t = 0; t < 16; t++) {
    schedule[t * 2] = input.getUint32(offset + t * 8, false);
    schedule[t * 2 + 1] = input.getUint32(offset + t * 8 + 4, false);
  }
  for (let t = 16; t < 80; t++) {
    const xHi = word(schedule, (t - 15) * 2);
    const xLo = word(schedule, (t - 15) * 2 + 1);
    // σ0 = rotr(x, 1) ^ rotr(x, 8) ^ shr(x, 7)
    const s0Hi = ((xHi >>> 1) | (xLo << 31)) ^ ((xHi >>> 8) | (xLo << 24)) ^ (xHi >>> 7);
    const s0Lo =
      ((xLo >>> 1) | (xHi << 31)) ^ ((xLo >>> 8) | (xHi << 24)) ^ ((xLo >>> 7) | (xHi << 25));
    const yHi = word(schedule, (t - 2) * 2);
    const yLo = word(schedule, (t - 2) * 2 + 1);
    // σ1 = rotr(y, 19) ^ rotr(y, 61) ^ shr(y, 6)
    const s1Hi = ((yHi >>> 19) | (yLo << 13)) ^ ((yLo >>> 29) | (yHi << 3)) ^ (yHi >>> 6);
    const s1Lo =
      ((yLo >>> 19) | (yHi << 13)) ^ ((yHi >>> 29) | (yLo << 3)) ^ ((yLo >>> 6) | (yHi << 26));
    const lo =
      (s1Lo >>> 0) +
      word(schedule, (t - 7) * 2 + 1) +
      (s0Lo >>> 0) +
      word(schedule, (t - 16) * 2 + 1);
    const hi =
      (s1Hi >>> 0) +
      word(schedule, (t - 7) * 2) +
      (s0Hi >>> 0) +
      word(schedule, (t - 16) * 2) +
      Math.floor(lo / 0x1_0000_0000);
    schedule[t * 2] = hi >>> 0;
    schedule[t * 2 + 1] = lo >>> 0;
  }

  let aHi = word(h, 0);
  let aLo = word(h, 1);
  let bHi = word(h, 2);
  let bLo = word(h, 3);
  let cHi = word(h, 4);
  let cLo = word(h, 5);
  let dHi = word(h, 6);
  let dLo = word(h, 7);
  let eHi = word(h, 8);
  let eLo = word(h, 9);
  let fHi = word(h, 10);
  let fLo = word(h, 11);
  let gHi = word(h, 12);
  let gLo = word(h, 13);
  let hHi = word(h, 14);
  let hLo = word(h, 15);

  for (let t = 0; t < 80; t++) {
    // Σ1 = rotr(e, 14) ^ rotr(e, 18) ^ rotr(e, 41)
    const bigS1Hi =
      ((eHi >>> 14) | (eLo << 18)) ^ ((eHi >>> 18) | (eLo << 14)) ^ ((eLo >>> 9) | (eHi << 23));
    const bigS1Lo =
      ((eLo >>> 14) | (eHi << 18)) ^ ((eLo >>> 18) | (eHi << 14)) ^ ((eHi >>> 9) | (eLo << 23));
    // ch = (e & f) ^ (~e & g)
    const chHi = (eHi & fHi) ^ (~eHi & gHi);
    const chLo = (eLo & fLo) ^ (~eLo & gLo);
    const t1Lo =
      hLo + (bigS1Lo >>> 0) + (chLo >>> 0) + word(K, t * 2 + 1) + word(schedule, t * 2 + 1);
    const t1Hi =
      (hHi +
        (bigS1Hi >>> 0) +
        (chHi >>> 0) +
        word(K, t * 2) +
        word(schedule, t * 2) +
        Math.floor(t1Lo / 0x1_0000_0000)) >>>
      0;
    const t1LoWrapped = t1Lo >>> 0;

    // Σ0 = rotr(a, 28) ^ rotr(a, 34) ^ rotr(a, 39)
    const bigS0Hi =
      ((aHi >>> 28) | (aLo << 4)) ^ ((aLo >>> 2) | (aHi << 30)) ^ ((aLo >>> 7) | (aHi << 25));
    const bigS0Lo =
      ((aLo >>> 28) | (aHi << 4)) ^ ((aHi >>> 2) | (aLo << 30)) ^ ((aHi >>> 7) | (aLo << 25));
    // maj = (a & b) ^ (a & c) ^ (b & c)
    const majHi = (aHi & bHi) ^ (aHi & cHi) ^ (bHi & cHi);
    const majLo = (aLo & bLo) ^ (aLo & cLo) ^ (bLo & cLo);
    const t2Lo = (bigS0Lo >>> 0) + (majLo >>> 0);
    const t2Hi = ((bigS0Hi >>> 0) + (majHi >>> 0) + Math.floor(t2Lo / 0x1_0000_0000)) >>> 0;
    const t2LoWrapped = t2Lo >>> 0;

    hHi = gHi;
    hLo = gLo;
    gHi = fHi;
    gLo = fLo;
    fHi = eHi;
    fLo = eLo;
    const nextELo = dLo + t1LoWrapped;
    eHi = (dHi + t1Hi + Math.floor(nextELo / 0x1_0000_0000)) >>> 0;
    eLo = nextELo >>> 0;
    dHi = cHi;
    dLo = cLo;
    cHi = bHi;
    cLo = bLo;
    bHi = aHi;
    bLo = aLo;
    const nextALo = t1LoWrapped + t2LoWrapped;
    aHi = (t1Hi + t2Hi + Math.floor(nextALo / 0x1_0000_0000)) >>> 0;
    aLo = nextALo >>> 0;
  }

  addInto(h, 0, aHi, aLo);
  addInto(h, 2, bHi, bLo);
  addInto(h, 4, cHi, cLo);
  addInto(h, 6, dHi, dLo);
  addInto(h, 8, eHi, eLo);
  addInto(h, 10, fHi, fLo);
  addInto(h, 12, gHi, gLo);
  addInto(h, 14, hHi, hLo);
}

/** `h[index..index + 1] += (hi, lo)`, as one 64-bit addition with the hi word first. */
function addInto(h: Uint32Array, index: number, hi: number, lo: number): void {
  const sumLo = word(h, index + 1) + lo;
  h[index] = (word(h, index) + hi + Math.floor(sumLo / 0x1_0000_0000)) >>> 0;
  h[index + 1] = sumLo >>> 0;
}
