import assert from 'node:assert/strict';
import {createHash, randomBytes} from 'node:crypto';
import {test} from 'node:test';

import {sha512} from './sha512.ts';

// `node:crypto` is the oracle here, not the implementation: this module exists so that nothing on
// the package's entry graph imports it (see `sha512.ts`), and a test file is free to, because it
// runs in Node by definition. Comparing against it is what keeps a hand-written hash honest.
function reference(data: Uint8Array): string {
  return createHash('sha512').update(data).digest('hex');
}

function hex(bytes: Uint8Array): string {
  return [...bytes].map((byte) => byte.toString(16).padStart(2, '0')).join('');
}

test('the FIPS 180-4 sample vectors', () => {
  assert.equal(
    hex(sha512(new TextEncoder().encode(''))),
    'cf83e1357eefb8bdf1542850d66d8007d620e4050b5715dc83f4a921d36ce9ce' +
      '47d0d13c5d85f2b0ff8318d2877eec2f63b931bd47417a81a538327af927da3e',
  );
  assert.equal(
    hex(sha512(new TextEncoder().encode('abc'))),
    'ddaf35a193617abacc417349ae20413112e6fa4e89a97ea20a9eeee64b55d39a' +
      '2192992a274fc1a836ba3c23a3feebbd454d4423643ce80e2a9ac94fa54ca49f',
  );
});

// The padding rule is where a hash implementation goes wrong quietly: a message of 111, 112 or 113
// bytes is the boundary at which the 16-byte length field stops fitting in the final block and one
// more block is owed. Each length here is checked against Node's answer for the same bytes.
test('every message length around a block boundary agrees with node:crypto', () => {
  for (const length of [0, 1, 55, 63, 64, 65, 110, 111, 112, 113, 127, 128, 129, 255, 256, 1000]) {
    const data = new Uint8Array(randomBytes(length));
    assert.equal(hex(sha512(data)), reference(data), `length ${length}`);
  }
});

test('random messages agree with node:crypto', () => {
  for (let i = 0; i < 200; i++) {
    const data = new Uint8Array(randomBytes(Math.floor(Math.random() * 4096)));
    assert.equal(hex(sha512(data)), reference(data), hex(data).slice(0, 32));
  }
});
