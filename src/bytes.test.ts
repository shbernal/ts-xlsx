import assert from 'node:assert/strict';
import {test} from 'node:test';

import {concat, toBase64} from './bytes.ts';

test('chunks join in order, empties and all', () => {
  assert.deepEqual(
    concat([Uint8Array.of(1, 2), new Uint8Array(), Uint8Array.of(3)]),
    Uint8Array.of(1, 2, 3),
  );
  assert.deepEqual(concat([]), new Uint8Array());
});

test('a known total is taken on trust, not recomputed', () => {
  const chunks = [Uint8Array.of(1, 2), Uint8Array.of(3)];
  assert.deepEqual(concat(chunks, 3), Uint8Array.of(1, 2, 3));
  // A total larger than the chunks leaves the tail zeroed rather than reallocating: the callers that
  // pass one read it off a length-prefixed header, which is the authority on the part's size.
  assert.deepEqual(concat(chunks, 5), Uint8Array.of(1, 2, 3, 0, 0));
});

test('a lone chunk is handed back rather than copied', () => {
  const only = Uint8Array.of(7, 8, 9);
  assert.equal(concat([only]), only, 'the common inflate case allocates nothing');
  assert.equal(concat([only], 3), only);
});

test('a lone chunk that is a view keeps its own bounds, not the buffer it views', () => {
  const backing = Uint8Array.of(1, 2, 3, 4, 5);
  assert.deepEqual(concat([backing.subarray(1, 3)]), Uint8Array.of(2, 3));
});

// `Buffer` is the oracle, not the implementation: `toBase64` exists so that a browser-safe module
// never has to reach for that Node global, and a test file runs in Node by definition.
test('base64 matches Buffer for every remainder, including the padded tails', () => {
  for (let length = 0; length <= 32; length++) {
    const bytes = Uint8Array.from({length}, (_, i) => (i * 37 + 11) & 0xff);
    assert.equal(toBase64(bytes), Buffer.from(bytes).toString('base64'), `length ${length}`);
  }
});

test('base64 covers the whole alphabet, high bytes and all', () => {
  const bytes = Uint8Array.from({length: 256}, (_, i) => i);
  assert.equal(toBase64(bytes), Buffer.from(bytes).toString('base64'));
});
