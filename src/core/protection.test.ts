import assert from 'node:assert/strict';
import {createHash} from 'node:crypto';
import {test} from 'node:test';

import {deriveCredential} from './protection.ts';

// An independent re-implementation of OOXML's agile hash, used to prove deriveCredential's output
// is genuinely derived from its own reported salt via the documented algorithm — catching a wrong
// concat order or a big-endian iteration counter, which a self-comparison could not.
function referenceHash(password: string, saltBase64: string, spinCount: number): string {
  const salt = Buffer.from(saltBase64, 'base64');
  const secret = Buffer.from(password, 'utf16le');
  let hash = createHash('sha512')
    .update(Buffer.concat([salt, secret]))
    .digest();
  const iteration = Buffer.alloc(4);
  for (let i = 0; i < spinCount; i++) {
    iteration.writeUInt32LE(i, 0);
    hash = createHash('sha512')
      .update(Buffer.concat([hash, iteration]))
      .digest();
  }
  return hash.toString('base64');
}

test('deriveCredential emits SHA-512, a 16-byte salt, a 64-byte hash, and the default spin count', () => {
  const credential = deriveCredential('secret');
  assert.equal(credential.algorithmName, 'SHA-512');
  assert.equal(credential.spinCount, 100000);
  assert.equal(Buffer.from(credential.saltValue, 'base64').length, 16);
  assert.equal(Buffer.from(credential.hashValue, 'base64').length, 64);
});

test('the hash is genuinely derived from the reported salt via the OOXML-agile algorithm', () => {
  const credential = deriveCredential('correct horse', 25);
  assert.equal(credential.hashValue, referenceHash('correct horse', credential.saltValue, 25));
});

test('a custom spin count is honored', () => {
  const credential = deriveCredential('pw', 7);
  assert.equal(credential.spinCount, 7);
  assert.equal(credential.hashValue, referenceHash('pw', credential.saltValue, 7));
});

test('the salt is real randomness — two derivations of one password differ', () => {
  const a = deriveCredential('same', 5);
  const b = deriveCredential('same', 5);
  assert.notEqual(a.saltValue, b.saltValue);
  assert.notEqual(a.hashValue, b.hashValue);
});
