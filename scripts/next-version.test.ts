import assert from 'node:assert/strict';
import {test} from 'node:test';

import {nextVersion} from './next-version.ts';

const SHA = '87333ef8a1b2c3d4e5f60718293a4b5c6d7e8f90';

test('a next build leads up to the patch after the stable manifest version', () => {
  assert.equal(nextVersion('3.2.1', 12, SHA), '3.2.2-next.12.g87333ef');
});

test('a prerelease manifest keeps its version core', () => {
  assert.equal(nextVersion('4.0.0-beta.1', 3, SHA), '4.0.0-next.3.g87333ef');
});

test('the hash prefix keeps an all-digit abbreviation valid SemVer', () => {
  assert.equal(nextVersion('3.2.1', 1, '0123456789'), '3.2.2-next.1.g0123456');
});

test('invalid inputs fail before a registry publish can use them', () => {
  assert.throws(() => nextVersion('3.2', 1, SHA), /not SemVer/);
  assert.throws(() => nextVersion('3.2.1', 0, SHA), /positive integer/);
  assert.throws(() => nextVersion('3.2.1', 1, 'main'), /lowercase hex hash/);
});
