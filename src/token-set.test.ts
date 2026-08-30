import assert from 'node:assert/strict';
import {test} from 'node:test';

import {tokenSet} from './token-set.ts';

type Orientation = 'portrait' | 'landscape';

const isOrientation = tokenSet<Orientation>({portrait: true, landscape: true});

test('a member narrows and a foreign token does not', () => {
  assert.equal(isOrientation('portrait'), true);
  assert.equal(isOrientation('landscape'), true);
  assert.equal(isOrientation('default'), false);
  assert.equal(isOrientation(''), false);
});

test('an inherited property is not a member', () => {
  // The guard reads a `Set`, so the prototype chain a plain `in` would walk is out of reach: a
  // hostile file naming `toString` or `constructor` narrows to nothing.
  assert.equal(isOrientation('toString'), false);
  assert.equal(isOrientation('constructor'), false);
  assert.equal(isOrientation('__proto__'), false);
});

test('the table is read once, not on every call', () => {
  const members: Record<'a', true> = {a: true};
  const isMember = tokenSet<'a'>(members);
  Reflect.set(members, 'b', true);
  assert.equal(isMember('b'), false, 'a table mutated after the fact does not widen the guard');
});
