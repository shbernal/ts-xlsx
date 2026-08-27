import assert from 'node:assert/strict';
import {test} from 'node:test';

import {formatBytes, formatCount, formatMillis} from './format.ts';

test('formatBytes reports whole bytes below a kibibyte', () => {
  assert.equal(formatBytes(0), '0 B');
  assert.equal(formatBytes(1), '1 B');
  assert.equal(formatBytes(1023), '1023 B');
});

test('formatBytes climbs a unit exactly at the boundary', () => {
  assert.equal(formatBytes(1024), '1.00 KiB');
  assert.equal(formatBytes(1024 * 1024), '1.00 MiB');
});

test('formatBytes drops a digit once the number is wide enough to carry it', () => {
  assert.equal(formatBytes(1024 * 9.5), '9.50 KiB');
  assert.equal(formatBytes(1024 * 42), '42.0 KiB');
});

test('formatBytes refuses to invent a figure for a value that is not one', () => {
  assert.equal(formatBytes(Number.NaN), '-');
  assert.equal(formatBytes(-1), '-');
});

test('formatMillis reports a sub-tenth measurement as a bound, never as zero', () => {
  // A write that finished faster than the clock can see must not read as "did not run",
  // and must not claim digits the browser's coarsened timer does not have.
  assert.equal(formatMillis(0), '<0.1 ms');
  assert.equal(formatMillis(0.04), '<0.1 ms');
});

test('formatMillis loses precision as the number grows, in three steps', () => {
  assert.equal(formatMillis(1.25), '1.3 ms');
  assert.equal(formatMillis(9.99), '10.0 ms');
  assert.equal(formatMillis(42.4), '42 ms');
  assert.equal(formatMillis(1500), '1.50 s');
});

test('formatCount groups so four thousand cannot be read as forty', () => {
  assert.equal(formatCount(1), '1');
  assert.equal(formatCount(4000), '4,000');
  assert.equal(formatCount(1234567), '1,234,567');
});
