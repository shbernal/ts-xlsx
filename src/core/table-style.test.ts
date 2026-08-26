import assert from 'node:assert/strict';
import {test} from 'node:test';

import {AuthoringError, XlsxError} from '../errors.ts';
import {checkTableStyle} from './table-style.ts';

const style = (overrides: Parameters<typeof checkTableStyle>[0]) => () =>
  checkTableStyle(overrides);

test('a table style with a name and no sized element passes', () => {
  assert.doesNotThrow(
    style({name: 'Branded', elements: {firstRowStripe: {size: 2}, headerRow: {}}}),
  );
});

test('a style claim about the whole style stays in the library taxonomy', () => {
  // Both of these are claims about how the parts fit together, which is where AuthoringError starts.
  assert.throws(style({name: '', elements: {}}), AuthoringError);
  assert.throws(style({name: 'S', elements: {headerRow: {size: 2}}}), {
    name: 'AuthoringError',
    message: /cannot carry a size/,
  });
});

test('a band size that is not a positive integer is a native RangeError', () => {
  // One number out of range, so it follows the taxonomy's own line and stays native. Asserted as
  // *not* an XlsxError too, so a later re-wrap reddens the suite rather than quietly changing what
  // a caller catches.
  for (const size of [0, -1, 1.5]) {
    assert.throws(style({name: 'S', elements: {firstRowStripe: {size}}}), {
      name: 'RangeError',
      message: /positive integer/,
    });
    assert.throws(
      style({name: 'S', elements: {firstRowStripe: {size}}}),
      (error: unknown) => !(error instanceof XlsxError),
    );
  }
});
