import assert from 'node:assert/strict';
import {test} from 'node:test';

import {estimateWrappedLines} from './text-metrics.ts';

test('a cell always occupies at least one line', () => {
  assert.equal(estimateWrappedLines('', 10), 1, 'the empty string is one line, not zero');
  assert.equal(estimateWrappedLines('short', 10), 1);
});

test('text wraps at the stated width', () => {
  assert.equal(estimateWrappedLines('0123456789', 10), 1, 'exactly full stays one line');
  assert.equal(estimateWrappedLines('0123456789a', 10), 2, 'one over opens the next');
  assert.equal(estimateWrappedLines('a'.repeat(95), 10), 10);
});

test('a hard break opens a line of its own, and what follows wraps independently', () => {
  assert.equal(estimateWrappedLines('a\nb', 10), 2);
  assert.equal(estimateWrappedLines('a\n\nb', 10), 3, 'an empty segment is still a line');
  assert.equal(
    estimateWrappedLines(`${'x'.repeat(25)}\ny`, 10),
    4,
    'three wrapped lines then one more',
  );
  assert.equal(estimateWrappedLines('a\r\nb', 10), 2, 'a CRLF is one break, not two');
  assert.equal(estimateWrappedLines('a\rb', 10), 2, 'a lone CR too');
});

test('a fractional width is honoured rather than rounded first', () => {
  assert.equal(estimateWrappedLines('abcdefghi', 8.43), 2, '9 characters over 8.43 units');
  assert.equal(estimateWrappedLines('abcdefgh', 8.43), 1);
});

test('a width that wraps nothing is refused rather than answered with Infinity or NaN', () => {
  for (const width of [0, -1, Number.NaN, Number.POSITIVE_INFINITY]) {
    assert.throws(() => estimateWrappedLines('text', width), RangeError, `width ${width}`);
  }
});
