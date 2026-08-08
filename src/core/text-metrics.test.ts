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

// The one string here whose true answer was measured rather than reasoned about: Excel auto-fitted
// this cell, in a column of stated width 40, to 87 points - six lines of a 14.5-point default -
// where the character count says five. Pinned so that a change making the count word-aware (which
// would raise this to 6 and is a live question) fails here and finds
// docs/knowledge/specs/rows-with-no-stated-height-are-autofitted-on-open.md.
test('counting characters reads a line low against Excel, which breaks at word boundaries', () => {
  const measuredAgainstExcel =
    'R1 len=200: tolerate boundary pipeline anchored boundary anchored anchored anchored boundary ' +
    'tolerate anchored canonical anchored tolerate anchored anchored boundary tolerate anchored ' +
    'anchored pipelin';
  assert.equal(measuredAgainstExcel.length, 200);
  assert.equal(estimateWrappedLines(measuredAgainstExcel, 40), 5, 'Excel laid this out in 6');
});

test('a width that wraps nothing is refused rather than answered with Infinity or NaN', () => {
  for (const width of [0, -1, Number.NaN, Number.POSITIVE_INFINITY]) {
    assert.throws(() => estimateWrappedLines('text', width), RangeError, `width ${width}`);
  }
});
