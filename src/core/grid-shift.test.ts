import assert from 'node:assert/strict';
import {test} from 'node:test';

import {isDeletedSpan, shiftIndex} from './grid-shift.ts';

// The module the architecture doc singles out as *the* single-source rule, exercised directly for
// the first time. Everything else in the model re-anchors through these two functions, so a drift
// here re-points a merge, a table, a validation or a filter at cells nobody chose, silently, and
// survives a re-write. Each case names the position a coordinate holds relative to the edit, since
// that, and not the numbers, is what the arithmetic is about.
//
// The scenario throughout: the splice deletes `count` lines at `start` and puts `count + delta`
// back, so `delta` is the net movement of everything after the edit.

test('shiftIndex leaves a coordinate before the edit exactly where it was', () => {
  // Row 3, insert two rows at 5. Nothing above the cut can move: that is what "before" means.
  assert.equal(shiftIndex(3, 5, 0, 2, 'row'), 3);
  assert.equal(shiftIndex(4, 5, 3, -3, 'row'), 4, 'a delete does not reach backwards either');
  assert.equal(shiftIndex(1, 2, 1, 0, 'row'), 1);
});

test('shiftIndex shifts a coordinate at or after the edited span by delta', () => {
  // The first line past the deleted span is the boundary an off-by-one lands on.
  assert.equal(shiftIndex(5, 5, 0, 2, 'row'), 7, 'a pure insert pushes the cut line itself down');
  assert.equal(
    shiftIndex(8, 5, 3, -3, 'row'),
    5,
    'the first line after a 3-row delete pulls up to 5',
  );
  assert.equal(shiftIndex(9, 5, 3, -3, 'row'), 6);
});

test('a shift never leaves the grid: the last line of each axis is the ceiling', () => {
  // The measured defect this closes. A whole-column region is written bounded, `B1:B1048576`, so an
  // insert above it would push its bottom edge to a row that cannot exist, and Excel meets the
  // resulting file with the "we found a problem with some content" repair prompt. Excel's own answer
  // to the same edit is this clamp: it leaves the edge on the last row.
  assert.equal(
    shiftIndex(1_048_576, 1, 0, 2, 'row'),
    1_048_576,
    'the bottom edge stays on the grid',
  );
  assert.equal(shiftIndex(1_048_575, 1, 0, 2, 'row'), 1_048_576, 'and so does the line above it');
  assert.equal(
    shiftIndex(16_384, 1, 0, 2, 'col'),
    16_384,
    'the column axis has its own, lower, wall',
  );
  assert.equal(shiftIndex(16_384, 1, 0, 2, 'row'), 16_386, 'which does not bind the row axis');
  assert.equal(
    shiftIndex(1_048_570, 1, 0, 2, 'row'),
    1_048_572,
    'a shift with room to land is untouched by the ceiling',
  );
});

test('shiftIndex clamps a coordinate inside a deleted span to the cut line', () => {
  // The best effort for a geometry straddling the cut. A caller that must instead *drop* what the
  // delete swallowed asks isDeletedSpan first, which is the division of labour these two encode.
  assert.equal(shiftIndex(5, 5, 3, -3, 'row'), 5, 'the first deleted line');
  assert.equal(shiftIndex(6, 5, 3, -3, 'row'), 5, 'and every line inside the span with it');
  assert.equal(shiftIndex(7, 5, 3, -3, 'row'), 5, 'up to the last');
});

test('shiftIndex with delta 0 is a replacement: the tail keeps its numbers', () => {
  // Replacing three rows with three others moves nothing after them, but still clamps what was
  // inside. The two halves are independent, which a test that only exercised inserts would miss.
  assert.equal(shiftIndex(4, 5, 3, 0, 'row'), 4);
  assert.equal(shiftIndex(6, 5, 3, 0, 'row'), 5, 'inside still clamps');
  assert.equal(shiftIndex(8, 5, 3, 0, 'row'), 8);
});

test('a count of 0 makes shiftIndex a pure insertion at the cut line', () => {
  // No line is deleted, so nothing can be clamped and `start` itself is already "at or after".
  for (const v of [1, 4]) assert.equal(shiftIndex(v, 5, 0, 3, 'row'), v);
  for (const v of [5, 6]) assert.equal(shiftIndex(v, 5, 0, 3, 'row'), v + 3);
});

test('isDeletedSpan is true only for a span wholly inside the deleted lines', () => {
  // Rows 5,6,7 deleted.
  assert.equal(isDeletedSpan(5, 7, 5, 3), true, 'exactly the deleted span');
  assert.equal(isDeletedSpan(6, 6, 5, 3), true, 'a single line inside it');
  assert.equal(isDeletedSpan(5, 5, 5, 3), true, 'the first deleted line');
  assert.equal(isDeletedSpan(7, 7, 5, 3), true, 'the last');
});

test('isDeletedSpan is false for a span that straddles either edge of the cut', () => {
  // Straddling means part of the geometry survives, so it moves rather than being dropped: the
  // caller clamps it through shiftIndex instead.
  assert.equal(isDeletedSpan(4, 7, 5, 3), false, 'starts above the cut');
  assert.equal(isDeletedSpan(5, 8, 5, 3), false, 'ends below it');
  assert.equal(isDeletedSpan(4, 8, 5, 3), false, 'spans the whole edit and more');
});

test('isDeletedSpan is false for a span entirely outside the cut on either side', () => {
  assert.equal(isDeletedSpan(1, 4, 5, 3), false);
  assert.equal(isDeletedSpan(8, 9, 5, 3), false);
});

test('an insert deletes nothing, so isDeletedSpan can never be true for it', () => {
  // count 0 is the insertion case, and an insertion has no casualties. Without this the empty span
  // `start..start-1` could read as containing a coordinate at `start`.
  for (const [lo, hi] of [
    [4, 4],
    [5, 5],
    [5, 9],
    [1, 1_000],
  ] as const) {
    assert.equal(isDeletedSpan(lo, hi, 5, 0), false, `${lo}..${hi}`);
  }
});
