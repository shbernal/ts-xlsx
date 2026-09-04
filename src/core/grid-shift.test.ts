import assert from 'node:assert/strict';
import {test} from 'node:test';

import {
  type AxisSplice,
  isDeletedSpan,
  shiftIndex,
  shiftPoint,
  shiftRect,
  shiftSpan,
} from './grid-shift.ts';

// The module the architecture doc singles out as *the* single-source rule, exercised directly for
// the first time. Everything else in the model re-anchors through these functions, so a drift here
// re-points a merge, a table, a validation or a filter at cells nobody chose, silently, and survives
// a re-write. Each case names the position a coordinate holds relative to the edit, since that, and
// not the numbers, is what the arithmetic is about.
//
// The scenario throughout: the splice deletes `count` lines at `start` and puts `count + delta`
// back, so `delta` is the net movement of everything after the edit.

const insertTwoAtRowFive: AxisSplice = {axis: 'row', start: 5, count: 0, delta: 2};
const deleteThreeAtRowFive: AxisSplice = {axis: 'row', start: 5, count: 3, delta: -3};
const replaceThreeAtRowFive: AxisSplice = {axis: 'row', start: 5, count: 3, delta: 0};

test('shiftIndex leaves a coordinate before the edit exactly where it was', () => {
  // Row 3, insert two rows at 5. Nothing above the cut can move: that is what "before" means.
  assert.equal(shiftIndex(3, insertTwoAtRowFive), 3);
  assert.equal(shiftIndex(4, deleteThreeAtRowFive), 4, 'a delete does not reach backwards either');
  assert.equal(shiftIndex(1, {axis: 'row', start: 2, count: 1, delta: 0}), 1);
});

test('shiftIndex shifts a coordinate at or after the edited span by delta', () => {
  // The first line past the deleted span is the boundary an off-by-one lands on.
  assert.equal(shiftIndex(5, insertTwoAtRowFive), 7, 'an insert pushes the cut line itself down');
  assert.equal(
    shiftIndex(8, deleteThreeAtRowFive),
    5,
    'the first line after a 3-row delete pulls up to 5',
  );
  assert.equal(shiftIndex(9, deleteThreeAtRowFive), 6);
});

test('a shift never leaves the grid: the last line of each axis is the ceiling', () => {
  // The measured defect this closes. A whole-column region is written bounded, `B1:B1048576`, so an
  // insert above it would push its bottom edge to a row that cannot exist, and Excel meets the
  // resulting file with the "we found a problem with some content" repair prompt. Excel's own answer
  // to the same edit is this clamp: it leaves the edge on the last row.
  const insertTwoAtTop: AxisSplice = {axis: 'row', start: 1, count: 0, delta: 2};
  assert.equal(
    shiftIndex(1_048_576, insertTwoAtTop),
    1_048_576,
    'the bottom edge stays on the grid',
  );
  assert.equal(shiftIndex(1_048_575, insertTwoAtTop), 1_048_576, 'and so does the line above it');
  assert.equal(
    shiftIndex(16_384, {...insertTwoAtTop, axis: 'col'}),
    16_384,
    'the column axis has its own, lower, wall',
  );
  assert.equal(shiftIndex(16_384, insertTwoAtTop), 16_386, 'which does not bind the row axis');
  assert.equal(
    shiftIndex(1_048_570, insertTwoAtTop),
    1_048_572,
    'a shift with room to land is untouched by the ceiling',
  );
});

test('shiftIndex clamps a coordinate inside a deleted span to the cut line', () => {
  // The best effort for a geometry straddling the cut. A caller that must instead *drop* what the
  // delete swallowed asks isDeletedSpan first, which is the division of labour these two encode.
  assert.equal(shiftIndex(5, deleteThreeAtRowFive), 5, 'the first deleted line');
  assert.equal(shiftIndex(6, deleteThreeAtRowFive), 5, 'and every line inside the span with it');
  assert.equal(shiftIndex(7, deleteThreeAtRowFive), 5, 'up to the last');
});

test('shiftIndex with delta 0 is a replacement: the tail keeps its numbers', () => {
  // Replacing three rows with three others moves nothing after them, but still clamps what was
  // inside. The two halves are independent, which a test that only exercised inserts would miss.
  assert.equal(shiftIndex(4, replaceThreeAtRowFive), 4);
  assert.equal(shiftIndex(6, replaceThreeAtRowFive), 5, 'inside still clamps');
  assert.equal(shiftIndex(8, replaceThreeAtRowFive), 8);
});

test('a count of 0 makes shiftIndex a pure insertion at the cut line', () => {
  // No line is deleted, so nothing can be clamped and `start` itself is already "at or after".
  const insertThreeAtRowFive: AxisSplice = {axis: 'row', start: 5, count: 0, delta: 3};
  for (const v of [1, 4]) assert.equal(shiftIndex(v, insertThreeAtRowFive), v);
  for (const v of [5, 6]) assert.equal(shiftIndex(v, insertThreeAtRowFive), v + 3);
});

test('isDeletedSpan is true only for a span wholly inside the deleted lines', () => {
  // Rows 5,6,7 deleted.
  assert.equal(isDeletedSpan(5, 7, deleteThreeAtRowFive), true, 'exactly the deleted span');
  assert.equal(isDeletedSpan(6, 6, deleteThreeAtRowFive), true, 'a single line inside it');
  assert.equal(isDeletedSpan(5, 5, deleteThreeAtRowFive), true, 'the first deleted line');
  assert.equal(isDeletedSpan(7, 7, deleteThreeAtRowFive), true, 'the last');
});

test('isDeletedSpan is false for a span that straddles either edge of the cut', () => {
  // Straddling means part of the geometry survives, so it moves rather than being dropped: the
  // caller clamps it through shiftIndex instead.
  assert.equal(isDeletedSpan(4, 7, deleteThreeAtRowFive), false, 'starts above the cut');
  assert.equal(isDeletedSpan(5, 8, deleteThreeAtRowFive), false, 'ends below it');
  assert.equal(isDeletedSpan(4, 8, deleteThreeAtRowFive), false, 'spans the whole edit and more');
});

test('isDeletedSpan is false for a span entirely outside the cut on either side', () => {
  assert.equal(isDeletedSpan(1, 4, deleteThreeAtRowFive), false);
  assert.equal(isDeletedSpan(8, 9, deleteThreeAtRowFive), false);
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
    assert.equal(isDeletedSpan(lo, hi, insertTwoAtRowFive), false, `${lo}..${hi}`);
  }
});

test('shiftSpan moves both edges together and drops a span the delete swallowed', () => {
  // The two answers are exclusive: a span straddling the cut clamps and survives, one wholly inside
  // it is gone. A caller reading only the moved edges would clamp its way out of the drop.
  assert.deepEqual(shiftSpan(8, 9, deleteThreeAtRowFive), {lo: 5, hi: 6}, 'wholly after the cut');
  assert.deepEqual(shiftSpan(4, 8, deleteThreeAtRowFive), {lo: 4, hi: 5}, 'straddling it clamps');
  assert.equal(shiftSpan(5, 7, deleteThreeAtRowFive), undefined, 'wholly inside it is dropped');
});

test('shiftRect moves only the spliced axis and passes the other through', () => {
  // The projection every range-bound thing in the model re-anchors by. The axis not being spliced
  // must come back with the numbers it went in with: a rectangle whose columns moved on a row insert
  // is a merge, a dropdown or a highlight sitting over cells nobody chose.
  const rect = {top: 8, left: 2, bottom: 9, right: 4};
  assert.deepEqual(shiftRect(rect, insertTwoAtRowFive), {top: 10, left: 2, bottom: 11, right: 4});
  assert.deepEqual(
    shiftRect(rect, {axis: 'col', start: 1, count: 0, delta: 2}),
    {top: 8, left: 4, bottom: 9, right: 6},
    'a column splice moves the horizontal edges and leaves the rows alone',
  );
});

test('shiftRect returns undefined for a rectangle the delete swallowed on the spliced axis', () => {
  // Only the spliced axis decides: deleting rows 5..7 takes a rectangle covering exactly those rows,
  // whatever columns it spans, and clamps one that merely overlaps them.
  assert.equal(shiftRect({top: 5, left: 1, bottom: 7, right: 40}, deleteThreeAtRowFive), undefined);
  assert.deepEqual(
    shiftRect({top: 4, left: 1, bottom: 7, right: 2}, deleteThreeAtRowFive),
    {top: 4, left: 1, bottom: 5, right: 2},
    'a rectangle straddling the cut clamps instead of dropping',
  );
});

test('shiftPoint moves a cell on the spliced axis and drops one on a deleted line', () => {
  assert.deepEqual(shiftPoint({col: 2, row: 8}, insertTwoAtRowFive), {col: 2, row: 10});
  assert.deepEqual(
    shiftPoint({col: 2, row: 8}, {axis: 'col', start: 1, count: 0, delta: 3}),
    {col: 5, row: 8},
    'a column splice moves the column and leaves the row',
  );
  assert.equal(shiftPoint({col: 2, row: 6}, deleteThreeAtRowFive), undefined, 'its line is gone');
  assert.equal(
    shiftPoint({col: 6, row: 6}, {axis: 'col', start: 5, count: 3, delta: -3}),
    undefined,
    'and the same coordinate on the other axis',
  );
});
