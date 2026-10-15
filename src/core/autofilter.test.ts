import assert from 'node:assert/strict';
import {test} from 'node:test';

import {type AutoFilter, canonicalizeAutoFilter, shiftAutoFilter} from './autofilter.ts';

// The most involved re-anchoring in the model, and the one with the least visible failure: a
// filter's criteria are addressed by their offset from the filter's own left edge, not by an
// absolute column. A column splice therefore has to move the range *and* re-base every offset
// against the new edge. Left alone, the offsets keep their old numbers and each criterion silently
// re-points at a neighbouring column, which no round-trip notices because the file is still valid.

const filterOn = (ref: string, colIds: readonly number[]): AutoFilter =>
  canonicalizeAutoFilter({
    ref,
    columns: colIds.map((colId) => ({
      colId,
      criteria: {kind: 'values' as const, values: ['x'], blank: false},
    })),
  });

const colIdsOf = (filter: AutoFilter | undefined) => filter?.columns.map((c) => c.colId);

test('canonicalizeAutoFilter normalises a bare range string to a criteria-free filter', () => {
  assert.deepEqual(canonicalizeAutoFilter('$B$2:$D$10'), {ref: 'B2:D10', columns: []});
});

test('a column insert before the filter re-bases every criterion against the new left edge', () => {
  // B2:D10 with criteria on B (offset 0) and D (offset 2). Two columns inserted at A push the whole
  // filter right, and the offsets must stay pointing at the same *columns*, so they do not change.
  const moved = shiftAutoFilter(filterOn('B2:D10', [0, 2]), {
    axis: 'col',
    start: 1,
    count: 0,
    delta: 2,
  });
  assert.equal(moved?.ref, 'D2:F10');
  assert.deepEqual(colIdsOf(moved), [0, 2]);
});

test('a column insert inside the filter widens it and pushes the criteria past the cut', () => {
  // B2:D10, one column inserted at C. B keeps offset 0; D was offset 2 and is now offset 3, because
  // a column it does not own moved in front of it.
  const moved = shiftAutoFilter(filterOn('B2:D10', [0, 2]), {
    axis: 'col',
    start: 3,
    count: 0,
    delta: 1,
  });
  assert.equal(moved?.ref, 'B2:E10');
  assert.deepEqual(colIdsOf(moved), [0, 3]);
});

test('a filter straddling the cut survives with its edges clamped', () => {
  // Deleting rows 1..4 takes the filter's top edge but not its bottom, so the region shrinks to the
  // cut line rather than disappearing.
  const moved = shiftAutoFilter(filterOn('B2:D10', [0]), {
    axis: 'row',
    start: 1,
    count: 4,
    delta: -4,
  });
  assert.equal(moved?.ref, 'B1:D6');
});
