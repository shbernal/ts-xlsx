import assert from 'node:assert/strict';
import {test} from 'node:test';

import {MergeIndex} from './merge-index.ts';
import type {MergeRect} from './merge.ts';
import {Worksheet} from './worksheet.ts';

// The index holds the region list by reference and rebuilds from it on the first query, so a list
// that already has regions in it needs no announcement. A test adding regions afterwards reports
// them the way the sheet does.
function indexOver(rects: MergeRect[]): MergeIndex {
  return new MergeIndex(rects);
}

function rect(top: number, left: number, bottom: number, right: number): MergeRect {
  return {top, left, bottom, right};
}

test('masterOf resolves a covered position to its region top-left and any other to itself', () => {
  const index = indexOver([rect(2, 2, 3, 4)]);
  assert.deepEqual(index.masterOf(3, 4), {row: 2, col: 2}, 'the far corner of the region');
  assert.deepEqual(index.masterOf(2, 2), {row: 2, col: 2}, 'the anchor itself');
  assert.deepEqual(index.masterOf(4, 2), {row: 4, col: 2}, 'a row below the region');
  assert.deepEqual(indexOver([]).masterOf(3, 4), {row: 3, col: 4}, 'no regions at all');
});

test('masterOf takes a covering region, since a sheet admits no overlap', () => {
  const index = indexOver([rect(1, 1, 5, 5), rect(2, 2, 3, 3)]);
  assert.deepEqual(index.masterOf(2, 2), {row: 1, col: 1});
});

test('overlap is found at a corner, along an edge, and by containment, in both orders', () => {
  const established = rect(10, 10, 20, 20);
  const cases: {name: string; other: MergeRect}[] = [
    {name: 'a shared corner cell', other: rect(20, 20, 25, 25)},
    {name: 'a shared edge', other: rect(15, 20, 15, 30)},
    {name: 'the candidate inside the region', other: rect(12, 12, 14, 14)},
    {name: 'the region inside the candidate', other: rect(1, 1, 40, 40)},
  ];
  for (const {name, other} of cases) {
    assert.notEqual(indexOver([established]).overlapping(other), undefined, `${name}, added first`);
    assert.notEqual(
      indexOver([other]).overlapping(established),
      undefined,
      `${name}, added second`,
    );
  }
});

test('a region that only touches without sharing a cell is not an overlap', () => {
  const index = indexOver([rect(10, 10, 20, 20)]);
  assert.equal(index.overlapping(rect(21, 10, 25, 20)), undefined, 'directly below');
  assert.equal(index.overlapping(rect(10, 21, 20, 25)), undefined, 'directly right');
  assert.equal(index.overlapping(rect(21, 21, 25, 25)), undefined, 'diagonally past the corner');
});

test('a region taller than a band is found from a query far below its top row', () => {
  // The bucket holds a region under its top row alone, so a tall region is reached through the
  // tallest-region window rather than by being listed in every band it covers.
  const index = indexOver([rect(1, 1, 5000, 3)]);
  assert.notEqual(index.overlapping(rect(4000, 2, 4000, 2)), undefined, 'deep inside the region');
  assert.equal(index.overlapping(rect(4000, 9, 4000, 9)), undefined, 'beside it, not inside');
  assert.equal(index.overlapping(rect(5001, 1, 5001, 3)), undefined, 'one row past its bottom');
});

test('a region reported after construction is found without a rebuild', () => {
  const rects: MergeRect[] = [];
  const index = indexOver(rects);
  assert.equal(
    index.overlapping(rect(3, 3, 4, 4)),
    undefined,
    'the first query rebuilds from empty',
  );
  const added = rect(3, 3, 4, 4);
  rects.push(added);
  index.note(added);
  assert.equal(index.overlapping(rect(4, 4, 6, 6)), added);
});

test('invalidating rereads a list that was rewritten behind the index', () => {
  const rects: MergeRect[] = [rect(10, 10, 20, 20)];
  const index = indexOver(rects);
  assert.notEqual(index.overlapping(rect(15, 15, 16, 16)), undefined);
  rects.length = 0;
  index.invalidate();
  assert.equal(index.overlapping(rect(15, 15, 16, 16)), undefined, 'the removed region is gone');
});

// The reader calls mergeCells once per <mergeCell> in the part, so overlap-checking each new region
// against every existing one made loading a sheet's merges quadratic. A ratio rather than a
// millisecond budget, taken as the fastest of several runs because a single timed run is at the
// mercy of whenever the collector pauses.
function fastestRun(runs: number, work: () => void): number {
  let best = Infinity;
  for (let i = 0; i < runs; i++) {
    const start = performance.now();
    work();
    best = Math.min(best, performance.now() - start);
  }
  return best;
}

function addMerges(count: number): void {
  const sheet = new Worksheet('S', 1);
  for (let i = 1; i <= count; i++) sheet.mergeCells(`A${i * 2}:C${i * 2}`);
}

test('adding non-overlapping merges scales linearly, not quadratically', () => {
  const small = fastestRun(5, () => addMerges(5000));
  const large = fastestRun(5, () => addMerges(20000));
  // Four times the merges. Linear measures around 4x; the scan it replaced measured 16x.
  assert.ok(
    large < small * 9,
    `20k merges took ${large.toFixed(1)}ms against ${small.toFixed(1)}ms for 5k: that is ${(large / small).toFixed(1)}x, which is not linear`,
  );
});
