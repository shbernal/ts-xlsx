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
// against every existing one made loading a sheet's merges quadratic. What follows counts the
// regions a query actually compares against, rather than timing a large one.
//
// It used to be timed, as a ratio between two sizes taken as the fastest of several runs. A ratio
// does tell linear from quadratic without depending on how fast the machine is, but it does depend
// on the machine being *free*: one `verify` run had another gate holding the CPU and the budget
// failed, then passed on three serial re-runs. Counting is exact, states the invariant directly
// (a query visits the bands its window spans, not the sheet), and no neighbouring process can move
// the number.

/**
 * A region that reports being compared against. `rectsOverlap` reads at least one of a candidate's
 * four bounds, and *which* one depends on where its conjunction short-circuits, so any read marks
 * the region once: the question is how many regions a query touched, not how many field reads it
 * took to touch them.
 */
function reporting(rect: MergeRect, touched: Set<MergeRect>): MergeRect {
  const view: MergeRect = {
    get top() {
      touched.add(view);
      return rect.top;
    },
    get left() {
      touched.add(view);
      return rect.left;
    },
    get bottom() {
      touched.add(view);
      return rect.bottom;
    },
    get right() {
      touched.add(view);
      return rect.right;
    },
  };
  return view;
}

/** `count` disjoint single-row regions, one every other row, each reporting when it is compared. */
function reportingBand(count: number, touched: Set<MergeRect>): MergeRect[] {
  return Array.from({length: count}, (_, i) =>
    reporting(rect(i * 2 + 1, 1, i * 2 + 1, 3), touched),
  );
}

test('a query compares against its own band, not against every region on the sheet', () => {
  // Four times the regions, and the query still reads the same handful: the index buckets a region
  // under its top row alone and the tallest region here is one row, so the window a query opens is
  // one band wide however far down the sheet the regions run. The scan this replaced compared every
  // region, which is what made loading a file's merges quadratic.
  const counts = [5_000, 20_000];
  const compared = counts.map((count) => {
    const touched = new Set<MergeRect>();
    const index = new MergeIndex(reportingBand(count, touched));
    // Warm the rebuild first: it reads every region's bounds to place them, which is the O(n) pass
    // the index is entitled to and not the per-query cost under test.
    index.overlapping(rect(1, 1, 1, 1));
    touched.clear();
    assert.equal(index.overlapping(rect(4001, 1, 4001, 3))?.top, 4001, `over ${count} regions`);
    return touched.size;
  });
  assert.deepEqual(
    compared,
    [compared[0], compared[0]],
    `a query compared against ${compared[1]} regions at ${counts[1]} where it compared against ${compared[0]} at ${counts[0]}`,
  );
  // A band is 64 rows and these regions sit on every other row, so a full band holds 32 of them.
  assert.ok(
    (compared[0] as number) <= 32,
    `a query compared against ${compared[0]} regions, which is more than one band holds`,
  );
});

test('adding a merge through the sheet asks the index once, whatever the sheet already holds', () => {
  // The end-to-end half: `mergeCells` overlap-checks, and with the query above bounded, the loading
  // path is bounded too. Asserted as a result rather than a duration - every one of these regions
  // is admitted, which a quadratic re-scan would also manage, but slowly enough to have been the
  // original bug report.
  const sheet = new Worksheet('S', 1);
  for (let i = 1; i <= 20_000; i++) sheet.mergeCells(`A${i * 2}:C${i * 2}`);
  assert.equal(sheet.merges.length, 20_000);
  assert.throws(() => sheet.mergeCells('A2:C2'), /overlap/i, 'and an overlap is still refused');
});
