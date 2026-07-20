// Cluster: core-model
//
// Real-world scenario: a program wants to drop one or more columns from a built
// sheet — e.g. remove a spacer column, or strip trailing empties — with a single
// spliceColumns(start, count). Historically this path was fragile: an internal
// assumption that every stored row is a dense array meant splicing columns could
// throw ("r is not a function"/"is not an array") on sheets whose rows are sparse.
//
// The invariant this locks: spliceColumns removes exactly the requested columns and
// shifts the columns to their right leftward, without throwing, for both a
// single-column and a multi-column removal.

import type {Assert, Case, CorpusApi} from '../case.ts';

// One row across five columns: A1..E1 hold 'c1'..'c5'.
const CELLS = ['A', 'B', 'C', 'D', 'E'].map((L, i) => ({ref: `${L}1`, value: `c${i + 1}`}));

export default {
  id: 'splice-columns-shifts-remaining',
  provenance: {source: 'upstream-issue', repo: 'exceljs/exceljs', ref: 670},
  cluster: 'core-model',
  description:
    'spliceColumns(start, count) removes exactly the requested columns and shifts the ' +
    'columns to their right leftward, without throwing — for both single- and ' +
    'multi-column removals.',

  behavior: [
    {
      name: 'removing one column shifts the rest left and does not throw',
      baseline: 'pass',
      expect(api: CorpusApi, assert: Assert) {
        const {cells, error} = api.mutateWorksheet({
          cells: CELLS,
          ops: [{op: 'spliceColumns', start: 2, count: 1}],
          read: ['A1', 'B1', 'C1', 'D1'],
        });
        assert.strictEqual(error, null, 'spliceColumns must not throw');
        assert.strictEqual(cells.A1, 'c1', 'column before the splice is untouched');
        assert.strictEqual(cells.B1, 'c3', 'c2 removed, c3 shifted into column B');
        assert.strictEqual(cells.C1, 'c4', 'c4 shifted into column C');
        assert.strictEqual(cells.D1, 'c5', 'c5 shifted into column D');
      },
    },
    {
      name: 'removing multiple columns shifts the remainder left correctly',
      baseline: 'pass',
      expect(api: CorpusApi, assert: Assert) {
        const {cells, error} = api.mutateWorksheet({
          cells: CELLS,
          ops: [{op: 'spliceColumns', start: 2, count: 2}],
          read: ['A1', 'B1', 'C1'],
        });
        assert.strictEqual(error, null, 'spliceColumns must not throw');
        assert.strictEqual(cells.A1, 'c1', 'column before the splice is untouched');
        assert.strictEqual(cells.B1, 'c4', 'c2 and c3 removed, c4 shifted into column B');
        assert.strictEqual(cells.C1, 'c5', 'c5 shifted into column C');
      },
    },
  ],
} satisfies Case;
