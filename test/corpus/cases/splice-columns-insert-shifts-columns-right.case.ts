// Cluster: core-model
//
// Real-world scenario: a program wants to open a workbook and *insert* one or more blank columns into
// an existing sheet — spliceColumns's insertion mode, spliceColumns(start, 0, ...columns) with a zero
// delete-count and one array per inserted column. Users reach for this to add spacer or fill-in
// columns without rebuilding the sheet. The insertion must shift every column at and after `start`
// rightward by the number of inserted columns, preserving their values, and must not throw — the
// removal-side splice invariants (splice-columns-shifts-remaining) have an insertion analog that is
// just as important and was historically the fragile, untested direction.
//
// This locks the insertion direction as a regression guard: inserting blank columns shifts the
// existing data right by exactly the inserted count.

import type {Assert, Case, CorpusApi} from '../case.ts';

// One row across five columns: A1..E1 hold 'c1'..'c5'.
const CELLS = ['A', 'B', 'C', 'D', 'E'].map((L, i) => ({ref: `${L}1`, value: `c${i + 1}`}));

export default {
  id: 'splice-columns-insert-shifts-columns-right',
  provenance: {source: 'upstream-issue'},
  cluster: 'core-model',
  description:
    'spliceColumns(start, 0, ...columns) inserts the given blank columns at `start` and shifts the ' +
    'columns at and after `start` rightward by the inserted count, preserving their values without ' +
    'throwing — the insertion analog of the removal-side column splice.',

  behavior: [
    {
      name: 'inserting two blank columns shifts the columns at and after the point right by two',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const {cells, error} = await api.mutateWorksheet({
          cells: CELLS,
          // insert two blank columns before column 3 (C): start=3, delete-count=0, two empty inserts
          ops: [{op: 'spliceColumns', start: 3, count: 0, inserts: [[], []]}],
          read: ['A1', 'B1', 'C1', 'D1', 'E1', 'F1', 'G1'],
        });
        assert.strictEqual(error, null, `insertion must not throw; got ${JSON.stringify(error)}`);
        assert.strictEqual(cells.A1, 'c1', 'columns before the insertion point are untouched');
        assert.strictEqual(cells.B1, 'c2', 'columns before the insertion point are untouched');
        assert.strictEqual(cells.C1, null, 'the first inserted column is blank');
        assert.strictEqual(cells.D1, null, 'the second inserted column is blank');
        assert.strictEqual(cells.E1, 'c3', 'c3 shifts right by two into column E');
        assert.strictEqual(cells.F1, 'c4', 'c4 shifts right by two into column F');
        assert.strictEqual(cells.G1, 'c5', 'c5 shifts right by two into column G');
      },
    },
  ],
} satisfies Case;
