// Cluster: tables
//
// Real-world scenario: a worksheet has a merged range and populated cells to the right of a block of
// columns the user deletes via a column splice. After the splice, columns to the right of the removed
// block must shift left by the removed count, keeping their values, and any merged range entirely to
// the right of the removed block must be re-anchored to its new columns. The splice shifts the cell
// data correctly, but a merged range to the right is left stranded at its pre-splice columns. (The
// companion row-splice merge behavior is locked by splice-rows-preserves-merged-cells; this is the
// column-splice analog.)

import type {Assert, Case, CorpusApi} from '../case.ts';

// Values across the row; a merge F1:G1 to the right of column B (which will be removed).
const sheet = {
  cells: [
    {ref: 'A1', value: 'A'},
    {ref: 'B1', value: 'B'},
    {ref: 'C1', value: 'C'},
    {ref: 'F1', value: 'F'},
    {ref: 'H1', value: 'H'},
  ],
  ops: [
    {op: 'mergeCells', range: 'F1:G1'},
    {op: 'spliceColumns', start: 2, count: 1},
  ],
  read: ['A1', 'B1', 'E1', 'G1'],
};

export default {
  id: 'splice-columns-preserves-merges-and-trailing-data',
  provenance: {source: 'upstream-issue'},
  cluster: 'tables',
  description:
    'Splicing out a column shifts the columns to its right left by one, preserving their values, and ' +
    're-anchors a merged range that lies entirely to the right of the removed column to its new ' +
    'position rather than stranding it at the old columns.',

  behavior: [
    {
      name: 'columns to the right of the removed block keep their values at shifted positions',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const {cells} = await api.mutateWorksheet(sheet);
        assert.strictEqual(cells.B1, 'C', 'the value formerly at C1 shifts left to B1');
        assert.strictEqual(
          cells.E1,
          'F',
          'the value formerly at F1 shifts left to E1 (trailing data preserved)',
        );
        assert.strictEqual(cells.G1, 'H', 'the value formerly at H1 shifts left to G1');
      },
    },
    {
      name: 'a merged range to the right of the removed column is re-anchored to its new position',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const {merges} = await api.mutateWorksheet(sheet);
        assert.ok(
          merges.includes('E1:F1'),
          `the F1:G1 merge must shift left to E1:F1 after removing a column to its left; got ${JSON.stringify(merges)}`,
        );
      },
    },
  ],
} satisfies Case;
