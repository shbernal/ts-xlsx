// Cluster: tables
//
// Real-world scenario: a sheet carries a formatted table over two columns, and someone deletes those
// two columns. The table has nothing left to occupy. A row splice that deletes every row of a table
// removes it; the column axis did not ask the question at all, so the table survived, still carrying
// the names of columns that no longer exist, now declared over whatever slid into their place. The
// writer then emits that table part, and the file opens with a repair prompt.
//
// The other half of the same asymmetry: a column insert to the left of a table moved its anchor by an
// unbounded increment, so a table on the last column could be pushed off the grid, after which the
// range it reports cannot even be read.

import type {Assert, Case, CorpusApi} from '../case.ts';

export default {
  id: 'column-splice-drops-a-table-it-deletes-whole',
  provenance: {source: 'grid-edit-audit'},
  cluster: 'tables',
  description:
    'A column splice re-anchors a table on the same terms a row splice does: a splice that deletes ' +
    'every column of a table removes the table, one to its left moves it, and neither can push its ' +
    'anchor outside the grid.',

  behavior: [
    {
      name: "a splice deleting a table's every column removes the table",
      expect(api: CorpusApi, assert: Assert) {
        assert.equal(api.tableThroughColumnSplice().wholeDelete.tableCount, 0);
      },
    },
    {
      name: "a splice to the table's left moves it and keeps it",
      expect(api: CorpusApi, assert: Assert) {
        const {leftInsert} = api.tableThroughColumnSplice();
        assert.equal(leftInsert.tableCount, 1);
        assert.equal(leftInsert.range, 'C1:D2');
      },
    },
    {
      name: 'a table on the last column keeps an anchor whose range can still be read',
      expect(api: CorpusApi, assert: Assert) {
        const {rightEdge} = api.tableThroughColumnSplice();
        assert.equal(rightEdge.range, 'XFD1:XFD2', 'clamped, not pushed past the last column');
      },
    },
  ],
} satisfies Case;
