// Cluster: tables
//
// Real-world scenario: a worksheet has a known number of populated rows. Deleting a run of rows that
// reaches the last populated row — trimming the tail off a sheet — should actually remove those rows
// and shrink the sheet's dimension, just as deleting rows in the middle shifts the rows below up. In
// the buggy behavior, a delete-splice whose range includes the final row silently leaves the trailing
// rows in place: the sheet still reports the old row count and the "deleted" cells still hold their
// values. The same defect applies to deleting trailing columns. A splice that touches the interior
// works; only the boundary at the end of the used range is mishandled.

/** @typedef {{ name: string, baseline: 'pass'|'fail', expect: (api: any, assert: any) => Promise<void>|void }} Behavior */

const rowCells = () => Array.from({length: 5}, (_, i) => ({ref: `A${i + 1}`, value: `r${i + 1}`}));
const colCells = () => Array.from({length: 5}, (_, i) => ({ref: `${String.fromCharCode(65 + i)}1`, value: `c${i + 1}`}));

export default {
  id: 'splice-rows-delete-at-worksheet-end',
  provenance: {source: 'upstream-issue'},
  cluster: 'tables',
  description:
    'A row/column delete-splice whose range includes the last populated row/column must actually ' +
    'remove it and shrink the sheet dimension — not silently leave the trailing entries in place — ' +
    'the same way an interior delete-splice shifts the following rows/columns up/left.',

  /** @type {Behavior[]} */
  behavior: [
    {
      name: 'deleting an interior row shifts the rows below up and reduces the populated row count',
      baseline: 'pass',
      expect(api, assert) {
        const {rowCount, cells} = api.mutateWorksheet({
          cells: rowCells(),
          ops: [{op: 'spliceRows', start: 3, count: 1}],
          read: ['A3', 'A4'],
        });
        assert.strictEqual(rowCount, 4, 'the sheet shrinks from 5 to 4 rows');
        assert.strictEqual(cells.A3, 'r4', 'the row below the deleted one shifts up into its place');
        assert.strictEqual(cells.A4, 'r5', 'subsequent rows shift up too');
      },
    },
    {
      name: 'deleting rows whose range includes the last populated row actually removes them',
      baseline: 'pass',
      expect(api, assert) {
        const {rowCount, cells} = api.mutateWorksheet({
          cells: rowCells(),
          ops: [{op: 'spliceRows', start: 4, count: 2}],
          read: ['A4', 'A5'],
        });
        assert.strictEqual(rowCount, 3, 'deleting the trailing two rows must leave only the first three');
        assert.strictEqual(cells.A4, null, 'the deleted trailing rows must not retain their cell values');
        assert.strictEqual(cells.A5, null, 'the deleted trailing rows must not retain their cell values');
      },
    },
    {
      name: 'deleting trailing columns removes them and shrinks the column count',
      baseline: 'pass',
      expect(api, assert) {
        const {columnCount, cells} = api.mutateWorksheet({
          cells: colCells(),
          ops: [{op: 'spliceColumns', start: 4, count: 2}],
          read: ['C1'],
        });
        assert.strictEqual(columnCount, 3, 'deleting the trailing two columns must leave only the first three');
        assert.strictEqual(cells.C1, 'c3', 'the columns before the cut are untouched');
      },
    },
  ],
};
