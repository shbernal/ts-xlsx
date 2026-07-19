// Cluster: rows
//
// Real-world scenario: code builds a worksheet, then removes rows with a splice, and afterward asks
// the worksheet for its last row (to append below it, or to read the final record). Deleting rows
// shifts the surviving rows up but can leave trailing EMPTY slots behind the real last populated row
// in the sheet's internal row list. `lastRow` must still resolve to that last populated row — its
// number and its cells reachable — rather than indexing into a trailing empty slot and returning a
// row whose cells are all empty (effectively undefined data), which breaks the "find the end of the
// data" idiom right after a delete.

/** @typedef {{ name: string, baseline: 'pass'|'fail', expect: (api: any, assert: any) => Promise<void>|void }} Behavior */

// Build five rows a..e, then splice.
const cells = [
  {ref: 'A1', value: 'a'},
  {ref: 'A2', value: 'b'},
  {ref: 'A3', value: 'c'},
  {ref: 'A4', value: 'd'},
  {ref: 'A5', value: 'e'},
];

export default {
  id: 'last-row-tracks-tail-after-splice',
  provenance: {source: 'upstream-issue'},
  cluster: 'rows',
  description:
    'After splicing rows out of a worksheet, lastRow resolves to the last POPULATED row (its number ' +
    'and value reachable), rather than falling into a trailing empty slot left behind by the delete.',

  /** @type {Behavior[]} */
  behavior: [
    {
      name: 'with no delete, lastRow points at the true final row (control)',
      baseline: 'pass',
      async expect(api, assert) {
        const {lastRow} = api.mutateWorksheet({cells, ops: []});
        assert.ok(lastRow, 'lastRow must resolve for a fully-populated sheet');
        assert.strictEqual(
          lastRow.value,
          'e',
          `lastRow of the untouched sheet is row 5 'e'; got ${JSON.stringify(lastRow?.value)}`,
        );
      },
    },
    {
      name: 'deleting the final row leaves lastRow on the new final populated row, not an empty slot',
      baseline: 'pass',
      async expect(api, assert) {
        const {lastRow} = api.mutateWorksheet({
          cells,
          ops: [{op: 'spliceRows', start: 5, count: 1}],
        });
        assert.ok(
          lastRow,
          `lastRow must resolve after removing the final row; got ${JSON.stringify(lastRow)}`,
        );
        assert.strictEqual(
          lastRow.value,
          'd',
          `lastRow should carry the last populated value 'd', not an emptied slot; got ${JSON.stringify(lastRow?.value)}`,
        );
        assert.strictEqual(
          lastRow.number,
          4,
          `lastRow should be row 4 after removing row 5; got #${lastRow?.number}`,
        );
      },
    },
    {
      name: 'deleting an interior block shifts data up and lastRow follows to the last populated row',
      baseline: 'pass',
      async expect(api, assert) {
        // Remove rows 2..3 (b,c): surviving data is a(1), d(2), e(3); the true last row is 'e'.
        const {lastRow} = api.mutateWorksheet({
          cells,
          ops: [{op: 'spliceRows', start: 2, count: 2}],
        });
        assert.ok(
          lastRow,
          `lastRow must resolve after an interior delete; got ${JSON.stringify(lastRow)}`,
        );
        assert.strictEqual(
          lastRow.value,
          'e',
          `lastRow should be the shifted-up last row 'e', not a trailing empty slot; got ${JSON.stringify(lastRow?.value)}`,
        );
      },
    },
  ],
};
