// Cluster: core-model
//
// Real-world scenario: a program builds a sheet, then wants to trim it down to a
// header — "remove every row after row 2". The natural call is a single
// spliceRows(start, count) where count spans all the remaining rows. Users reach for
// one bulk splice precisely because doing it one row at a time in a loop is
// pathologically slow on large sheets (tens of thousands of rows).
//
// The trap real users hit: when `count` is large — in particular when it reaches or
// exceeds the number of rows actually present from `start` onward — the splice
// removes *nothing* and the row count is unchanged, silently. A small count works; a
// count that would clear the tail is a no-op. The number passed must mean "remove
// this many rows" for every value, not just small ones.

/** @typedef {{ name: string, baseline: 'pass'|'fail', expect: (api: any, assert: any) => Promise<void>|void }} Behavior */

// A 10-row sheet: A1..A10 hold 'r1'..'r10'.
const CELLS = Array.from({length: 10}, (_, i) => ({ref: `A${i + 1}`, value: `r${i + 1}`}));

export default {
  id: 'splice-rows-removes-requested-count',
  provenance: {source: 'upstream-issue', repo: 'exceljs/exceljs', ref: 674},
  cluster: 'core-model',
  description:
    'spliceRows(start, count) removes exactly `count` rows for every count, including ' +
    'a count large enough to clear all rows from `start` to the end — not just small ' +
    'counts. A bulk removal must not silently become a no-op.',

  /** @type {Behavior[]} */
  behavior: [
    {
      name: 'a small count removes exactly that many rows and shifts the tail up',
      baseline: 'pass',
      expect(api, assert) {
        const {rowCount, cells} = api.mutateWorksheet({
          cells: CELLS,
          ops: [{op: 'spliceRows', start: 3, count: 2}],
          read: ['A2', 'A3', 'A4'],
        });
        assert.strictEqual(rowCount, 8, 'removing 2 of 10 rows leaves 8');
        assert.strictEqual(cells.A2, 'r2', 'rows before the splice are untouched');
        assert.strictEqual(cells.A3, 'r5', 'r3 and r4 removed, r5 shifted into row 3');
        assert.strictEqual(cells.A4, 'r6', 'r6 shifted into row 4');
      },
    },
    {
      name: 'a count that spans all rows from start to the end clears the tail',
      baseline: 'pass',
      expect(api, assert) {
        // 10 rows, remove from row 3 with count 8 (rows 3..10) → only r1, r2 remain.
        const {rowCount, cells} = api.mutateWorksheet({
          cells: CELLS,
          ops: [{op: 'spliceRows', start: 3, count: 8}],
          read: ['A1', 'A2', 'A3'],
        });
        assert.strictEqual(rowCount, 2, 'removing rows 3..10 leaves 2 rows, not all 10');
        assert.strictEqual(cells.A3, null, 'no row survives past row 2');
      },
    },
    {
      name: 'a count larger than the rows present still clears the tail (never a no-op)',
      baseline: 'pass',
      expect(api, assert) {
        // Over-large count must clamp to "remove the rest", not silently remove nothing.
        const {rowCount} = api.mutateWorksheet({
          cells: CELLS,
          ops: [{op: 'spliceRows', start: 3, count: 200}],
          read: [],
        });
        assert.strictEqual(rowCount, 2, 'an over-large count removes the tail, not zero rows');
      },
    },
  ],
};
