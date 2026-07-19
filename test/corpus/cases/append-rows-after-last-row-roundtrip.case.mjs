// Cluster: streaming
//
// Real-world scenario: a user opens an existing workbook that already has a block of data rows on a
// sheet, appends new rows immediately after the last populated row, and saves. After the round-trip
// the original rows must be untouched and the appended rows must sit at the next contiguous indices —
// no blank gap before them, no overwrite of the existing data. The load-bearing fact is that a
// loaded worksheet reports its last populated row so `addRow` places new content at N+1.

/** @typedef {{ name: string, baseline: 'pass'|'fail', expect: (api: any, assert: any) => Promise<void>|void }} Behavior */

const INITIAL = [
  ['a', 1],
  ['b', 2],
  ['c', 3],
];
const APPEND = [
  ['d', 4],
  ['e', 5],
];

export default {
  id: 'append-rows-after-last-row-roundtrip',
  provenance: {source: 'upstream-issue'},
  cluster: 'streaming',
  description:
    'Appending rows to a loaded workbook places them at contiguous indices immediately after the ' +
    'last populated row — no gap, no overwrite — and both the original and appended rows survive a ' +
    'write/reload with correct values.',

  /** @type {Behavior[]} */
  behavior: [
    {
      name: 'a loaded worksheet reports its last populated row so new content lands at N+1',
      baseline: 'pass',
      async expect(api, assert) {
        const {loadedRowCount} = await api.appendRowsAfterReload(INITIAL, APPEND);
        assert.strictEqual(
          loadedRowCount,
          INITIAL.length,
          'the loaded sheet reports 3 populated rows',
        );
      },
    },
    {
      name: 'appended rows land contiguously after the existing data with no gap or overwrite',
      baseline: 'pass',
      async expect(api, assert) {
        const {finalRowCount, rows} = await api.appendRowsAfterReload(INITIAL, APPEND);
        assert.strictEqual(
          finalRowCount,
          INITIAL.length + APPEND.length,
          'row count grows by exactly the appended rows',
        );
        assert.deepStrictEqual(
          rows[3],
          ['d', 4],
          'the first appended row is at index 4, right after the originals',
        );
        assert.deepStrictEqual(rows[4], ['e', 5], 'the second appended row is at index 5');
      },
    },
    {
      name: 'the original rows retain their values unchanged after the round-trip',
      baseline: 'pass',
      async expect(api, assert) {
        const {rows} = await api.appendRowsAfterReload(INITIAL, APPEND);
        assert.deepStrictEqual(
          [rows[0], rows[1], rows[2]],
          INITIAL,
          'the original three rows are untouched',
        );
      },
    },
  ],
};
