// Cluster: merged-cells
//
// Real-world scenario: a workbook merges a rectangular range (e.g. a 2×2 block). In OOXML a merged
// region is a single master (top-left) cell carrying the visible value; the other covered cells are
// subordinate and must not carry an independent value. A user sets a value by addressing a non-master
// (slave) cell inside the merge — e.g. the bottom-right cell. That write must resolve to the region's
// master, not silently create a competing value on a slave cell (which is malformed and can be
// dropped or rejected by spreadsheet applications). After a round-trip the merge span is intact with
// exactly one value on the master.

/** @typedef {{ name: string, baseline: 'pass'|'fail', expect: (api: any, assert: any) => Promise<void>|void }} Behavior */

export default {
  id: 'merged-range-slave-cell-write-resolves-to-master',
  provenance: {source: 'upstream-issue'},
  cluster: 'merged-cells',
  description:
    'Setting a value by addressing a non-master cell inside a merged range resolves to the master: ' +
    'only the master carries a value in the worksheet XML, the merge span is preserved, and the value ' +
    'is read back on both the master and (via the merge) the addressed cell.',

  /** @type {Behavior[]} */
  behavior: [
    {
      name: 'only the master cell carries a value in the worksheet XML',
      baseline: 'pass',
      async expect(api, assert) {
        const {cellsWithValue} = await api.mergeSlaveWrite();
        assert.deepStrictEqual(cellsWithValue, ['A1'], `only the master A1 must carry a value; got ${JSON.stringify(cellsWithValue)}`);
      },
    },
    {
      name: 'the merge span is declared over the same range',
      baseline: 'pass',
      async expect(api, assert) {
        const {merges} = await api.mergeSlaveWrite();
        assert.deepStrictEqual(merges, ['A1:B2'], 'the merge span survives');
      },
    },
    {
      name: 'the value written via the slave cell is read back on the master',
      baseline: 'pass',
      async expect(api, assert) {
        const {masterValue, slaveValue} = await api.mergeSlaveWrite();
        assert.strictEqual(masterValue, 'slave-write', 'the master carries the value written through the slave');
        assert.strictEqual(slaveValue, 'slave-write', 'reading the slave returns the merged region’s value');
      },
    },
  ],
};
