// Cluster: streaming
//
// Real-world scenario: streaming a worksheet to disk, a producer wants to append a batch
// of rows in one call instead of looping row-by-row. The batch add must behave like adding
// each row individually — same values, same sequential row numbers — and must not commit or
// lock the sheet, so more rows can still be added after. In the streaming writer this batch
// convenience is declared in the types but absent at runtime, so the call throws
// "addRows is not a function". The single-row add is the working control.

/** @typedef {{ name: string, baseline: 'pass'|'fail', expect: (api: any, assert: any) => Promise<void>|void }} Behavior */

export default {
  id: 'streaming-write-add-rows-batch',
  provenance: {source: 'upstream-issue', repo: 'exceljs/exceljs', ref: 1277},
  cluster: 'streaming',
  description:
    'The streaming worksheet writer can append a batch of rows in one call, materializing ' +
    'each row’s values with correct sequential row numbers, identically to adding rows one ' +
    'at a time.',

  /** @type {Behavior[]} */
  behavior: [
    {
      name: 'adding a batch of rows to a streaming worksheet does not throw',
      baseline: 'fail',
      async expect(api, assert) {
        const {ok, error} = await api.streamWriteSheet({
          ops: [{op: 'addRows', value: [['a', 1], ['b', 2]]}],
          read: ['A1'],
        });
        assert.ok(ok, `streaming batch add must be supported; got error ${JSON.stringify(error)}`);
      },
    },
    {
      name: 'the batch-added rows carry their values in order',
      baseline: 'fail',
      async expect(api, assert) {
        const {ok, cells, rowCount} = await api.streamWriteSheet({
          ops: [{op: 'addRows', value: [['a', 1], ['b', 2]]}],
          read: ['A1', 'B1', 'A2', 'B2'],
        });
        assert.ok(ok, 'the batch add must complete to inspect the rows');
        assert.strictEqual(rowCount, 2, 'two rows appended');
        assert.deepStrictEqual([cells.A1, cells.B1], ['a', 1], 'first row values');
        assert.deepStrictEqual([cells.A2, cells.B2], ['b', 2], 'second row values');
      },
    },
    {
      name: 'adding rows one at a time works (control)',
      baseline: 'pass',
      async expect(api, assert) {
        const {ok, cells, rowCount} = await api.streamWriteSheet({
          ops: [{op: 'addRow', value: ['a', 1]}, {op: 'addRow', value: ['b', 2]}],
          read: ['A1', 'B2'],
        });
        assert.ok(ok, 'single-row streaming add works');
        assert.strictEqual(rowCount, 2, 'two rows written');
        assert.strictEqual(cells.A1, 'a', 'first row value');
        assert.strictEqual(cells.B2, 2, 'second row value');
      },
    },
  ],
};
