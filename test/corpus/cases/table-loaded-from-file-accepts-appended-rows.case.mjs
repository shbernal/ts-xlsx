// Cluster: tables
//
// Real-world scenario: a program opens an existing `.xlsx` that already contains a defined
// table, fetches it by name, and appends data rows. Today this throws "Cannot read
// properties of undefined (reading 'length')" because a table rehydrated from a file has no
// rows array populated — the loader reconstructs the table's name, columns, and range but
// never repopulates its data rows from the on-sheet cells. Every table-height or append
// operation on a loaded table then dereferences undefined. A table read from a file must
// expose its data rows so appends work exactly as they do on a freshly-created table.

/** @typedef {{ name: string, baseline: 'pass'|'fail', expect: (api: any, assert: any) => Promise<void>|void }} Behavior */

const SPEC = {
  sheets: [
    {name: 'S', tables: [{name: 'T', ref: 'A1', headers: ['H1', 'H2'], rows: [['a', 1], ['b', 2]]}]},
  ],
};

export default {
  id: 'table-loaded-from-file-accepts-appended-rows',
  provenance: {source: 'upstream-issue', repo: 'exceljs/exceljs', ref: 1062},
  cluster: 'tables',
  description:
    'A table read back from a file exposes its data rows and accepts appended rows, ' +
    'exactly like a freshly-created table — a loaded table must not be a half-hydrated ' +
    'model that throws when its height is read or a row is appended.',

  /** @type {Behavior[]} */
  behavior: [
    {
      name: 'the table is fetchable by name after a round-trip',
      baseline: 'pass',
      async expect(api, assert) {
        const {hasTable} = await api.roundtripTableAppend(SPEC, {tableName: 'T', appendRows: []});
        assert.ok(hasTable, 'getTable returns the reloaded table');
      },
    },
    {
      name: 'the reloaded table exposes its data rows',
      baseline: 'fail',
      async expect(api, assert) {
        const {loadedRowCount} = await api.roundtripTableAppend(SPEC, {tableName: 'T', appendRows: []});
        assert.strictEqual(loadedRowCount, 2, 'the two written data rows are rehydrated on load');
      },
    },
    {
      name: 'appending a row to the reloaded table succeeds',
      baseline: 'fail',
      async expect(api, assert) {
        const {addError, committed, finalRowCount} = await api.roundtripTableAppend(SPEC, {
          tableName: 'T',
          appendRows: [['c', 3]],
        });
        assert.strictEqual(addError, null, `append must not throw; got ${JSON.stringify(addError)}`);
        assert.ok(committed, 'the append commits');
        assert.strictEqual(finalRowCount, 3, 'the appended row is present after re-writing');
      },
    },
  ],
};
