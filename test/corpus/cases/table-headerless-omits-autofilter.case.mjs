// Cluster: tables
//
// Real-world scenario: a user adds a table with the header row turned off — a headerless table that
// shows only data rows. In OOXML a table's autoFilter and its header-row count are only meaningful
// when the table actually has a header row: the autoFilter anchors its filter dropdowns on the
// header cells. A headerless table must therefore set headerRowCount="0" and emit NO autoFilter
// element. The writer already zeroes the header-row count, but it still injects an autoFilter for
// the headerless table — an internally inconsistent table part that Excel flags as corrupt and
// repairs by stripping the AutoFilter (and sometimes the whole table) on open. A header-bearing
// table, by contrast, legitimately carries both an autoFilter and a header-row count of 1.

/** @typedef {{ name: string, baseline: 'pass'|'fail', expect: (api: any, assert: any) => Promise<void>|void }} Behavior */

const headerlessSpec = {
  sheets: [{name: 'S', tables: [{name: 'Headerless', ref: 'A1', headerRow: false, headers: ['H1', 'H2'], rows: [['a', 1], ['b', 2]]}]}],
};
const headeredSpec = {
  sheets: [{name: 'S', tables: [{name: 'Headered', ref: 'A1', headerRow: true, headers: ['H1', 'H2'], rows: [['a', 1], ['b', 2]]}]}],
};

export default {
  id: 'table-headerless-omits-autofilter',
  provenance: {source: 'upstream-issue'},
  cluster: 'tables',
  description:
    'A table written with the header row disabled sets headerRowCount to 0 and emits no autoFilter ' +
    '(autoFilter is only valid over a header row), so the table part stays valid instead of being ' +
    'repaired away; a header-bearing table still carries its autoFilter and a header-row count of 1.',

  /** @type {Behavior[]} */
  behavior: [
    {
      name: 'a headerless table sets its header-row count to 0',
      baseline: 'pass',
      async expect(api, assert) {
        const {tables} = await api.inspectPackage(headerlessSpec);
        assert.strictEqual(tables.length, 1, 'precondition: one table part is written');
        assert.strictEqual(tables[0].headerRowCount, '0', 'a headerless table declares headerRowCount="0"');
      },
    },
    {
      name: 'a headerless table emits no autoFilter (autoFilter is only valid with a header row)',
      baseline: 'fail',
      async expect(api, assert) {
        const {tables} = await api.inspectPackage(headerlessSpec);
        assert.strictEqual(
          tables[0].autoFilterRef,
          null,
          `a headerless table must not emit an autoFilter; got ref ${JSON.stringify(tables[0].autoFilterRef)}`
        );
      },
    },
    {
      name: 'the emitted headerless table part is well-formed XML',
      baseline: 'pass',
      async expect(api, assert) {
        const {tables} = await api.inspectPackage(headerlessSpec);
        assert.strictEqual(tables[0].xmlWellFormed, true, 'the table part must be well-formed');
      },
    },
    {
      name: 'a header-bearing table still carries its autoFilter and a header-row count of 1',
      baseline: 'pass',
      async expect(api, assert) {
        const {tables} = await api.inspectPackage(headeredSpec);
        assert.strictEqual(tables[0].headerRowCount, '1', 'a headered table declares headerRowCount="1"');
        assert.ok(tables[0].autoFilterRef, 'a headered table emits an autoFilter over its header row');
      },
    },
  ],
};
