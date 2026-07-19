// Cluster: tables
//
// Real-world scenario: a worksheet defines a table with more than three columns. On load, every
// declared column must be exposed with its correct header name — the column collection length equals
// the number of columns in the table definition, not a fixed cap. A historical defect truncated a
// loaded table's columns to three; this locks the full-width read as a regression guard. (A loaded
// table also exposing its data rows is covered separately by loaded-table-exposes-data-rows.)

/** @typedef {{ name: string, baseline: 'pass'|'fail', expect: (api: any, assert: any) => Promise<void>|void }} Behavior */

export default {
  id: 'loaded-wide-table-exposes-all-columns',
  provenance: {source: 'upstream-issue'},
  cluster: 'tables',
  description:
    'A table with five columns, written and read back, exposes all five columns with their correct ' +
    'header names in order — not truncated to a fixed cap of three.',

  /** @type {Behavior[]} */
  behavior: [
    {
      name: 'all five columns survive a write/read round-trip',
      baseline: 'pass',
      async expect(api, assert) {
        const {colCount} = await api.wideTableColumnReadReport();
        assert.strictEqual(
          colCount,
          5,
          'the loaded table exposes all five columns, not a cap of three',
        );
      },
    },
    {
      name: 'each loaded column carries its correct header name in order',
      baseline: 'pass',
      async expect(api, assert) {
        const {colNames} = await api.wideTableColumnReadReport();
        assert.deepStrictEqual(
          colNames,
          ['C1', 'C2', 'C3', 'C4', 'C5'],
          'the header names match the source order',
        );
      },
    },
  ],
};
