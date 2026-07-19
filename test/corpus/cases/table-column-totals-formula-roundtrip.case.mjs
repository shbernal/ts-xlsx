// Cluster: tables
//
// Real-world scenario: a table has one or more columns carrying a totals-row function (e.g. a SUM in
// the totals row), which the table part serializes as a non-leaf <tableColumn> element (it holds a
// child totalsRowFunction/formula rather than being self-closing). When such a workbook is read back,
// the loader must recognize every declared column and keep the auto-filter's per-column state (the
// filter-button flags) aligned to the same number of columns. A reported failure treated a column
// with a child element as a leaf and dropped it, yielding fewer parsed columns than declared and a
// filter-button index overrun that crashed the load.

/** @typedef {{ name: string, baseline: 'pass'|'fail', expect: (api: any, assert: any) => Promise<void>|void }} Behavior */

const SPEC = {
  sheets: [
    {
      name: 'S',
      tables: [
        {
          name: 'T',
          ref: 'A1',
          totalsRow: true,
          columnDefs: [
            {name: 'Item', totalsRowLabel: 'Total'},
            {name: 'Amount', totalsRowFunction: 'sum'},
          ],
          rows: [
            ['a', 1],
            ['b', 2],
          ],
        },
      ],
    },
  ],
};

export default {
  id: 'table-column-totals-formula-roundtrip',
  provenance: {source: 'upstream-issue'},
  cluster: 'tables',
  description:
    'A table with a totals-row function column (a non-leaf tableColumn element) loads without ' +
    'crashing, every declared column is recovered, and the totals-row function survives into the ' +
    'written table XML.',

  /** @type {Behavior[]} */
  behavior: [
    {
      name: 'a table with a totals-row formula column writes and round-trips without throwing',
      baseline: 'pass',
      async expect(api, assert) {
        const result = await api.tryWriteWorkbook(SPEC);
        assert.strictEqual(
          result.ok,
          true,
          `writing/reading must not throw; got ${JSON.stringify(result.error)}`,
        );
      },
    },
    {
      name: 'every declared table column is present in the written table part',
      baseline: 'pass',
      async expect(api, assert) {
        const {tables} = await api.inspectPackage(SPEC);
        assert.strictEqual(tables.length, 1, 'the table part is written');
        assert.strictEqual(
          tables[0].columnCount,
          2,
          'both columns are emitted, none dropped as a leaf',
        );
      },
    },
    {
      name: 'the table part is well-formed with a totals-row function',
      baseline: 'pass',
      async expect(api, assert) {
        const {tables} = await api.inspectPackage(SPEC);
        assert.strictEqual(tables[0].xmlWellFormed, true, 'the table XML is well-formed');
      },
    },
  ],
};
