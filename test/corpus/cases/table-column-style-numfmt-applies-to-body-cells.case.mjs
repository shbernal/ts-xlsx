// Cluster: tables
//
// Real-world scenario: a table is created with a per-column style — a number format on one column
// (e.g. a thousands-separated decimal) — intending that format to apply to that column's body cells.
// The per-column style must be merged into each affected cell's style so the produced package is
// valid and the column's body cells render with the requested format. A reported failure had the
// per-column style corrupt the written file; the correct behavior is a clean write where the styled
// column's body cells carry the numFmt and other columns are unaffected.

/** @typedef {{ name: string, baseline: 'pass'|'fail', expect: (api: any, assert: any) => Promise<void>|void }} Behavior */

const NUMFMT = '#,##0.00';

export default {
  id: 'table-column-style-numfmt-applies-to-body-cells',
  provenance: {source: 'upstream-issue'},
  cluster: 'tables',
  description:
    'A table column with a numFmt style writes a valid package whose body cells for that column carry ' +
    'the requested number format, while columns without a per-column style are unaffected.',

  /** @type {Behavior[]} */
  behavior: [
    {
      name: 'a table with a styled column writes and reloads without error',
      baseline: 'pass',
      async expect(api, assert) {
        const {writeOk, reloadOk, writeError} = await api.tableColumnStyleReport(NUMFMT);
        assert.strictEqual(
          writeOk,
          true,
          `writing must not throw; got ${JSON.stringify(writeError)}`,
        );
        assert.strictEqual(reloadOk, true, 'the package reloads without error');
      },
    },
    {
      name: 'the styled column’s body cells carry the requested number format',
      baseline: 'pass',
      async expect(api, assert) {
        const {styledBody} = await api.tableColumnStyleReport(NUMFMT);
        assert.deepStrictEqual(
          styledBody,
          [NUMFMT, NUMFMT],
          `both body cells of the styled column carry the numFmt; got ${JSON.stringify(styledBody)}`,
        );
      },
    },
    {
      name: 'columns without a per-column style are unaffected',
      baseline: 'pass',
      async expect(api, assert) {
        const {unstyledBody} = await api.tableColumnStyleReport(NUMFMT);
        assert.ok(
          unstyledBody.every((f) => f !== NUMFMT),
          `the unstyled column must not pick up the numFmt; got ${JSON.stringify(unstyledBody)}`,
        );
      },
    },
  ],
};
