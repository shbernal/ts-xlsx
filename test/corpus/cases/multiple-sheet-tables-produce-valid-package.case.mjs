// Cluster: tables
//
// Real-world scenario: a program loops over a collection, adding one worksheet per item and a table
// (plus a data validation) to each. The generated file must open cleanly — not in Excel's protected /
// "needs repair" view — and the data validations must not be stripped during repair. The classic
// cause of this failure is table part id collisions: if the tables across sheets do not get unique
// ids, Excel rejects the package. Building many sheets-with-tables in a loop must produce a valid
// package with unique table ids and every sheet's validation intact.

/** @typedef {{ name: string, baseline: 'pass'|'fail', expect: (api: any, assert: any) => Promise<void>|void }} Behavior */

export default {
  id: 'multiple-sheet-tables-produce-valid-package',
  provenance: {source: 'upstream-issue'},
  cluster: 'tables',
  description:
    'Adding a table (and a data validation) to each of many worksheets in a loop produces a valid ' +
    'package with unique table part ids that reloads cleanly, with each sheet’s validation preserved.',

  /** @type {Behavior[]} */
  behavior: [
    {
      name: 'many sheets each with a table write and reload without error',
      baseline: 'pass',
      async expect(api, assert) {
        const {writeOk, writeError, reloadOk, tableCount} = await api.multiSheetTableReport();
        assert.strictEqual(writeOk, true, `the write must succeed; got ${JSON.stringify(writeError)}`);
        assert.strictEqual(reloadOk, true, 'the package reloads');
        assert.strictEqual(tableCount, 5, 'one table part per sheet is emitted');
      },
    },
    {
      name: 'table part ids are unique across the package',
      baseline: 'pass',
      async expect(api, assert) {
        const {idsUnique} = await api.multiSheetTableReport();
        assert.strictEqual(idsUnique, true, 'each table gets a distinct id (a collision forces repair)');
      },
    },
    {
      name: 'a per-sheet data validation survives the round-trip',
      baseline: 'pass',
      async expect(api, assert) {
        const {firstSheetDvSurvives} = await api.multiSheetTableReport();
        assert.strictEqual(firstSheetDvSurvives, true, 'the list validation is not stripped');
      },
    },
  ],
};
