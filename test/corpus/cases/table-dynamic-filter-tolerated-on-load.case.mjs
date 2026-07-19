// Cluster: tables
//
// Real-world scenario: an Excel table can carry a column-level dynamic filter — for example the
// built-in "Above Average" filter — persisted in the table part's autoFilter as a self-closing
// <dynamicFilter type="aboveAverage" val="…"/> inside a <filterColumn>. A workbook carrying such a
// table must open: reading it must not abort on the dynamicFilter node. The file is otherwise valid
// and opens cleanly in Excel; a reader that rejects the exotic filter node would refuse a whole
// legitimate workbook. (Round-tripping the filter data itself is a separate concern — this case
// locks that the load is tolerant and the table/worksheet survive.)
//
// The fixture is authored by building a normal table and injecting a column-level dynamicFilter into
// its table-part autoFilter, reproducing the shape Excel emits without depending on Excel to write it.

/** @typedef {{ name: string, baseline: 'pass'|'fail', expect: (api: any, assert: any) => Promise<void>|void }} Behavior */

const FIXTURE = 'table-dynamic-filter-tolerated/sample.xlsx';

export default {
  id: 'table-dynamic-filter-tolerated-on-load',
  provenance: {source: 'upstream-issue'},
  cluster: 'tables',
  description:
    'Loading a workbook whose table carries a column-level dynamicFilter (e.g. Above Average) ' +
    'completes without throwing and the worksheet survives, rather than the exotic filter node ' +
    'aborting the entire read.',

  /** @type {Behavior[]} */
  behavior: [
    {
      name: 'a table with a column-level dynamicFilter loads without throwing',
      baseline: 'pass',
      async expect(api, assert) {
        const {ok, error} = await api.readFixtureReport(FIXTURE);
        assert.strictEqual(
          ok,
          true,
          `the dynamicFilter must be tolerated, not abort the load; got ${JSON.stringify(error)}`,
        );
      },
    },
    {
      name: 'the worksheet is recovered intact',
      baseline: 'pass',
      async expect(api, assert) {
        const {sheetNames} = await api.readFixtureReport(FIXTURE);
        assert.deepStrictEqual(sheetNames, ['Sheet1'], 'the sheet survives the tolerant read');
      },
    },
  ],
};
