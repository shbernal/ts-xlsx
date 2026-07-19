// Cluster: tables
//
// Real-world scenario: a program opens a real .xlsx that already contains a defined table and
// wants to read its data by table name. The table is found and its column names are exposed, but
// its data rows come back undefined — the loader reconstructs the table's name, columns, and
// range but never repopulates its data rows from the on-sheet cells. So reading a column's values
// or the table height on a loaded table yields nothing (or throws). A table read from a file must
// expose its data rows, populated from the sheet, exactly like a freshly-created table.

/** @typedef {{ name: string, baseline: 'pass'|'fail', expect: (api: any, assert: any) => Promise<void>|void }} Behavior */

const FIXTURE = 'loaded-table-exposes-data-rows/sample.xlsx';
const TABLE = 'tbl_dictionary';

export default {
  id: 'loaded-table-exposes-data-rows',
  provenance: {source: 'upstream-issue', repo: 'exceljs/exceljs', ref: 2297},
  cluster: 'tables',
  description:
    'A table loaded from a real file is findable by name and exposes its declared column names, ' +
    'and — the open bug — must expose its data rows populated from the on-sheet cells rather than ' +
    'an undefined rows array.',

  /** @type {Behavior[]} */
  behavior: [
    {
      name: 'the table is findable by name and exposes its column names',
      baseline: 'pass',
      async expect(api, assert) {
        const {found, columns} = await api.readFixtureTable(FIXTURE, TABLE);
        assert.ok(found, 'the table is located by name');
        assert.ok(columns && columns.length >= 1, 'the declared column names are exposed');
        assert.ok(columns.includes('Name'), 'a known column is present');
      },
    },
    {
      name: 'the loaded table exposes its data rows populated from the sheet',
      baseline: 'pass',
      async expect(api, assert) {
        const {rowCount} = await api.readFixtureTable(FIXTURE, TABLE);
        assert.ok(
          typeof rowCount === 'number' && rowCount > 0,
          `the loaded table must expose its data rows; got rowCount ${JSON.stringify(rowCount)}`,
        );
      },
    },
  ],
};
