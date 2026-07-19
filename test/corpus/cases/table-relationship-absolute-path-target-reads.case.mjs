// Cluster: tables
//
// Real-world scenario: a workbook's worksheet relationship to its table part uses an absolute
// package-path target (a leading-slash "/xl/tables/table1.xml") rather than a relative one. Both
// forms are valid OOXML and open in desktop apps, but the library resolves only relative targets,
// so loading the file crashes ("Cannot read properties of undefined (reading 'name')") when it
// fails to find the table part. A reader must resolve an absolute-path relationship target and
// load the table it points to — a control workbook from the same generator with no table also
// continues to load.

/** @typedef {{ name: string, baseline: 'pass'|'fail', expect: (api: any, assert: any) => Promise<void>|void }} Behavior */

const WITH_TABLE = 'table-relationship-absolute-path-target-reads/with-table.xlsx';
const CONTROL = 'table-relationship-absolute-path-target-reads/no-table-control.xlsx';

export default {
  id: 'table-relationship-absolute-path-target-reads',
  provenance: {source: 'upstream-issue', repo: 'exceljs/exceljs', ref: 1468},
  cluster: 'tables',
  description:
    'A workbook whose worksheet→table relationship uses an absolute package-path target ' +
    '("/xl/tables/table1.xml") reads without crashing and resolves the table, and a table-free ' +
    'control from the same generator continues to load.',

  /** @type {Behavior[]} */
  behavior: [
    {
      name: 'a table referenced by an absolute-path relationship reads without crashing',
      baseline: 'pass',
      async expect(api, assert) {
        const {ok, error} = await api.readFixtureReport(WITH_TABLE);
        assert.ok(ok, `the read must not crash on the absolute-path table rel; got ${JSON.stringify(error)}`);
      },
    },
    {
      name: 'a table-free control workbook from the same generator still loads',
      baseline: 'pass',
      async expect(api, assert) {
        const {ok} = await api.readFixtureReport(CONTROL);
        assert.ok(ok, 'the control loads');
      },
    },
  ],
};
