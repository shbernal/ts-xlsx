// Cluster: xlsx-io
//
// Real-world scenario: a workbook produced by a foreign generator (HanCell) omits optional
// document properties — its docProps/app.xml has no <Company> element. A reader that assumes the
// element is present crashes dereferencing it ("Cannot read properties of undefined (reading
// 'company')"), so the whole file is unreadable over a missing optional property. Absent optional
// document properties must read back as unset, not as a fatal error, and the worksheets must load.

/** @typedef {{ name: string, baseline: 'pass'|'fail', expect: (api: any, assert: any) => Promise<void>|void }} Behavior */

const FIXTURE = 'foreign-sheet-missing-company-property/sample.xlsx';

export default {
  id: 'foreign-sheet-missing-company-property',
  provenance: {source: 'upstream-issue', repo: 'exceljs/exceljs', ref: 3014},
  cluster: 'xlsx-io',
  description:
    'A workbook whose docProps/app.xml omits optional properties (e.g. Company) reads without ' +
    'crashing — a missing optional document property is treated as unset, and the worksheets load.',

  /** @type {Behavior[]} */
  behavior: [
    {
      name: 'a workbook missing an optional Company property reads without throwing',
      baseline: 'pass',
      async expect(api, assert) {
        const {ok, error} = await api.readFixtureReport(FIXTURE);
        assert.ok(ok, `a missing optional property must not be fatal; got ${JSON.stringify(error)}`);
      },
    },
    {
      name: 'the worksheets are accessible after loading',
      baseline: 'pass',
      async expect(api, assert) {
        const {sheetNames} = await api.readFixtureReport(FIXTURE);
        assert.ok(sheetNames && sheetNames.length >= 1, 'at least one worksheet is exposed');
      },
    },
  ],
};
