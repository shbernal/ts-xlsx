// Cluster: xlsx-io
//
// Real-world scenario: an existing .xlsx contains a worksheet named "History" — a Microsoft-reserved
// worksheet name Excel forbids users from creating manually, yet a file that already has such a sheet
// opens fine in Excel. When the library loads this workbook it throws, because it applies its
// reserved-name validation guard to the incoming name. The guard is meant to protect a user who is
// creating or renaming a sheet through the API — not to reject sheets that already exist in a
// foreign-generated file. Loading a pre-existing file must never fail on name validation.
//
// The fixture is a workbook whose sheet was renamed to "History" in the stored XML, reproducing the
// foreign file without the API's own guard blocking its authoring.

/** @typedef {{ name: string, baseline: 'pass'|'fail', expect: (api: any, assert: any) => Promise<void>|void }} Behavior */

const FIXTURE = 'load-workbook-with-reserved-sheet-name/sample.xlsx';

export default {
  id: 'load-workbook-with-reserved-sheet-name',
  provenance: {source: 'upstream-issue'},
  cluster: 'xlsx-io',
  description:
    'Loading a workbook that already contains a worksheet with a Microsoft-reserved name ("History") ' +
    'succeeds without throwing and exposes the sheet with its original name — the reserved-name guard ' +
    'applies to API-driven creation, not to reading an existing file.',

  /** @type {Behavior[]} */
  behavior: [
    {
      name: 'a workbook with a reserved-named sheet loads without throwing',
      baseline: 'fail',
      async expect(api, assert) {
        const {ok, error} = await api.readFixtureReport(FIXTURE);
        assert.strictEqual(ok, true, `a pre-existing reserved sheet name must not abort the load; got ${JSON.stringify(error)}`);
      },
    },
    {
      name: 'the reserved-named sheet is exposed with its original name preserved',
      baseline: 'fail',
      async expect(api, assert) {
        const {sheetNames} = await api.readFixtureReport(FIXTURE);
        assert.ok(sheetNames && sheetNames.includes('History'), `the History sheet must survive; got ${JSON.stringify(sheetNames)}`);
      },
    },
  ],
};
