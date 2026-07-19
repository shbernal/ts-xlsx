// Cluster: xlsx-io
//
// Real-world scenario: an existing .xlsx contains a worksheet named "History" — a Microsoft-reserved
// worksheet name Excel blocks in its UI (a Track-Changes/shared-workbook nicety), yet a file that
// already has such a sheet opens fine in Excel and "History" is a perfectly valid OOXML sheet name.
// The library over-applies that UI restriction to the document model: it throws both when loading a
// foreign file that contains a "History" sheet AND when a caller constructs one through the API.
// Neither should fail — only genuinely-invalid names (illegal characters, over-length, empty) belong
// in the reject set. Loading a pre-existing file must never fail on name validation, and adding a
// "History" sheet is a legitimate document-model operation.
//
// The load fixture is a workbook whose sheet was renamed to "History" in the stored XML, reproducing
// the foreign file without the API's own guard blocking its authoring. The add-side behaviors
// exercise the guard directly on construction.

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
      baseline: 'pass',
      async expect(api, assert) {
        const {ok, error} = await api.readFixtureReport(FIXTURE);
        assert.strictEqual(
          ok,
          true,
          `a pre-existing reserved sheet name must not abort the load; got ${JSON.stringify(error)}`,
        );
      },
    },
    {
      name: 'the reserved-named sheet is exposed with its original name preserved',
      baseline: 'pass',
      async expect(api, assert) {
        const {sheetNames} = await api.readFixtureReport(FIXTURE);
        assert.ok(
          sheetNames?.includes('History'),
          `the History sheet must survive; got ${JSON.stringify(sheetNames)}`,
        );
      },
    },
    {
      name: 'constructing a worksheet named "History" through the API is permitted',
      baseline: 'pass',
      async expect(api, assert) {
        const {addThrew, roundtripName, addError} = await api.addReservedSheetNameReport();
        assert.strictEqual(
          addThrew,
          false,
          `adding a "History" sheet must not throw; got ${JSON.stringify(addError)}`,
        );
        assert.strictEqual(
          roundtripName,
          'History',
          'the constructed History sheet round-trips with its name',
        );
      },
    },
    {
      name: 'a genuinely-invalid sheet name is still rejected',
      baseline: 'pass',
      async expect(api, assert) {
        const {invalidRejected} = await api.addReservedSheetNameReport();
        assert.strictEqual(
          invalidRejected,
          true,
          'a name with an illegal character (a/b) is still rejected',
        );
      },
    },
  ],
};
