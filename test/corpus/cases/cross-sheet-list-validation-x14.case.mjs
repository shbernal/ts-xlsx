// Cluster: data-validation
//
// Real-world scenario: a "list" data validation on a range (e.g. B2:B16) has a dropdown source on
// a DIFFERENT worksheet (e.g. Dropdown!$D$4:$D$8). A legacy inline <dataValidations> element
// cannot reference another sheet, so Excel writes the validation only in the worksheet's
// extension list — an x14:dataValidations block whose <xm:f> holds the cross-sheet source and
// <xm:sqref> the target range — with NO legacy element at all. A reader that understands only the
// legacy element sees no validation, so the dropdown is dropped on read and lost on save. The
// cross-sheet list validation must be detected on read and preserved on round-trip.

/** @typedef {{ name: string, baseline: 'pass'|'fail', expect: (api: any, assert: any) => Promise<void>|void }} Behavior */

const FIXTURE = 'cross-sheet-list-validation-x14/sample.xlsx';

export default {
  id: 'cross-sheet-list-validation-x14',
  provenance: {source: 'upstream-issue', repo: 'exceljs/exceljs', ref: 1900},
  cluster: 'data-validation',
  description:
    'A list data validation whose source references another worksheet — stored only in the x14 ' +
    'extension list — is detected on read and preserved on a round-trip, instead of being ' +
    'silently dropped because only the legacy element is understood.',

  /** @type {Behavior[]} */
  behavior: [
    {
      name: 'the cross-sheet list validation is detected on read',
      baseline: 'pass',
      async expect(api, assert) {
        const {count} = await api.readFixtureValidations(FIXTURE);
        assert.ok(count >= 1, `the x14 list validation must be read; got ${count} validations`);
      },
    },
    {
      name: 'the x14 validation survives a read→write round-trip',
      baseline: 'pass',
      async expect(api, assert) {
        const {totalExt} = await api.roundtripFixtureValidationXml(FIXTURE);
        assert.ok(
          totalExt >= 1,
          `the extended (x14) validation must be re-serialized; got ${totalExt}`,
        );
      },
    },
  ],
};
