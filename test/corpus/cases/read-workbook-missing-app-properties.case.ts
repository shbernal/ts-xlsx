// Cluster: address-decoding
//
// Real-world scenario: spreadsheet editors and export libraries other than Excel frequently produce
// .xlsx packages that omit the docProps/app.xml part entirely, or ship a minimal one without the
// extended application properties (company, manager). Opening such a workbook must degrade
// gracefully — surface whatever core content is present and leave the absent app-properties fields
// unset — rather than dereferencing an undefined properties object and throwing. Re-saving through
// Excel (which injects a full app.xml) happens to sidestep the crash, confirming the missing part is
// the trigger; the library must read valid foreign-generated files without needing that round-trip.
//
// The fixture is a workbook with its docProps/app.xml part (and the matching content-type override)
// removed, reproducing the foreign-generator shape.

import type {Assert, Case, CorpusApi} from '../case.ts';

const FIXTURE = 'read-workbook-missing-app-properties/sample.xlsx';

export default {
  id: 'read-workbook-missing-app-properties',
  provenance: {source: 'upstream-issue'},
  cluster: 'address-decoding',
  description:
    'Reading an .xlsx package that has no docProps/app.xml part completes without throwing and the ' +
    'worksheet content is accessible — the absent application-properties are left unset rather than ' +
    'crashing the load or being fabricated.',

  behavior: [
    {
      name: 'a workbook missing docProps/app.xml loads without throwing',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const {ok, error} = await api.readFixtureReport(FIXTURE);
        assert.strictEqual(
          ok,
          true,
          `a missing app.xml must not abort the load; got ${JSON.stringify(error)}`,
        );
      },
    },
    {
      name: 'the worksheets are recovered intact',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const {sheetNames} = await api.readFixtureReport(FIXTURE);
        assert.deepStrictEqual(
          sheetNames,
          ['Sheet1', 'Data'],
          'both worksheets survive the tolerant read',
        );
      },
    },
    {
      name: 'cell values are accessible despite the missing application properties',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const cells = await api.readFixtureCells(FIXTURE, ['A1']);
        assert.strictEqual(cells.A1.value, 'hi', 'the first sheet’s cell value reads back');
      },
    },
  ],
} satisfies Case;
