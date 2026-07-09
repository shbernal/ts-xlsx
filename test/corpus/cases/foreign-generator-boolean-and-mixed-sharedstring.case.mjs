// Cluster: foreign-generator-tolerance
//
// Real-world scenario: a spreadsheet exported by a non-Excel cloud application (a collaborative
// online sheet tool) opens fine in desktop Excel but makes a JS spreadsheet library throw during
// parse. Opening it in Excel and re-saving unchanged produces a file that loads without issue; the
// two packages carry the same logical content and differ only in serialization style. The foreign
// generator writes things that are schema-valid but atypical of Excel: boolean attributes spelled
// "true"/"false" instead of "1"/"0", and a shared-string <si> whose leading empty <t/> is
// immediately followed by rich-text <r> runs (a mixed-content shape). The reader must tolerate both
// and load the same logical content as the Excel-normalized copy.

/** @typedef {{ name: string, baseline: 'pass'|'fail', expect: (api: any, assert: any) => Promise<void>|void }} Behavior */

const LARK = 'foreign-generator-boolean-and-mixed-sharedstring/lark-export.xlsx';
const NORMALIZED = 'foreign-generator-boolean-and-mixed-sharedstring/excel-normalized.xlsx';

export default {
  id: 'foreign-generator-boolean-and-mixed-sharedstring',
  provenance: {source: 'upstream-issue'},
  cluster: 'foreign-generator-tolerance',
  description:
    'A workbook from a foreign generator — boolean attributes spelled "true"/"false" and a ' +
    'shared-string <si> with a leading empty <t/> before rich-text runs — loads without throwing, ' +
    'yielding the same worksheet as the Excel-normalized copy of the same document.',

  /** @type {Behavior[]} */
  behavior: [
    {
      name: 'the Excel-normalized copy of the same document loads (control)',
      baseline: 'pass',
      async expect(api, assert) {
        const {ok, sheetNames} = await api.readFixtureReport(NORMALIZED);
        assert.ok(ok, 'the Excel-normalized copy must load');
        assert.ok(sheetNames && sheetNames.length >= 1, 'and expose its worksheet');
      },
    },
    {
      name: 'the foreign-exported original loads without throwing on the mixed shared-string shape',
      baseline: 'fail',
      async expect(api, assert) {
        const {ok, error} = await api.readFixtureReport(LARK);
        assert.ok(
          ok,
          `a schema-valid foreign export must load, not throw on its mixed <si>/boolean shapes; got ${JSON.stringify(error)}`
        );
      },
    },
    {
      name: 'the foreign export exposes the same worksheet as the normalized copy',
      baseline: 'fail',
      async expect(api, assert) {
        const foreign = await api.readFixtureReport(LARK);
        const normalized = await api.readFixtureReport(NORMALIZED);
        assert.deepStrictEqual(
          foreign.sheetNames,
          normalized.sheetNames,
          'the two interchangeable inputs must expose the same sheets'
        );
      },
    },
  ],
};
