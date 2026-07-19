// Cluster: security
//
// Real-world scenario: a caller hands the OOXML reader a legacy binary .xls file — the old
// OLE2 / Compound File Binary format (magic bytes D0 CF 11 E0), not a ZIP. The reader only
// understands the ZIP-based .xlsx container, so it extracts no worksheets. Instead of signalling
// that the input is the wrong format, it resolves to a workbook whose worksheets array is empty,
// giving the caller no clue the file was unreadable. Silent success on unparseable input is a
// hostile-input-facing defect: a non-ZIP payload must be rejected loudly with a catchable error,
// never presented as a valid-but-empty workbook.

/** @typedef {{ name: string, baseline: 'pass'|'fail', expect: (api: any, assert: any) => Promise<void>|void }} Behavior */

const FIXTURE = 'legacy-xls-binary-rejected-loudly/sample.xls';

export default {
  id: 'legacy-xls-binary-rejected-loudly',
  provenance: {source: 'upstream-issue', repo: 'exceljs/exceljs', ref: 2692},
  cluster: 'security',
  description:
    'Loading a legacy binary .xls (OLE2/CFB, non-ZIP) into the OOXML reader fails loudly with a ' +
    'catchable error rather than silently resolving to a workbook with an empty worksheets array.',

  /** @type {Behavior[]} */
  behavior: [
    {
      name: 'a non-ZIP legacy .xls payload does not resolve to a silently-empty workbook',
      baseline: 'pass',
      async expect(api, assert) {
        const {ok, sheetNames} = await api.readFixtureReport(FIXTURE);
        assert.ok(
          !(ok && (!sheetNames || sheetNames.length === 0)),
          'the reader must not succeed with zero sheets on a wrong-format file — it must reject it'
        );
      },
    },
  ],
};
