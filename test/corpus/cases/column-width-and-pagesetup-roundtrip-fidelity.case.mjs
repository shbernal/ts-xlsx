// Cluster: styles
//
// Real-world scenario: a user opens a real workbook with carefully-set column widths and print
// settings (scale, fit-to-width, page order, orientation), makes no changes, and saves. Every
// column width must come back to the exact fractional character-unit value, and the pageSetup
// attributes must survive — the "open a template and re-save" path must be faithful, independent
// of the library's assumed default font/DPI.

/** @typedef {{ name: string, baseline: 'pass'|'fail', expect: (api: any, assert: any) => Promise<void>|void }} Behavior */

const FIXTURE = 'column-width-and-pagesetup-roundtrip-fidelity/sample.xlsx';

export default {
  id: 'column-width-and-pagesetup-roundtrip-fidelity',
  provenance: {source: 'upstream-issue', repo: 'exceljs/exceljs', ref: 816},
  cluster: 'styles',
  description:
    'A no-op read→write round-trip preserves each column’s exact fractional width and the ' +
    'worksheet pageSetup (scale, fit-to-width/height, page order, orientation) rather than ' +
    'corrupting or dropping them.',

  /** @type {Behavior[]} */
  behavior: [
    {
      name: 'every column width survives the round-trip exactly',
      baseline: 'pass',
      async expect(api, assert) {
        const {source, rewritten} = await api.roundtripFixtureStyleFacts(FIXTURE);
        assert.ok(
          source.columnWidths.length >= 1,
          'precondition: the file has custom column widths',
        );
        assert.deepStrictEqual(
          rewritten.columnWidths,
          source.columnWidths,
          'the fractional widths are preserved exactly',
        );
      },
    },
    {
      name: 'the pageSetup print attributes survive the round-trip',
      baseline: 'pass',
      async expect(api, assert) {
        const {source, rewritten} = await api.roundtripFixtureStyleFacts(FIXTURE);
        assert.deepStrictEqual(
          rewritten.pageSetup,
          source.pageSetup,
          'scale/fit/pageOrder/orientation are preserved',
        );
      },
    },
  ],
};
