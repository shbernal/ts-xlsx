// Cluster: xlsx-io
//
// Real-world scenario: a program opens a styled template `.xlsx` authored in desktop
// Excel — banded solid fills, bold headers, custom column widths, merged-cell layout
// across several sheets — fills in some cells, and saves it back. The saved copy must
// render the same: column widths, fills, fonts and number formats a template declares
// must survive a read→write round-trip. "Fill in a formatted template and re-save" is a
// mainstream use, and a lossy round-trip that resets widths to zero or drops fills makes
// the library unusable for it.
//
// Fixture `template.xlsx` (authored in Excel) has three sheets, a wide first column
// (width ~35.5, customWidth) beside many narrow custom-width columns (~4.898), merged
// header blocks, and cells carrying solid theme-tinted fills and bold fonts.

/** @typedef {{ name: string, baseline: 'pass'|'fail', expect: (api: any, assert: any) => Promise<void>|void }} Behavior */

const FIXTURE = 'template-styles-survive-read-write-roundtrip/template.xlsx';
const SHEET = 'Americas';

export default {
  id: 'template-styles-survive-read-write-roundtrip',
  provenance: {source: 'upstream-issue', repo: 'exceljs/exceljs', ref: 62},
  cluster: 'xlsx-io',
  description:
    'A styled template read from disk and written back unchanged keeps its sheet names, ' +
    'custom column widths, and per-cell styling (fills, fonts, number formats) — the ' +
    'format-preserving "open a template, fill it in, save it" path must not reset widths ' +
    'or drop styles.',

  /** @type {Behavior[]} */
  behavior: [
    {
      name: 'sheet names survive the no-op round-trip',
      baseline: 'pass',
      async expect(api, assert) {
        const {sheetNames, sheetNamesBefore} = await api.roundtripFixture(FIXTURE);
        assert.deepStrictEqual(sheetNamesBefore, ['Americas', 'Asia-Pacific', 'EMEA'], 'names as read');
        assert.deepStrictEqual(sheetNames, sheetNamesBefore, 'names must be identical after round-trip');
      },
    },
    {
      name: 'custom column widths survive the round-trip',
      baseline: 'pass',
      async expect(api, assert) {
        const {columns, columnsBefore} = await api.roundtripFixture(FIXTURE);
        const b = columnsBefore[SHEET];
        const a = columns[SHEET];
        assert.ok(b['1'] && b['1'].customWidth, 'template declares a custom-width first column');
        assert.strictEqual(a['1'].width, b['1'].width, 'the wide first column width must be preserved');
        assert.strictEqual(a['1'].customWidth, true, 'and stay flagged as a custom width');
        assert.strictEqual(a['2'].width, b['2'].width, 'a narrow custom column width must be preserved too');
      },
    },
    {
      name: 'per-cell fills and fonts survive the round-trip unchanged',
      baseline: 'pass',
      async expect(api, assert) {
        const {styleSurvival} = await api.roundtripFixture(FIXTURE);
        assert.ok(styleSurvival.checked > 0, 'the template must actually carry styled cells to check');
        assert.strictEqual(
          styleSurvival.mismatches,
          0,
          `every styled cell must round-trip; first drift: ${JSON.stringify(styleSurvival.sample)}`
        );
      },
    },
  ],
};
