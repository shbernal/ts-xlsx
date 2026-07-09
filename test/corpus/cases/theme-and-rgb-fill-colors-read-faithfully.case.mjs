// Cluster: styles
//
// Real-world scenario: a workbook authored in a desktop app colors cells two ways — some with an
// explicit RGB fill, some with a theme color plus a tint (a shade derived from a palette entry).
// A reader must surface both faithfully: an RGB fill's foreground carries the concrete argb, and
// a theme+tint fill's foreground carries the theme index and its tint, so no fill color is lost
// or flattened on read.

/** @typedef {{ name: string, baseline: 'pass'|'fail', expect: (api: any, assert: any) => Promise<void>|void }} Behavior */

const FIXTURE = 'theme-and-rgb-fill-colors-read-faithfully/sample.xlsx';
const SHEET = '3 - Week Look Ahead';

export default {
  id: 'theme-and-rgb-fill-colors-read-faithfully',
  provenance: {source: 'upstream-issue', repo: 'exceljs/exceljs', ref: 1690},
  cluster: 'styles',
  description:
    'Cell fill colors are read faithfully whether expressed as an explicit RGB value (concrete ' +
    'argb on fgColor) or as a theme color with a tint (theme index + tint on fgColor) — a ' +
    'themed fill is not dropped as "missing color".',

  /** @type {Behavior[]} */
  behavior: [
    {
      name: 'an explicit RGB fill exposes a concrete argb foreground',
      baseline: 'pass',
      async expect(api, assert) {
        const styles = await api.readFixtureCellStyles(FIXTURE, [`${SHEET}!A1`]);
        assert.strictEqual(styles[`${SHEET}!A1`].fill.fgColor.argb, 'FFFFFF00', 'the RGB fill argb is read');
      },
    },
    {
      name: 'a theme+tint fill exposes its theme index and tint',
      baseline: 'pass',
      async expect(api, assert) {
        const styles = await api.readFixtureCellStyles(FIXTURE, [`${SHEET}!C2`]);
        const fg = styles[`${SHEET}!C2`].fill.fgColor;
        assert.strictEqual(fg.theme, 3, 'the theme index is read');
        assert.ok(typeof fg.tint === 'number' && fg.tint !== 0, `the tint is read (got ${JSON.stringify(fg.tint)})`);
      },
    },
  ],
};
