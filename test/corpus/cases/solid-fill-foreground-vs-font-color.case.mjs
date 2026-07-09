// Cluster: styles
//
// Real-world scenario: a cell has a solid background fill AND a distinct font color (e.g. a
// yellow fill behind red text). In the OOXML solid-pattern fill model the *visible* background
// color is the pattern foreground (`fgColor`), while `bgColor` is an automatic placeholder
// (indexed 64). The font color is an entirely separate style facet. A reader must surface the
// visible fill color on `fill.fgColor`, keep the automatic `bgColor` distinct, and expose the
// font color independently — never conflating fill and font color.

/** @typedef {{ name: string, baseline: 'pass'|'fail', expect: (api: any, assert: any) => Promise<void>|void }} Behavior */

const FIXTURE = 'solid-fill-foreground-vs-font-color/sample.xlsx';

export default {
  id: 'solid-fill-foreground-vs-font-color',
  provenance: {source: 'upstream-issue', repo: 'exceljs/exceljs', ref: 2325},
  cluster: 'styles',
  description:
    'A solid-pattern fill exposes its visible color on fill.fgColor with an automatic indexed ' +
    'bgColor, and the cell font color is surfaced independently on font.color — fill color and ' +
    'font color are never conflated.',

  /** @type {Behavior[]} */
  behavior: [
    {
      name: 'the visible solid-fill color is on fgColor, with an automatic indexed bgColor',
      baseline: 'pass',
      async expect(api, assert) {
        const {'Sheet1!A1': a1} = await api.readFixtureCellStyles(FIXTURE, ['Sheet1!A1']);
        assert.strictEqual(a1.fill.pattern, 'solid', 'the fill is a solid pattern');
        assert.strictEqual(a1.fill.fgColor.argb, 'FFFFFF00', 'the visible yellow is on fgColor');
        assert.strictEqual(a1.fill.bgColor.indexed, 64, 'bgColor stays the automatic indexed placeholder');
      },
    },
    {
      name: 'the font color is surfaced independently of the fill color',
      baseline: 'pass',
      async expect(api, assert) {
        const {'Sheet1!A1': a1} = await api.readFixtureCellStyles(FIXTURE, ['Sheet1!A1']);
        assert.strictEqual(a1.fontColor.argb, 'FFFF0000', 'the red font color is separate from the yellow fill');
      },
    },
  ],
};
