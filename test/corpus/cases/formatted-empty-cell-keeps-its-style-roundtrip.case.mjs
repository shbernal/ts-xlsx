// Cluster: xlsx-io
//
// Real-world scenario: a program paints a cell — a background fill, a border — to lay out a form or
// a coloured grid, but never puts a value in it. On write-then-read the formatting is lost because
// the writer serialises only cells that hold a value, dropping the style-only ones. A formatted-but-
// empty cell is a real cell to Excel (`<c r=".." s=".."/>` with no `<v>`): its style must survive the
// round-trip, the cell must stay value-less, and a cell that carries neither a value nor a style must
// not be fabricated into existence.

/** @typedef {{ name: string, baseline: 'pass'|'fail', expect: (api: any, assert: any) => Promise<void>|void }} Behavior */

export default {
  id: 'formatted-empty-cell-keeps-its-style-roundtrip',
  provenance: {source: 'upstream-issue'},
  cluster: 'xlsx-io',
  description:
    'A cell given a fill or border but no value keeps that style across a write→read round-trip and ' +
    'stays value-less; a cell with neither a value nor a style of its own is not written at all.',

  /** @type {Behavior[]} */
  behavior: [
    {
      name: 'a fill on a value-less cell survives the round-trip',
      baseline: 'pass',
      async expect(api, assert) {
        const {filledArgb} = await api.styledEmptyCellReport();
        assert.strictEqual(filledArgb, 'FF00FF00', 'the fill on the empty cell must not be dropped on write');
      },
    },
    {
      name: 'the formatted empty cell reads back empty, with no value invented',
      baseline: 'pass',
      async expect(api, assert) {
        const {filledValue} = await api.styledEmptyCellReport();
        assert.strictEqual(filledValue, null, 'a styled-but-empty cell stays empty, not fabricated a value');
      },
    },
    {
      name: 'a border on a value-less cell survives while the cell stays empty',
      baseline: 'pass',
      async expect(api, assert) {
        const {borderedStyle, borderedValue} = await api.styledEmptyCellReport();
        assert.strictEqual(borderedStyle, 'thin', 'the border edge on the empty cell must survive');
        assert.strictEqual(borderedValue, null, 'the bordered empty cell stays empty');
      },
    },
    {
      name: 'a cell with neither a value nor a style is not fabricated',
      baseline: 'pass',
      async expect(api, assert) {
        const {untouched} = await api.styledEmptyCellReport();
        assert.strictEqual(untouched, false, 'a merely-touched cell must not gain a value or style');
      },
    },
  ],
};
