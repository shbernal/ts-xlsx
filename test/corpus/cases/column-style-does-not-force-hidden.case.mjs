// Cluster: styles
//
// Real-world scenario: a template applies a number format (or other style) to a
// whole column so every value in it renders consistently, without touching the
// column's visibility. Attaching a style must not, as a side effect, mark the
// column hidden — a reported failure mode where styled columns vanished on open.
// The style must apply and the column must stay visible unless visibility was set
// explicitly.

/** @typedef {{ name: string, baseline: 'pass'|'fail', expect: (api: any, assert: any) => Promise<void> }} Behavior */

const STYLED = {sheets: [{name: 'S', columns: [{index: 2, numFmt: '#,##0'}]}]};
const STYLED_VISIBLE = {sheets: [{name: 'S', columns: [{index: 2, numFmt: '#,##0', hidden: false}]}]};

export default {
  id: 'column-style-does-not-force-hidden',
  provenance: {source: 'upstream-issue', repo: 'exceljs/exceljs', ref: 458},
  cluster: 'styles',
  description:
    'Applying a style (e.g. a number format) to a column does not, as a side effect, ' +
    'mark the column hidden; the format applies and the column stays visible unless ' +
    'hidden was set explicitly.',

  /** @type {Behavior[]} */
  behavior: [
    {
      name: 'a column given a number format is not silently hidden',
      baseline: 'pass',
      async expect(api, assert) {
        const col = (await api.roundtripWorkbook(STYLED)).sheets.S.columns[2];
        assert.strictEqual(col.hidden, false, 'a styled column must not become hidden');
      },
    },
    {
      name: 'the applied number format survives the round-trip',
      baseline: 'pass',
      async expect(api, assert) {
        const col = (await api.roundtripWorkbook(STYLED)).sheets.S.columns[2];
        assert.strictEqual(col.numFmt, '#,##0', 'the column number format should persist');
      },
    },
    {
      name: 'an explicit hidden:false with a style stays visible',
      baseline: 'pass',
      async expect(api, assert) {
        const col = (await api.roundtripWorkbook(STYLED_VISIBLE)).sheets.S.columns[2];
        assert.strictEqual(col.hidden, false);
      },
    },
  ],
};
