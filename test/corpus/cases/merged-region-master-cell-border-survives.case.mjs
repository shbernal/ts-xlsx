// Cluster: styles
//
// Real-world scenario: a cell carries border styling (a thin top border, a medium bottom border) and
// is then made the top-left/master cell of a merged region. The master cell must retain the border
// (it renders the merged region's outline) and its other style facets (number format, font) through
// the merge and a write/read round-trip, rather than the merge clearing them.

/** @typedef {{ name: string, baseline: 'pass'|'fail', expect: (api: any, assert: any) => Promise<void>|void }} Behavior */

export default {
  id: 'merged-region-master-cell-border-survives',
  provenance: {source: 'upstream-issue'},
  cluster: 'styles',
  description:
    'The master (top-left) cell of a merged region keeps its border and other style facets (numFmt, ' +
    'font) after the merge and a round-trip, so the merged region renders its intended outline.',

  /** @type {Behavior[]} */
  behavior: [
    {
      name: 'the master cell keeps its top and bottom borders after the merge',
      baseline: 'pass',
      async expect(api, assert) {
        const {hasTopBorder, hasBottomBorder, topStyle, bottomStyle} =
          await api.mergeMasterBorderReport();
        assert.strictEqual(hasTopBorder, true, 'the top border survives the merge');
        assert.strictEqual(hasBottomBorder, true, 'the bottom border survives the merge');
        assert.strictEqual(topStyle, 'thin', 'the top border style is preserved');
        assert.strictEqual(bottomStyle, 'medium', 'the bottom border style is preserved');
      },
    },
    {
      name: 'the master cell keeps its non-border style (numFmt, font) through the merge',
      baseline: 'pass',
      async expect(api, assert) {
        const {numFmt, fontBold} = await api.mergeMasterBorderReport();
        assert.strictEqual(numFmt, '0.00', 'the number format survives');
        assert.strictEqual(fontBold, true, 'the font survives');
      },
    },
    {
      name: 'the merge is declared over the range',
      baseline: 'pass',
      async expect(api, assert) {
        const {merges} = await api.mergeMasterBorderReport();
        assert.ok(merges.includes('A1:B2'), `the merge is present; got ${JSON.stringify(merges)}`);
      },
    },
  ],
};
