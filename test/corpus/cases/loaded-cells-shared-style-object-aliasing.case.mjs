// Cluster: styles
//
// Real-world scenario: a spreadsheet is written so several cells reference the same style index (the
// on-disk format deduplicates identical formatting into one shared style record). After the workbook
// is loaded, each cell exposes a `style` object — but because those cells resolved to the same index,
// the reader hands them all the *same* in-memory style object rather than independent copies. So
// mutating one loaded cell's style (setting a fill, font, alignment, or number format) silently
// bleeds into every other cell that happened to share that style index on disk, potentially across
// rows and sheets. The user expects each loaded cell's style to be independent.

/** @typedef {{ name: string, baseline: 'pass'|'fail', expect: (api: any, assert: any) => Promise<void>|void }} Behavior */

export default {
  id: 'loaded-cells-shared-style-object-aliasing',
  provenance: {source: 'upstream-issue'},
  cluster: 'styles',
  description:
    'Cells that shared a style index on disk must expose independent style objects once loaded, so ' +
    'mutating one cell\'s fill leaves sibling cells untouched — both in memory and after write-back.',

  /** @type {Behavior[]} */
  behavior: [
    {
      name: 'mutating one loaded cell\'s fill does not bleed into a sibling that shared its style index',
      baseline: 'fail',
      async expect(api, assert) {
        const {sibling, original, bled} = await api.loadMutateCellStyle();
        assert.ok(!bled, `the sibling cell must keep its own fill; it changed to ${sibling}`);
        assert.strictEqual(sibling, original, 'the sibling retains the original shared fill color');
      },
    },
    {
      name: 'writing back after a single-cell style edit changes only that cell on disk',
      baseline: 'fail',
      async expect(api, assert) {
        const {diskSibling, original, diskBled} = await api.loadMutateCellStyle();
        assert.ok(!diskBled, `only the edited cell should change on disk; sibling became ${diskSibling}`);
        assert.strictEqual(diskSibling, original, 'the sibling keeps its original fill in the written file');
      },
    },
  ],
};
