// Cluster: images
//
// Real-world scenario: images are anchored to single cells (C2, C3) while rows are appended in the
// same loop that adds the images. Each image's cell-range anchor must resolve to the coordinates of
// that exact cell, independent of the interleaving order of addRow and addImage — no off-by-one row
// drift, no phantom extra row, and a one-to-one mapping from each image to its intended cell. Locks
// the cell-anchored image position against the reported misalignment.

/** @typedef {{ name: string, baseline: 'pass'|'fail', expect: (api: any, assert: any) => Promise<void>|void }} Behavior */

export default {
  id: 'cell-anchored-image-position-stable-under-row-adds',
  provenance: {source: 'upstream-issue'},
  cluster: 'images',
  description:
    'Images anchored to single cells (C2, C3) with addRow calls interleaved between them resolve ' +
    'each to its exact cell — from-anchor col 2 at rows 1 and 2 (zero-based) — mapping one-to-one ' +
    'with no off-by-one drift or phantom row.',

  /** @type {Behavior[]} */
  behavior: [
    {
      name: 'each single-cell-anchored image produces exactly one anchor',
      baseline: 'pass',
      async expect(api, assert) {
        const {anchorCount} = await api.cellAnchoredImagePositionReport();
        assert.strictEqual(anchorCount, 2, 'two images anchored to two cells produce two anchors');
      },
    },
    {
      name: 'each image resolves to its intended cell with no row drift',
      baseline: 'pass',
      async expect(api, assert) {
        const {froms} = await api.cellAnchoredImagePositionReport();
        assert.deepStrictEqual(froms[0], {col: 2, row: 1}, 'the C2 image anchors at col 2, row 1 (zero-based)');
        assert.deepStrictEqual(froms[1], {col: 2, row: 2}, 'the C3 image anchors at col 2, row 2 — no off-by-one');
      },
    },
  ],
};
