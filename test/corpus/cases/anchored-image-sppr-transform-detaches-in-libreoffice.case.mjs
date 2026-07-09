// Cluster: images
//
// Real-world scenario: a fixed-size image is placed with a single-cell anchor (a from col/row plus
// an explicit extent). The anchor itself is correct, but the emitted drawing XML also writes the
// picture's own absolute shape transform (spPr/xfrm) as a zeroed placeholder: offset 0,0 and
// extent 0,0. Excel ignores this transform for an anchored drawing and positions purely from the
// anchor, so the file looks fine there — but a strict consumer (notably LibreOffice Calc) honors
// the zeroed absolute transform and renders the picture detached from its anchor cell, at the wrong
// place. The picture must not carry a competing zeroed transform that would override its anchor.

/** @typedef {{ name: string, baseline: 'pass'|'fail', expect: (api: any, assert: any) => Promise<void>|void }} Behavior */

const oneCellImageSpec = {
  sheets: [
    {
      name: 'S',
      images: [{range: {tl: {col: 2, row: 11}, ext: {width: 100, height: 80}, editAs: 'oneCell'}}],
    },
  ],
};

export default {
  id: 'anchored-image-sppr-transform-detaches-in-libreoffice',
  provenance: {source: 'upstream-issue'},
  cluster: 'images',
  description:
    'A single-cell-anchored image is emitted with its anchor from-cell and extent intact and ' +
    'without a competing zeroed absolute shape transform (spPr off 0,0 + ext 0,0) that a strict ' +
    'consumer would honor, detaching the picture from its anchor cell.',

  /** @type {Behavior[]} */
  behavior: [
    {
      name: 'the single-cell anchor preserves its from-cell and extent (control)',
      baseline: 'pass',
      async expect(api, assert) {
        const {anchors} = await api.inspectImageAnchors(oneCellImageSpec);
        assert.strictEqual(anchors.length, 1, 'exactly one anchored image');
        assert.strictEqual(anchors[0].anchorType, 'oneCell', 'it is a single-cell anchor');
        assert.strictEqual(anchors[0].from.col, 2, 'the anchor from-column is preserved');
        assert.strictEqual(anchors[0].from.row, 11, 'the anchor from-row is preserved');
        assert.ok(anchors[0].ext && anchors[0].ext.cx > 0 && anchors[0].ext.cy > 0, 'the anchor extent is non-zero');
      },
    },
    {
      name: 'the emitted picture carries no zeroed absolute transform that would override the anchor',
      baseline: 'fail',
      async expect(api, assert) {
        const {anchors} = await api.inspectImageAnchors(oneCellImageSpec);
        assert.strictEqual(
          anchors[0].spPr.zeroedTransform,
          false,
          `an anchored picture must not emit a zeroed spPr transform (off 0,0 + ext 0,0); got ${JSON.stringify(anchors[0].spPr)}`
        );
      },
    },
  ],
};
