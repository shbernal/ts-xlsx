// Cluster: images
//
// Real-world scenario: a user adds a picture positioned by a top-left cell plus an explicit pixel
// extent (width/height) and asks for it to be pinned in place with an "absolute" movement mode so it
// neither moves nor resizes when surrounding cells change. Supplying a top-left + extent produces a
// single-anchor drawing (xdr:oneCellAnchor), but in the OOXML drawing schema the editAs movement
// attribute is defined only on the two-cell anchor (xdr:twoCellAnchor); oneCellAnchor and
// absoluteAnchor carry no editAs. Emitting editAs on a oneCellAnchor is schema-invalid — the mode is
// not honored and a strict consumer rejects or repairs it. The requested absolute pin must be
// expressed through a construct that can legally carry it, not stamped onto an anchor kind that
// cannot. When a bottom-right cell IS supplied the drawing is a twoCellAnchor and editAs is valid.

/** @typedef {{ name: string, baseline: 'pass'|'fail', expect: (api: any, assert: any) => Promise<void>|void }} Behavior */

const TWO_CELL = {sheets: [{name: 'S', images: [{range: {tl: {col: 1, row: 1}, br: {col: 4, row: 6}, editAs: 'absolute'}}]}]};
const ONE_CELL = {sheets: [{name: 'S', images: [{range: {tl: {col: 1, row: 1}, ext: {width: 100, height: 80}, editAs: 'absolute'}}]}]};

export default {
  id: 'image-editas-only-valid-on-two-cell-anchor',
  provenance: {source: 'upstream-issue'},
  cluster: 'images',
  description:
    'The editAs movement mode is emitted only where the OOXML drawing schema allows it — on a ' +
    'two-cell anchor. A top-left + extent placement produces a one-cell anchor, which cannot carry ' +
    'editAs; the requested absolute pin must not be stamped onto it as a schema-invalid attribute.',

  /** @type {Behavior[]} */
  behavior: [
    {
      name: 'a top-left + bottom-right placement with editAs:absolute yields a two-cell anchor carrying editAs',
      baseline: 'pass',
      async expect(api, assert) {
        const {anchors, xmlWellFormed} = await api.inspectImageAnchors(TWO_CELL);
        assert.ok(xmlWellFormed, 'drawing XML must be well-formed');
        assert.strictEqual(anchors.length, 1, 'one image → one anchor');
        assert.strictEqual(anchors[0].anchorType, 'twoCell', 'top-left + bottom-right is a two-cell anchor');
        assert.strictEqual(anchors[0].editAs, 'absolute', 'editAs is valid and preserved on a two-cell anchor');
      },
    },
    {
      name: 'a top-left + extent placement produces a one-cell anchor',
      baseline: 'pass',
      async expect(api, assert) {
        const {anchors} = await api.inspectImageAnchors(ONE_CELL);
        assert.strictEqual(anchors.length, 1, 'one image → one anchor');
        assert.strictEqual(anchors[0].anchorType, 'oneCell', 'top-left + extent (no bottom-right) is a one-cell anchor');
      },
    },
    {
      name: 'editAs is not stamped onto the one-cell anchor, which the drawing schema cannot carry it on',
      baseline: 'fail',
      async expect(api, assert) {
        const {anchors} = await api.inspectImageAnchors(ONE_CELL);
        assert.strictEqual(
          anchors[0].editAs,
          null,
          `editAs is a two-cell-anchor attribute; a one-cell anchor must not carry it (schema-invalid). Got editAs=${anchors[0].editAs}`
        );
      },
    },
  ],
};
