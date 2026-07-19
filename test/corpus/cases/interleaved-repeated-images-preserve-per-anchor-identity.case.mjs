// Cluster: images
//
// Real-world scenario: a workbook registers two distinct images (A and B) and a worksheet places
// three anchors in the order B, A, A. After a write/read round-trip each anchor must reference the
// image it was placed with — the rendered sequence must stay B, A, A. The defect: the drawing
// serializer resolves each anchor's image relationship with a fragile "same as the previous anchor"
// heuristic instead of a stable image-id → relationship-id mapping. When one image is reused across
// non-adjacent anchors while another interleaves, the heuristic collides and the third anchor
// (which should be A) points at B's relationship instead.

/** @typedef {{ name: string, baseline: 'pass'|'fail', expect: (api: any, assert: any) => Promise<void>|void }} Behavior */

export default {
  id: 'interleaved-repeated-images-preserve-per-anchor-identity',
  provenance: {source: 'upstream-issue'},
  cluster: 'images',
  description:
    'Placing two distinct images interleaved across anchors (B, A, A) resolves each anchor to the ' +
    'image it was placed with after serialization; a reused image maps to a single stable ' +
    'relationship by identity rather than by adjacency to the previous anchor.',

  /** @type {Behavior[]} */
  behavior: [
    {
      name: 'each distinct image used by the drawing gets exactly one relationship',
      baseline: 'pass',
      async expect(api, assert) {
        const {distinctMediaCount, distinctRelTargets} = await api.interleavedImageAnchors('BAA');
        assert.strictEqual(distinctMediaCount, 2, 'two distinct images produce two media parts');
        assert.strictEqual(
          distinctRelTargets,
          2,
          'each distinct image maps to exactly one relationship target',
        );
      },
    },
    {
      name: 'the first two anchors (B, then A) resolve to the images they were placed with',
      baseline: 'pass',
      async expect(api, assert) {
        const {resolvedLetter} = await api.interleavedImageAnchors('BAA');
        assert.strictEqual(resolvedLetter[0], 'B', 'anchor 1 renders B');
        assert.strictEqual(resolvedLetter[1], 'A', 'anchor 2 renders A');
      },
    },
    {
      name: 'the third anchor, reusing image A, does not collide onto the other image',
      baseline: 'pass',
      async expect(api, assert) {
        const {resolvedLetter} = await api.interleavedImageAnchors('BAA');
        assert.strictEqual(
          resolvedLetter[2],
          'A',
          `anchor 3 was placed with A and must render A, not the interleaved image; got ${JSON.stringify(resolvedLetter)}`,
        );
      },
    },
  ],
};
