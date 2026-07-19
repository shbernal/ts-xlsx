// Cluster: images
//
// Real-world scenario: a worksheet drawing contains a vector shape (an xdr:sp rectangle) or a
// text box — not a picture. A library that models only cell-anchored images drops the whole
// drawing on a no-op load→save: the worksheet's drawing relationship, the drawing part, and the
// shape all disappear, so the reopened file has lost its shapes. Unmodeled drawing content must
// be preserved through a round-trip rather than silently discarded.

/** @typedef {{ name: string, baseline: 'pass'|'fail', expect: (api: any, assert: any) => Promise<void>|void }} Behavior */

const FIXTURE = 'vector-shape-drawing-survives-roundtrip/sample.xlsx';

export default {
  id: 'vector-shape-drawing-survives-roundtrip',
  provenance: {source: 'upstream-issue', repo: 'exceljs/exceljs', ref: 1147},
  cluster: 'images',
  description:
    'A no-op load→save preserves a worksheet drawing that contains a vector shape (xdr:sp) — the ' +
    'drawing relationship, the drawing part, and the shape survive instead of the whole drawing ' +
    'being dropped.',

  /** @type {Behavior[]} */
  behavior: [
    {
      name: 'the drawing part and its worksheet reference survive the round-trip',
      baseline: 'pass',
      async expect(api, assert) {
        const {source, rewritten} = await api.roundtripFixturePackageParts(FIXTURE);
        assert.ok(
          source.drawings >= 1 && source.hasDrawingRef,
          'precondition: source has a drawing',
        );
        assert.ok(rewritten.drawings >= 1, 'the drawing part survives');
        assert.ok(rewritten.hasDrawingRef, 'the worksheet still references the drawing');
      },
    },
    {
      name: 'the vector shape element survives inside the drawing',
      baseline: 'pass',
      async expect(api, assert) {
        const {source, rewritten} = await api.roundtripFixturePackageParts(FIXTURE);
        assert.ok(source.drawingHasShape, 'precondition: source drawing has an xdr:sp shape');
        assert.ok(rewritten.drawingHasShape, 'the vector shape survives the round-trip');
      },
    },
  ],
};
