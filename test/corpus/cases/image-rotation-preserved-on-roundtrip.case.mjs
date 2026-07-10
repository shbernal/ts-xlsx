// Cluster: images
//
// Real-world scenario: a workbook contains an image that was rotated in the drawing — the OOXML
// drawing anchor carries a rotation transform (`rot` on the picture's `<a:xfrm>` shape properties,
// in 1/60000-degree units). Reading and writing the workbook back must retain that rotation. In the
// reported failure the rotation was dropped on read, so the re-written file rendered the image
// un-rotated. The rotation must survive a load followed by a save.
//
// The fixture is a workbook whose single image anchor was given a rotation transform (rot="2700000",
// i.e. 45°), reproducing an application-rotated image without needing an interactive editor.

/** @typedef {{ name: string, baseline: 'pass'|'fail', expect: (api: any, assert: any) => Promise<void>|void }} Behavior */

const FIXTURE = 'image-rotation-preserved-on-roundtrip/sample.xlsx';
const ROT = 2700000; // 45° in OOXML 1/60000-degree units

export default {
  id: 'image-rotation-preserved-on-roundtrip',
  provenance: {source: 'upstream-issue'},
  cluster: 'images',
  description:
    'An image whose drawing anchor declares a rotation transform keeps that rotation through a ' +
    'load/save round-trip, rather than the rotation being dropped so the re-written image renders ' +
    'un-rotated.',

  /** @type {Behavior[]} */
  behavior: [
    {
      name: 'the source fixture image carries a rotation transform (precondition)',
      baseline: 'pass',
      async expect(api, assert) {
        const {sourceRot} = await api.roundtripFixtureImageRotation(FIXTURE);
        assert.strictEqual(sourceRot, ROT, 'the fixture drawing declares the rotation');
      },
    },
    {
      name: 'the image rotation survives a load/save round-trip',
      baseline: 'fail',
      async expect(api, assert) {
        const {rewrittenRot} = await api.roundtripFixtureImageRotation(FIXTURE);
        assert.strictEqual(
          rewrittenRot,
          ROT,
          `the rotation must be re-emitted on save; got ${JSON.stringify(rewrittenRot)} (dropped)`
        );
      },
    },
  ],
};
