// Cluster: images
//
// Real-world scenario: a caller adds several images to a worksheet, then decides one should no longer
// appear in the output — a logo swapped out, a placeholder dropped. The library exposes a way to add
// and anchor an image but no way to remove a previously-added one, so the only recourse is to rebuild
// the worksheet from scratch. A worksheet must offer image removal: dropping a specific image leaves
// the written workbook with only the images the caller still wants, each remaining image preserved at
// its original anchor.

/** @typedef {{ name: string, baseline: 'pass'|'fail', expect: (api: any, assert: any) => Promise<void>|void }} Behavior */

export default {
  id: 'worksheet-image-removal',
  provenance: {source: 'upstream-issue'},
  cluster: 'images',
  description:
    'A worksheet exposes a way to remove a previously-added image: removing one image drops exactly ' +
    'that image and preserves the rest at their original anchors, rather than forcing a rebuild.',

  /** @type {Behavior[]} */
  behavior: [
    {
      name: 'a worksheet exposes an image-removal method',
      baseline: 'pass',
      async expect(api, assert) {
        const {supported} = await api.removeImageReport();
        assert.strictEqual(
          supported,
          true,
          'the worksheet must expose a way to remove an added image',
        );
      },
    },
    {
      name: 'removing one image drops it and leaves the others',
      baseline: 'pass',
      async expect(api, assert) {
        const {supported, before, after, removedGone, othersSurvive} =
          await api.removeImageReport();
        assert.ok(supported, 'removal must be supported for this behavior to hold');
        assert.strictEqual(after, before - 1, 'exactly one image is removed');
        assert.ok(removedGone, 'the targeted image is gone');
        assert.ok(othersSurvive, 'the remaining image is preserved');
      },
    },
  ],
};
