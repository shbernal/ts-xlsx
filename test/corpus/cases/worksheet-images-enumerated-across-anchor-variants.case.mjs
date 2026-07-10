// Cluster: images
//
// Real-world scenario: a user loads an .xlsx with embedded images and enumerates the worksheet's
// images to recover each image's binary and its top-left cell anchor, so it can be re-associated with
// the surrounding row data. The enumeration must return one entry per embedded image — for both a
// two-cell (from/to) anchor and a one-cell (from + extent) anchor — with its anchor coordinates,
// rather than an empty result when the media parts are present. This locks that the image-enumeration
// read surface works across anchor variants.

/** @typedef {{ name: string, baseline: 'pass'|'fail', expect: (api: any, assert: any) => Promise<void>|void }} Behavior */

export default {
  id: 'worksheet-images-enumerated-across-anchor-variants',
  provenance: {source: 'upstream-issue'},
  cluster: 'images',
  description:
    'Enumerating a worksheet’s images after a round-trip returns one entry per embedded image across ' +
    'anchor variants (a two-cell and a one-cell anchor), each with its top-left cell coordinates, ' +
    'rather than an empty result.',

  /** @type {Behavior[]} */
  behavior: [
    {
      name: 'both anchor variants are enumerated (not an empty result)',
      baseline: 'pass',
      async expect(api, assert) {
        const {count} = await api.enumerateImagesAfterRoundtrip();
        assert.strictEqual(count, 2, 'both the two-cell and one-cell anchored images are enumerated');
      },
    },
    {
      name: 'each enumerated image reports its top-left cell anchor',
      baseline: 'pass',
      async expect(api, assert) {
        const {images} = await api.enumerateImagesAfterRoundtrip();
        const tls = images.map(i => i.tl && `${i.tl.col},${i.tl.row}`).sort();
        assert.deepStrictEqual(tls, ['1,1', '5,5'], `each image reports its from-cell; got ${JSON.stringify(tls)}`);
      },
    },
    {
      name: 'the media binaries backing the images are present',
      baseline: 'pass',
      async expect(api, assert) {
        const {mediaCount} = await api.enumerateImagesAfterRoundtrip();
        assert.ok(mediaCount >= 1, 'the package carries the image media');
      },
    },
  ],
};
