// Cluster: images
//
// Real-world scenario: a user adds an image and supplies the file extension as the raw result of a
// path-extension helper — ".png" (with a leading dot) rather than "png". The media part is then
// written with a doubled separator in its filename ("xl/media/image1..png"), and on read the
// media-matching logic (which assumes a single well-formed extension) fails to recognize the
// doubled-dot part, so getImages() comes back empty and the picture is lost. A leading-dot extension
// must be normalized so the media filename is well-formed and the image survives the round-trip
// identically to a dot-less extension.

import type {Assert, Case, CorpusApi} from '../case.ts';

export default {
  id: 'image-dotted-extension-roundtrip',
  provenance: {source: 'upstream-issue'},
  cluster: 'images',
  description:
    'An image added with a leading-dot extension (".png") produces a well-formed media filename (no ' +
    'doubled separator) and survives a write/read round-trip so the worksheet still reports one ' +
    'image — behaving identically to a dot-less extension.',

  behavior: [
    {
      name: 'a leading-dot extension does not produce a doubled-separator media filename',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const {mediaParts, doubledSeparator} = await api.imageExtensionRoundtrip('.png');
        assert.strictEqual(
          doubledSeparator,
          false,
          `the media filename must be well-formed; got ${JSON.stringify(mediaParts)}`,
        );
      },
    },
    {
      name: 'an image added with a leading-dot extension is still discoverable after a round-trip',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const {reloadedImageCount} = await api.imageExtensionRoundtrip('.png');
        assert.strictEqual(reloadedImageCount, 1, 'the worksheet reports its one image, not zero');
      },
    },
    {
      name: 'a dot-less extension round-trips the image (control)',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const {reloadedImageCount, doubledSeparator} = await api.imageExtensionRoundtrip('png');
        assert.strictEqual(
          doubledSeparator,
          false,
          'a dot-less extension yields a well-formed media filename',
        );
        assert.strictEqual(reloadedImageCount, 1, 'the image is discoverable');
      },
    },
  ],
} satisfies Case;
