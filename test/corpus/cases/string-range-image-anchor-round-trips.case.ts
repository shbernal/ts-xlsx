// Cluster: images
//
// Real-world scenario: images are commonly anchored to a *cell range* expressed as a
// string — `sheet.addImage(id, "B2:D6")` — and real `.xlsx` files authored by other tools
// carry drawing anchors that decode to such ranges (e.g. a banner spanning "B198:BN198").
// Both adding a string-range image and reading a file full of range-anchored images must
// succeed and yield a normalized two-cell anchor, never throw. Historically, reading such
// a file threw `TypeError: Cannot create property 'editAs' on string 'B198:BN198'` because
// the range string was treated as an object.

import type {Assert, Case, CorpusApi} from '../case.ts';

const FIXTURE = 'string-range-image-anchor-round-trips/sample.xlsx';

export default {
  id: 'string-range-image-anchor-round-trips',
  provenance: {source: 'upstream-issue', repo: 'exceljs/exceljs', ref: 680},
  cluster: 'images',
  description:
    'An image anchored to a cell range given as a string is a first-class anchor form on ' +
    'both write and read: adding "B2:D6" produces a two-cell anchor, and reading a file ' +
    'whose images use range anchors exposes an object range with integer cell coordinates ' +
    'without throwing.',

  behavior: [
    {
      name: 'reading a file whose images use range anchors exposes object ranges without throwing',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const {images, count} = await api.readFixtureImageAnchors(FIXTURE);
        assert.strictEqual(count, 2, 'both images in the file are exposed');
        for (const im of images) {
          assert.ok(
            im.tl && Number.isInteger(im.tl.col),
            'top-left is an integer cell coordinate, not a string',
          );
          assert.ok(
            im.br && Number.isInteger(im.br.col),
            'bottom-right is an integer cell coordinate',
          );
          assert.strictEqual(
            im.editAs,
            'oneCell',
            'the anchor editAs is normalized, not left undefined on a string',
          );
        }
      },
    },
    {
      name: 'adding an image with a string range produces a two-cell anchor and does not throw',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const {anchors} = await api.inspectImageAnchors({
          sheets: [{name: 'S', images: [{range: 'B2:D6'}]}],
        });
        const a = anchors[0];
        assert.ok(a, 'the string-range image is serialized as an anchor');
        assert.strictEqual(a.anchorType, 'twoCell', 'a range becomes a two-cell anchor');
        assert.strictEqual(a.from.col, 1, 'from column is B (zero-based 1)');
        assert.strictEqual(a.from.row, 1, 'from row is 2 (zero-based 1)');
        assert.ok(a.to && a.to.col >= a.from.col, 'the anchor has a valid bottom-right corner');
      },
    },
  ],
} satisfies Case;
