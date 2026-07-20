// Cluster: images
//
// Real-world scenario: a user anchors an image with an explicit extent given in pixels
// (e.g. 191×47) expecting it to appear at exactly that size. The rendered size must depend
// only on the requested pixels and the fixed 96-DPI pixel→EMU conversion (1 px = 9525 EMU),
// never on the source image's own embedded DPI/resolution metadata. A file whose picture ext
// is rescaled by the source PNG's advertised DPI comes out at the wrong size. So 191×47 px
// must serialize to ext cx=1819275, cy=447675 EMU whatever DPI the image bytes declare.

import type {Assert, Case, CorpusApi} from '../case.ts';

// A single image anchored at the top-left with an explicit pixel extent — a oneCellAnchor.
const SPEC = {
  sheets: [{name: 'S', images: [{range: {tl: {col: 0, row: 0}, ext: {width: 191, height: 47}}}]}],
};

const PX_TO_EMU = 9525;

export default {
  id: 'image-pixel-extent-converts-to-emu-independent-of-dpi',
  provenance: {source: 'upstream-issue', repo: 'exceljs/exceljs', ref: 1036},
  cluster: 'images',
  description:
    'An image anchored with an explicit pixel extent serializes to a oneCellAnchor whose ext ' +
    'cx/cy are the requested pixels converted at the fixed 96-DPI rate (1 px = 9525 EMU) — the ' +
    'rendered size follows the requested pixels only, never the source image’s embedded DPI.',

  behavior: [
    {
      name: 'a pixel extent becomes a oneCellAnchor with EMU = pixels × 9525',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const {anchors} = await api.inspectImageAnchors(SPEC);
        assert.strictEqual(anchors.length, 1, 'exactly one image anchor');
        const {anchorType, ext} = anchors[0];
        assert.strictEqual(anchorType, 'oneCell', 'an explicit extent is a oneCellAnchor');
        assert.strictEqual(ext.cx, 191 * PX_TO_EMU, 'width 191 px → 1819275 EMU');
        assert.strictEqual(ext.cy, 47 * PX_TO_EMU, 'height 47 px → 447675 EMU');
      },
    },
  ],
} satisfies Case;
