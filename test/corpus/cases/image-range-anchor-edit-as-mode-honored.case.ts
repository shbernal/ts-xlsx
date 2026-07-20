// Cluster: images
//
// Real-world scenario: a user anchors an image over a range of cells (a top-left and bottom-right
// cell anchor) and chooses how it reacts to cell resizing/moving via an anchor-edit option: a
// two-cell anchor that moves and resizes with its cells, a one-cell anchor that moves but does not
// resize, or an absolute anchor that neither moves nor resizes. Whichever mode the caller explicitly
// requests must be honored in the drawing XML's editAs attribute, so the image behaves as intended
// when the sheet's rows/columns are resized.

import type {Assert, Case, CorpusApi} from '../case.ts';

const imageSpec = (editAs: CorpusApi) => ({
  sheets: [{name: 'S', images: [{range: {tl: {col: 1, row: 1}, br: {col: 3, row: 3}, editAs}}]}],
});

const firstEditAs = (anchors: CorpusApi) => (anchors[0] ? anchors[0].editAs : undefined);

export default {
  id: 'image-range-anchor-edit-as-mode-honored',
  provenance: {source: 'upstream-issue'},
  cluster: 'images',
  description:
    'A range-anchored image honors the caller’s explicit anchor-edit mode — twoCell (move+resize), ' +
    'oneCell (move only), or absolute (neither) — emitting the requested editAs in the drawing XML.',

  behavior: [
    {
      name: 'an explicit twoCell request emits editAs="twoCell"',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const {anchors} = await api.inspectImageAnchors(imageSpec('twoCell'));
        assert.strictEqual(
          firstEditAs(anchors),
          'twoCell',
          'the two-cell move+resize mode is honored',
        );
      },
    },
    {
      name: 'an explicit oneCell request emits editAs="oneCell"',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const {anchors} = await api.inspectImageAnchors(imageSpec('oneCell'));
        assert.strictEqual(
          firstEditAs(anchors),
          'oneCell',
          'the one-cell move-only mode is honored',
        );
      },
    },
    {
      name: 'an explicit absolute request emits editAs="absolute"',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const {anchors} = await api.inspectImageAnchors(imageSpec('absolute'));
        assert.strictEqual(
          firstEditAs(anchors),
          'absolute',
          'the absolute (fixed) mode is honored',
        );
      },
    },
    {
      // With no explicit mode (omitted, or passed as undefined) the anchor must resolve to a stable
      // default — oneCell (move-but-don't-resize), in the singular OpenXML spelling — never absolute,
      // an error, or a literal "undefined". This pins the semantics an inconsistent doc left ambiguous.
      name: 'an image with no explicit editAs defaults to editAs="oneCell"',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const {anchors} = await api.inspectImageAnchors(imageSpec(undefined));
        assert.strictEqual(
          firstEditAs(anchors),
          'oneCell',
          'the omitted-mode default is oneCell (singular spelling)',
        );
      },
    },
  ],
} satisfies Case;
