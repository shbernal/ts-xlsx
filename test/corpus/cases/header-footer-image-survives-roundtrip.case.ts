// Cluster: images
//
// Real-world scenario: a worksheet places an image in its page header/footer — encoded by a
// picture token (&G) in the header/footer text, a <legacyDrawingHF> relationship pointing at a
// VML drawing, that VML drawing (a v:shape referencing the image), and the image media part. The
// same sheet also carries a cell comment (its own VML drawing + comments part). A library that
// does not model header/footer images drops them on a no-op load→save: the header image
// vanishes, while the comment must still survive — and the two VML drawings must keep distinct,
// non-colliding relationship ids. A round-trip must preserve the header/footer image alongside
// the comment.

import type {Assert, Case, CorpusApi} from '../case.ts';

const FIXTURE = 'header-footer-image-survives-roundtrip/sample.xlsx';

export default {
  id: 'header-footer-image-survives-roundtrip',
  provenance: {source: 'upstream-issue', repo: 'exceljs/exceljs', ref: 2563},
  cluster: 'images',
  description:
    'A no-op load→save preserves a page header/footer image — its &G token, the legacyDrawingHF ' +
    'relationship, the header/footer VML drawing, and the image media — while a coexisting cell ' +
    'comment also survives, instead of the header image being silently dropped.',

  behavior: [
    {
      name: 'the header/footer image token and its legacyDrawingHF wiring survive the round-trip',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const {source, rewritten} = await api.roundtripFixturePackageParts(FIXTURE);
        assert.ok(
          source.hasLegacyDrawingHF && source.hasHeaderFooterImageToken,
          'precondition: source has an HF image',
        );
        assert.ok(
          rewritten.hasHeaderFooterImageToken,
          'the &G header/footer picture token survives',
        );
        assert.ok(rewritten.hasLegacyDrawingHF, 'the legacyDrawingHF relationship survives');
      },
    },
    {
      name: 'both VML drawings (header image + comment) survive with the image media',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const {source, rewritten} = await api.roundtripFixturePackageParts(FIXTURE);
        assert.strictEqual(
          rewritten.vml,
          source.vml,
          `both VML drawings must survive (source had ${source.vml})`,
        );
        assert.ok(rewritten.media >= source.media, 'the header/footer image media survives');
        assert.strictEqual(rewritten.comments, source.comments, 'the coexisting comment survives');
      },
    },
  ],
} satisfies Case;
