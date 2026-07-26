// Cluster: comment
//
// Real-world scenario: a workbook carries Excel's modern threaded comments (the 2018 review-style
// conversations — author, timestamp, replies). Those live in per-sheet
// `xl/threadedComments/threadedComment{n}.xml` parts wired by a
// `.../2017/10/relationships/threadedComment` sheet relationship, with the authors in a workbook-level
// `xl/persons/person.xml` registry wired by a `.../relationships/person` relationship. Neither part is
// modeled yet, so before this preservation a no-op load→save dropped both — the conversation and its
// authors vanished from a fill-and-save workflow. Until the full thread model exists, these unmodeled
// parts must survive a round-trip.
//
// NOTE: this is the interim safety net only. It does NOT yet fix the companion data-quality bug — Excel
// also writes a legacy fallback `<comment>` (the "[Threaded comment] Your version of Excel..."
// boilerplate) into `comments{n}.xml`, which today still surfaces as a garbage `cell.note`. Suppressing
// that fallback belongs to the read-model phase, not here.

import type {Assert, Case, CorpusApi} from '../case.ts';

const FIXTURE = 'threaded-comment-parts-survive-roundtrip/sample.xlsx';

export default {
  id: 'threaded-comment-parts-survive-roundtrip',
  provenance: {source: 'excel-desktop-verification'},
  cluster: 'comment',
  description:
    'A no-op load→save preserves modern threaded-comment parts (per-sheet threadedComment parts and ' +
    'the workbook-level persons author registry) rather than dropping them — a threaded-comment-bearing ' +
    'workbook survives a fill-and-save. Interim preservation, ahead of the full thread model.',

  behavior: [
    {
      name: 'per-sheet threadedComment parts survive the round-trip',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const {source, rewritten} = await api.roundtripFixturePackageParts(FIXTURE);
        assert.ok(source.threadedComments >= 1, 'precondition: source has threaded-comment parts');
        assert.strictEqual(
          rewritten.threadedComments,
          source.threadedComments,
          `all ${source.threadedComments} threadedComment parts survive`,
        );
      },
    },
    {
      name: 'workbook-level persons author registry survives the round-trip',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const {source, rewritten} = await api.roundtripFixturePackageParts(FIXTURE);
        assert.ok(source.persons >= 1, 'precondition: source has a persons part');
        assert.strictEqual(rewritten.persons, source.persons, 'persons part survives');
      },
    },
  ],
} satisfies Case;
