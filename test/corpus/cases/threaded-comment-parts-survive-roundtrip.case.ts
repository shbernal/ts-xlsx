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
// NOTE: this is the interim safety net only, and preserving the parts is necessary but NOT sufficient —
// see the known-open behavior at the end. Excel also writes a legacy fallback `<comment>` (the
// "[Threaded comment] Your version of Excel..." boilerplate) into `comments{n}.xml`, which today still
// surfaces as a garbage `cell.note`; re-serialising it as a plain note destroys the `tc=`/`xr:uid`
// binding, so Excel itself no longer recognises the preserved threads. Both halves — suppressing the
// fallback on read and owning it on write — belong to the read-model phase, not here.

import type {Assert, Case, CorpusApi} from '../case.ts';

const FIXTURE = 'threaded-comment-parts-survive-roundtrip/sample.xlsx';

// A second real-world shape: a *resolved* thread whose reply is by a different author, plus a genuine
// legacy note on another cell of the same sheet. Excel refuses to put a note and a thread on the SAME
// cell (AddComment and AddCommentThreaded each reject the other's cell), so co-existence is per sheet,
// not per cell — which is exactly what makes this file the interesting one: its `comments1.xml` mixes
// two synthetic `tc={guid}` thread fallbacks with one real note author.
const RESOLVED_MULTI_AUTHOR = 'threaded-comment-parts-survive-roundtrip/resolved-multi-author.xlsx';

// A third shape: a message that @mentions someone. A mention is not decoration — it carries the
// mentioned person's id plus the character span of the message text that renders as the mention chip, so
// dropping it loses who was asked and shifting it highlights the wrong words. Excel's own save is what
// produced this file: an injected mention was re-resolved by Excel, which re-pointed it at a NEW person
// entry it added with `providerId="PeoplePicker"` (the same human as an existing `providerId="AD"`
// author, but a separate registry entry), then rendered `@Grace Hopper` as a chip over exactly the
// 13 characters `startIndex="0" length="13"` names — the leading `@` included.
const MENTION_IN_THREAD = 'threaded-comment-parts-survive-roundtrip/mention-in-thread.xlsx';

export default {
  id: 'threaded-comment-parts-survive-roundtrip',
  provenance: {source: 'excel-desktop-verification'},
  cluster: 'comment',
  description:
    'A no-op load→save preserves modern threaded-comment parts (per-sheet threadedComment parts and ' +
    'the workbook-level persons author registry) rather than dropping them — a threaded-comment-bearing ' +
    'workbook survives a fill-and-save. Preservation is asserted on the conversation itself too: message ' +
    'count, reply structure, resolved (`done`) state, and each author of a multi-author thread. Interim ' +
    'preservation, ahead of the full thread model.',

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
    {
      name: 'the conversation inside the parts survives, not just the parts',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const {source, rewritten} = await api.roundtripFixturePackageParts(RESOLVED_MULTI_AUTHOR);
        assert.strictEqual(source.threadedCommentMessages, 3, 'precondition: three messages');
        assert.strictEqual(
          source.threadedCommentReplies,
          1,
          'precondition: one of them is a reply',
        );
        assert.strictEqual(
          rewritten.threadedCommentMessages,
          source.threadedCommentMessages,
          'every message survives',
        );
        assert.strictEqual(
          rewritten.threadedCommentReplies,
          source.threadedCommentReplies,
          'the reply keeps its parentId, so thread structure is not flattened',
        );
      },
    },
    {
      name: 'a resolved thread stays resolved across the round-trip',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const {source, rewritten} = await api.roundtripFixturePackageParts(RESOLVED_MULTI_AUTHOR);
        assert.strictEqual(source.resolvedThreadHeads, 1, 'precondition: one head is marked done');
        assert.strictEqual(
          rewritten.resolvedThreadHeads,
          source.resolvedThreadHeads,
          'the done flag is not dropped, so a resolved thread does not reopen itself',
        );
      },
    },
    {
      name: 'each author of a multi-author thread survives with its messages still mapped to it',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const {source, rewritten} = await api.roundtripFixturePackageParts(RESOLVED_MULTI_AUTHOR);
        assert.strictEqual(source.personEntries, 2, 'precondition: two registered persons');
        assert.strictEqual(
          source.threadedCommentAuthorIds.length,
          2,
          'precondition: messages are split across both persons',
        );
        assert.strictEqual(rewritten.personEntries, source.personEntries, 'both persons survive');
        assert.deepStrictEqual(
          rewritten.threadedCommentAuthorIds,
          source.threadedCommentAuthorIds,
          'every message still resolves to the same person id, so no author is re-pointed or orphaned',
        );
      },
    },
    {
      name: 'an @mention inside a message survives the round-trip with its mentioned person',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const {source, rewritten} = await api.roundtripFixturePackageParts(MENTION_IN_THREAD);
        assert.strictEqual(source.threadedCommentMentions, 1, 'precondition: one mention');
        assert.strictEqual(
          source.threadedCommentMentionPersonIds.length,
          1,
          'precondition: the mention names a person',
        );
        assert.strictEqual(
          rewritten.threadedCommentMentions,
          source.threadedCommentMentions,
          'the mention is not dropped, so who was asked is not lost',
        );
        assert.deepStrictEqual(
          rewritten.threadedCommentMentionPersonIds,
          source.threadedCommentMentionPersonIds,
          'the mention still points at the same person, so it is not re-pointed or orphaned',
        );
      },
    },
    {
      name: 'the mention keeps the text span it highlights, so it stays over the mentioned name',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const {source, rewritten} = await api.roundtripFixturePackageParts(MENTION_IN_THREAD);
        // `@Grace Hopper` is 13 characters at offset 0 — Excel renders the chip over exactly this span,
        // so the pair is asserted literally rather than only compared before/after.
        assert.deepStrictEqual(source.threadedCommentMentionSpans, ['0:13'], 'precondition: span');
        assert.deepStrictEqual(
          rewritten.threadedCommentMentionSpans,
          source.threadedCommentMentionSpans,
          'startIndex and length are unchanged, so the chip does not slide onto the wrong words',
        );
      },
    },
    {
      name: 'the separate PeoplePicker person entry a mention resolves through survives',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const {source, rewritten} = await api.roundtripFixturePackageParts(MENTION_IN_THREAD);
        assert.strictEqual(
          source.personEntries,
          3,
          'precondition: two authors plus the mentioned identity Excel interned separately',
        );
        assert.deepStrictEqual(
          source.personProviderIds,
          ['AD', 'PeoplePicker'],
          'precondition: the mentioned identity is registered by a different provider than the authors',
        );
        assert.strictEqual(
          rewritten.personEntries,
          source.personEntries,
          'all three entries survive — the mentioned identity is not merged into its author twin',
        );
        assert.deepStrictEqual(
          rewritten.personProviderIds,
          source.personProviderIds,
          'each entry keeps its providerId, so the mention still resolves through the right one',
        );
      },
    },
    {
      // KNOWN-OPEN, and the reason preserving the parts is necessary but not yet sufficient: verified
      // against desktop Excel (2026-07-26) on this fixture's round-tripped output. Excel read back ZERO
      // threaded comments and three ordinary notes, even though threadedComment1.xml, person.xml, both
      // relationships and both content-type overrides all survived intact and the package validates
      // clean against OpenXmlValidator. The break is in `comments{n}.xml`, which is re-serialised from
      // the note model: the `<authors>` list collapses to one empty `<author/>` (the synthetic
      // `tc={headId}` entries are gone) and every `<comment>` loses its `xr:uid`. Those are exactly the
      // two things Excel matches on, so the thread part is orphaned and the conversation is invisible in
      // the app. Flipping this behavior green is the read model's job (it must own the fallback rather
      // than round-tripping it as a plain note); until then the bytes survive for us but not for Excel.
      name: 'the fallback comments keep the tc= authors and xr:uids that bind them to their threads',
      baseline: 'fail',
      async expect(api: CorpusApi, assert: Assert) {
        const {source, rewritten} = await api.roundtripFixturePackageParts(RESOLVED_MULTI_AUTHOR);
        assert.strictEqual(
          source.commentFallbackThreadAuthors,
          2,
          'precondition: one tc= fallback author per thread head',
        );
        assert.strictEqual(
          rewritten.commentFallbackThreadAuthors,
          source.commentFallbackThreadAuthors,
          'each thread head keeps its tc= fallback author, so Excel still sees the cell as threaded',
        );
        assert.deepStrictEqual(
          rewritten.commentUids,
          source.commentUids,
          'each fallback comment keeps the xr:uid that points at its thread head',
        );
      },
    },
  ],
} satisfies Case;
