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
// Preserving the parts turned out to be necessary but not sufficient, which is what the last three
// behaviors are about. Excel also writes a legacy fallback `<comment>` (the "[Threaded comment] Your
// version of Excel..." boilerplate) into `comments{n}.xml`, and it binds a cell to its thread through
// that comment's synthetic `tc={headId}` author and `xr:uid` — so re-serialising the fallback as an
// ordinary note left every preserved thread orphaned and invisible in the app. The fallback is therefore
// owned: suppressed on read, and rebuilt from the thread model on write.

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
      // The behavior that makes preservation mean something *in Excel*, and the one this case was long
      // open on. Verified against desktop Excel: before this, our round-tripped output was read back as
      // ZERO threaded comments and three ordinary notes — even though threadedComment1.xml, person.xml,
      // both relationships and both content-type overrides all survived intact and the package validated
      // clean. The break was in `comments{n}.xml`, re-serialised from the note model: the `<authors>`
      // list collapsed to one empty `<author/>` (losing the synthetic `tc={headId}` entries) and every
      // `<comment>` lost its `xr:uid`. Excel matches a cell to its thread on exactly those two, so the
      // thread part was orphaned and the conversation invisible in the app. Re-verified after the fix on
      // this fixture's own round-tripped output: two threaded comments, B1 resolved with its reply, B2
      // open, and D4 still an ordinary note.
      name: 'the fallback comments keep the tc= authors and xr:uids that bind them to their threads',
      baseline: 'pass',
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
          rewritten.commentFallbackUids,
          source.commentFallbackUids,
          'each fallback comment keeps the xr:uid that points at its thread head',
        );
      },
    },
    {
      name: 'the text a pre-2018 reader sees for a conversation is regenerated word for word',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        // The fallback is not carried through — it is rebuilt from the thread model — so matching Excel
        // is a claim about our own wording and reply layout, not about copying bytes. Excel folds a whole
        // conversation into one comment: fixed boilerplate, `Comment:` and the opening message, then a
        // repeated `Reply:` per reply, each body indented four spaces. Compared after the line-end
        // normalisation every XML reader performs, since Excel writes CRLF where we write LF and no
        // consumer can tell the two apart.
        const {source, rewritten} = await api.roundtripFixturePackageParts(RESOLVED_MULTI_AUTHOR);
        assert.strictEqual(source.commentFallbackTexts.length, 2, 'precondition: two fallbacks');
        assert.match(
          source.commentFallbackTexts[0],
          /^\[Threaded comment\]\n\nYour version of Excel/,
          'precondition: the boilerplate Excel actually writes',
        );
        assert.deepStrictEqual(
          rewritten.commentFallbackTexts,
          source.commentFallbackTexts,
          'every fallback reads back identically, replies and indentation included',
        );
      },
    },
    {
      name: 'a conversation does not multiply into extra comments, and every comment keeps its shape',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        // Two threads plus one genuine note is three comments and three VML shapes. Re-emitting the
        // fallback as well as a note per threaded cell would inflate both counts, and a comment without
        // its shape reads as text but renders nothing at all.
        const {source, rewritten} = await api.roundtripFixturePackageParts(RESOLVED_MULTI_AUTHOR);
        assert.strictEqual(source.commentEntries, 3, 'precondition: two fallbacks and one note');
        assert.strictEqual(rewritten.commentEntries, source.commentEntries);
        assert.strictEqual(
          rewritten.commentVmlShapes,
          rewritten.commentEntries,
          'one VML shape per comment, so each still has a box to render into',
        );
      },
    },
  ],
} satisfies Case;
