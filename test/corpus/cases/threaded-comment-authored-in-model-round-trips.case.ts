// Cluster: comment
//
// Real-world scenario: a report generator wants to leave review comments in the workbook it produces —
// Excel's modern threaded comments (the 2018 conversations: an author, a timestamp, replies, a resolved
// state, `@mentions`), not the anonymous legacy notes that predate them. Nothing here comes from an input
// file: the conversation, its participants and its mention are built in the model and written from it.
//
// That makes this the case that distinguishes *serialising* the feature from *carrying* it. A conversation
// read from a file could survive a round-trip on preserved bytes alone while the library understood nothing
// about it; an authored one cannot. Everything the package says about the conversation — the per-sheet
// threadedComment part, the workbook-level person registry, and the legacy fallback `<comment>` that is
// how Excel binds a cell to its thread — has to be built from the model, and then read back into the same
// model.
//
// Verified against desktop Excel on exactly this package (2026-07-26): it opens clean with no repair
// prompt and no repair log, reports one threaded comment on B2 with its reply, renders the purple
// threaded-comment indicator on B2 and the red note triangle on D4, and validates clean against
// OpenXmlValidator. The facets are deliberately combined rather than tested apart, because they interact:
// the fallback text folds the reply in, the mention spans the message text by character offset, and the
// note on D4 must not be mistaken for the thread's fallback (nor the fallback for a note).

import type {Assert, Case, CorpusApi} from '../case.ts';
import type {Untyped} from '../untyped.ts';

// The two ids for one human that Excel itself writes: an authoring identity, and the separate entry it
// interns when that person is @mentioned. Registering both is what a faithful writer has to allow — merging
// them by name would silently re-point the mention.
const GRACE_MENTIONED = '{BA397017-DD76-4496-AA75-59ADB199950C}';
const HEAD = '{11111111-2222-3333-4444-555555555555}';

export default {
  id: 'threaded-comment-authored-in-model-round-trips',
  provenance: {source: 'excel-desktop-verification'},
  cluster: 'comment',
  description:
    'A threaded conversation authored in the model — a resolved thread with a reply by a second author ' +
    'and an @mention, beside a legacy note on another cell — is serialised into the threadedComment part, ' +
    'the workbook person registry and the legacy fallback comment, and reads back as the same model. ' +
    'Nothing is carried through from an input file, and the same model always writes the same bytes.',

  behavior: [
    {
      name: 'an authored conversation is written as its own part, wired to the workbook person registry',
      async expect(api: CorpusApi, assert: Assert) {
        const {parts} = await api.authoredCommentThreadRoundtrip();
        assert.strictEqual(parts.threadedComments, 1, 'one threadedComment part for the one sheet');
        assert.strictEqual(
          parts.persons,
          1,
          'one workbook-level registry, singular and unnumbered',
        );
        assert.strictEqual(parts.threadedCommentMessages, 2, 'the opening message and its reply');
        assert.strictEqual(
          parts.threadedCommentReplies,
          1,
          'the reply names its parent, so the thread is a thread and not two',
        );
        assert.strictEqual(
          parts.resolvedThreadHeads,
          1,
          'resolved is written on the head alone, where Excel reads it from',
        );
      },
    },
    {
      name: 'every participant is registered, including the separate identity a mention resolves through',
      async expect(api: CorpusApi, assert: Assert) {
        const {parts} = await api.authoredCommentThreadRoundtrip();
        assert.strictEqual(parts.personEntries, 3, 'two authors plus the mentioned identity');
        assert.deepStrictEqual(
          parts.personProviderIds,
          ['AD', 'PeoplePicker'],
          'the mentioned identity keeps its own provider rather than being merged into its author twin',
        );
        assert.strictEqual(parts.threadedCommentMentions, 1, 'the mention is written');
        assert.deepStrictEqual(
          parts.threadedCommentMentionPersonIds,
          [GRACE_MENTIONED],
          'and points at the mentioned identity, not at the same human’s authoring entry',
        );
        assert.deepStrictEqual(
          parts.threadedCommentMentionSpans,
          ['0:13'],
          '`@Grace Hopper` is 13 characters at offset 0 — the span Excel draws the chip over',
        );
      },
    },
    {
      name: 'the conversation is given the legacy fallback comment that binds its cell to it',
      async expect(api: CorpusApi, assert: Assert) {
        // Without this, Excel ignores the thread part however intact it is: the cell is bound to its
        // conversation through the fallback comment's synthetic `tc={headId}` author and its `xr:uid`.
        const {parts} = await api.authoredCommentThreadRoundtrip();
        assert.strictEqual(
          parts.commentFallbackThreadAuthors,
          1,
          'a tc= author for the one thread',
        );
        assert.deepStrictEqual(
          parts.commentFallbackUids,
          [HEAD],
          'the fallback names the thread head, which is what Excel matches the cell on',
        );
        assert.strictEqual(
          parts.commentEntries,
          2,
          'the fallback and the note on the other cell — a conversation is one comment, not one per message',
        );
        assert.strictEqual(
          parts.commentVmlShapes,
          parts.commentEntries,
          'each has a box to render into; a comment with no shape reads as text but draws nothing',
        );
      },
    },
    {
      name: 'the fallback text a pre-2018 reader sees folds the reply into the opening message',
      async expect(api: CorpusApi, assert: Assert) {
        const {parts} = await api.authoredCommentThreadRoundtrip();
        assert.deepStrictEqual(
          parts.commentFallbackTexts,
          [
            '[Threaded comment]\n\nYour version of Excel allows you to read this threaded comment; ' +
              'however, any edits to it will get removed if the file is opened in a newer version of ' +
              'Excel. Learn more: https://go.microsoft.com/fwlink/?linkid=870924\n\n' +
              'Comment:\n    @Grace Hopper is this gross or net?\n' +
              'Reply:\n    Gross. Confirmed with finance.',
          ],
          'the boilerplate verbatim, then Comment: and the head, then one Reply: per reply, bodies ' +
            'indented four spaces — exactly what Excel writes',
        );
      },
    },
    {
      name: 'the authored conversation reads back as the same model, participants resolved',
      async expect(api: CorpusApi, assert: Assert) {
        const {model} = await api.authoredCommentThreadRoundtrip();
        assert.strictEqual(model.sheets.length, 1);
        const sheet = model.sheets[0];
        assert.strictEqual(sheet.threads.length, 1, 'one conversation, not one thread per message');
        const thread = sheet.threads[0];
        assert.strictEqual(thread.ref, 'B2', 'the anchor is canonical — it was authored as `$B$2`');
        assert.strictEqual(thread.resolved, true);
        assert.deepStrictEqual(
          thread.comments.map((comment: Untyped) => [
            comment.id,
            comment.author?.displayName,
            comment.date,
            comment.text,
          ]),
          [
            [HEAD, 'Ada Lovelace', '2026-07-26T12:00:00.00', '@Grace Hopper is this gross or net?'],
            [
              '{66666666-7777-8888-9999-AAAAAAAAAAAA}',
              'Grace Hopper',
              '2026-07-26T12:05:00.00',
              'Gross. Confirmed with finance.',
            ],
          ],
          'each message keeps its own id, author, timestamp and text, in the order it was written',
        );
        assert.deepStrictEqual(
          thread.comments[0].mentions,
          [
            {
              person: {
                id: GRACE_MENTIONED,
                displayName: 'Grace Hopper',
                userId: null,
                providerId: 'PeoplePicker',
              },
              personId: GRACE_MENTIONED,
              startIndex: 0,
              length: 13,
              span: '@Grace Hopper',
            },
          ],
          'the mention resolves to its registered identity and still covers the mentioned name',
        );
      },
    },
    {
      name: 'the conversation is not surfaced as a note, and the real note beside it is untouched',
      async expect(api: CorpusApi, assert: Assert) {
        // The fallback comment is boilerplate wrapping a copy of the conversation. Reading it back as
        // `cell.note` would hand the caller garbage — and re-writing it as a plain note would destroy the
        // binding and leave Excel unable to see the thread at all.
        const {model} = await api.authoredCommentThreadRoundtrip();
        assert.deepStrictEqual(
          model.sheets[0].notes,
          {D4: 'an ordinary note beside the conversation'},
          'the threaded cell has no note; the genuine one survives verbatim',
        );
        assert.deepStrictEqual(
          model.sheets[0].at,
          {B2: 'B2', $B$2: 'B2', D4: null},
          'the conversation is found by its anchor either way, and the noted cell is not a threaded one',
        );
      },
    },
    {
      name: 'the same model always writes the same bytes, so an authored conversation is reproducible',
      async expect(api: CorpusApi, assert: Assert) {
        // Threaded comments need guids and timestamps, and the writer supplies neither: every id and date
        // is the caller's. A writer that reached for a clock or a random source here would make every save
        // of an unchanged workbook a different file.
        const {deterministic} = await api.authoredCommentThreadRoundtrip();
        assert.strictEqual(deterministic, true);
      },
    },
  ],
} satisfies Case;
