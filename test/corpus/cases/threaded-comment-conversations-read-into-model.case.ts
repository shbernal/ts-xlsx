// Cluster: comment
//
// Real-world scenario: a reviewer opens a workbook to find out what the conversation on a cell actually
// said — who asked, who answered, when, whether it was settled, and who was pulled in by name. Excel's
// modern threaded comments carry all of that, but on the wire it is scattered: a flat list of messages
// bound into threads only by `parentId`, each naming its author by a GUID that resolves through a
// separate workbook-level registry, and a resolved flag that only the thread's first message carries.
// Reading the parts is not enough — a reader that hands back that shape has moved the work to its caller.
// So this locks the reconstruction: threads with their replies in place, authors and @mentions resolved
// to real identities, and the resolved state read from where it is actually written.
//
// The reader also has to tell a conversation apart from a note. Excel writes a legacy fallback
// `<comment>` beside every thread — fixed boilerplate wrapping a copy of the conversation — so a reader
// that takes the comments part at face value reports a garbage note on every threaded cell. What that
// fallback becomes on the way back out is asserted by `threaded-comment-parts-survive-roundtrip`.

import type {Assert, Case, CorpusApi} from '../case.ts';

// Two single-author threads: A1 asked and answered, C3 a lone remark.
const SAMPLE = 'threaded-comment-parts-survive-roundtrip/sample.xlsx';

// A resolved two-message thread on B1 whose reply is by a SECOND author, an open thread on B2, and a
// genuine legacy note on D4 — the mix that separates "a thread" from "a note" on one sheet.
const RESOLVED_MULTI_AUTHOR = 'threaded-comment-parts-survive-roundtrip/resolved-multi-author.xlsx';

// The same file plus an @mention, including the separate `providerId="PeoplePicker"` registry entry
// Excel interns for a mentioned identity.
const MENTION_IN_THREAD = 'threaded-comment-parts-survive-roundtrip/mention-in-thread.xlsx';

export default {
  id: 'threaded-comment-conversations-read-into-model',
  provenance: {source: 'excel-desktop-verification'},
  cluster: 'comment',
  description:
    "Reading a workbook reconstructs Excel's modern threaded comments as conversations: each thread " +
    'anchored to its cell with its replies in order, every message resolved to the identity that wrote ' +
    'it, the resolved state taken from the thread head that carries it, and every @mention resolved to ' +
    'the person it names along with the run of text it highlights.',

  behavior: [
    {
      name: 'each conversation is read back as one thread anchored to its cell',
      baseline: 'pass',
      expect(api: CorpusApi, assert: Assert) {
        const {sheets} = api.readFixtureCommentThreads(SAMPLE);
        const [sheet] = sheets;
        assert.deepStrictEqual(
          sheet.threads.map((thread: CorpusApi) => thread.ref),
          ['A1', 'C3'],
          'both threads are found, each on the cell it hangs off',
        );
      },
    },
    {
      name: 'a reply belongs to the thread it answers instead of becoming a second conversation',
      baseline: 'pass',
      expect(api: CorpusApi, assert: Assert) {
        const {sheets} = api.readFixtureCommentThreads(SAMPLE);
        const [sheet] = sheets;
        // Three messages on the wire, two conversations: flattening the reply out would report three.
        assert.deepStrictEqual(
          sheet.threads.map((thread: CorpusApi) =>
            thread.comments.map((comment: CorpusApi) => comment.text),
          ),
          [
            ['Should this number include tax?', 'Yes, gross of tax.'],
            ['Confirmed against the ledger.'],
          ],
          'the reply follows the message it answers, in the order it was written',
        );
      },
    },
    {
      name: 'a resolved thread reads back resolved, and an open one does not',
      baseline: 'pass',
      expect(api: CorpusApi, assert: Assert) {
        const {sheets} = api.readFixtureCommentThreads(RESOLVED_MULTI_AUTHOR);
        const [sheet] = sheets;
        // Resolved is a property of the THREAD: only the head carries `done` on the wire (an open one
        // omits it entirely rather than writing `done="0"`), so reading it per message would report the
        // reply of a settled thread as unsettled.
        assert.deepStrictEqual(
          sheet.threads.map((thread: CorpusApi) => [thread.ref, thread.resolved]),
          [
            ['B1', true],
            ['B2', false],
          ],
        );
      },
    },
    {
      name: 'every message resolves to the identity that wrote it, not to a raw id',
      baseline: 'pass',
      expect(api: CorpusApi, assert: Assert) {
        const {sheets} = api.readFixtureCommentThreads(RESOLVED_MULTI_AUTHOR);
        const [thread] = sheets[0].threads;
        assert.deepStrictEqual(
          thread.comments.map((comment: CorpusApi) => comment.author.displayName),
          ['Ada Lovelace', 'Grace Hopper'],
          'the question and its answer are attributed to different people',
        );
        assert.notStrictEqual(
          thread.comments[0].authorId,
          thread.comments[1].authorId,
          'precondition: they really are distinct registry entries, not one author twice',
        );
      },
    },
    {
      name: 'the registry of identities a conversation resolves through is readable in its own right',
      baseline: 'pass',
      expect(api: CorpusApi, assert: Assert) {
        const {persons} = api.readFixtureCommentThreads(RESOLVED_MULTI_AUTHOR);
        assert.deepStrictEqual(
          persons.map((person: CorpusApi) => [person.displayName, person.providerId]),
          [
            ['Grace Hopper', 'AD'],
            ['Ada Lovelace', 'AD'],
          ],
          'both authors are registered, each with the provider that registered it',
        );
      },
    },
    {
      name: 'a message keeps the timestamp the file wrote, verbatim',
      baseline: 'pass',
      expect(api: CorpusApi, assert: Assert) {
        const {sheets} = api.readFixtureCommentThreads(RESOLVED_MULTI_AUTHOR);
        const [thread] = sheets[0].threads;
        // Excel writes local wall-clock with fractional seconds and NO timezone, which is not a
        // round-trippable instant. Reading it as one would invent a zone the file never stated and shift
        // the time; the string is what the file actually said.
        assert.strictEqual(thread.comments[0].date, '2026-07-26T10:54:00.01');
        assert.strictEqual(thread.comments[1].date, '2026-07-26T10:54:00.04');
      },
    },
    {
      name: 'an @mention resolves to the person it names and to the text it highlights',
      baseline: 'pass',
      expect(api: CorpusApi, assert: Assert) {
        const {sheets} = api.readFixtureCommentThreads(MENTION_IN_THREAD);
        const mentions = sheets[0].threads[1].comments[0].mentions;
        assert.strictEqual(mentions.length, 1, 'precondition: the message mentions one person');
        const [mention] = mentions;
        assert.strictEqual(mention.person.displayName, 'Grace Hopper', 'who was asked');
        // The offsets are only meaningful against the exact message text, so the span is asserted as
        // the text it actually covers: read them wrong and this is "@Grace Hopp" or "Grace Hopper ".
        assert.strictEqual(
          mention.span,
          '@Grace Hopper',
          'the chip covers the name, leading @ included',
        );
      },
    },
    {
      name: 'a mentioned person resolves through their own registry entry, not their author twin',
      baseline: 'pass',
      expect(api: CorpusApi, assert: Assert) {
        const {persons, sheets} = api.readFixtureCommentThreads(MENTION_IN_THREAD);
        // Excel interns a mentioned identity SEPARATELY from that person's authoring identity: same
        // display name, same userId, different id, `providerId="PeoplePicker"`. Both entries must
        // survive as themselves — resolving a mention by name or userId would pick the author entry, and
        // deduplicating the registry that way would delete the one the mention points at.
        const graces = persons.filter((person: CorpusApi) => person.displayName === 'Grace Hopper');
        assert.strictEqual(graces.length, 2, 'precondition: one human, two registry entries');
        assert.deepStrictEqual(
          graces.map((person: CorpusApi) => person.providerId),
          ['AD', 'PeoplePicker'],
          'precondition: they differ only by id and registering provider',
        );
        const [mention] = sheets[0].threads[1].comments[0].mentions;
        assert.strictEqual(
          mention.person.providerId,
          'PeoplePicker',
          'the mention resolves to the entry it names',
        );
        assert.strictEqual(
          mention.person.id,
          mention.personId,
          'and that entry is the one the file pointed at',
        );
      },
    },
    {
      name: 'a cell with a genuine legacy note is not reported as carrying a conversation',
      baseline: 'pass',
      expect(api: CorpusApi, assert: Assert) {
        // D4 holds a real note; B1 a thread. Excel refuses to put both on one cell, so a reader that
        // conflated the two features would answer this question wrong in both directions.
        const {sheets} = api.readFixtureCommentThreads(RESOLVED_MULTI_AUTHOR, ['B1', 'D4', 'A5']);
        assert.deepStrictEqual(sheets[0].at, {B1: 'B1', D4: null, A5: null});
      },
    },
    {
      name: 'a threaded cell reads back with no note, and the genuine note beside it is untouched',
      baseline: 'pass',
      expect(api: CorpusApi, assert: Assert) {
        // Excel writes a legacy fallback `<comment>` beside every thread so a pre-2018 reader sees
        // something: boilerplate ("[Threaded comment] Your version of Excel allows you to read...")
        // wrapping a copy of the conversation. It is not an annotation anybody wrote, so it must not
        // surface as one — a caller iterating notes would otherwise get that paragraph on B1 and B2
        // while the real conversation sat elsewhere. D4's note is a real one and stays.
        const {sheets} = api.readFixtureCommentThreads(RESOLVED_MULTI_AUTHOR);
        assert.deepStrictEqual(sheets[0].notes, {D4: 'A genuine legacy note.'});
      },
    },
    {
      name: 'a sheet whose only comments are conversations reads back with no notes at all',
      baseline: 'pass',
      expect(api: CorpusApi, assert: Assert) {
        // Every comment in this fixture is a thread fallback, so nothing at all survives as a note.
        // The pair matters: suppressing per *cell* satisfies both, while suppressing the whole part
        // whenever a sheet has threads satisfies this one and loses D4 above.
        const {sheets} = api.readFixtureCommentThreads(SAMPLE);
        assert.deepStrictEqual(sheets[0].notes, {});
      },
    },
  ],
} satisfies Case;
