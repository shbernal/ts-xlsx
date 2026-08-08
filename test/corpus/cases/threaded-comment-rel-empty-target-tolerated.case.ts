// Cluster: comment
//
// Real-world scenario: a sheet declares a threaded-comment relationship (type
// `.../2017/10/relationships/threadedComment`) whose Target attribute is empty — it announces a
// conversation without saying where the part is. Foreign generators emit relationships
// unconditionally, and a hand-edited or partially-repaired package lands in the same shape. The
// relationship is the *only* wiring for these parts (no worksheet element names one, the way none names
// a pivot table), so a blank target is the whole difference between reachable and not.
//
// Two things must hold. The load must survive: dereferencing the empty target and then reading a
// property off the missing part is the exact crash the sibling `worksheet-comment-rel-empty-target`
// case pins for the notes relationship. And the *text* must survive, which is the less obvious half.
// Excel writes a legacy fallback `<comment>` beside every conversation — the "[Threaded comment] Your
// version of Excel..." boilerplate wrapping a copy of the messages — and we normally suppress that on
// read, since surfacing it as `cell.note` would hand the caller garbage and re-emitting it as a plain
// note would destroy the `tc=`/`xr:uid` binding Excel resolves the thread through. Suppression is
// therefore conditioned on holding the conversation the fallback *names*, not merely on the cell having
// been threaded once: with the thread part out of reach the boilerplate is the last remaining record of
// what anyone said, so it is kept as a note rather than dropped.
//
// The fixture is the Excel-authored `resolved-multi-author.xlsx` (a resolved thread with a reply by a
// second author, an open thread, and a genuine legacy note on a third cell) with one edit: its sheet
// rels carry `Target=""` on the threadedComment relationship. Every part is otherwise intact, including
// the thread part itself — which is the point. A part nothing can reach is not adopted by filename.

import type {Assert, Case, CorpusApi} from '../case.ts';

const FIXTURE = 'threaded-comment-rel-empty-target-tolerated/empty-threaded-comment-rel.xlsx';

// The boilerplate opening, and the two threaded cells and one genuinely-noted cell of the fixture.
const BOILERPLATE = /^\[Threaded comment\]\n\nYour version of Excel/;
const THREADED_CELLS = ['B1', 'B2'];
const NOTED_CELL = 'D4';

export default {
  id: 'threaded-comment-rel-empty-target-tolerated',
  provenance: {source: 'foreign-generator-probe'},
  cluster: 'comment',
  description:
    'Loading a workbook whose worksheet declares a threadedComment relationship with an empty Target ' +
    'completes without throwing and recovers the sheet. No conversation is invented from the ' +
    'unreachable part, and because the thread is not held, the legacy fallback comment is kept as a ' +
    'note — the last remaining record of what was said — rather than suppressed. The re-written ' +
    'package carries no thread part, no author registry and no tc= fallback author, so it never emits ' +
    'half a representation.',

  behavior: [
    {
      name: 'a blank threadedComment relationship target does not abort the load',
      async expect(api: CorpusApi, assert: Assert) {
        const {ok, error, sheetNames} = await api.readFixtureReport(FIXTURE);
        assert.strictEqual(
          ok,
          true,
          `a blank threadedComment-rel target must not abort the load; got ${JSON.stringify(error)}`,
        );
        assert.deepStrictEqual(sheetNames, ['Review'], 'the worksheet survives the tolerant read');
      },
    },
    {
      name: 'no conversation is read from a thread part the sheet cannot reach',
      async expect(api: CorpusApi, assert: Assert) {
        // The part is still in the package. Finding it would mean guessing by filename instead of
        // following the relationship, and a guess is how a conversation gets attached to the wrong
        // sheet in a multi-sheet workbook.
        const facts = await api.readFixtureCommentThreads(FIXTURE, THREADED_CELLS);
        assert.deepStrictEqual(facts.sheets[0]!.threads, [], 'no threads');
        assert.deepStrictEqual(
          facts.sheets[0]!.at,
          {B1: null, B2: null},
          'and neither cell reports one',
        );
      },
    },
    {
      name: 'the author registry is still read, since its own relationship is intact',
      async expect(api: CorpusApi, assert: Assert) {
        // The persons part is workbook-level and wired independently, so one sheet's broken
        // relationship must not take the identities down with it — they are what a repaired or
        // re-linked conversation would resolve through.
        const facts = await api.readFixtureCommentThreads(FIXTURE);
        assert.strictEqual(facts.persons.length, 2, 'both registered persons survive');
      },
    },
    {
      name: 'the fallback boilerplate is kept as a note, so the words are not lost with the thread',
      async expect(api: CorpusApi, assert: Assert) {
        const facts = await api.readFixtureCommentThreads(FIXTURE);
        const {notes} = facts.sheets[0]!;
        assert.deepStrictEqual(
          Object.keys(notes).sort(),
          [...THREADED_CELLS, NOTED_CELL].sort(),
          'every comment in the part surfaces, the two fallbacks included',
        );
        for (const ref of THREADED_CELLS) {
          assert.match(notes[ref] ?? '', BOILERPLATE, `${ref} keeps the fallback text`);
        }
        assert.match(
          notes.B1 ?? '',
          /Comment:\n {4}Is this gross or net of tax\?\nReply:\n {4}Gross\. Confirmed with finance\./,
          'including the conversation copied inside it — both messages, not just the first',
        );
        assert.strictEqual(
          notes[NOTED_CELL],
          'A genuine legacy note.',
          'and a real note is unaffected either way',
        );
      },
    },
    {
      name: 'the re-written package emits neither the thread part nor an orphaned tc= fallback',
      async expect(api: CorpusApi, assert: Assert) {
        // The two halves of a conversation are emitted together or not at all: verified against
        // desktop Excel, a `tc=` fallback whose thread part is absent shows as neither a thread nor a
        // note — the text disappears entirely. Holding no thread, the writer emits plain notes, which
        // is the shape that still renders.
        const {source, rewritten} = await api.roundtripFixturePackageParts(FIXTURE);
        assert.strictEqual(source.threadedComments, 1, 'precondition: the part is in the package');
        assert.strictEqual(source.commentFallbackThreadAuthors, 2, 'precondition: two tc= authors');
        assert.strictEqual(
          rewritten.threadedComments,
          0,
          'an unreachable conversation is not re-emitted as if it had been read',
        );
        assert.strictEqual(
          rewritten.persons,
          0,
          'and the registry is not written beside no conversation',
        );
        assert.strictEqual(
          rewritten.commentFallbackThreadAuthors,
          0,
          'no fallback claims a thread that is not there',
        );
      },
    },
    {
      name: 'every comment still gets exactly one box to render into',
      async expect(api: CorpusApi, assert: Assert) {
        const {source, rewritten} = await api.roundtripFixturePackageParts(FIXTURE);
        assert.strictEqual(source.commentEntries, 3, 'precondition: two fallbacks and one note');
        assert.strictEqual(rewritten.commentEntries, 3, 'all three are written back');
        assert.strictEqual(
          rewritten.commentVmlShapes,
          rewritten.commentEntries,
          'one VML shape per comment — a comment without one reads as text but renders nothing',
        );
      },
    },
  ],
} satisfies Case;
