// Cluster: comment
//
// Real-world scenario: a reviewer's comment carries a character XML 1.0 has no syntax for, a control
// character pasted out of a database column, a lone surrogate from a mangled name. In a legacy note that
// text survives, because a note's body is a `<t>` and `<t>` has SpreadsheetML's `_xHHHH_` convention. A
// threaded comment's `<text>` is a different element in the 2018 extension namespace, and nothing in that
// schema documents an escape, so the same string used to be refused there: one comment system accepting
// what the other rejected, for a reason no caller could see.
//
// Excel Desktop settled it (2026-08-25, Microsoft 365 16.0 build 20228): a `<text>` patched to hold
// `_x0041_` reads back over COM as `A`, `_x005F_x0041_` as the literal seven-character text, and
// `_xZZZZ_`/`_x041_`/`_x00041_` untouched: the same closed grammar and the same single left-to-right
// pass as cell text. Excel then re-saved the package and wrote a control character back out as
// `_x0001_`, so the escape is its own representation here and not merely tolerated on read. The
// measurement is written up in `docs/knowledge/specs/spreadsheetml-xhhhh-escape-is-decoded-on-read.md`;
// the fixture below is the exact package Excel was shown.

import type {Assert, Case, CorpusApi} from '../case.ts';

const FIXTURE = 'threaded-comment-text-decodes-xhhhh-escape/xhhhh-escapes-in-threads.xlsx';
const REFS = ['A1', 'A2', 'A3', 'A4', 'A5', 'A6', 'A7', 'A8'];

export default {
  id: 'threaded-comment-text-decodes-xhhhh-escape',
  cluster: 'comment',
  provenance: {source: 'excel-desktop-verification'},
  description:
    "A threaded comment's `<text>` carries the `_xHHHH_` convention exactly as cell text does: Excel " +
    'decodes it on read and writes it back on save, so this library must escape on write and decode on ' +
    'read. The grammar is closed: four hex digits, no more and no fewer, and resolved in one ' +
    'left-to-right pass, so `_x005F_` in front of an escape yields the literal text rather than the ' +
    'character it resembles. The consequence for callers is that the two comment systems finally agree ' +
    'on the same string.',

  behavior: [
    {
      name: 'a `_x0041_` in a message reads back as the character it names',
      expect(api: CorpusApi, assert: Assert) {
        const {eager} = api.threadedCommentEscapeReport(FIXTURE, REFS);
        assert.strictEqual(eager.A1, 'A');
      },
    },
    {
      name: 'a character XML cannot carry survives the escape Excel itself writes for it',
      expect(api: CorpusApi, assert: Assert) {
        // A8 is the case that matters most: Excel wrote `_x0001_` here unprompted when it re-saved the
        // package, so a reader that skipped the decode would report seven characters of noise as the
        // reviewer's words.
        const {eager} = api.threadedCommentEscapeReport(FIXTURE, REFS);
        assert.strictEqual(eager.A8, '\u0001');
      },
    },
    {
      name: '`_x005F_x0041_` is the escaped underscore, so the message holds the literal `_x0041_`',
      expect(api: CorpusApi, assert: Assert) {
        // The one decision a second decoding pass would silently get wrong: it would collapse this to
        // `A` and destroy every message that legitimately reads like an escape.
        const {eager} = api.threadedCommentEscapeReport(FIXTURE, REFS);
        assert.strictEqual(eager.A2, '_x0041_');
      },
    },
    {
      name: 'text that only resembles the escape shape is left exactly as written',
      expect(api: CorpusApi, assert: Assert) {
        const {eager} = api.threadedCommentEscapeReport(FIXTURE, REFS);
        assert.strictEqual(eager.A3, '_xZZZZ_');
        assert.strictEqual(eager.A4, '_x041_');
        assert.strictEqual(eager.A5, '_x00041_');
      },
    },
    {
      name: 'the escape decodes even for a character XML could have carried literally',
      expect(api: CorpusApi, assert: Assert) {
        const {eager} = api.threadedCommentEscapeReport(FIXTURE, REFS);
        assert.strictEqual(eager.A6, 'a\tb');
      },
    },
    {
      name: 'a message with nothing to decode is untouched by the decoder',
      expect(api: CorpusApi, assert: Assert) {
        const {eager} = api.threadedCommentEscapeReport(FIXTURE, REFS);
        assert.strictEqual(eager.A7, 'plain');
      },
    },
    {
      name: 'every message survives a write through our own writer and a read back',
      expect(api: CorpusApi, assert: Assert) {
        // The assertion that would catch the escape and the decode landing out of step: each is the
        // other's inverse, so re-emitting what we read must not change a single message.
        const {eager, roundtrip} = api.threadedCommentEscapeReport(FIXTURE, REFS);
        for (const ref of REFS) assert.strictEqual(roundtrip[ref], eager[ref], `at ${ref}`);
      },
    },
    {
      name: 'a control character authored into a message is written as its escape, never raw',
      expect(api: CorpusApi, assert: Assert) {
        const text = 'before\u0001after';
        const report = api.authoredThreadedCommentEscape(text);
        assert.strictEqual(report.emittedText, 'before_x0001_after');
        // The legacy fallback the writer builds beside the thread carries the same words through a
        // `<t>`, so this also proves the two halves of one package agree.
        assert.deepStrictEqual(report.rawInPart, []);
        assert.strictEqual(report.readText, text);
      },
    },
    {
      name: 'a lone surrogate in a message survives rather than becoming U+FFFD',
      expect(api: CorpusApi, assert: Assert) {
        const text = 'a\uD800b';
        const report = api.authoredThreadedCommentEscape(text);
        assert.strictEqual(report.emittedText, 'a_xD800_b');
        assert.strictEqual(report.readText, text);
      },
    },
    {
      name: 'a message that already reads like an escape round-trips as itself',
      expect(api: CorpusApi, assert: Assert) {
        const report = api.authoredThreadedCommentEscape('a_x0041_b');
        assert.strictEqual(report.emittedText, 'a_x005F_x0041_b');
        assert.strictEqual(report.readText, 'a_x0041_b');
      },
    },
  ],
} satisfies Case;
