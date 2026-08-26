// Cluster: page-setup
//
// Real-world scenario: a print header is prose a human typed into Page Setup, a document title, a
// client name, a path pasted out of a system that put a control character in it. It reaches the file as
// the text of an `<oddHeader>`, which is `ST_Xstring` exactly as a `<t>` is, so the same question
// arises: does SpreadsheetML's `_xHHHH_` convention apply here, or is a header plain text? The answer
// was assumed from the format's shape: the `&`-prefixed section codes look like the only in-band
// syntax a header has, and the assumption cost callers a header that could not carry what a note
// could.
//
// Excel Desktop settled it (2026-08-25, Microsoft 365 16.0 build 20228): a `<headerFooter>` patched to
// hold the eight grammar rows the cell-text probe used reads back over COM through `PageSetup` with the
// identical verdict on every one: `_x0041_` decodes to `A`, `_x005F_x0041_` yields the literal seven
// characters, `_xZZZZ_`/`_x041_`/`_x00041_` come back untouched, and a tab is decoded even though XML
// could have carried it. Excel then re-saved the package and wrote a control character back as
// `_x0001_` and the literal back as `_x005F_x0041_`, so the escape is its own representation here, not
// merely tolerated on read. The measurement is written up in
// `docs/knowledge/specs/spreadsheetml-xhhhh-escape-is-decoded-on-read.md`; the fixture below is the
// exact package Excel was shown.

import type {Assert, Case, CorpusApi} from '../case.ts';

const FIXTURE = 'print-header-text-decodes-xhhhh-escape/xhhhh-escapes-in-headers.xlsx';

export default {
  id: 'print-header-text-decodes-xhhhh-escape',
  cluster: 'page-setup',
  provenance: {source: 'excel-desktop-verification'},
  description:
    'A print header/footer definition carries the `_xHHHH_` convention exactly as cell text does: ' +
    'Excel decodes it on read and writes it back on save, so this library must escape on write and ' +
    'decode on read. The grammar is closed: four hex digits, no more and no fewer, and resolved in ' +
    'one left-to-right pass, so `_x005F_` in front of an escape yields the literal text rather than ' +
    'the character it resembles. The consequence for callers is that a header may carry any character ' +
    'a note or a cell may, instead of being refused for one XML cannot spell.',

  behavior: [
    {
      name: 'a `_x0041_` in a header section reads back as the character it names',
      expect(api: CorpusApi, assert: Assert) {
        const {eager} = api.headerFooterEscapeReport(FIXTURE);
        assert.strictEqual(eager.oddHeader, '&LA&C_x0041_&R_xZZZZ_');
      },
    },
    {
      name: 'the section codes around an escape are untouched by the decode',
      expect(api: CorpusApi, assert: Assert) {
        // The decode must not disturb the `&`-prefixed syntax the header is made of: an `&L`/`&C`/`&R`
        // that shifted or vanished would silently move the text to another corner of the page.
        const {eager} = api.headerFooterEscapeReport(FIXTURE);
        assert.match(String(eager.oddHeader), /^&L.*&C.*&R/);
        assert.match(String(eager.oddFooter), /^&L.*&C.*&R/);
      },
    },
    {
      name: '`_x005F_x0041_` is the escaped underscore, so the header holds the literal `_x0041_`',
      expect(api: CorpusApi, assert: Assert) {
        // The one decision a second decoding pass would silently get wrong: it would collapse this to
        // `A` and destroy every header that legitimately looks like an escape.
        const {eager} = api.headerFooterEscapeReport(FIXTURE);
        assert.ok(String(eager.oddHeader).includes('&C_x0041_&R'));
      },
    },
    {
      name: 'text that only resembles the escape shape is left exactly as written',
      expect(api: CorpusApi, assert: Assert) {
        const {eager} = api.headerFooterEscapeReport(FIXTURE);
        assert.ok(String(eager.oddHeader).includes('&R_xZZZZ_'));
        assert.ok(String(eager.oddFooter).includes('&L_x041_'));
        assert.ok(String(eager.oddFooter).includes('&C_x00041_'));
      },
    },
    {
      name: 'the escape decodes even for a character XML could have carried literally',
      expect(api: CorpusApi, assert: Assert) {
        const {eager} = api.headerFooterEscapeReport(FIXTURE);
        assert.ok(String(eager.oddFooter).endsWith('&Ra\tb'));
      },
    },
    {
      name: 'a character XML cannot carry survives the escape Excel itself writes for it',
      expect(api: CorpusApi, assert: Assert) {
        // The row that matters most: Excel wrote `_x0001_` here unprompted when it re-saved the
        // package, so a reader that skipped the decode would report seven characters of noise as the
        // words printed on every even page.
        const {eager} = api.headerFooterEscapeReport(FIXTURE);
        assert.strictEqual(eager.evenFooter, '&L\u0001');
      },
    },
    {
      name: 'a header with nothing to decode is untouched by the decoder',
      expect(api: CorpusApi, assert: Assert) {
        const {eager} = api.headerFooterEscapeReport(FIXTURE);
        assert.strictEqual(eager.evenHeader, '&Lplain');
      },
    },
    {
      name: 'a header slot the file never declared stays absent rather than becoming empty text',
      expect(api: CorpusApi, assert: Assert) {
        const {eager} = api.headerFooterEscapeReport(FIXTURE);
        assert.strictEqual(eager.firstHeader, null);
        assert.strictEqual(eager.firstFooter, null);
      },
    },
    {
      name: 'every header slot survives a write through our own writer and a read back',
      expect(api: CorpusApi, assert: Assert) {
        // The assertion that would catch the escape and the decode landing out of step: each is the
        // other's inverse, so re-emitting what we read must not change a single header.
        const {eager, roundtrip} = api.headerFooterEscapeReport(FIXTURE);
        assert.deepStrictEqual(roundtrip, eager);
      },
    },
    {
      name: 'a control character authored into a header is written as its escape, never raw',
      expect(api: CorpusApi, assert: Assert) {
        const report = api.authoredHeaderFooterEscape('&Cbefore\u0001after');
        assert.strictEqual(report.emitted, '&amp;Cbefore_x0001_after');
        assert.deepStrictEqual(report.rawInPart, []);
        assert.strictEqual(report.read, '&Cbefore\u0001after');
      },
    },
    {
      name: 'a lone surrogate in a header survives rather than becoming U+FFFD',
      expect(api: CorpusApi, assert: Assert) {
        const text = '&Ca\uD800b';
        const report = api.authoredHeaderFooterEscape(text);
        assert.strictEqual(report.emitted, '&amp;Ca_xD800_b');
        assert.strictEqual(report.read, text);
      },
    },
    {
      name: 'a header that already reads like an escape round-trips as itself',
      expect(api: CorpusApi, assert: Assert) {
        const report = api.authoredHeaderFooterEscape('&Ca_x0041_b');
        assert.strictEqual(report.emitted, '&amp;Ca_x005F_x0041_b');
        assert.strictEqual(report.read, '&Ca_x0041_b');
      },
    },
  ],
} satisfies Case;
