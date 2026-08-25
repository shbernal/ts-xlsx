import type {Assert, Case, CorpusApi} from '../case.ts';

const FIXTURE = 'escaped-characters-in-cell-text-decode-on-read/xhhhh-escapes.xlsx';
const REFS = ['A1', 'A2', 'A3', 'A4', 'A5', 'A6', 'A7', 'B1'];

export default {
  id: 'escaped-characters-in-cell-text-decode-on-read',
  cluster: 'xlsx-io',
  description:
    'SpreadsheetML spells a character its XML cannot carry as `_xHHHH_`, so a workbook Excel ' +
    'authored with a control character in a cell holds the seven-character text `_x0001_` on disk ' +
    'and must read back as the character. The escape is a closed grammar: exactly four hex digits ' +
    'between `_x` and `_`, decoded in one left-to-right pass so `_x005F_` in front of an escape ' +
    'yields the literal text rather than collapsing to the character it looks like. Anything that ' +
    'merely resembles the shape is ordinary text and must survive untouched.',
  provenance: {source: 'excel-desktop-verification'},
  behavior: [
    {
      name: 'a `_x0041_` in an inline string reads back as the character it names',
      expect(api: CorpusApi, assert: Assert) {
        const {eager} = api.escapeDecodeReport(FIXTURE, REFS);
        assert.strictEqual(eager.A1, 'A');
      },
    },
    {
      name: 'a `_x0041_` in the shared-strings pool decodes the same as one in an inline string',
      expect(api: CorpusApi, assert: Assert) {
        const {eager} = api.escapeDecodeReport(FIXTURE, REFS);
        assert.strictEqual(eager.A7, 'A');
      },
    },
    {
      name: 'the cached `<v>` of a t=str formula decodes too, so Excel reads that cell as A',
      expect(api: CorpusApi, assert: Assert) {
        const {eager} = api.escapeDecodeReport(FIXTURE, REFS);
        assert.strictEqual(eager.B1, 'A');
      },
    },
    {
      name: '`_x005F_x0041_` is the escaped underscore, so the cell holds the literal text `_x0041_`',
      expect(api: CorpusApi, assert: Assert) {
        // The one decision a second decoding pass would silently get wrong: it would collapse this
        // to 'A' and lose every value that legitimately reads like an escape.
        const {eager} = api.escapeDecodeReport(FIXTURE, REFS);
        assert.strictEqual(eager.A2, '_x0041_');
      },
    },
    {
      name: 'text that only resembles the escape shape is left exactly as written',
      expect(api: CorpusApi, assert: Assert) {
        const {eager} = api.escapeDecodeReport(FIXTURE, REFS);
        assert.strictEqual(eager.A3, '_xZZZZ_');
        assert.strictEqual(eager.A4, '_x041_');
        assert.strictEqual(eager.A5, '_x00041_');
      },
    },
    {
      name: 'the escape decodes even for a character XML could have carried literally',
      expect(api: CorpusApi, assert: Assert) {
        // Excel writes `_x0009_` for a tab and reads it back as a tab, though a literal tab is
        // perfectly legal in XML. A decoder limited to the unrepresentable range would disagree
        // with Excel on Excel's own files.
        const {eager} = api.escapeDecodeReport(FIXTURE, REFS);
        assert.strictEqual(eager.A6, 'a\tb');
      },
    },
    {
      name: 'the streaming reader decodes identically to the buffered one, cell for cell',
      expect(api: CorpusApi, assert: Assert) {
        const {eager, streaming} = api.escapeDecodeReport(FIXTURE, REFS);
        for (const ref of REFS) assert.strictEqual(streaming[ref], eager[ref], `at ${ref}`);
      },
    },
    {
      name: 'every value survives a write through our own writer and a read back',
      expect(api: CorpusApi, assert: Assert) {
        // The assertion that would have caught the escape and the unescape landing out of step:
        // each is the other's inverse, so re-emitting what we read must not change any value.
        const {eager, roundtrip} = api.escapeDecodeReport(FIXTURE, REFS);
        for (const ref of REFS) assert.strictEqual(roundtrip[ref], eager[ref], `at ${ref}`);
      },
    },
    {
      name: 'a control character authored into cell text survives write then read unchanged',
      expect(api: CorpusApi, assert: Assert) {
        const text = 'a\u0001b';
        const report = api.xmlCharacterSafetyReport('cell-text', text);
        assert.strictEqual(report.writeOk, true);
        assert.deepStrictEqual(report.partsWithRawChar, []);
        assert.strictEqual(report.emittedText, 'a_x0001_b');
        assert.strictEqual(report.readValue, text);
      },
    },
    {
      name: 'a lone surrogate survives the same round trip rather than becoming U+FFFD',
      expect(api: CorpusApi, assert: Assert) {
        const text = 'a\uD800b';
        const report = api.xmlCharacterSafetyReport('cell-text', text);
        assert.strictEqual(report.emittedText, 'a_xD800_b');
        assert.strictEqual(report.readValue, text);
      },
    },
    {
      name: 'text that already reads like an escape round-trips as itself, not as its character',
      expect(api: CorpusApi, assert: Assert) {
        const report = api.xmlCharacterSafetyReport('cell-text', 'a_x0041_b');
        assert.strictEqual(report.emittedText, 'a_x005F_x0041_b');
        assert.strictEqual(report.readValue, 'a_x0041_b');
      },
    },
    {
      name: "a control character in a formula's cached result survives write then read",
      expect(api: CorpusApi, assert: Assert) {
        const text = 'a\u0001b';
        const report = api.xmlCharacterSafetyReport('formula-result', text);
        assert.strictEqual(report.emittedText, 'a_x0001_b');
        assert.strictEqual(report.readValue, text);
      },
    },
    {
      name: 'the other cell-text carriers — a note body, a rich-text run — round-trip the same way',
      expect(api: CorpusApi, assert: Assert) {
        const text = 'a\u0001b';
        const {note, runs} = api.escapedTextCarrierRoundtrip(text);
        assert.strictEqual(note, text);
        assert.deepStrictEqual(runs, [text, `${text}!`]);
      },
    },
  ],
} satisfies Case;
