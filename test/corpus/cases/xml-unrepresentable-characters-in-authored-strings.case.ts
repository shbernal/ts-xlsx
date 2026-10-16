// Cluster: xlsx-io
//
// Real-world scenario: the strings an application writes into a workbook come from a database, a CSV
// export or a user form, and some of them carry characters XML 1.0 cannot hold: a C0 control left by
// a legacy field terminator, a U+FFFE noncharacter from a mis-decoded byte order mark, half of a
// surrogate pair from a truncated UTF-16 field. Passed through raw, they produce a package the
// application reports as damaged and refuses to open.
//
// SpreadsheetML has one convention for this and it covers cell values only: `_xHHHH_`, which Excel
// writes and reads back. So a cell value carrying such a character is escaped and survives intact,
// while a sheet name, which has no such convention and whose tab would otherwise read
// `Sheet_x0001_A`, is refused outright. Because `_xHHHH_` now means something, a value that
// legitimately contains the
// literal text `_x0041_` must have its underscore escaped as well, or it would come back as `A`.

import type {Assert, Case, CorpusApi} from '../case.ts';

const CONTROL = '\u0001';

export default {
  id: 'xml-unrepresentable-characters-in-authored-strings',
  cluster: 'xlsx-io',
  provenance: {source: 'writer-boundary-probe'},
  description:
    'A string carrying a character XML 1.0 cannot represent is never emitted raw: a cell value is ' +
    'escaped with the SpreadsheetML _xHHHH_ convention, and a sheet name, which has no such ' +
    'convention, is refused with an error naming the code point.',

  behavior: [
    {
      name: 'a U+FFFE noncharacter in cell text is escaped, not passed through',
      expect(api: CorpusApi, assert: Assert) {
        const {partsWithRawChar, emittedText} = api.xmlCharacterSafetyReport(
          'cell-text',
          'a\uFFFEb',
        );
        assert.deepStrictEqual(partsWithRawChar, []);
        assert.strictEqual(emittedText, 'a_xFFFE_b');
      },
    },
    {
      name: 'text that only resembles an escape is left alone',
      expect(api: CorpusApi, assert: Assert) {
        const {emittedText} = api.xmlCharacterSafetyReport('cell-text', 'a_ b_x c_xZZZZ_ d');
        assert.strictEqual(emittedText, 'a_ b_x c_xZZZZ_ d');
      },
    },
    {
      name: 'tab, line feed and carriage return remain legal cell text',
      expect(api: CorpusApi, assert: Assert) {
        const {writeOk, emittedText} = api.xmlCharacterSafetyReport('cell-text', 'a\tb\nc\rd');
        assert.strictEqual(writeOk, true);
        assert.strictEqual(emittedText, 'a\tb\nc\rd', 'these three are valid XML 1.0 characters');
      },
    },
    {
      name: 'U+007F (DEL) remains legal cell text',
      expect(api: CorpusApi, assert: Assert) {
        const {writeOk, emittedText} = api.xmlCharacterSafetyReport('cell-text', 'a\u007Fb');
        assert.strictEqual(writeOk, true);
        assert.strictEqual(emittedText, 'a\u007Fb', 'DEL is forbidden by XML 1.1, not by XML 1.0');
      },
    },
    {
      name: 'a control character in a cached string formula result is escaped, not emitted raw',
      expect(api: CorpusApi, assert: Assert) {
        const {writeOk, partsWithRawChar, emittedText} = api.xmlCharacterSafetyReport(
          'formula-result',
          `a${CONTROL}b`,
        );
        assert.strictEqual(writeOk, true);
        assert.deepStrictEqual(partsWithRawChar, []);
        assert.strictEqual(emittedText, 'a_x0001_b');
      },
    },
    {
      name: 'a control character in a sheet name is refused rather than written',
      expect(api: CorpusApi, assert: Assert) {
        const {writeOk, partsWithRawChar} = api.xmlCharacterSafetyReport(
          'sheet-name',
          `Sheet${CONTROL}A`,
        );
        assert.strictEqual(
          writeOk,
          false,
          'a sheet tab has no _xHHHH_ convention, so there is no faithful way to write this name',
        );
        assert.deepStrictEqual(partsWithRawChar, []);
      },
    },
    {
      name: 'the refusal names the code point it could not write',
      expect(api: CorpusApi, assert: Assert) {
        const {writeError} = api.xmlCharacterSafetyReport('sheet-name', `Sheet${CONTROL}A`);
        assert.match(
          String(writeError),
          /U\+0001/,
          `the error must identify the offending character; got ${JSON.stringify(writeError)}`,
        );
      },
    },
  ],
} satisfies Case;
