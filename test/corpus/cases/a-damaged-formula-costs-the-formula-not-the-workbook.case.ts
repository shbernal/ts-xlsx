// Cluster: formulas
//
// Real-world scenario: a `.xlsb` arrives with one cell's token stream damaged -- a last token that
// declares an operand past the end of its own `rgce`. The BIFF12 decoder's contract already covers
// the neighbouring case in as many words: "a token this decoder does not know makes the whole formula
// undecodable, by design... the decoder returns `undefined` and its caller keeps the cached result".
// A truncated token is the same damage, and it did not degrade the same way: it threw out of the
// record reader, uncaught, so one damaged formula in one cell discarded an otherwise fully
// recoverable workbook -- every sheet, every value, every other formula.
//
// The rule this locks: the record-level bound is what fails a malformed file closed; a formula that
// runs off its own end is one undecodable formula. Excel opens such a file and shows the cached
// result, which is what the cache is for.

import type {Assert, Case, CorpusApi} from '../case.ts';

export default {
  id: 'a-damaged-formula-costs-the-formula-not-the-workbook',
  provenance: {source: 'hostile-input-probe'},
  cluster: 'formulas',
  description:
    'A BIFF12 token stream whose last token runs past the end of its own rgce leaves that formula ' +
    'undecodable and its cached result in place; the workbook, its sheets and its other formulas ' +
    'read exactly as they would have.',

  behavior: [
    {
      name: 'the workbook still reads',
      expect(api: CorpusApi, assert: Assert) {
        const report = api.xlsbTruncatedFormulaToken();
        assert.equal(report.threw, false, `the read survives it (${report.errorName})`);
        assert.equal(report.sheets, 4, 'every sheet is still there');
      },
    },
    {
      name: 'the damaged cell keeps the result Excel cached beside its formula',
      expect(api: CorpusApi, assert: Assert) {
        const report = api.xlsbTruncatedFormulaToken();
        assert.equal(report.damagedFormula, null, 'the formula is undecodable');
        assert.equal(report.damagedCell, 7, 'and the cached result is what the cell reads as');
      },
    },
    {
      name: 'no other formula is touched',
      expect(api: CorpusApi, assert: Assert) {
        const report = api.xlsbTruncatedFormulaToken();
        assert.equal(report.siblingFormula, 'A1*2', 'the cell beside it decodes normally');
        assert.equal(
          report.formulas,
          40,
          'and the workbook keeps every formula but the damaged one',
        );
      },
    },
  ],
} satisfies Case;
