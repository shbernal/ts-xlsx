// Cluster: formulas
//
// Real-world scenario: a worksheet has formula cells whose cached results are falsy — a numeric 0, a
// boolean false, or an empty string. When such a cell is read back (or copied/cloned as part of
// normal model handling), the formula's result must survive exactly. A reported failure: copy logic
// that decides which fields to carry over with a truthiness test silently drops any result equal to
// 0, false, or "", so the cell comes back as a formula with no result even though the source recorded
// one. A truthy result is unaffected — which is exactly what pinpoints the truthiness bug.

import type {Assert, Case, CorpusApi} from '../case.ts';

export default {
  id: 'formula-cell-preserves-falsy-result-values',
  provenance: {source: 'upstream-issue'},
  cluster: 'formulas',
  description:
    'A formula cell whose cached result is falsy (0, false, or "") round-trips with that exact result ' +
    'preserved, rather than a truthiness test dropping it; a truthy result is unaffected.',

  behavior: [
    {
      name: 'a truthy formula result is preserved (control)',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const {truthy} = await api.formulaFalsyResultReport();
        assert.strictEqual(truthy.hasResult, true, 'the truthy result is present');
        assert.strictEqual(truthy.result, 2, 'the truthy result value is preserved');
      },
    },
    {
      name: 'a formula result of 0 is preserved, not dropped',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const {zero} = await api.formulaFalsyResultReport();
        assert.strictEqual(zero.hasResult, true, 'the numeric-zero result must survive');
        assert.strictEqual(zero.result, 0, 'the result reads back as exactly 0');
      },
    },
    {
      name: 'a formula result of false is preserved, not dropped',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const {boolFalse} = await api.formulaFalsyResultReport();
        assert.strictEqual(boolFalse.hasResult, true, 'the boolean-false result must survive');
        assert.strictEqual(boolFalse.result, false, 'the result reads back as exactly false');
      },
    },
    {
      name: 'a formula result of empty string is preserved, not dropped',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const {emptyString} = await api.formulaFalsyResultReport();
        assert.strictEqual(emptyString.hasResult, true, 'the empty-string result must survive');
        assert.strictEqual(
          emptyString.result,
          '',
          'the result reads back as exactly the empty string',
        );
      },
    },
  ],
} satisfies Case;
