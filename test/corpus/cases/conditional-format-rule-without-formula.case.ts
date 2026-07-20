// Cluster: formulas
//
// Real-world scenario: a worksheet carries a conditional-formatting rule whose formula list is absent
// — for example an expression-type rule authored by another tool that emitted the rule without a
// formula element. When the library serializes the worksheet, the conditional-formatting writer
// assumes every rule has at least one formula and indexes into the formula list, throwing a TypeError
// ("Cannot read properties of undefined (reading '0')") instead of emitting XML. Writing such a
// workbook must succeed and preserve the rule rather than crashing on the missing formula.

import type {Assert, Case, CorpusApi} from '../case.ts';

export default {
  id: 'conditional-format-rule-without-formula',
  provenance: {source: 'upstream-issue'},
  cluster: 'formulas',
  description:
    'Serializing a worksheet whose conditional-formatting rule has no formula must not throw — the ' +
    'writer must tolerate a formula-less rule rather than indexing into an absent formula list.',

  behavior: [
    {
      name: 'an expression rule WITH a formula writes successfully (control)',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const {writeOk, writeError} = await api.authorConditionalFormatting({
          ref: 'A1:A10',
          rules: [{type: 'expression', formulae: ['$A1>0'], style: {}}],
        });
        assert.strictEqual(
          writeOk,
          true,
          `a well-formed expression rule must write; error: ${writeError}`,
        );
      },
    },
    {
      name: 'an expression rule with NO formula does not crash serialization',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const {writeOk, writeError} = await api.authorConditionalFormatting({
          ref: 'A1:A10',
          rules: [{type: 'expression', style: {}}],
        });
        assert.strictEqual(
          writeOk,
          true,
          `a formula-less conditional-formatting rule must serialize, not throw; got ${writeError}`,
        );
      },
    },
  ],
} satisfies Case;
