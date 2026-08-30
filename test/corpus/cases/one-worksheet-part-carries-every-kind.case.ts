import type {Assert, Case, CorpusApi} from '../case.ts';

export default {
  id: 'one-worksheet-part-carries-every-kind',
  cluster: 'core-model',
  provenance: {source: 'reader-composition-audit'},
  description:
    'A worksheet part carries several unrelated kinds of content at once: a hyperlink, a standard ' +
    'data validation, a cross-sheet validation that can only live in the 2009 extension block, a ' +
    'classic conditional format, and a data bar whose gradient and negative-fill facets live in ' +
    'that same extension. A sheet built this way is ordinary, and every one of the five must survive ' +
    'a write and come back, none of them shadowed by another.',
  behavior: [
    {
      name: 'the hyperlink and its visible label both come back',
      expect(api: CorpusApi, assert: Assert) {
        const report = api.worksheetPartCarryingEveryKind();
        assert.strictEqual(report.hyperlink, 'https://example.com/');
        assert.strictEqual(report.cellText, 'go');
      },
    },
    {
      name: 'the standard and extended validations both come back, each in its own form',
      expect(api: CorpusApi, assert: Assert) {
        const report = api.worksheetPartCarryingEveryKind();
        assert.strictEqual(report.standardValidation, '"one,two,three"');
        assert.strictEqual(
          report.extendedValidation,
          'Other!$A$1:$A$1',
          'the cross-sheet list is read back out of the extension block, not dropped',
        );
      },
    },
    {
      name: 'the classic rule and the data bar both come back, the bar with its extension facets',
      expect(api: CorpusApi, assert: Assert) {
        const report = api.worksheetPartCarryingEveryKind();
        assert.strictEqual(report.classicRuleType, 'cellIs');
        assert.strictEqual(
          report.dataBarGradient,
          false,
          'a non-gradient bar is only expressible in the extension, so this proves it was married back',
        );
        assert.strictEqual(report.dataBarNegativeFill, 'FFFF0000');
      },
    },
  ],
} satisfies Case;
