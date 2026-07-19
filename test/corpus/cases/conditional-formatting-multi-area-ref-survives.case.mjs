// Cluster: conditional-formatting
//
// Real-world scenario: a user applies a single conditional-formatting rule (e.g. a colorScale) to
// several non-contiguous blocks at once — rows 1, 3, and 5 of a grid — so the scale is computed
// across the union of all those cells. In OOXML this is one conditionalFormatting element whose
// sqref lists multiple space-separated areas ("A1:C1 A3:C3 A5:C5"), exactly what Excel produces when
// you select multiple ranges and apply one rule. Supplying the multiple areas in one ref must emit
// that single multi-area rule, not silently drop it to zero conditionalFormatting elements.

/** @typedef {{ name: string, baseline: 'pass'|'fail', expect: (api: any, assert: any) => Promise<void>|void }} Behavior */

const CF = {
  ref: 'A1:C1 A3:C3 A5:C5',
  rules: [
    {
      type: 'colorScale',
      cfvo: [{type: 'min'}, {type: 'max'}],
      color: [{argb: 'FFFF0000'}, {argb: 'FF00FF00'}],
    },
  ],
};

export default {
  id: 'conditional-formatting-multi-area-ref-survives',
  provenance: {source: 'upstream-issue'},
  cluster: 'conditional-formatting',
  description:
    'A conditional-formatting rule whose ref names multiple non-contiguous areas emits exactly one ' +
    'conditionalFormatting element whose sqref carries every area, with the rule preserved rather ' +
    'than silently discarded.',

  /** @type {Behavior[]} */
  behavior: [
    {
      name: 'a multi-area ref produces exactly one conditionalFormatting element with a rule',
      baseline: 'pass',
      async expect(api, assert) {
        const {writeOk, xml} = await api.authorConditionalFormatting(CF);
        assert.strictEqual(writeOk, true, 'writing the multi-area rule does not throw');
        assert.strictEqual(
          xml.blockCount,
          1,
          'exactly one conditionalFormatting element is emitted',
        );
        assert.ok(xml.ruleCount >= 1, 'the rule is present, not discarded');
      },
    },
    {
      name: 'the emitted sqref carries all supplied areas as one space-separated multi-area reference',
      baseline: 'pass',
      async expect(api, assert) {
        const {xml} = await api.authorConditionalFormatting(CF);
        assert.deepStrictEqual(
          xml.sqrefs,
          ['A1:C1 A3:C3 A5:C5'],
          `the multi-area sqref must survive; got ${JSON.stringify(xml.sqrefs)}`,
        );
      },
    },
    {
      name: 'a valid multi-area rule never yields a sheet with zero conditionalFormatting elements',
      baseline: 'pass',
      async expect(api, assert) {
        const {xml} = await api.authorConditionalFormatting(CF);
        assert.ok(xml.blockCount > 0, 'the rule is not dropped to zero blocks');
      },
    },
  ],
};
