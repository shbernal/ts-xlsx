// Cluster: styles
//
// Real-world scenario: an author wants a cell's fill to visually represent its numeric value as a
// proportional bar — a data-bar conditional-formatting rule over a range, carrying a bar color, a
// gradient flag, and two conditional-format value objects (cfvo) defining the low and high anchors
// of the bar scale (e.g. num 0 to num 1). The rule must write valid worksheet XML (a dataBar element
// with its cfvo anchors and color) that opens without a repair prompt, and it must round-trip: after
// write→read the rule is present on the same range with its color, gradient flag, and both cfvo
// anchors intact. A malformed shape — a truncated/empty cfvo or dataBar element — corrupts the sheet.

/** @typedef {{ name: string, baseline: 'pass'|'fail', expect: (api: any, assert: any) => Promise<void>|void }} Behavior */

const CF = {
  ref: 'A1:A3',
  rules: [
    {
      type: 'dataBar',
      gradient: true,
      color: {argb: 'FF638EC6'},
      cfvo: [{type: 'num', value: 0}, {type: 'num', value: 1}],
    },
  ],
};

export default {
  id: 'databar-conditional-formatting-roundtrip',
  provenance: {source: 'upstream-issue'},
  cluster: 'styles',
  description:
    'A data-bar conditional-formatting rule with two cfvo anchors (min and max) and a bar color ' +
    'emits well-formed worksheet XML and round-trips: after write→read the rule is on the same range ' +
    'as a dataBar with its color, gradient flag, and both cfvo anchors preserved.',

  /** @type {Behavior[]} */
  behavior: [
    {
      name: 'writing a two-cfvo dataBar rule emits well-formed XML with a dataBar element',
      baseline: 'pass',
      async expect(api, assert) {
        const {writeOk, xml} = await api.authorConditionalFormatting(CF);
        assert.strictEqual(writeOk, true, 'writing the dataBar rule does not throw');
        assert.strictEqual(xml.hasDataBar, true, 'a dataBar element is emitted');
        assert.strictEqual(xml.cfvoCount, 2, 'both cfvo anchors are present (not truncated to one)');
        assert.strictEqual(xml.hasColor, true, 'the bar color is emitted');
        assert.strictEqual(xml.wellFormed, true, 'the conditionalFormatting block is well-formed XML');
      },
    },
    {
      name: 'the rule reads back as a dataBar on the same range with its bar color',
      baseline: 'pass',
      async expect(api, assert) {
        const {reload} = await api.authorConditionalFormatting(CF);
        assert.ok(reload, 'a conditional-formatting rule reads back');
        assert.strictEqual(reload.type, 'dataBar', 'the rule type is dataBar');
        assert.strictEqual(reload.color, 'FF638EC6', 'the bar color survives');
      },
    },
    {
      // The gradient flag is written but the reader does not surface it on the round-tripped rule,
      // so it comes back unset — a fidelity gap distinct from the color/anchor preservation above.
      name: 'the gradient flag survives the round-trip',
      baseline: 'fail',
      async expect(api, assert) {
        const {reload} = await api.authorConditionalFormatting(CF);
        assert.strictEqual(reload.gradient, true, 'the gradient flag reads back as set');
      },
    },
    {
      name: 'both cfvo anchors survive in order',
      baseline: 'pass',
      async expect(api, assert) {
        const {reload} = await api.authorConditionalFormatting(CF);
        assert.strictEqual(reload.cfvo.length, 2, 'two cfvo anchors read back');
        assert.strictEqual(reload.cfvo[0].value, 0, 'the low anchor is 0');
        assert.strictEqual(reload.cfvo[1].value, 1, 'the high anchor is 1');
      },
    },
  ],
};
