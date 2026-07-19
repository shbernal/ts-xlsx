// Cluster: conditional-formatting
//
// Real-world scenario: a conditional-formatting rule sets `stopIfTrue` — when the rule matches a
// cell, evaluation of all lower-priority rules for that cell halts. This is a first-class OOXML
// cfRule attribute and is essential for layering rules (a match on an early rule must be able to
// suppress later ones). A user who sets stopIfTrue on a rule needs it emitted in the sheet XML and
// preserved on read. The observed defect: the flag is silently dropped on write, so the layered rules
// all evaluate and the intended precedence is lost.

/** @typedef {{ name: string, baseline: 'pass'|'fail', expect: (api: any, assert: any) => Promise<void>|void }} Behavior */

export default {
  id: 'conditional-formatting-stop-if-true-survives',
  provenance: {source: 'upstream-issue'},
  cluster: 'conditional-formatting',
  description:
    'A conditional-formatting rule that sets stopIfTrue serializes the flag on the cfRule and ' +
    'preserves it on read, so rule precedence (a match halting lower-priority rules) is honored.',

  /** @type {Behavior[]} */
  behavior: [
    {
      name: 'a stopIfTrue rule serializes the stopIfTrue attribute on the cfRule',
      baseline: 'pass',
      async expect(api, assert) {
        const {xmlHasStopIfTrue} = await api.conditionalFormattingStopIfTrue();
        assert.strictEqual(xmlHasStopIfTrue, true, 'the written cfRule must carry stopIfTrue="1"');
      },
    },
    {
      name: 'the stopIfTrue flag round-trips onto the reloaded rule',
      baseline: 'pass',
      async expect(api, assert) {
        const {reloadStopIfTrue} = await api.conditionalFormattingStopIfTrue();
        assert.strictEqual(reloadStopIfTrue, true, 'the reloaded rule keeps stopIfTrue');
      },
    },
  ],
};
