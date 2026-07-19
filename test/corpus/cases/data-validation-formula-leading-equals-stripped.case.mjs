// Cluster: validations
//
// Real-world scenario: a list-type data validation's source is supplied as a formula string with a
// leading '=' — e.g. "=$AA$1:$AA$2". In OOXML the dataValidation formula1 element carries the formula
// WITHOUT a leading '=' (the '=' is a UI/authoring convention, not part of the stored formula). When
// the writer emits the '=' verbatim into formula1, the application does not apply the validation until
// the file is reopened/repaired. The writer must strip exactly one leading '=' from a validation
// formula.

/** @typedef {{ name: string, baseline: 'pass'|'fail', expect: (api: any, assert: any) => Promise<void>|void }} Behavior */

export default {
  id: 'data-validation-formula-leading-equals-stripped',
  provenance: {source: 'upstream-issue'},
  cluster: 'validations',
  description:
    'A data-validation formula supplied with a leading "=" serializes into formula1 without the "=" ' +
    '(OOXML formula1 carries no equals sign), so the application applies the validation immediately ' +
    'rather than only after a reopen/repair.',

  /** @type {Behavior[]} */
  behavior: [
    {
      name: 'a validation formula supplied with a leading = serializes without it',
      baseline: 'pass',
      async expect(api, assert) {
        const {formula1, hasLeadingEquals} = await api.dvFormulaLeadingEquals('=$AA$1:$AA$2');
        assert.strictEqual(
          hasLeadingEquals,
          false,
          `formula1 must not keep the leading "="; got ${JSON.stringify(formula1)}`,
        );
      },
    },
    {
      name: 'the reference after the = is preserved verbatim',
      baseline: 'pass',
      async expect(api, assert) {
        const {formula1} = await api.dvFormulaLeadingEquals('=$AA$1:$AA$2');
        assert.strictEqual(
          formula1,
          '$AA$1:$AA$2',
          `the range reference survives without the "="; got ${JSON.stringify(formula1)}`,
        );
      },
    },
    {
      name: 'a formula supplied without a leading = is unchanged (control)',
      baseline: 'pass',
      async expect(api, assert) {
        const {formula1} = await api.dvFormulaLeadingEquals('$AA$1:$AA$2');
        assert.strictEqual(formula1, '$AA$1:$AA$2', 'a formula with no "=" is emitted verbatim');
      },
    },
  ],
};
