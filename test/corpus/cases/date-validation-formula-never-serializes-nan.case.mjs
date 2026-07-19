// Cluster: data-validation
//
// Real-world scenario: a user adds a date-type data validation to a range, expecting the validation
// bound to be written as a real date. When the operand is a genuine Date, the worksheet XML must
// carry a valid date serial in the formula. When the operand is not a coercible date, the writer must
// never emit the literal token "NaN" into the formula element — Excel then treats the bound as broken
// and the validation silently fails. A real Date works today; a non-coercible operand emits NaN.

/** @typedef {{ name: string, baseline: 'pass'|'fail', expect: (api: any, assert: any) => Promise<void>|void }} Behavior */

export default {
  id: 'date-validation-formula-never-serializes-nan',
  provenance: {source: 'upstream-issue'},
  cluster: 'data-validation',
  description:
    'A date-type data validation writes a real date serial for a genuine Date operand, and never ' +
    'emits the literal "NaN" into the validation formula for a non-coercible operand (which would ' +
    'silently break the bound in Excel).',

  /** @type {Behavior[]} */
  behavior: [
    {
      name: 'a genuine Date operand writes a valid date serial, not NaN',
      baseline: 'pass',
      async expect(api, assert) {
        const {formula1, hasNaN} = await api.authorDateValidation('2020-01-01T00:00:00.000Z');
        assert.strictEqual(hasNaN, false, 'no NaN token is emitted for a real date');
        assert.ok(formula1 && /^\d+(\.\d+)?$/.test(formula1), `the bound is a numeric serial; got ${JSON.stringify(formula1)}`);
      },
    },
    {
      name: 'a non-coercible operand never emits the literal NaN into the formula',
      baseline: 'pass',
      async expect(api, assert) {
        const {formula1, hasNaN} = await api.authorDateValidation('invalid');
        assert.strictEqual(hasNaN, false, `a non-coercible operand must not serialize "NaN"; got ${JSON.stringify(formula1)}`);
      },
    },
  ],
};
