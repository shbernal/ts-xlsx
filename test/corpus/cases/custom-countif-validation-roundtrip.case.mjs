// Cluster: validations
//
// Real-world scenario: a column is guarded against duplicate entries with a *custom* data
// validation whose formula is a COUNTIF expression (e.g. COUNTIF($A$2:$A$20,A3)<2) over a large
// target range, paired with an error title and message Excel shows when the rule is violated.
// (The original complaint — that Excel does not block a *pasted* duplicate — is Excel's own
// runtime paste behavior, which neither the OOXML format nor this library governs; it is out of
// scope.) The durable, checkable guarantee is faithful persistence: the validation must survive a
// load/save round-trip with its type, formula (relative + absolute references), target range, and
// error strings intact, and a worksheet-level xr:uid extension attribute must not cause the
// validation to be dropped or the file to fail to parse.

/** @typedef {{ name: string, baseline: 'pass'|'fail', expect: (api: any, assert: any) => Promise<void>|void }} Behavior */

const FIXTURE = 'custom-countif-validation-roundtrip/source.xlsx';

export default {
  id: 'custom-countif-validation-roundtrip',
  provenance: {source: 'upstream-issue'},
  cluster: 'validations',
  description:
    'A custom-type data validation with a COUNTIF formula, a target range, and error strings ' +
    'is read back with those facts intact and survives a load/save round-trip without loss — ' +
    'a worksheet-level xr:uid extension attribute does not cause it to be dropped.',

  /** @type {Behavior[]} */
  behavior: [
    {
      name: 'the custom COUNTIF validation is read from the model with its formula and error strings',
      baseline: 'pass',
      async expect(api, assert) {
        const {sheets} = await api.readFixtureValidationRules(FIXTURE);
        const rules = Object.values(sheets).flatMap((s) => s.rules.map((r) => r.rule));
        const custom = rules.find((r) => r.type === 'custom');
        assert.ok(custom, `expected a custom-type validation; got ${JSON.stringify(rules)}`);
        assert.ok(
          (custom.formulae || []).some((f) => /COUNTIF/i.test(f)),
          `the custom formula must retain its COUNTIF text; got ${JSON.stringify(custom.formulae)}`,
        );
        assert.strictEqual(
          custom.errorTitle,
          'Duplicate Value',
          'the error title must be read back',
        );
      },
    },
    {
      name: 'the custom validation survives a round-trip with type, formula, and error strings',
      baseline: 'pass',
      async expect(api, assert) {
        const {sheets} = await api.roundtripFixtureValidationXml(FIXTURE);
        const rules = Object.values(sheets).flatMap((s) => s.standardRules);
        const custom = rules.find((r) => r.type === 'custom');
        assert.ok(
          custom,
          `the custom validation must survive re-serialization; got ${JSON.stringify(rules)}`,
        );
        assert.ok(
          /COUNTIF/i.test(custom.formula1 || ''),
          `the COUNTIF formula must survive; got ${custom.formula1}`,
        );
        assert.strictEqual(
          custom.errorTitle,
          'Duplicate Value',
          'the error title must survive the round-trip',
        );
        assert.ok(custom.sqref, 'the validation must keep a target range');
      },
    },
  ],
};
