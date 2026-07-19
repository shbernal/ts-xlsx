// Cluster: types
//
// Real-world scenario: a workbook has data validations whose formula operands are not literal numbers
// but expressions Excel evaluates at validation time — a cell reference like L26, or a defined name
// like a named range backing a list. In the sheet XML these are <formula1>L26</formula1> and
// <formula1>myNames</formula1>. Numeric-typed validations (whole, decimal, date) run the operand
// through numeric parsing to convert literals; for a non-numeric expression that parse fails and the
// reader stores null/NaN instead of the original reference — losing it. The reader must preserve the
// original formula/reference text for any validation whose operand is not a pure numeric literal.
//
// The fixture declares a whole-type validation whose formula1 is a cell reference (L26), a list-type
// validation whose formula1 is a name (myNames), and a whole-type validation with a literal (10).

/** @typedef {{ name: string, baseline: 'pass'|'fail', expect: (api: any, assert: any) => Promise<void>|void }} Behavior */

const FIXTURE = 'datavalidation-reference-formula/sample.xlsx';

export default {
  id: 'datavalidation-reference-formula-not-lost-on-read',
  provenance: {source: 'upstream-issue'},
  cluster: 'types',
  description:
    'A numeric-typed data validation whose formula operand is a cell reference (not a numeric ' +
    'literal) reads back preserving the reference string rather than coercing it to null/NaN; a ' +
    'list-type name and a numeric literal both survive.',

  /** @type {Behavior[]} */
  behavior: [
    {
      name: 'a whole-type validation with a cell-reference operand keeps the reference (not null/NaN)',
      baseline: 'pass',
      async expect(api, assert) {
        const {cells} = await api.readFixtureValidations(FIXTURE);
        const dv = cells['Sheet1!A1'];
        assert.ok(dv, 'the A1 validation is read');
        assert.strictEqual(
          dv.formulae?.[0],
          'L26',
          `the reference must survive; got ${JSON.stringify(dv.formulae)}`,
        );
      },
    },
    {
      name: 'a list-type validation with a defined-name operand keeps the name (control)',
      baseline: 'pass',
      async expect(api, assert) {
        const {cells} = await api.readFixtureValidations(FIXTURE);
        const dv = cells['Sheet1!B1'];
        assert.strictEqual(
          dv.formulae?.[0],
          'myNames',
          `the list name must survive; got ${JSON.stringify(dv?.formulae)}`,
        );
      },
    },
  ],
};
