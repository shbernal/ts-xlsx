// Cluster: formulas
//
// Real-world scenario: a cell holds a formula with comparison operators and string
// concatenation, e.g. IF(ROUND(A1,1)<ROUND(B1,1),"x",""). The '<', '>', and '&'
// characters are XML-special and must be escaped inside the worksheet's <f> element,
// or the produced package is malformed and Excel refuses it. Reading the workbook
// back must yield the identical formula string.

/** @typedef {{ name: string, baseline: 'pass'|'fail', expect: (api: any, assert: any) => void }} Behavior */

const LT = 'IF(ROUND(A1,1)<ROUND(B1,1),"x","")';
const GT = 'IF(A1>B1,"a"&"b","")';

export default {
  id: 'formula-xml-special-chars-escaping',
  provenance: {source: 'upstream-issue', repo: 'exceljs/exceljs', ref: 1355},
  cluster: 'formulas',
  description:
    'Formulas containing XML-special characters (<, >, &) must be escaped in the ' +
    '<f> element so the package is well-formed, and must round-trip verbatim.',

  /** @type {Behavior[]} */
  behavior: [
    {
      name: 'a formula with a "<" operator round-trips verbatim',
      baseline: 'pass',
      async expect(api, assert) {
        const {sheets} = await api.roundtripWorkbook({sheets: [{name: 'S', cells: [{ref: 'C1', formula: LT}]}]});
        assert.strictEqual(sheets.S.cells.C1.formula, LT);
      },
    },
    {
      name: 'a formula with ">" and "&" round-trips verbatim',
      baseline: 'pass',
      async expect(api, assert) {
        const {sheets} = await api.roundtripWorkbook({sheets: [{name: 'S', cells: [{ref: 'C1', formula: GT}]}]});
        assert.strictEqual(sheets.S.cells.C1.formula, GT);
      },
    },
    {
      name: 'special characters produce well-formed worksheet XML',
      baseline: 'pass',
      async expect(api, assert) {
        const {sheets} = await api.inspectPackage({sheets: [{name: 'S', cells: [{ref: 'C1', formula: LT}]}]});
        assert.ok(sheets.S.xmlWellFormed, 'worksheet XML not well-formed (unescaped special char)');
        assert.ok('C1' in sheets.S.formulas, 'expected a <f> element for C1');
      },
    },
  ],
};
