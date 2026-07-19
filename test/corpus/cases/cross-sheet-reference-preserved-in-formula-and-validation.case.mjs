// Cluster: formulas
//
// Real-world scenario: a workbook has two worksheets. On the first sheet a cell holds a formula
// that references a range on the second sheet using the OOXML sheet-qualified form (e.g.
// `MixedCase!A10`), and a data-validation list whose allowed values come from a range on that other
// sheet (`Levels!$A$2:$A$9999`). When the file is written and reopened, the cross-sheet reference
// must survive verbatim: the referenced sheet name keeps its exact casing (not lowercased) and the
// `!` sheet/cell separator is preserved. Reporters saw references broken on reopen — the sheet name
// came back lowercased and the formula/validation failed to resolve — because the reference text
// was mangled on write.

/** @typedef {{ name: string, baseline: 'pass'|'fail', expect: (api: any, assert: any) => Promise<void>|void }} Behavior */

const SPEC = {
  sheets: [
    {name: 'SheetA', cells: [{ref: 'B2', formula: 'MixedCase!A10', result: 42}]},
    {name: 'MixedCase', cells: [{ref: 'A10', value: 42}]},
  ],
};

export default {
  id: 'cross-sheet-reference-preserved-in-formula-and-validation',
  provenance: {source: 'upstream-issue'},
  cluster: 'formulas',
  description:
    'A cross-sheet reference in a cell formula and in a data-validation list survives a write→read ' +
    'round-trip with the referenced sheet name kept at its exact casing and the ! separator intact.',

  /** @type {Behavior[]} */
  behavior: [
    {
      name: 'a cross-sheet cell formula round-trips with the sheet name casing and ! separator intact',
      baseline: 'pass',
      async expect(api, assert) {
        const out = await api.roundtripFormulas(SPEC);
        assert.strictEqual(
          out.B2.formula,
          'MixedCase!A10',
          'the exact cross-sheet reference survives',
        );
      },
    },
    {
      name: 'the written formula XML keeps the mixed-case sheet name (not lowercased)',
      baseline: 'pass',
      async expect(api, assert) {
        const {sheets} = await api.inspectPackage(SPEC);
        assert.strictEqual(
          sheets.SheetA.formulas.B2,
          'MixedCase!A10',
          'the serialized formula text preserves the sheet-name casing and ! separator',
        );
      },
    },
    {
      name: 'a data-validation list referencing another sheet keeps that sheet-name casing on round-trip',
      baseline: 'pass',
      async expect(api, assert) {
        const {readBack} = await api.authorListValidations([
          {ref: 'A1', formula: 'Levels!$A$2:$A$9999'},
        ]);
        assert.strictEqual(
          readBack.A1.formulae[0],
          'Levels!$A$2:$A$9999',
          'the cross-sheet validation reference survives with its sheet name casing intact',
        );
      },
    },
    {
      // Not just the sheet name — the COLUMN LETTERS inside the reference must stay uppercase. A
      // reported defect lower-cased "$A$2" to "$a$2", producing a subtly wrong/invalid reference.
      name: 'the column letters of a cross-sheet reference stay uppercase (not lower-cased)',
      baseline: 'pass',
      async expect(api, assert) {
        const {xml} = await api.authorListValidations([{ref: 'A1', formula: 'Levels!$A$2:$A$8'}]);
        assert.ok(
          xml.formula1.some((f) => f === 'Levels!$A$2:$A$8'),
          `the reference must keep uppercase column letters; got ${JSON.stringify(xml.formula1)}`,
        );
      },
    },
  ],
};
