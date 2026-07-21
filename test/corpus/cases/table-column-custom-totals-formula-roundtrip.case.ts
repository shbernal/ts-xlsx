// Cluster: tables
//
// Real-world scenario: a table's totals row carries a column whose aggregate is not a built-in
// (sum/count/…) but an arbitrary formula the user typed — Excel records it as totalsRowFunction="custom"
// with a <totalsRowFormula> child holding the formula text (no leading "="). A library that models only
// the built-in functions drops that child on read and re-emits a bare totalsRowFunction="custom" with no
// formula: the file still validates (the child is optional in the schema) but the user's custom total is
// silently lost, and Excel shows a blank/zero in that cell on reopen.
//
// So the library round-trips <totalsRowFormula> verbatim and materializes it into the totals cell as the
// cell's own formula — the same treatment a built-in aggregate gets, minus the SUBTOTAL synthesis. The
// formula caches no result (the library is not a calc engine; Excel computes it on open). A custom column
// with no stored formula has nothing to write and stays blank, exactly as it did before.

import type {Assert, Case, CorpusApi} from '../case.ts';

// Header row 1, data rows 2–3, totals row 4. Column A labelled "Total"; column B is a custom total whose
// formula grosses the SUM up by 10% — a formula Excel cannot express as any built-in SUBTOTAL code.
const CUSTOM_TOTAL = {
  sheets: [
    {
      name: 'S',
      tables: [
        {
          name: 'T',
          ref: 'A1',
          totalsRow: true,
          columnDefs: [
            {name: 'Item', totalsRowLabel: 'Total'},
            {name: 'Amount', totalsRowFunction: 'custom', totalsRowFormula: 'SUM(T[Amount])*1.1'},
          ],
          rows: [
            ['a', 1],
            ['b', 2],
          ],
        },
      ],
    },
  ],
};

// A custom column that carries no <totalsRowFormula> (function set to "custom" with nothing behind it)
// has nothing to materialize and must stay blank — the pre-existing behaviour for a formula-less column.
// Totals row is row 3 (header 1, data 2, totals 3), columns A/B.
const CUSTOM_WITHOUT_FORMULA = {
  sheets: [
    {
      name: 'S',
      tables: [
        {
          name: 'T',
          ref: 'A1',
          totalsRow: true,
          columnDefs: [
            {name: 'Item', totalsRowLabel: 'Total'},
            {name: 'Amount', totalsRowFunction: 'custom'},
          ],
          rows: [['a', 1]],
        },
      ],
    },
  ],
};

export default {
  id: 'table-column-custom-totals-formula-roundtrip',
  provenance: {source: 'excel-desktop-verification'},
  cluster: 'tables',
  description:
    'A table column with totalsRowFunction="custom" round-trips its <totalsRowFormula> child verbatim ' +
    "and materializes the formula into the totals cell (uncached), matching Excel's on-open rendering. " +
    'The formula survives a write→read→write cycle at both the table part and the grid cell. A custom ' +
    'column with no stored formula writes no child and leaves its totals cell blank.',

  behavior: [
    {
      name: 'a custom totals column writes its <totalsRowFormula> into the table part',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const {tables} = await api.inspectPackage(CUSTOM_TOTAL);
        assert.strictEqual(tables.length, 1, 'the table part is written');
        assert.deepStrictEqual(
          tables[0].totalsRowFormulas,
          ['SUM(T[Amount])*1.1'],
          `the custom formula should be emitted as a child, got ${JSON.stringify(tables[0].totalsRowFormulas)}`,
        );
        assert.strictEqual(tables[0].xmlWellFormed, true, 'the table XML is well-formed');
      },
    },
    {
      name: 'a custom totals column materializes its formula into the totals-row cell',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const {formulas, cellText} = (await api.inspectPackage(CUSTOM_TOTAL)).sheets.S;
        assert.strictEqual(
          formulas.B4,
          'SUM(T[Amount])*1.1',
          `B4 should hold the custom totals formula, got ${formulas.B4}`,
        );
        // No cached <v>: a cached formula result would surface in cellText. Excel recomputes on open.
        assert.strictEqual(
          cellText.B4,
          undefined,
          `B4 should carry a formula with no cached value, got cached ${cellText.B4}`,
        );
      },
    },
    {
      name: 'the custom totals formula survives a write→read→write round-trip at the part and the cell',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const facts = await api.roundtripInspectPackage(CUSTOM_TOTAL);
        // Part-level: proves the reader recovered the column's formula rather than the cell alone
        // surviving because it was materialized on the first write.
        assert.deepStrictEqual(
          facts.tables[0].totalsRowFormulas,
          ['SUM(T[Amount])*1.1'],
          `the <totalsRowFormula> should survive the round-trip, got ${JSON.stringify(facts.tables[0].totalsRowFormulas)}`,
        );
        assert.strictEqual(
          facts.sheets.S.formulas.B4,
          'SUM(T[Amount])*1.1',
          `the materialized cell formula should survive the round-trip, got ${facts.sheets.S.formulas.B4}`,
        );
      },
    },
    {
      name: 'a custom column with no stored formula writes no child and leaves its totals cell blank',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const {tables, sheets} = await api.inspectPackage(CUSTOM_WITHOUT_FORMULA);
        assert.deepStrictEqual(
          tables[0].totalsRowFormulas,
          [],
          `no <totalsRowFormula> should be emitted, got ${JSON.stringify(tables[0].totalsRowFormulas)}`,
        );
        assert.strictEqual(
          sheets.S.cellText.B3,
          undefined,
          `B3 (custom, no formula) should be blank, got ${sheets.S.cellText.B3}`,
        );
        assert.strictEqual(
          sheets.S.formulas.B3,
          undefined,
          `B3 (custom, no formula) should carry no formula, got ${sheets.S.formulas.B3}`,
        );
      },
    },
  ],
} satisfies Case;
