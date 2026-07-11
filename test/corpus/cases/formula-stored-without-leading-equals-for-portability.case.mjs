// Cluster: formulas
//
// Real-world scenario: a caller sets a cell to a formula. OOXML stores the formula text in <f>
// WITHOUT a leading '=' (e.g. `1+2`, not `=1+2`). If a caller passes the expression with a leading
// '=' and the writer stores that character verbatim, the document is inconsistently tolerated: Excel
// accepts it, but stricter consumers (Google Sheets, WPS) reject a file whose stored formula begins
// with '='. The library must normalize the stored formula so it never carries a leading '=',
// regardless of how the caller supplied it. A formula supplied without the '=' is the working
// control.

/** @typedef {{ name: string, baseline: 'pass'|'fail', expect: (api: any, assert: any) => Promise<void>|void }} Behavior */

// Self-contained constant formulas: each cell holds only an <f>, with no preceding value cells, so
// the stored formula text can be read back per-cell unambiguously.
const SHEET = {
  sheets: [
    {
      name: 'S',
      cells: [
        {ref: 'A1', formula: '=1+2'},
        {ref: 'B1', formula: '1+2'},
      ],
    },
  ],
};

export default {
  id: 'formula-stored-without-leading-equals-for-portability',
  provenance: {source: 'upstream-issue'},
  cluster: 'formulas',
  description:
    'A cell formula is stored in the sheet XML without a leading "=" (OOXML <f> carries no equals ' +
    'sign), even when the caller supplied one — so the document is portable to strict consumers ' +
    '(Google Sheets/WPS), not just Excel; a formula supplied without "=" is unchanged.',

  /** @type {Behavior[]} */
  behavior: [
    {
      name: 'a formula supplied without a leading = is stored verbatim (control)',
      baseline: 'pass',
      async expect(api, assert) {
        const {sheets} = await api.inspectPackage(SHEET);
        assert.strictEqual(sheets.S.formulas.B1, '1+2', 'the plain formula stores unchanged');
      },
    },
    {
      name: 'a formula supplied with a leading = is stored without it',
      baseline: 'fail',
      async expect(api, assert) {
        const {sheets} = await api.inspectPackage(SHEET);
        assert.strictEqual(
          sheets.S.formulas.A1,
          '1+2',
          `the stored <f> text must not begin with "="; got ${JSON.stringify(sheets.S.formulas.A1)} — a ` +
            'leading = makes the file unreadable to strict consumers'
        );
      },
    },
    {
      name: 'the round-tripped formula does not retain a leading =',
      baseline: 'fail',
      async expect(api, assert) {
        const {A1} = await api.roundtripFormulas(SHEET);
        assert.ok(A1.formula && !A1.formula.startsWith('='), `read-back formula must have no leading =; got ${JSON.stringify(A1.formula)}`);
      },
    },
  ],
};
