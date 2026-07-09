// Cluster: formulas
//
// Real-world scenario: a user writes a cell formula using a modern Excel function that post-dates
// the original OOXML function grammar — FILTER, XLOOKUP, LET, SEQUENCE — by its plain, readable name
// (e.g. `=FILTER(A1:D1, A2:D2 = N1)`). When the file is opened in current desktop Excel, the formula
// is silently dropped or shown as removed, even though typing the identical formula into Excel works
// fine. The cause is an OOXML storage convention: functions introduced after the frozen legacy
// grammar must be stored in the sheet XML with a `_xlfn.` name-mangling prefix (and LET's defined
// names with a `_xlpm.` prefix). The prefix is an on-disk detail — the reader must strip it back to
// the plain name — but if the writer never applies it, Excel rejects the formula.

/** @typedef {{ name: string, baseline: 'pass'|'fail', expect: (api: any, assert: any) => Promise<void>|void }} Behavior */

export default {
  id: 'modern-function-xlfn-prefix-roundtrip',
  provenance: {source: 'upstream-issue'},
  cluster: 'formulas',
  description:
    'A modern function written by its plain name (FILTER, XLOOKUP, …) is persisted in the sheet XML ' +
    'with the required _xlfn. name-mangling prefix so Excel accepts it, while a formula whose input ' +
    'already carries the prefix is not double-prefixed.',

  /** @type {Behavior[]} */
  behavior: [
    {
      name: 'a modern function written by its plain name is stored with the _xlfn. prefix on disk',
      baseline: 'fail',
      async expect(api, assert) {
        const {sheets} = await api.inspectPackage({
          sheets: [{name: 'S', cells: [{ref: 'A1', formula: 'FILTER(B1:D1,B2:D2=1)', result: 0}]}],
        });
        assert.ok(
          /_xlfn\.FILTER/.test(sheets.S.formulas.A1 || ''),
          `FILTER must be stored as _xlfn.FILTER for Excel to accept it; got: ${sheets.S.formulas.A1}`
        );
      },
    },
    {
      name: 'a formula whose input already carries an explicit _xlfn. prefix is not double-prefixed',
      baseline: 'pass',
      async expect(api, assert) {
        const {sheets} = await api.inspectPackage({
          sheets: [{name: 'S', cells: [{ref: 'A1', formula: '_xlfn.XLOOKUP(1,B:B,C:C)', result: 0}]}],
        });
        const f = sheets.S.formulas.A1 || '';
        assert.ok(/_xlfn\.XLOOKUP/.test(f), 'the explicit prefix survives');
        assert.ok(!/_xlfn\._xlfn/.test(f), `the prefix must not be doubled; got: ${f}`);
      },
    },
  ],
};
