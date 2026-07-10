// Cluster: csv
//
// Real-world scenario: a workbook is read from XLSX and the caller wants to emit one worksheet as
// CSV. The CSV writer accepts a sheet-selection option. When the caller supplies a name that matches
// a worksheet, that sheet's rows are written. When no selector is given, the first worksheet is
// written. But when the name matches no worksheet (e.g. the caller passed an intended output filename
// by mistake), the writer silently emits an empty CSV for a non-empty workbook, hiding the error —
// it should error or have a clearly-defined fallback instead of yielding zero rows.

/** @typedef {{ name: string, baseline: 'pass'|'fail', expect: (api: any, assert: any) => Promise<void>|void }} Behavior */

export default {
  id: 'csv-write-sheet-selection-by-name',
  provenance: {source: 'upstream-issue'},
  cluster: 'csv',
  description:
    'CSV write selects a worksheet by name (its rows are emitted) and defaults to the first sheet ' +
    'with no selector; a name matching no worksheet does not silently produce empty output for a ' +
    'non-empty workbook.',

  /** @type {Behavior[]} */
  behavior: [
    {
      name: 'a matching sheet name emits that worksheet’s rows',
      baseline: 'pass',
      async expect(api, assert) {
        const {text} = await api.csvWriteSheetSelection('Second');
        assert.strictEqual(text, 'b,2\nc,3', 'the named sheet is written, not the first');
      },
    },
    {
      name: 'no selector emits the first worksheet',
      baseline: 'pass',
      async expect(api, assert) {
        const {text} = await api.csvWriteSheetSelection(undefined);
        assert.strictEqual(text, 'a,1', 'the default is the first worksheet');
      },
    },
    {
      name: 'a name matching no worksheet does not silently yield empty output',
      baseline: 'fail',
      async expect(api, assert) {
        const {ok, rowCount} = await api.csvWriteSheetSelection('Nope');
        assert.ok(!ok || rowCount > 0, 'a non-matching selector must error or fall back, not emit zero rows for a non-empty workbook');
      },
    },
  ],
};
