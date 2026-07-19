// Cluster: formulas
//
// Real-world scenario: an Excel workbook built with the What-If Analysis "Data Table" feature carries
// cells whose formula is the OOXML data-table kind — `<f t="dataTable" ref="…" r1="…"/>` with row/
// column input-cell references — rather than a plain formula string. A user opens such a workbook,
// changes an unrelated cell, and saves. After the round trip the data-table formula is gone: the cell
// no longer recalculates against its inputs, and downstream tools that referenced it show #REF!. A
// document reader/writer must preserve exotic formula kinds it does not itself evaluate — reading one
// must surface it, and writing must re-emit it verbatim, even when the edit touches a different cell.

/** @typedef {{ name: string, baseline: 'pass'|'fail', expect: (api: any, assert: any) => Promise<void>|void }} Behavior */

export default {
  id: 'data-table-formula-survives-roundtrip',
  provenance: {source: 'upstream-issue'},
  cluster: 'formulas',
  description:
    'A What-If-Analysis data-table formula (`<f t="dataTable">` with input-cell references) is ' +
    'recognized on read (its kind, range, and cached result are surfaced) and re-emitted on write, ' +
    'so a read-modify-write cycle does not silently drop the data-table kind.',

  /** @type {Behavior[]} */
  behavior: [
    {
      name: 'reading a data-table formula surfaces its kind, range, and cached result',
      baseline: 'pass',
      async expect(api, assert) {
        const {reloadOk, readShareType, readRef, readResult} =
          await api.dataTableFormulaRoundtrip();
        assert.strictEqual(reloadOk, true, 'the workbook with a data-table formula loads');
        assert.strictEqual(
          readShareType,
          'dataTable',
          'the formula kind is recognized as a data table',
        );
        assert.strictEqual(readRef, 'B2:B5', 'the data-table range is surfaced');
        assert.strictEqual(readResult, 99, 'the cached result is surfaced');
      },
    },
    {
      name: 'writing the workbook back preserves the data-table formula kind',
      baseline: 'pass',
      async expect(api, assert) {
        const {outHasDataTable} = await api.dataTableFormulaRoundtrip();
        assert.strictEqual(
          outHasDataTable,
          true,
          'the re-written worksheet must still declare the data-table formula (t="dataTable"), not drop it',
        );
      },
    },
  ],
};
