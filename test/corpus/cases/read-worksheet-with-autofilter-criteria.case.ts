// Cluster: tables
//
// Real-world scenario: a spreadsheet application saves a worksheet whose column auto-filter has
// actual *criteria* applied — not just the outer filter range, but concrete selections. In OOXML
// these live inside the `<autoFilter>` element as `<filterColumn>` children holding either a
// `<filters>` list of `<filter val="…"/>` values (a value selection) or a `<customFilters>` block
// of `<customFilter operator="…" val="…"/>` comparisons (a "greater than N" style rule). A reader
// that understood only the outer range historically threw on these children — "Unexpected xml node
// in parseOpen: filter" / "parseClose: customFilters" — making the entire workbook unreadable.
//
// The filter criteria are metadata about which rows are hidden; a robust reader must at minimum
// tolerate them and still surface every cell, through both the buffered and the streaming path.
// The fixture is a minimal foreign-shaped package carrying both a value filter and a custom
// comparison filter, exactly the shapes a real save produces.

import type {Assert, Case, CorpusApi} from '../case.ts';

const FIXTURE = 'read-worksheet-with-autofilter-criteria/sample.xlsx';

export default {
  id: 'read-worksheet-with-autofilter-criteria',
  provenance: {source: 'upstream-issue'},
  cluster: 'tables',
  description:
    'Reading a worksheet whose autoFilter carries filter criteria — a value list (filters/filter) ' +
    'and a custom comparison (customFilters/customFilter) — must not throw, and every cell must ' +
    'remain accessible, via both the buffered and the streaming reader.',

  behavior: [
    {
      name: 'a buffered read of an autoFilter with value and custom criteria does not throw',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const {ok, error, sheetNames} = await api.readFixtureReport(FIXTURE);
        assert.ok(ok, `the read must succeed, got error: ${error}`);
        assert.deepStrictEqual(sheetNames, ['Data'], 'the sheet is surfaced by its real name');
      },
    },
    {
      name: 'the filtered column and its neighbours load all cell values (rows not dropped)',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const cells = await api.readFixtureCells(FIXTURE, ['A2', 'A4', 'B3']);
        assert.strictEqual(cells.A2.value, 'apple', 'the first filtered value is present');
        assert.strictEqual(cells.A4.value, 'plum', 'a value excluded by the filter still loads');
        assert.strictEqual(
          cells.B3.value,
          12,
          'the custom-filtered numeric column loads its value',
        );
      },
    },
    {
      name: 'the streaming reader also tolerates the criteria and delivers the cells',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const cells = await api.streamReadFixture(FIXTURE, ['A2', 'A4']);
        assert.ok(cells.A2 && cells.A4, 'the streaming path yields the requested cells');
        assert.strictEqual(
          cells.A2.value,
          'apple',
          'streaming surfaces the same value the buffered read did',
        );
        assert.strictEqual(
          cells.A4.value,
          'plum',
          'streaming does not drop rows past the filtered ones',
        );
      },
    },
  ],
} satisfies Case;
