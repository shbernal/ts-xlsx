// Cluster: types
//
// Real-world scenario: a sheet's header row declares five columns, but a data row leaves the last two
// blank. A consumer building fixed-width positional arrays (row → array aligned to the header) iterates
// each row's cells with the "include empty cells" option so that every column position is represented
// and the arrays line up column-for-column. Interior blanks (a gap between populated cells) are
// surfaced correctly, but the trailing run of empty cells at the end of a row is dropped from
// iteration — even though the sheet's declared column count includes them. The result is an internal
// inconsistency: the row says it is five columns wide, but iteration yields only three cells, so
// positional reconstruction misaligns for any row whose populated cells stop short of the last column.

import type {Assert, Case, CorpusApi} from '../case.ts';

// Header spans A–E; row 2 populates A–C (D,E trailing-empty); row 3 populates A,C (B interior-empty; D,E trailing-empty).
const spec = {
  sheets: [
    {
      name: 'S',
      cells: [
        {ref: 'A1', value: 'h1'},
        {ref: 'B1', value: 'h2'},
        {ref: 'C1', value: 'h3'},
        {ref: 'D1', value: 'h4'},
        {ref: 'E1', value: 'h5'},
        {ref: 'A2', value: 'a'},
        {ref: 'B2', value: 'b'},
        {ref: 'C2', value: 'c'},
        {ref: 'A3', value: 'x'},
        {ref: 'C3', value: 'z'},
      ],
    },
  ],
};

export default {
  id: 'row-iteration-includes-trailing-empty-cells',
  provenance: {source: 'upstream-issue'},
  cluster: 'types',
  description:
    'Iterating a row with the include-empty option must surface every column position up to the ' +
    "sheet's declared width — including a trailing run of empty cells — so positional row " +
    'reconstruction aligns column-for-column with a wider header, consistently with interior blanks.',

  behavior: [
    {
      name: 'an interior empty cell is surfaced by an include-empty iteration',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const {rows} = await api.readRowCellPresence(spec, [3]);
        assert.ok(
          rows[3].cols.includes(2),
          'the interior blank (column 2) is yielded between the populated cells',
        );
      },
    },
    {
      name: "trailing empty cells are surfaced up to the sheet's declared column width",
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const {rows, columnCount} = await api.readRowCellPresence(spec, [2]);
        assert.strictEqual(columnCount, 5, 'the header makes the sheet five columns wide');
        assert.ok(
          rows[2].cols.includes(4) && rows[2].cols.includes(5),
          `the trailing empties (columns 4,5) must be yielded; got columns ${JSON.stringify(rows[2].cols)}`,
        );
      },
    },
    {
      name: 'every data row reconstructs to the header width so rows align column-for-column',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const {rows, columnCount} = await api.readRowCellPresence(spec, [2]);
        assert.strictEqual(
          rows[2].cols.length,
          columnCount,
          `a row must yield one cell per declared column; yielded ${rows[2].cols.length} of ${columnCount}`,
        );
      },
    },
  ],
} satisfies Case;
