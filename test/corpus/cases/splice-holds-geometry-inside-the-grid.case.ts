// Cluster: tables
//
// Real-world scenario: a sheet carries the regions a template normally carries. A dropdown down a
// whole column, a data-bar rule down another, a filter over the data, and a merged banner at the
// bottom. A whole-column region is stored as a bounded range to the very last row (1048576), because
// that is the spelling Excel itself writes, so those regions already sit on the grid's edge. The user
// then inserts rows near the top. Every edge below the insert wants to move down, and the bottom ones
// have nowhere to go: the grid has no row 1048577. A file that names one is not a file with a small
// mistake in it. Excel opens it with "we found a problem with some content in ...", the repair prompt,
// where the same workbook with its edges left inside the grid opens clean, so the whole sheet's
// content is put at the user's mercy by an edit that touched none of it. The same arithmetic anchors
// merges, tables, images and the autofilter, and the column axis has the same wall at XFD.

import type {Assert, Case, CorpusApi} from '../case.ts';

const LAST_ROW = 1_048_576;

/** Every row/column coordinate a reference names, as numbers, so a case can bound-check them. */
function coordinates(refs: readonly string[]): {rows: number[]; columns: string[]} {
  const rows: number[] = [];
  const columns: string[] = [];
  for (const ref of refs) {
    for (const part of ref.split(/[\s:]+/)) {
      const match = /^\$?([A-Z]{1,3})?\$?(\d+)?$/.exec(part);
      if (match?.[1] !== undefined) columns.push(match[1]);
      if (match?.[2] !== undefined) rows.push(Number(match[2]));
    }
  }
  return {rows, columns};
}

export default {
  id: 'splice-holds-geometry-inside-the-grid',
  provenance: {source: 'excel-desktop-verification'},
  cluster: 'tables',
  description:
    'A row or column insert never pushes an anchored region past the edge of the grid. A ' +
    'validation, conditional format, autofilter or merge whose far edge already sits on the last ' +
    'row (or the last column) keeps that edge there, rather than being re-anchored to a line the ' +
    'format has no room for, which is a package Excel meets with its repair prompt.',

  behavior: [
    {
      name: 'no region an insert re-anchors names a row past the last one',
      async expect(api: CorpusApi, assert: Assert) {
        const {rows} = await api.spliceHoldsGeometryInsideTheGrid();
        const refs = [
          ...rows.validationRefs,
          ...rows.formattingRefs,
          ...(rows.autoFilterRef === null ? [] : [rows.autoFilterRef]),
          ...rows.merges,
        ];
        assert.ok(refs.length >= 4, `expected every region back; got ${JSON.stringify(rows)}`);
        for (const row of coordinates(refs).rows) {
          assert.ok(row <= LAST_ROW, `${JSON.stringify(refs)} names row ${row}, past the grid`);
        }
      },
    },
    {
      name: 'a whole-column region keeps exactly the bottom edge Excel gives it',
      async expect(api: CorpusApi, assert: Assert) {
        const {rows} = await api.spliceHoldsGeometryInsideTheGrid();
        assert.deepStrictEqual(
          rows.validationRefs,
          [`B1:B${LAST_ROW}`],
          'inserting rows above a whole-column dropdown leaves it covering the whole column',
        );
        assert.strictEqual(
          rows.autoFilterRef,
          `A1:A${LAST_ROW}`,
          'a filter headed above the insert keeps both edges: the top has not moved, the bottom cannot',
        );
      },
    },
    {
      name: 'the column axis has the same wall, at the last column',
      async expect(api: CorpusApi, assert: Assert) {
        const {columns} = await api.spliceHoldsGeometryInsideTheGrid();
        const refs = [
          ...columns.validationRefs,
          ...(columns.autoFilterRef === null ? [] : [columns.autoFilterRef]),
        ];
        for (const column of coordinates(refs).columns) {
          assert.ok(
            column.length < 3 || (column.length === 3 && column <= 'XFD'),
            `${JSON.stringify(refs)} names column ${column}, past the grid`,
          );
        }
        assert.deepStrictEqual(columns.validationRefs, ['A1:XFD1']);
      },
    },
  ],
} satisfies Case;
