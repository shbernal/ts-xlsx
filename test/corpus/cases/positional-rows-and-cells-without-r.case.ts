// Cluster: rows
//
// Real-world scenario: `r` is `use="optional"` on both `sml:CT_Row` and `sml:CT_Cell`. A producer may
// rely on document position instead, and several non-Excel writers do: the nth `<row>` of
// `<sheetData>` is row n, and the nth `<c>` of a row is that row's nth column. Excel opens such a
// file without comment and reads every cell.
//
// A reader that requires the attribute does not lose formatting or an edge case, it loses *every
// cell of every row*, silently, and reports a workbook full of blanks. That is a whole class of
// foreign file, and it is invisible to a corpus built only from files this library wrote, because
// this library always writes the attribute.
//
// The fixture mixes the spellings on purpose, and its expected addresses are Excel Desktop's own
// reading of it: a wholly positional row, a row whose cells step over a declared `r` (which
// re-anchors the count, so the cell after `D4` is `E4`), and a row that names itself while leaving
// its cells positional.

import type {Assert, Case, CorpusApi} from '../case.ts';

const FIXTURE = 'positional-rows-and-cells-without-r/no-r.xlsx';
const CELLS = ['A1', 'B1', 'C1', 'A3', 'B3', 'A4', 'D4', 'E4', 'A7', 'B7'];

export default {
  id: 'positional-rows-and-cells-without-r',
  provenance: {source: 'excel-desktop-verification'},
  cluster: 'rows',
  description:
    'A worksheet whose `<row>` and `<c>` elements carry no `r` is read by document position, as the ' +
    'schema allows and Excel does: every cell lands where its position puts it, a declared `r` ' +
    're-anchors the count, and the streaming reader agrees with the buffered one throughout.',

  behavior: [
    {
      name: 'a wholly positional row places its cells across the columns in order',
      async expect(api: CorpusApi, assert: Assert) {
        const cells = await api.readFixtureCells(FIXTURE, CELLS);
        assert.strictEqual(cells.A1.value, 'first');
        assert.strictEqual(cells.B1.value, 2);
        assert.strictEqual(
          cells.C1.type,
          'date',
          'and its style still resolves through its column',
        );
      },
    },
    {
      name: 'a formula cell keeps its formula and cached result at its positional address',
      async expect(api: CorpusApi, assert: Assert) {
        const cells = await api.readFixtureCells(FIXTURE, CELLS);
        assert.strictEqual(cells.A3.value, 'third row');
        assert.deepStrictEqual(cells.B3.value, {formula: 'B1*2', result: 4});
      },
    },
    {
      name: 'a declared `r` re-anchors the count for the cells after it',
      async expect(api: CorpusApi, assert: Assert) {
        // The two spellings are mixable, and this is where a reader that merely counts `<c>` elements
        // parts company with Excel: after `D4`, position resumes at E, not at B.
        const cells = await api.readFixtureCells(FIXTURE, CELLS);
        assert.strictEqual(cells.A4.value, 'positional');
        assert.strictEqual(cells.D4.value, 'declared');
        assert.strictEqual(cells.E4.value, 'after the declared one');
      },
    },
    {
      name: 'a row that names itself still places its positional cells',
      async expect(api: CorpusApi, assert: Assert) {
        const cells = await api.readFixtureCells(FIXTURE, CELLS);
        assert.strictEqual(cells.A7.value, 'row says seven');
        assert.strictEqual(cells.B7.value, 7);
      },
    },
    {
      name: 'the streaming reader numbers the same rows and columns as the buffered one',
      async expect(api: CorpusApi, assert: Assert) {
        // The two readers each held their own copy of the row-number inference, and the copies had
        // drifted; this is the assertion that keeps one shared answer honest.
        const streamed = await api.streamReadFixture(FIXTURE, CELLS);
        const cells = await api.readFixtureCells(FIXTURE, CELLS);
        for (const ref of CELLS) {
          assert.deepStrictEqual(
            streamed[ref].value,
            cells[ref].value,
            `${ref}: a row-by-row read must place a cell where a whole read does`,
          );
        }
      },
    },
  ],
} satisfies Case;
