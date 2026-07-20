// Cluster: styles
//
// Real-world scenario: a user applies an alignment property (e.g. textRotation) to a single column,
// or to individual cells, expecting only that target to be affected. A reported failure was that
// every cell in the sheet ended up with the same rotation, because cells/rows/columns shared one
// alignment/style object by reference — mutating a nested property on one entity mutated the shared
// instance and bled across unrelated cells. Correct behavior is copy-on-write isolation: assigning
// alignment to one column applies to that column's own cells only, and setting it on one cell leaves
// every other cell untouched.

import type {Assert, Case, CorpusApi} from '../case.ts';

// Column 1 gets a text rotation; columns 2/3 do not. A1 and A2 are in column 1; B1 and C1 are not.
const COL_SPEC = {
  sheets: [
    {
      name: 'S',
      columns: [{index: 1, alignment: {textRotation: 45}}],
      cells: [
        {ref: 'A1', value: 'a'},
        {ref: 'A2', value: 'c'},
        {ref: 'B1', value: 'b'},
        {ref: 'C1', value: 'd'},
      ],
    },
  ],
};

// A single cell gets an alignment; its row/column neighbours must not.
const CELL_SPEC = {
  sheets: [
    {
      name: 'S',
      cells: [
        {ref: 'A1', value: 'a', alignment: {textRotation: 90}},
        {ref: 'B1', value: 'b'},
        {ref: 'A2', value: 'c'},
      ],
    },
  ],
};

export default {
  id: 'alignment-does-not-leak-across-cells',
  provenance: {source: 'upstream-issue'},
  cluster: 'styles',
  description:
    'Assigning alignment (e.g. textRotation) to one column applies to that column’s cells only, and ' +
    'setting alignment on a single cell leaves its neighbours untouched — no shared-style leak that ' +
    'bleeds the alignment across the whole sheet.',

  behavior: [
    {
      name: 'a column’s alignment applies to that column’s cells',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const model = await api.roundtripWorkbook(COL_SPEC);
        const cells = model.sheets.S.cells;
        assert.strictEqual(
          cells.A1.alignment?.textRotation,
          45,
          'A1 (column 1) carries the rotation',
        );
        assert.strictEqual(
          cells.A2.alignment?.textRotation,
          45,
          'A2 (column 1) carries the rotation',
        );
      },
    },
    {
      name: 'the column’s alignment does not leak to other columns’ cells',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const model = await api.roundtripWorkbook(COL_SPEC);
        const cells = model.sheets.S.cells;
        assert.ok(
          !cells.B1.alignment || cells.B1.alignment.textRotation === undefined,
          `B1 must not inherit the rotation; got ${JSON.stringify(cells.B1.alignment)}`,
        );
        assert.ok(
          !cells.C1.alignment || cells.C1.alignment.textRotation === undefined,
          `C1 must not inherit the rotation; got ${JSON.stringify(cells.C1.alignment)}`,
        );
      },
    },
    {
      name: 'setting alignment on one cell does not change its neighbours',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const model = await api.roundtripWorkbook(CELL_SPEC);
        const cells = model.sheets.S.cells;
        assert.strictEqual(
          cells.A1.alignment?.textRotation,
          90,
          'the targeted cell carries the rotation',
        );
        assert.ok(
          !cells.B1.alignment || cells.B1.alignment.textRotation === undefined,
          'the row neighbour is unaffected',
        );
        assert.ok(
          !cells.A2.alignment || cells.A2.alignment.textRotation === undefined,
          'the column neighbour is unaffected',
        );
      },
    },
  ],
} satisfies Case;
