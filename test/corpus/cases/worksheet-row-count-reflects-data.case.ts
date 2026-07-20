// Cluster: core-model
//
// Real-world scenario: code reads a sheet and asks it how many rows it has —
// `rowCount` (the index of the last row carrying anything) and `actualRowCount`
// (how many rows actually hold values) — to drive a loop. These must be concrete
// numbers that reflect the data, never `undefined`, and `actualRowCount` must
// exclude the gaps so a sparse sheet is measured honestly.

import type {Assert, Case, CorpusApi} from '../case.ts';

const DENSE = {
  sheets: [
    {
      name: 'S',
      cells: [
        {ref: 'A1', value: 1},
        {ref: 'A2', value: 2},
        {ref: 'A3', value: 3},
      ],
    },
  ],
};
const SPARSE = {
  sheets: [
    {
      name: 'S',
      cells: [
        {ref: 'A1', value: 1},
        {ref: 'A3', value: 3},
      ],
    },
  ],
};

export default {
  id: 'worksheet-row-count-reflects-data',
  provenance: {source: 'upstream-issue', repo: 'exceljs/exceljs', ref: 251},
  cluster: 'core-model',
  description:
    'After a write→read round-trip, a worksheet reports rowCount and actualRowCount ' +
    'as concrete numbers reflecting its data — rowCount spans to the last populated ' +
    'row and actualRowCount counts only rows that hold values.',

  behavior: [
    {
      name: 'rowCount is the last populated row index, not undefined',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const {rowCount} = (await api.roundtripWorkbook(DENSE)).sheets.S;
        assert.strictEqual(typeof rowCount, 'number', 'rowCount should be a number');
        assert.strictEqual(rowCount, 3);
      },
    },
    {
      name: 'actualRowCount counts populated rows, not undefined',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const {actualRowCount} = (await api.roundtripWorkbook(DENSE)).sheets.S;
        assert.strictEqual(typeof actualRowCount, 'number', 'actualRowCount should be a number');
        assert.strictEqual(actualRowCount, 3);
      },
    },
    {
      name: 'a gap between rows is reflected: rowCount spans it, actualRowCount excludes it',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const {rowCount, actualRowCount} = (await api.roundtripWorkbook(SPARSE)).sheets.S;
        assert.strictEqual(rowCount, 3, 'rowCount should span to the last populated row');
        assert.strictEqual(actualRowCount, 2, 'actualRowCount should skip the empty row');
      },
    },
  ],
} satisfies Case;
