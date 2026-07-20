// Cluster: styles
//
// Real-world scenario: a user writes numeric data across several columns and applies a distinct
// number format to each — percentage on two columns, currency on another. After saving, every column
// must keep its own format. The reported failure was that all columns collapsed to whichever format
// was assigned last, because cells/columns that started from the same default style shared one style
// object, so mutating numFmt through one handle leaked to its siblings. The user-facing expectation
// is that assigning a number format to one column affects only that column, and different columns
// hold different formats independently after a round-trip.

import type {Assert, Case, CorpusApi} from '../case.ts';

const PCT = '0.00%';
const CUR = '"$"#,##0.00';
const SPEC = {
  sheets: [
    {
      name: 'S',
      columns: [
        {index: 1, numFmt: PCT},
        {index: 2, numFmt: PCT},
        {index: 3, numFmt: CUR},
      ],
      cells: [
        {ref: 'A1', value: 0.1},
        {ref: 'B1', value: 0.2},
        {ref: 'C1', value: 3},
      ],
    },
  ],
};

export default {
  id: 'per-column-numfmt-stays-independent',
  provenance: {source: 'upstream-issue'},
  cluster: 'styles',
  description:
    'Assigning distinct number formats to different columns keeps each format independent after a ' +
    'round-trip — setting a format on one column does not leak to the others through a shared style ' +
    'object.',

  behavior: [
    {
      name: 'each column keeps its own number format after a round-trip',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const model = await api.roundtripWorkbook(SPEC);
        const cols = model.sheets.S.columns;
        assert.strictEqual(cols[1].numFmt, PCT, 'column 1 keeps its percentage format');
        assert.strictEqual(cols[2].numFmt, PCT, 'column 2 keeps its percentage format');
        assert.strictEqual(
          cols[3].numFmt,
          CUR,
          'column 3 keeps its currency format, not the last-assigned one',
        );
      },
    },
    {
      name: 'a cell under a currency column does not inherit the percentage columns’ format',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const model = await api.roundtripWorkbook(SPEC);
        assert.strictEqual(
          model.sheets.S.cells.C1.numFmt,
          CUR,
          'the currency cell keeps currency, not percentage',
        );
      },
    },
    {
      name: 'cells under the percentage columns keep the percentage format',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const model = await api.roundtripWorkbook(SPEC);
        assert.strictEqual(model.sheets.S.cells.A1.numFmt, PCT, 'A1 keeps percentage');
        assert.strictEqual(model.sheets.S.cells.B1.numFmt, PCT, 'B1 keeps percentage');
      },
    },
  ],
} satisfies Case;
