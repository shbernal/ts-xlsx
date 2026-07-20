// Cluster: tables
//
// Real-world scenario: a worksheet has styled rows (fonts, fills, number formats). When a row above
// them is deleted (or inserted) via a row-splice, every row below the splice point shifts up (or
// down). The style a cell carried must travel with it to its new row index — a shifted cell keeps
// its font, fill, and numFmt rather than coming back blank. A reported defect was that shifted rows
// lost all styling from the splice point onward; that must never regress. (Merge preservation across
// the same shift is locked separately by `splice-rows-preserves-merged-cells`; this case is about
// per-cell style travelling with the row.)

import type {Assert, Case, CorpusApi} from '../case.ts';

// A styled cell two rows below the top; distinct font + fill + numFmt so a lost style is unambiguous.
const styledCell = (ref: CorpusApi) => ({
  ref,
  value: 'styled',
  font: {bold: true, color: {argb: 'FFFF0000'}},
  fill: {type: 'pattern', pattern: 'solid', fgColor: {argb: 'FF00FF00'}},
  numFmt: '0.00',
});

export default {
  id: 'splice-rows-carries-styles-on-shifted-rows',
  provenance: {source: 'upstream-issue'},
  cluster: 'tables',
  description:
    'A row-splice that shifts a styled row carries the cell style to its new position: deleting a ' +
    'row above a styled cell shifts it up and it keeps its font, fill, and numFmt; inserting a row ' +
    'above shifts it down and the style is likewise preserved rather than blanked.',

  behavior: [
    {
      name: 'deleting a row above a styled cell shifts it up and preserves its value',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const {styles} = await api.mutateWorksheet({
          cells: [{ref: 'A1', value: 'top'}, styledCell('A3')],
          ops: [{op: 'spliceRows', start: 1, count: 1}],
          readStyles: ['A2'],
        });
        assert.strictEqual(
          styles.A2.value,
          'styled',
          'the styled cell shifts from A3 up to A2 with its value',
        );
      },
    },
    {
      name: 'the shifted cell keeps its font, fill, and number format',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const {styles} = await api.mutateWorksheet({
          cells: [{ref: 'A1', value: 'top'}, styledCell('A3')],
          ops: [{op: 'spliceRows', start: 1, count: 1}],
          readStyles: ['A2'],
        });
        assert.strictEqual(styles.A2.font?.bold, true, 'the font travels with the shifted cell');
        assert.strictEqual(
          styles.A2.fill?.fgColor?.argb,
          'FF00FF00',
          'the fill travels with the shifted cell',
        );
        assert.strictEqual(
          styles.A2.numFmt,
          '0.00',
          'the number format travels with the shifted cell',
        );
      },
    },
    {
      name: 'inserting a row above a styled cell shifts it down and preserves its style',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const {styles} = await api.mutateWorksheet({
          cells: [{ref: 'A1', value: 'top'}, styledCell('A3')],
          ops: [{op: 'spliceRows', start: 1, count: 0, inserts: [['inserted']]}],
          readStyles: ['A4'],
        });
        assert.strictEqual(
          styles.A4.value,
          'styled',
          'the styled cell shifts from A3 down to A4 with its value',
        );
        assert.strictEqual(
          styles.A4.font?.bold,
          true,
          'the font is preserved on the shifted-down cell',
        );
        assert.strictEqual(
          styles.A4.numFmt,
          '0.00',
          'the number format is preserved on the shifted-down cell',
        );
      },
    },
  ],
} satisfies Case;
