// Cluster: streaming
//
// Real-world scenario: this library ships two readers over one format, and the streaming one
// documents its decoded value as "identical to what readXlsx would produce for the same cell". That
// is the promise a caller picks the streaming reader on: same answers, less memory. There was no
// check on it, and it was not true.
//
// The corpus already reads one workbook as .xlsx and as .xlsb and requires identical models, and that
// property is what keeps those two codecs honest. This is the same property for the two readers of
// one format, and it is worth more than any single divergence it catches: the one it caught first was
// a cell whose format is inherited from its column rather than stated on the cell, where the buffered
// reader resolved cell then row then column and the streaming reader read the cell's own `s` alone.
// `numFmt` is what tells a date serial from a plain number, so 45000 came back as a Date from one
// reader and as 45000 from the other, silently, with no error on either side.

import type {Assert, Case, CorpusApi} from '../case.ts';

// A sheet whose formats live on the column and on the row rather than on each cell, plus ordinary
// cells of every type, so the comparison covers both the inheritance and the plain decode.
const SPEC = {
  sheets: [
    {
      name: 'Inherited',
      columns: [{index: 1, numFmt: 'yyyy-mm-dd'}],
      cells: [
        {ref: 'A1', value: 45000},
        {ref: 'A2', value: 45001},
        {ref: 'B1', value: 'text'},
        {ref: 'B2', value: 42},
        {ref: 'C1', value: true},
        {ref: 'C2', value: 3.5},
      ],
    },
    {
      name: 'Plain',
      cells: [
        {ref: 'A1', value: 'x'},
        {ref: 'A3', value: 7},
      ],
    },
  ],
};

export default {
  id: 'both-readers-produce-one-model',
  provenance: {source: 'reader-parity-property'},
  cluster: 'streaming',
  description:
    'One workbook read by the buffered reader and by the streaming reader yields the same cells ' +
    'with the same values and the same value types, including where a cell inherits its number ' +
    'format from its column rather than carrying its own.',

  behavior: [
    {
      name: 'both readers see the same sheets',
      expect(api: CorpusApi, assert: Assert) {
        const report = api.streamVsEagerCells(SPEC);
        assert.deepEqual(report.streamingSheets, report.eagerSheets);
      },
    },
    {
      name: 'both readers decode every cell to the same value and the same type',
      expect(api: CorpusApi, assert: Assert) {
        const report = api.streamVsEagerCells(SPEC);
        assert.deepEqual(report.streaming, report.eager);
      },
    },
    {
      name: 'and they still agree when a cell carries no style of its own',
      expect(api: CorpusApi, assert: Assert) {
        // A hand-written package a real producer emits and this writer never does: no `<c s>` at all,
        // so a cell's format can only come from its column. This is where the two readers disagreed.
        const report = api.streamVsEagerCells(SPEC, {stripCellStyles: true});
        assert.deepEqual(report.streaming, report.eager);
      },
    },
    {
      name: 'a cell inheriting a date format from its column decodes as a date in both',
      expect(api: CorpusApi, assert: Assert) {
        const report = api.streamVsEagerCells(SPEC, {stripCellStyles: true});
        assert.equal(report.eager['Inherited!A1']?.type, 'date', 'the buffered reader saw a date');
        assert.equal(
          report.streaming['Inherited!A1']?.type,
          'date',
          'and so did the streaming one',
        );
      },
    },
  ],
} satisfies Case;
