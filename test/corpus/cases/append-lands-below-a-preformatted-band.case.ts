import type {Assert, Case, CorpusApi} from '../case.ts';

export default {
  id: 'append-lands-below-a-preformatted-band',
  cluster: 'rows',
  description:
    'A template is laid out by styling the rows and columns that data will later fill. Those ' +
    'lines are used, so an append must land past them: a sheet whose row 2 is styled and empty ' +
    'appends to row 3, and the author never finds two of their rows collapsed into one.',
  provenance: {source: 'api-contract-audit'},
  behavior: [
    {
      name: 'a row carrying only its own formatting counts toward the append point',
      expect(api: CorpusApi, assert: Assert) {
        const report = api.appendOverPreformattedBandReport();
        assert.strictEqual(report.rowCountBeforeAppend, 2);
        assert.strictEqual(report.columnCountBeforeAppend, 2);
      },
    },
    {
      name: 'addRow lands below the styled row and leaves it empty and styled',
      expect(api: CorpusApi, assert: Assert) {
        const report = api.appendOverPreformattedBandReport();
        assert.strictEqual(report.appendedRowValue, 'appended');
        assert.strictEqual(report.styledRowValue, null);
        assert.strictEqual(report.styledRowKeptFill, 'pattern');
      },
    },
    {
      name: 'addColumn lands beyond the styled column and leaves it empty',
      expect(api: CorpusApi, assert: Assert) {
        const report = api.appendOverPreformattedBandReport();
        assert.strictEqual(report.appendedColumnValue, 'appended-col');
        assert.strictEqual(report.styledColumnValue, null);
      },
    },
    {
      name: 'the styled line bounds the used range without counting as a populated row',
      expect(api: CorpusApi, assert: Assert) {
        const report = api.appendOverPreformattedBandReport();
        assert.strictEqual(report.usedRangeAfterRoundtrip, 'A1:C3');
        assert.strictEqual(report.actualRowCount, 2);
      },
    },
  ],
} satisfies Case;
