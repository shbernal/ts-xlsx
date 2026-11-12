// Cluster: streaming
//
// Real-world scenario: a service exports a report too large to hold in memory, and streams it to an
// upload as it reads the rows from its database. Committing each row was documented as bounding
// memory, but a committed row only gave up its cell objects: its serialised XML was kept until the
// workbook committed, and only then did any byte of the sheet reach the upload. Memory grew with the
// report, and the upload saw nothing until the very end.
//
// A sheet is now written into the archive while its rows are committed. That fixes the order the
// rows go out in, so the writer refuses what it can no longer honour (a cell in a row above the last
// one written, a column width set after the first row) rather than dropping it from the file.

import type {Assert, Case, CorpusApi} from '../case.ts';

export default {
  id: 'a-streamed-sheet-is-written-as-its-rows-commit',
  provenance: {source: 'field-report', repo: 'shbernal/ts-xlsx', ref: 33},
  cluster: 'streaming',
  description:
    'Rows committed to the streaming writer reach the caller’s sink before the workbook commits, ' +
    'in row order, and anything that would have to be written before rows already in the archive ' +
    'is refused legibly rather than silently lost.',

  behavior: [
    {
      name: 'committed rows reach the sink before the workbook commits, and the package reads back whole',
      async expect(api: CorpusApi, assert: Assert) {
        const report = await api.streamInOrderReport({rows: 2000});
        assert.ok(
          report.bytesBeforeWorkbookCommit > 0,
          'the sheet went to the sink while its rows were committed',
        );
        assert.equal(report.reloadedRows, 2000);
        assert.equal(report.lastValue, 2000);
      },
    },
    {
      name: 'an uncommitted row above a committed one is written with it, in row order',
      async expect(api: CorpusApi, assert: Assert) {
        const {rowsAboveACommit} = await api.streamInOrderReport({rows: 1});
        assert.deepEqual(rowsAboveACommit, [
          {number: 1, value: 'above'},
          {number: 2, value: 'committed'},
        ]);
      },
    },
    {
      name: 'a cell in a row above the last one written is refused with a reason',
      async expect(api: CorpusApi, assert: Assert) {
        const {behindRefusal} = await api.streamInOrderReport({rows: 1});
        assert.match(behindRefusal ?? '', /already written/);
      },
    },
    {
      name: 'a column width set after the first row is written is refused, not dropped',
      async expect(api: CorpusApi, assert: Assert) {
        const {lateColumnRefusal} = await api.streamInOrderReport({rows: 1});
        assert.match(lateColumnRefusal ?? '', /first row/);
      },
    },
  ],
} satisfies Case;
