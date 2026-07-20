// Cluster: streaming
//
// Real-world scenario: an .xlsx is a ZIP, and OOXML does not mandate the order of entries within
// it. Tools such as openpyxl commonly emit archives where a worksheet part (xl/worksheets/sheet1
// .xml) appears physically before xl/workbook.xml. Read via the streaming reader, the reader meets
// a worksheet entry before it has seen the workbook metadata; it must defer/buffer so that once
// workbook.xml is processed, every worksheet is parsed and all rows are yielded. The historical
// failure dropped the early worksheet entry and reached end-of-archive before the workbook model
// was built, crashing on undefined sheet metadata. The reader must parse the file completely and
// yield all rows regardless of entry ordering.

import type {Assert, Case, CorpusApi} from '../case.ts';

const FIXTURE = 'streaming-read-worksheet-entry-before-workbook/source.xlsx';

export default {
  id: 'streaming-read-worksheet-entry-before-workbook',
  provenance: {source: 'upstream-issue'},
  cluster: 'streaming',
  description:
    'The streaming reader tolerates a package whose ZIP places a worksheet part before ' +
    'xl/workbook.xml — it yields every worksheet and all rows rather than crashing on an unbuilt ' +
    'workbook model.',

  behavior: [
    {
      name: 'streaming-reading a worksheet-before-workbook archive completes without error',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const {ok, error, sheetNames} = await api.streamReadReport(FIXTURE);
        assert.ok(ok, `the read must not crash on entry ordering; got ${JSON.stringify(error)}`);
        assert.ok(sheetNames.length >= 1, 'at least one worksheet is emitted');
      },
    },
    {
      name: 'all data rows are delivered regardless of ZIP entry order',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const {totalRows} = await api.streamReadReport(FIXTURE);
        assert.ok(totalRows > 0, `every worksheet's rows must be yielded; got ${totalRows} rows`);
      },
    },
  ],
} satisfies Case;
