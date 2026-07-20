// Cluster: styles
//
// Real-world scenario: a worksheet with grouped (outlined) rows/columns can place its summary rows
// above their detail (rather than the default below) and its summary columns to the left (rather than
// the default right). These are the OOXML sheetPr/outlinePr attributes summaryBelow and summaryRight.
// Setting them via the worksheet's outline properties must serialize into the sheet properties and
// read back, so an author who inverts summary placement gets a file that honors it. This locks the
// runtime write/read path (the public TypeScript type surface exposing outlineProperties is a
// separate, type-level concern).

import type {Assert, Case, CorpusApi} from '../case.ts';

export default {
  id: 'worksheet-outline-summary-position-round-trips',
  provenance: {source: 'upstream-issue'},
  cluster: 'styles',
  description:
    'Worksheet outline summary-position properties (summaryBelow, summaryRight) serialize into the ' +
    'sheet properties as an outlinePr element and read back on a round-trip, so an inverted summary ' +
    'placement is honored rather than dropped.',

  behavior: [
    {
      name: 'summaryBelow/summaryRight = false serialize into an outlinePr element',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const {outlinePrEmitted} = await api.outlinePropertiesRoundtrip();
        assert.strictEqual(
          outlinePrEmitted,
          true,
          'the sheet properties emit <outlinePr summaryBelow="0" summaryRight="0"/>',
        );
      },
    },
    {
      name: 'the summary-position flags read back after a round-trip',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const {reReadSummaryBelow, reReadSummaryRight} = await api.outlinePropertiesRoundtrip();
        assert.strictEqual(reReadSummaryBelow, false, 'summaryBelow round-trips as false');
        assert.strictEqual(reReadSummaryRight, false, 'summaryRight round-trips as false');
      },
    },
  ],
} satisfies Case;
