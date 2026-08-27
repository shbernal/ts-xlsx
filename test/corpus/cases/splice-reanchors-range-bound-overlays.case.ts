// Cluster: conditional-formatting
//
// Real-world scenario: a sheet has a dropdown, a highlight rule, a review conversation and a filter
// all pinned to the same block of rows, and someone inserts a header row above them. In a spreadsheet
// every one of those follows the cells it was attached to. Anything that stays behind is silent wrong
// output: the dropdown now constrains the row above the data, the highlight paints the wrong cells,
// the conversation hangs off a cell whose contents moved on, and the filter covers a header row it was
// never pointed at. Deleting the block instead must take all four with it rather than clamping them
// onto whatever survived at the cut.

import type {Assert, Case, CorpusApi} from '../case.ts';

export default {
  id: 'splice-reanchors-range-bound-overlays',
  cluster: 'conditional-formatting',
  provenance: {source: 'model-contract-audit'},
  description:
    'A row splice re-anchors every range-bound overlay, not just the cell grid: a data validation, a ' +
    'conditional format, a threaded comment and the sheet autofilter each follow the cells they cover ' +
    'through an insert, and are dropped when the splice deletes every row they were anchored to. The ' +
    'anchors are read back out of the written package, so the shift survives serialization.',

  behavior: [
    {
      name: 'an inserted row moves the anchored content, and the written package agrees',
      async expect(api: CorpusApi, assert: Assert) {
        const {inserted} = await api.spliceReanchorsRangeBoundOverlays();
        assert.strictEqual(inserted.movedValue, 'anchored', 'the covered cell moved down one row');
      },
    },
    {
      name: 'a data validation follows the cells it covers through an inserted row',
      async expect(api: CorpusApi, assert: Assert) {
        const {inserted} = await api.spliceReanchorsRangeBoundOverlays();
        assert.deepStrictEqual(
          inserted.validationRefs,
          ['B6:B7'],
          'the dropdown must cover the moved cells, not the row that took their place',
        );
      },
    },
    {
      name: 'a conditional format follows the cells it covers through an inserted row',
      async expect(api: CorpusApi, assert: Assert) {
        const {inserted} = await api.spliceReanchorsRangeBoundOverlays();
        assert.deepStrictEqual(inserted.formattingRefs, ['B6:B7']);
      },
    },
    {
      name: 'the sheet autofilter follows the region it filters through an inserted row',
      async expect(api: CorpusApi, assert: Assert) {
        const {inserted} = await api.spliceReanchorsRangeBoundOverlays();
        assert.strictEqual(inserted.autoFilterRef, 'B6:C7');
      },
    },
    {
      name: 'a threaded comment follows its cell through an inserted row',
      async expect(api: CorpusApi, assert: Assert) {
        const {inserted} = await api.spliceReanchorsRangeBoundOverlays();
        assert.deepStrictEqual(
          inserted.threadRefs,
          ['B6'],
          'a conversation anchored to a cell must land where the cell landed',
        );
      },
    },
    {
      name: 'deleting every anchored row drops the overlays rather than re-pointing them',
      async expect(api: CorpusApi, assert: Assert) {
        const {deleted} = await api.spliceReanchorsRangeBoundOverlays();
        assert.deepStrictEqual(deleted.validationRefs, []);
        assert.deepStrictEqual(deleted.formattingRefs, []);
        assert.strictEqual(deleted.autoFilterRef, null);
        assert.deepStrictEqual(deleted.threadRefs, []);
      },
    },
  ],
} satisfies Case;
