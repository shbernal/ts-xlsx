import type {Assert, Case, CorpusApi} from '../case.ts';

export default {
  id: 'frozen-header-view-state',
  cluster: 'sheet-views',
  description:
    'A workbook with a frozen top row over grouped columns must carry the view state a consumer ' +
    'initialises its layout from: a <bookViews> window rect, exactly one sheet marked tabSelected, ' +
    'and an outlineLevelCol sizing the column-outline bar that sits directly above the frozen row. ' +
    'Spreadsheet apps always write all three; a package that omits them positions the frozen pane ' +
    'against an uninitialised window, and the header row can stay unpainted until some later event ' +
    'forces a relayout.',
  provenance: {source: 'field-report'},
  behavior: [
    {
      name: 'the workbook declares exactly one window view, before its sheet list',
      baseline: 'pass',
      expect(api: CorpusApi, assert: Assert) {
        const report = api.frozenHeaderViewStateReport();
        assert.strictEqual(report.bookViewCount, 1);
        assert.strictEqual(report.bookViewsBeforeSheets, true);
      },
    },
    {
      name: 'the window view carries a positive width and height for panes to be laid out against',
      baseline: 'pass',
      expect(api: CorpusApi, assert: Assert) {
        const report = api.frozenHeaderViewStateReport();
        assert.ok(
          typeof report.windowWidth === 'number' && report.windowWidth > 0,
          `expected a positive windowWidth, got ${report.windowWidth}`,
        );
        assert.ok(
          typeof report.windowHeight === 'number' && report.windowHeight > 0,
          `expected a positive windowHeight, got ${report.windowHeight}`,
        );
      },
    },
    {
      name: 'exactly one sheet is marked selected, and it is the frozen one',
      baseline: 'pass',
      expect(api: CorpusApi, assert: Assert) {
        // Two selected sheets is a group selection, where an edit to one is applied to all of them;
        // none selected leaves the consumer with no sheet view to initialise.
        assert.deepStrictEqual(api.frozenHeaderViewStateReport().selectedSheets, ['Report']);
      },
    },
    {
      name: 'grouped columns report their depth as outlineLevelCol on the sheet format',
      baseline: 'pass',
      expect(api: CorpusApi, assert: Assert) {
        const report = api.frozenHeaderViewStateReport();
        assert.strictEqual(report.outlineLevelCol, 1);
        // Nothing groups rows here, so the row counterpart stays absent rather than being written 0.
        assert.strictEqual(report.outlineLevelRow, null);
      },
    },
    {
      name: 'the frozen top row still serializes its pane alongside the view state',
      baseline: 'pass',
      expect(api: CorpusApi, assert: Assert) {
        assert.strictEqual(api.frozenHeaderViewStateReport().paneEmitted, true);
      },
    },
  ],
} satisfies Case;
