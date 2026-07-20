// Cluster: xlsx-io
//
// Real-world scenario: a worksheet is created and given content (e.g. explicit row
// heights) but no explicit view options. OOXML worksheets are expected to carry a
// <sheetViews> block with at least one <sheetView>; without it, row/column
// dimensions can render inconsistently across Excel's display-scaling settings and
// some consumers flag the sheet. A writer should always emit a default sheetView.

import type {Assert, Case, CorpusApi} from '../case.ts';

export default {
  id: 'worksheet-always-emits-sheetview',
  provenance: {source: 'upstream-issue', repo: 'exceljs/exceljs', ref: 743},
  cluster: 'xlsx-io',
  description:
    'Every written worksheet must serialize a <sheetViews> block with at least one ' +
    '<sheetView>, even when no view options were set.',

  behavior: [
    {
      name: 'a worksheet with explicit row heights still emits a sheetView',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const {sheets} = await api.inspectPackage({
          sheets: [
            {
              name: 'Sheet1',
              rows: [
                {index: 1, height: 40},
                {index: 2, height: 40},
              ],
            },
          ],
        });
        assert.ok(sheets.Sheet1.hasSheetViews, 'expected a <sheetViews> block');
        assert.ok(sheets.Sheet1.sheetViewCount >= 1, 'expected at least one <sheetView>');
      },
    },
    {
      name: 'a default worksheet emits a default sheetView',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const {sheets} = await api.inspectPackage({
          sheets: [{name: 'Sheet1', cells: [{ref: 'A1', value: 'x'}]}],
        });
        assert.ok(sheets.Sheet1.hasSheetViews, 'expected a <sheetViews> block');
        assert.ok(sheets.Sheet1.sheetViewCount >= 1, 'expected at least one <sheetView>');
      },
    },
  ],
} satisfies Case;
