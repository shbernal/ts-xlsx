// Cluster: styles
//
// Real-world scenario: a cell holding a value that looks like a number but must stay text (a part
// code, a leading-zero identifier) carries the quote-prefix flag, and a cell styled through a theme's
// named style carries a link to it. Both live on the cell's `xf` record exactly as its fill and
// number format do, and both are written and read back the same way.
//
// They were not part of the tuple every copy path is driven by, so every copy path dropped them: a
// row splice, a column splice, a `duplicateRow`, a `dst.model = src.model`. The number format came
// through and the quote prefix did not, so inserting a row above a sheet turned '007' back into 7.

import type {Assert, Case, CorpusApi} from '../case.ts';

export default {
  id: 'cell-formatting-survives-every-structural-copy',
  provenance: {source: 'grid-edit-audit'},
  cluster: 'styles',
  description:
    "A cell's whole formatting, including the quote-prefix flag and its link to a named cell " +
    'style, survives every path that copies a cell: a row splice, a column splice, a duplicated ' +
    'row, and a worksheet-model assignment onto another sheet.',

  behavior: [
    {
      name: 'a row splice carries the quote prefix and the named-style link with the cell',
      expect(api: CorpusApi, assert: Assert) {
        const {spliceRows} = api.cellContentSurvivesCopies();
        assert.equal(spliceRows.value, '007');
        assert.equal(spliceRows.numFmt, '0.00');
        assert.equal(spliceRows.quotePrefix, true);
        assert.equal(spliceRows.namedStyle, 1);
      },
    },
    {
      name: 'a column splice does the same on the other axis',
      expect(api: CorpusApi, assert: Assert) {
        const {spliceColumns} = api.cellContentSurvivesCopies();
        assert.equal(spliceColumns.quotePrefix, true);
        assert.equal(spliceColumns.namedStyle, 1);
      },
    },
    {
      name: 'a duplicated row reproduces both',
      expect(api: CorpusApi, assert: Assert) {
        const {duplicateRow} = api.cellContentSurvivesCopies();
        assert.equal(duplicateRow.quotePrefix, true);
        assert.equal(duplicateRow.namedStyle, 1);
      },
    },
    {
      name: 'and a worksheet-model assignment onto another sheet carries both across',
      expect(api: CorpusApi, assert: Assert) {
        const {modelCopy} = api.cellContentSurvivesCopies();
        assert.equal(modelCopy.value, '007');
        assert.equal(modelCopy.quotePrefix, true);
        assert.equal(modelCopy.namedStyle, 1);
      },
    },
    {
      name: 'the spliced cell still writes both onto its xf record',
      expect(api: CorpusApi, assert: Assert) {
        const {written} = api.cellContentSurvivesCopies();
        assert.equal(written.quotePrefix, true, 'the flag reaches the emitted cellXfs entry');
        assert.equal(written.xfId, true, 'and so does the link to the named style');
      },
    },
  ],
} satisfies Case;
