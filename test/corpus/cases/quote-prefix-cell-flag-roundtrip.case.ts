// Cluster: styles
//
// Real-world scenario: spreadsheet applications support a "quote prefix" flag on a cell format. When
// set, the application stores the cell's content as literal text even when it looks like a formula or
// number, and shows a leading apostrophe in the formula bar without that apostrophe being part of the
// stored value. In the file format this is a boolean attribute (quotePrefix) on the cell-format
// record (the xf the cell references). A user who marks a cell whose content begins with a
// formula-like character (e.g. a leading equals sign) with quote-prefix expects it stored as literal
// text and the flag preserved on read/modify/write. Legacy neither emits the attribute on write nor
// preserves it on read, so the flag is silently lost.

import type {Assert, Case, CorpusApi} from '../case.ts';

export default {
  id: 'quote-prefix-cell-flag-roundtrip',
  provenance: {source: 'upstream-issue'},
  cluster: 'styles',
  description:
    'A cell marked with the quote-prefix flag writes a cell-format record carrying the quotePrefix ' +
    'attribute, and a read/modify/write round-trip preserves the flag — the mechanism a spreadsheet ' +
    'uses to force formula-like content to be stored as literal text.',

  behavior: [
    {
      name: 'writing a quote-prefixed cell emits the quotePrefix attribute on its cell-format record',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const {writtenQuotePrefix} = await api.quotePrefixReport();
        assert.ok(writtenQuotePrefix, 'the cell\'s xf must carry quotePrefix="1"');
      },
    },
    {
      name: 'the quote-prefix flag survives a read/modify/write round-trip',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const {reloaded} = await api.quotePrefixReport();
        assert.ok(reloaded, 'the reloaded cell must still report the quote-prefix flag');
      },
    },
  ],
} satisfies Case;
