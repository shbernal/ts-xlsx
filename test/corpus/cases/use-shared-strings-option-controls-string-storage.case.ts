// Cluster: streaming
//
// Real-world scenario: a workbook with string cells is written to a buffer. The writer exposes a
// useSharedStrings option meant to choose how in-cell text is stored — enabled, strings are
// deduplicated into a shared string table (sharedStrings.xml) and cells reference it by index;
// disabled, strings are stored inline in the worksheet cells. The reported failure is that the option
// has no effect: shared strings are emitted regardless. Both representations must read back to the
// same values, but the option must actually control the storage.

import type {Assert, Case, CorpusApi} from '../case.ts';

export default {
  id: 'use-shared-strings-option-controls-string-storage',
  provenance: {source: 'upstream-issue'},
  cluster: 'streaming',
  description:
    'The useSharedStrings write option controls string storage: enabled emits a shared string table ' +
    'with cell references, disabled stores strings inline with no shared-strings part — rather than ' +
    'always using shared strings regardless of the option.',

  behavior: [
    {
      name: 'useSharedStrings=true emits a shared string table and cell references',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const {hasSharedStringsPart, isSharedRef} = await api.sharedStringsOption(true);
        assert.strictEqual(hasSharedStringsPart, true, 'a sharedStrings part is written');
        assert.strictEqual(isSharedRef, true, 'the cell is stored as a shared-string reference');
      },
    },
    {
      name: 'useSharedStrings=false stores strings inline with no shared-strings part',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const {hasSharedStringsPart, isInline} = await api.sharedStringsOption(false);
        assert.strictEqual(
          hasSharedStringsPart,
          false,
          'no sharedStrings part when the option is disabled',
        );
        assert.strictEqual(isInline, true, 'the string is stored inline in the cell');
      },
    },
  ],
} satisfies Case;
