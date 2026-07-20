// Cluster: security
//
// Real-world scenario: a workbook is protected at the WORKBOOK level — its structure is locked so
// sheets cannot be added, deleted, reordered, or unhidden (the OOXML <workbookProtection
// lockStructure="1"> element). A user opens that workbook, reads it, and saves it back out. The
// workbook-level structure protection must survive that round-trip. The observed defect: the writer
// drops <workbookProtection> entirely, silently unlocking the structure of a file that was protected —
// even though worksheet-level protection (locked cells / sheetProtection) is preserved. Structure
// protection is a real, if weak, integrity signal and must not be lost on a passthrough save.

import type {Assert, Case, CorpusApi} from '../case.ts';

export default {
  id: 'workbook-structure-protection-survives-roundtrip',
  provenance: {source: 'upstream-issue'},
  cluster: 'security',
  description:
    'Workbook-level structure protection (lockStructure) present in a loaded workbook survives a ' +
    'read→write round-trip rather than being silently dropped, so a protected workbook stays ' +
    'structure-locked after a passthrough save.',

  behavior: [
    {
      name: 'workbook structure protection is re-emitted after a read→write round-trip',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const {sourceHadProtection, rewrittenHasProtection} =
          await api.workbookProtectionRoundtrip();
        assert.strictEqual(
          sourceHadProtection,
          true,
          'the source workbook declares workbook protection',
        );
        assert.strictEqual(
          rewrittenHasProtection,
          true,
          'the re-written workbook must still declare workbookProtection (not silently drop it)',
        );
      },
    },
    {
      name: 'the re-emitted protection still locks the workbook structure',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const {rewrittenLocksStructure} = await api.workbookProtectionRoundtrip();
        assert.strictEqual(
          rewrittenLocksStructure,
          true,
          'lockStructure="1" survives the round-trip',
        );
      },
    },
  ],
} satisfies Case;
