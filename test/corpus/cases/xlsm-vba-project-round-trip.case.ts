// Cluster: security
//
// Real-world scenario: a macro-enabled workbook (.xlsm) carries a `vbaProject.bin` part linked from
// `xl/_rels/workbook.xml.rels` by a `vbaProject` relationship, and `xl/workbook.xml` declares the
// macro-enabled content type so Excel and OPC readers recognise the package as macro-enabled rather
// than a plain workbook that merely happens to reference a foreign part. A library that reads such a
// file and writes it back out (e.g. to apply an unrelated edit) must not silently discard the VBA
// project — a caller who round-trips a `.xlsm` and re-saves it with that extension would otherwise
// lose the workbook's macros with no error and no indication anything was dropped.
//
// The writer cannot author a VBA project itself (its bytes are opaque and never modeled), so the case
// exercises the round-trip by splicing a synthetic vbaProject part onto an otherwise-plain written
// package before reloading it — the only way to produce a macro-enabled-shaped package without an
// interactive VBA editor.

import type {Assert, Case, CorpusApi} from '../case.ts';

export default {
  id: 'xlsm-vba-project-round-trip',
  cluster: 'security',
  provenance: {},
  description:
    "A macro-enabled workbook's vbaProject part, its workbook relationship, and the workbook's " +
    'macro-enabled content type must all survive being read into the model and written back out, ' +
    'rather than the vbaProject relationship being silently discarded as an unrecognised workbook ' +
    'relationship.',

  behavior: [
    {
      name: 'the synthetic macro-enabled package carries a vbaProject part (precondition)',
      async expect(api: CorpusApi, assert: Assert) {
        const {originalHasVba} = await api.xlsmVbaProjectRoundtrip();
        assert.strictEqual(originalHasVba, true);
      },
    },
    {
      name: 'reading the package captures the vbaProject relationship as a preserved workbook reference',
      async expect(api: CorpusApi, assert: Assert) {
        const {reloadedPreservedCount} = await api.xlsmVbaProjectRoundtrip();
        assert.strictEqual(
          reloadedPreservedCount,
          1,
          'the loaded workbook must expose exactly one preserved vbaProject reference',
        );
      },
    },
    {
      name: 'writing the loaded workbook back out re-emits xl/vbaProject.bin',
      async expect(api: CorpusApi, assert: Assert) {
        const {rewrittenHasVba} = await api.xlsmVbaProjectRoundtrip();
        assert.strictEqual(
          rewrittenHasVba,
          true,
          'xl/vbaProject.bin must survive a read/write cycle, not be silently dropped',
        );
      },
    },
    {
      name: 're-emitted workbook.xml declares the macro-enabled content type, not the plain one',
      async expect(api: CorpusApi, assert: Assert) {
        const {rewrittenIsMacroEnabled} = await api.xlsmVbaProjectRoundtrip();
        assert.strictEqual(
          rewrittenIsMacroEnabled,
          true,
          'a package carrying a VBA project must declare xl/workbook.xml as macro-enabled, or Excel ' +
            'flags the re-saved file as needing repair',
        );
      },
    },
  ],
} satisfies Case;
