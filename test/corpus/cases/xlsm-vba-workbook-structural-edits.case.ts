// Cluster: security
//
// Real-world scenario: a caller reads an existing macro-enabled workbook and wants to shape its macro
// project — drop a module that's no longer needed, wire up a COM reference — using the public Workbook
// API, not the low-level project-editor primitives. Workbook.removeVbaModule/addVbaReference exist
// precisely so a caller never has to reach past the model into raw vbaProject.bin bytes for this. Each
// method must inherit the same splice guarantees the underlying primitive proved in isolation (see
// xlsm-vba-remove-module-preserves-references, xlsm-vba-add-reference-preserves-modules) once routed
// through readXlsx -> Workbook -> writeXlsx: a real package round-trip, not just a bare-bin call.
//
// This case chains both edits on one workbook and asserts the result survives a full
// readXlsx/writeXlsx/readXlsx cycle: the final module set and kinds are correct, the pre-existing
// PROJECTREFERENCES record and the newly added reference both survive, an untouched module rides through
// byte-for-byte, the package stays declared macro-enabled, and _VBA_PROJECT is preserved untouched.

import type {Assert, Case, CorpusApi} from '../case.ts';

export default {
  id: 'xlsm-vba-workbook-structural-edits',
  cluster: 'security',
  provenance: {},
  description:
    'Workbook.removeVbaModule/addVbaReference, chained on a workbook read from a real package, must ' +
    'survive a full readXlsx/writeXlsx round-trip: the module set lands correctly, the ' +
    "project's reference and an untouched module survive, and the package stays macro-enabled.",

  behavior: [
    {
      name: 'the final module set and kinds reflect the remove',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const {moduleNames, moduleKinds} = await api.xlsmVbaWorkbookStructuralEdits();
        assert.deepEqual(moduleNames, ['ThisWorkbook', 'Class1']);
        assert.deepEqual(
          moduleKinds,
          [
            ['ThisWorkbook', 'document'],
            ['Class1', 'class'],
          ],
          'Module1 is gone; ThisWorkbook and Class1 keep their kinds',
        );
      },
    },
    {
      name: 'the pre-existing reference and the newly added reference both survive',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const {originalReferencePreserved, newReferencePresent} =
          await api.xlsmVbaWorkbookStructuralEdits();
        assert.strictEqual(
          originalReferencePreserved,
          true,
          'the hand-crafted PROJECTREFERENCES record must survive both chained edits, not just one',
        );
        assert.strictEqual(newReferencePresent, true, 'the newly added reference must be present');
      },
    },
    {
      name: 'an untouched module survives the chained edits verbatim',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const {untouchedModuleByteIdentical} = await api.xlsmVbaWorkbookStructuralEdits();
        assert.strictEqual(
          untouchedModuleByteIdentical,
          true,
          'Class1, touched by none of the three edits, must keep identical stream bytes',
        );
      },
    },
    {
      name: 'the package stays macro-enabled and _VBA_PROJECT is preserved untouched',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const {rewrittenIsMacroEnabled, vbaProjectStreamPreserved} =
          await api.xlsmVbaWorkbookStructuralEdits();
        assert.strictEqual(
          rewrittenIsMacroEnabled,
          true,
          'the rewritten workbook part must still declare the macro-enabled content type',
        );
        assert.strictEqual(
          vbaProjectStreamPreserved,
          true,
          '_VBA_PROJECT must be left byte-for-byte unchanged — Excel runs the remaining modules’ ' +
            'existing p-code, and resetting the cookie would crash the load',
        );
      },
    },
  ],
} satisfies Case;
