// Cluster: security
//
// Real-world scenario: a caller pruning a macro project incrementally — dropping a helper module that is
// no longer needed from an existing macro-enabled workbook — needs a way to shrink an already-loaded
// vbaProject.bin by one module without disturbing what remains. The project carries state that cannot be
// re-synthesized from a model: a reference to an external type library (PROJECTREFERENCES) and other
// modules whose bytes (p-code prefix included) must ride through untouched. A naive "rebuild the project"
// strategy would drop the reference and force every untouched module to be re-emitted (and thus
// re-verified) even though only one module actually left — so removing a module must SPLICE it out of the
// original bytes, mirroring how adding a reference already does (see
// xlsm-vba-add-reference-preserves-modules).
//
// The surface under test is the project-editor primitive removeVbaModule(bin, name): it parses the
// existing project fail-closed, resolves the module's kind and stream name, drops its MODULE record block
// from the dir stream (decrementing MODULES_COUNT), drops its declaration from PROJECT/PROJECTwm, and
// drops its compressed source stream from the VBA storage. It leaves _VBA_PROJECT completely untouched:
// Excel runs the surviving modules' existing compiled p-code, and resetting the stream to an "unmatchable
// version" cookie does not force a recompile from source — on a project that carries real p-code it
// actively crashes the load (ADR 0019). Only procedural/class modules can be removed this way —
// removing a document module (ThisWorkbook) or a designer module would break host linkage the primitive
// has no visibility into, so it is rejected fail-closed.
// This case asserts the removed module is gone from every stream it touched, the hand-crafted
// PROJECTREFERENCES record and an untouched module survive byte-for-byte, and _VBA_PROJECT is preserved
// untouched.

import type {Assert, Case, CorpusApi} from '../case.ts';

export default {
  id: 'xlsm-vba-remove-module-preserves-references',
  cluster: 'security',
  provenance: {},
  description:
    'Removing a standard module from an existing VBA project via removeVbaModule must drop it from ' +
    'every stream it touched, while the project’s reference and its other modules (byte-for-byte) are ' +
    'preserved and _VBA_PROJECT is preserved untouched.',

  behavior: [
    {
      name: 'the removed module is gone; the other modules and their kinds survive',
      async expect(api: CorpusApi, assert: Assert) {
        const {moduleNames, moduleKinds} = await api.xlsmVbaRemoveModule();
        assert.deepEqual(moduleNames, ['ThisWorkbook', 'Class1']);
        assert.deepEqual(
          moduleKinds,
          [
            ['ThisWorkbook', 'document'],
            ['Class1', 'class'],
          ],
          'the surviving modules keep their exact kind',
        );
      },
    },
    {
      name: 'the removed module’s stream is gone from the VBA storage',
      async expect(api: CorpusApi, assert: Assert) {
        const {removedModuleStreamGone} = await api.xlsmVbaRemoveModule();
        assert.strictEqual(
          removedModuleStreamGone,
          true,
          "the removed module's CFB stream must be gone",
        );
      },
    },
    {
      name: 'the project reference and an untouched module survive the removal verbatim',
      async expect(api: CorpusApi, assert: Assert) {
        const {referencePreserved, untouchedModuleByteIdentical} = await api.xlsmVbaRemoveModule();
        assert.strictEqual(
          referencePreserved,
          true,
          'the PROJECTREFERENCES record must be carried through the splice unchanged, not dropped',
        );
        assert.strictEqual(
          untouchedModuleByteIdentical,
          true,
          'a module the caller did not remove must keep identical stream bytes, p-code prefix included',
        );
      },
    },
    {
      name: 'PROJECT drops only the removed module’s declaration; PROJECTwm drops only its name pair',
      async expect(api: CorpusApi, assert: Assert) {
        const {removedDeclLineGone, otherDeclLinesSurvive, projectwmNoLongerHasModule1} =
          await api.xlsmVbaRemoveModule();
        assert.strictEqual(removedDeclLineGone, true, "the removed module's Module= line is gone");
        assert.strictEqual(
          otherDeclLinesSurvive,
          true,
          'the surviving modules’ Document=/Class= lines are untouched',
        );
        assert.strictEqual(
          projectwmNoLongerHasModule1,
          true,
          "the removed module's name pair is gone from PROJECTwm",
        );
      },
    },
    {
      name: '_VBA_PROJECT is preserved untouched',
      async expect(api: CorpusApi, assert: Assert) {
        const {vbaProjectStreamPreserved} = await api.xlsmVbaRemoveModule();
        assert.strictEqual(
          vbaProjectStreamPreserved,
          true,
          '_VBA_PROJECT must be left byte-for-byte unchanged — Excel runs the surviving modules’ ' +
            'existing p-code, and resetting the cookie would crash the load',
        );
      },
    },
  ],
} satisfies Case;
