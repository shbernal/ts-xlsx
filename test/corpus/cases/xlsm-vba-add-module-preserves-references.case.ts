// Cluster: security
//
// Real-world scenario: a caller building up a macro project incrementally — adding a helper module to
// an existing macro-enabled workbook rather than authoring the whole project from scratch — needs a way
// to grow an already-loaded vbaProject.bin by one module without disturbing what is already there. The
// project carries state that cannot be re-synthesized from a model: a reference to an external type
// library (PROJECTREFERENCES) and other modules whose bytes (p-code prefix included) must ride through
// untouched. A naive "rebuild the project" strategy would drop the reference and force every untouched
// module to be re-emitted (and thus re-verified) even though only one module actually changed — so
// adding a module must SPLICE a new module block into the original bytes, mirroring how editing an
// existing module's source already does (see xlsm-vba-edit-module-source-preserves-references).
//
// The surface under test is the project-editor primitive addVbaModule(bin, module): it parses the
// existing project fail-closed, inserts a new MODULE record into the dir stream (bumping MODULES_COUNT),
// adds the module's declaration to PROJECT/PROJECTwm, adds the new compressed source stream under the
// VBA storage, and resets _VBA_PROJECT to the recompile cookie so Excel compiles the new module in
// alongside the rest. This case asserts the new module reads back correctly, the hand-crafted
// PROJECTREFERENCES record and an untouched module survive byte-for-byte, and the project recompiles.

import type {Assert, Case, CorpusApi} from '../case.ts';

export default {
  id: 'xlsm-vba-add-module-preserves-references',
  cluster: 'security',
  provenance: {},
  description:
    'Adding a new standard module to an existing VBA project via addVbaModule must land the new ' +
    'module readably alongside the existing ones, while the project’s reference and its other ' +
    'modules (byte-for-byte) are preserved and _VBA_PROJECT is reset to the recompile cookie.',

  behavior: [
    {
      name: 'the new module reads back with its source and procedural kind, alongside the untouched ones',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const {moduleNames, moduleKinds, addedSourcePresent} = await api.xlsmVbaAddModule();
        assert.deepEqual(moduleNames, ['ThisWorkbook', 'Module1', 'Class1', 'AddedModule']);
        assert.deepEqual(
          moduleKinds,
          [
            ['ThisWorkbook', 'document'],
            ['Module1', 'procedural'],
            ['Class1', 'class'],
            ['AddedModule', 'procedural'],
          ],
          'existing modules keep their kind; the new one reads back procedural',
        );
        assert.strictEqual(
          addedSourcePresent,
          true,
          'the new module must read back with exactly the source it was added with',
        );
      },
    },
    {
      name: 'the project reference survives the add verbatim',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const {referencePreserved} = await api.xlsmVbaAddModule();
        assert.strictEqual(
          referencePreserved,
          true,
          'the PROJECTREFERENCES record must be carried through the splice unchanged, not dropped',
        );
      },
    },
    {
      name: 'an untouched module rides through byte-for-byte',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const {untouchedModuleByteIdentical} = await api.xlsmVbaAddModule();
        assert.strictEqual(
          untouchedModuleByteIdentical,
          true,
          'a module the caller did not add or edit must keep identical stream bytes, p-code prefix included',
        );
      },
    },
    {
      name: 'the recompile cookie is set so Excel compiles the new module in',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const {recompileCookieReset} = await api.xlsmVbaAddModule();
        assert.strictEqual(
          recompileCookieReset,
          true,
          '_VBA_PROJECT must be reset so Excel recompiles every module, including the newly added one',
        );
      },
    },
  ],
} satisfies Case;
