// Cluster: security
//
// Real-world scenario: incremental macro authoring often needs to add a library reference to an
// existing project — e.g. wiring up "Microsoft Scripting Runtime" (Dictionary, FileSystemObject) for a
// module that was just added. The project carries state that cannot be re-synthesized from a model: an
// existing reference (PROJECTREFERENCES) and every module's bytes, p-code prefix included, must ride
// through untouched while the new reference is spliced in. A naive "rebuild the project" strategy would
// drop the existing reference and force every module to be re-emitted (and re-verified) even though
// nothing about them changed — so adding a reference must SPLICE a new REFERENCENAME + REFERENCEREGISTERED
// record pair into the dir stream's original bytes, mirroring how removing a module already does (see
// xlsm-vba-remove-module-preserves-references).
//
// A non-obvious wrinkle, confirmed against a genuine Excel-authored project (not guessed from the spec
// alone): the plain-text PROJECT stream carries NO "Reference=" line for a registered library reference —
// references live only in the binary dir stream. So this case also locks that the PROJECT stream is left
// completely untouched by the add.
//
// The surface under test is the project-editor primitive addVbaReference(bin, ref): it parses the
// existing project fail-closed, assembles the [MS-OVBA] LibidReference string
// (`*\G{GUID}#Major.Minor#LCID#Path#Name`), and inserts REFERENCENAME + REFERENCEREGISTERED into the dir
// stream immediately before MODULES_COUNT (references have no count field of their own). It leaves
// _VBA_PROJECT completely untouched: Excel runs each module's existing compiled p-code, and resetting the
// stream to an "unmatchable version" cookie does not force a recompile from source — on a project that
// carries real p-code it actively crashes the load (ADR 0019).

import type {Assert, Case, CorpusApi} from '../case.ts';

export default {
  id: 'xlsm-vba-add-reference-preserves-modules',
  cluster: 'security',
  provenance: {},
  description:
    'Adding a registered library reference to an existing VBA project via addVbaReference must land ' +
    'the new REFERENCENAME/REFERENCEREGISTERED pair in the dir stream, while an existing reference and ' +
    'every module (byte-for-byte) are preserved, the PROJECT stream is left untouched, and ' +
    '_VBA_PROJECT is preserved untouched.',

  behavior: [
    {
      name: 'the module set is unaffected by adding a reference',
      async expect(api: CorpusApi, assert: Assert) {
        const {moduleNames} = await api.xlsmVbaAddReference();
        assert.deepEqual(moduleNames, ['ThisWorkbook', 'Module1', 'Class1']);
      },
    },
    {
      name: 'the pre-existing reference and the newly added reference are both present in the dir stream',
      async expect(api: CorpusApi, assert: Assert) {
        const {existingReferencePreserved, newReferencePresent} = await api.xlsmVbaAddReference();
        assert.strictEqual(
          existingReferencePreserved,
          true,
          'a REFERENCEREGISTERED record already in the project must survive the splice unchanged',
        );
        assert.strictEqual(
          newReferencePresent,
          true,
          'the new reference’s assembled Libid string must be present and correctly formed',
        );
      },
    },
    {
      name: 'an untouched module rides through byte-for-byte',
      async expect(api: CorpusApi, assert: Assert) {
        const {untouchedModuleByteIdentical} = await api.xlsmVbaAddReference();
        assert.strictEqual(
          untouchedModuleByteIdentical,
          true,
          'a module stream must keep identical bytes, p-code prefix included, when only a reference is added',
        );
      },
    },
    {
      name: 'the PROJECT stream is left completely untouched',
      async expect(api: CorpusApi, assert: Assert) {
        const {projectStreamUnchanged} = await api.xlsmVbaAddReference();
        assert.strictEqual(
          projectStreamUnchanged,
          true,
          'a registered reference has no Reference= line in a real Excel-authored PROJECT stream, so ' +
            'adding one must not touch PROJECT at all',
        );
      },
    },
    {
      name: '_VBA_PROJECT is preserved untouched',
      async expect(api: CorpusApi, assert: Assert) {
        const {vbaProjectStreamPreserved} = await api.xlsmVbaAddReference();
        assert.strictEqual(
          vbaProjectStreamPreserved,
          true,
          '_VBA_PROJECT must be left byte-for-byte unchanged — Excel runs the modules’ existing p-code, ' +
            'and resetting the cookie would crash the load',
        );
      },
    },
  ],
} satisfies Case;
