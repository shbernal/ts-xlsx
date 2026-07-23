// Cluster: security
//
// Real-world scenario: a caller loads an existing macro-enabled workbook (.xlsm) authored by Excel,
// changes the VBA source of one module, and re-saves — the "tweak an existing macro" workflow. The
// workbook's VBA project carries state the library cannot re-synthesize from a model: references to
// external type libraries (PROJECTREFERENCES), and `document` code-behind modules (ThisWorkbook, a
// sheet) whose linkage to the host is baked into the project's `dir`/`PROJECT` streams. A naive
// "re-author the project" strategy would silently drop every reference and cannot emit a document
// module at all — so editing one module's source must SPLICE into the original bytes, leaving every
// other stream untouched, rather than rebuilding the project from a read-back model.
//
// The public surface is Workbook.setVbaModuleSource(name, source): it reads the attached vbaProject.bin,
// replaces the named module's source stream, resets _VBA_PROJECT to the recompile cookie so Excel
// rebuilds p-code from the new source, and re-attaches — inheriting the macro content type and the
// stale-signature drop. This case drives it through a full read -> edit -> write -> re-read cycle and
// asserts the edit lands while references, untouched modules (including a code-page-1251 one), and the
// macro-enabled shape all survive.

import type {Assert, Case, CorpusApi} from '../case.ts';

export default {
  id: 'xlsm-vba-edit-module-source-preserves-references',
  cluster: 'security',
  provenance: {},
  description:
    'Editing one VBA module’s source in a loaded .xlsm via Workbook.setVbaModuleSource must land ' +
    'the new source and survive a write/re-read cycle, while the project’s references, its other ' +
    'modules (byte-for-byte, code page included), and its macro-enabled content type are preserved and ' +
    '_VBA_PROJECT is reset to the recompile cookie.',

  behavior: [
    {
      name: 'the fixture package loads as a three-module macro project (precondition)',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const {originalModuleNames} = await api.xlsmVbaEditModuleSource();
        assert.deepEqual(originalModuleNames, ['ThisWorkbook', 'Module1', 'Class1']);
      },
    },
    {
      name: 'the edited document module carries its new source after a write/re-read cycle',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const {editedSourcePresent} = await api.xlsmVbaEditModuleSource();
        assert.strictEqual(
          editedSourcePresent,
          true,
          'the ThisWorkbook document module must read back with exactly the edited source',
        );
      },
    },
    {
      name: 'every module keeps its kind, including the edited document module',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const {moduleKinds} = await api.xlsmVbaEditModuleSource();
        assert.deepEqual(moduleKinds, [
          ['ThisWorkbook', 'document'],
          ['Module1', 'procedural'],
          ['Class1', 'class'],
        ]);
      },
    },
    {
      name: 'the project reference survives the edit verbatim',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const {referencePreserved} = await api.xlsmVbaEditModuleSource();
        assert.strictEqual(
          referencePreserved,
          true,
          'the PROJECTREFERENCES record must be carried through the splice unchanged, not dropped',
        );
      },
    },
    {
      name: 'an untouched module rides through byte-for-byte, code page preserved',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const {untouchedModuleByteIdentical, codePageModulePreserved} =
          await api.xlsmVbaEditModuleSource();
        assert.strictEqual(
          untouchedModuleByteIdentical,
          true,
          'a module the caller did not edit must be re-emitted with identical stream bytes',
        );
        assert.strictEqual(
          codePageModulePreserved,
          true,
          'the code-page-1251 module must still decode its non-ASCII source, proving no re-encoding',
        );
      },
    },
    {
      name: 'the re-emitted project keeps the macro content type and resets the recompile cookie',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const {rewrittenIsMacroEnabled, rewrittenHasVba, recompileCookieReset} =
          await api.xlsmVbaEditModuleSource();
        assert.strictEqual(rewrittenHasVba, true, 'xl/vbaProject.bin must be re-emitted');
        assert.strictEqual(
          rewrittenIsMacroEnabled,
          true,
          'the package must stay macro-enabled or Excel flags the re-saved file for repair',
        );
        assert.strictEqual(
          recompileCookieReset,
          true,
          '_VBA_PROJECT must be reset so Excel recompiles p-code from the edited source',
        );
      },
    },
  ],
} satisfies Case;
