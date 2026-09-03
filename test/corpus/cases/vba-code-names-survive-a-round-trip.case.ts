// Cluster: xlsx-io
//
// Real-world scenario: a macro-enabled workbook binds its VBA project to its document through code
// names, not display names. `<workbookPr codeName="ThisWorkbook"/>` is what a macro means when it
// writes `ThisWorkbook`, and `<sheetPr codeName="Sheet1"/>` is what it means by `Sheet1`, which is
// deliberately *not* the sheet's tab name, so that renaming a tab does not break the code. Excel
// writes both once a project exists, and a French or German Excel writes names like `Feuil2` there
// while the tab reads something else entirely.
//
// A round trip that drops them leaves the macros referring to names the document no longer answers
// to. Nothing in the reading model consults a code name, which is exactly why it has to be preserved
// deliberately rather than by accident.

import type {Assert, Case, CorpusApi} from '../case.ts';

const FIXTURE = 'vba-code-names-survive-a-round-trip/code-names.xlsx';

export default {
  id: 'vba-code-names-survive-a-round-trip',
  provenance: {source: 'round-trip-fidelity-audit'},
  cluster: 'xlsx-io',
  description:
    'The workbook and sheet VBA code names (`<workbookPr codeName>`, `<sheetPr codeName>`) are read ' +
    'and written back unchanged, so a macro-enabled workbook keeps the binding between its project ' +
    'and its document across a read/write cycle.',

  behavior: [
    {
      name: 'the workbook and every sheet code name are read off the package',
      async expect(api: CorpusApi, assert: Assert) {
        const {before} = await api.codeNamesReport(FIXTURE);
        assert.strictEqual(before.workbook, 'ThisWorkbook');
        assert.strictEqual(before.sheets.Data.codeName, 'Sheet1');
        assert.strictEqual(before.sheets.Notes.codeName, 'Feuil2');
      },
    },
    {
      name: 'they survive a write and read back identically',
      async expect(api: CorpusApi, assert: Assert) {
        const {before, after} = await api.codeNamesReport(FIXTURE);
        // Named, not merely compared: a reader that drops code names and a writer that never emits
        // them agree with each other perfectly, and `after === before` would call that a pass.
        assert.strictEqual(after.workbook, 'ThisWorkbook');
        assert.strictEqual(after.sheets.Data.codeName, 'Sheet1');
        assert.strictEqual(after.sheets.Notes.codeName, 'Feuil2');
        assert.deepStrictEqual(after, before, 'and nothing else about them moved');
      },
    },
    {
      name: 'a code name does not displace what else `<sheetPr>` carries',
      async expect(api: CorpusApi, assert: Assert) {
        // The attribute and the element's children are written by one function, and the first sheet
        // carries both, so this is where "emit the attribute" could quietly drop a tab colour.
        const {before, after} = await api.codeNamesReport(FIXTURE);
        assert.deepStrictEqual(before.sheets.Data.tabColor, {argb: 'FFFF0000'});
        assert.deepStrictEqual(after.sheets.Data.tabColor, before.sheets.Data.tabColor);
      },
    },
  ],
} satisfies Case;
