// Cluster: tables
//
// Real-world scenario: a user opens a workbook that contains a table, sets the value of a cell inside
// the table's range (a body cell), and writes the workbook back. The result must remain a valid,
// openable spreadsheet — the table XML, its worksheet relationship, and its definition survive, and
// the edited value is present — not a truncated or structurally corrupt package.

import type {Assert, Case, CorpusApi} from '../case.ts';

export default {
  id: 'table-cell-edit-roundtrip-stays-valid',
  provenance: {source: 'upstream-issue'},
  cluster: 'tables',
  description:
    'Editing a cell inside a table’s range and writing the workbook produces a valid, openable ' +
    'package: the table part and its unique worksheet relationship survive and the edited value ' +
    'reads back.',

  behavior: [
    {
      name: 'the edited package writes and reloads without error',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const {writeOk, reloadOk, writeError} = await api.tableCellEditRoundtrip();
        assert.strictEqual(
          writeOk,
          true,
          `writing must not throw; got ${JSON.stringify(writeError)}`,
        );
        assert.strictEqual(reloadOk, true, 'the edited package reloads');
      },
    },
    {
      name: 'the table part survives and the edited value reads back',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const {tablePresent, hasTablePart, editedValue} = await api.tableCellEditRoundtrip();
        assert.strictEqual(hasTablePart, true, 'the table part is still present');
        assert.strictEqual(tablePresent, true, 'the table definition survives');
        assert.strictEqual(editedValue, 999, 'the edited cell value reads back');
      },
    },
    {
      name: 'the worksheet relationship ids remain unique after the edit',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const {relUnique} = await api.tableCellEditRoundtrip();
        assert.strictEqual(relUnique, true, 'the worksheet→table relationship ids stay unique');
      },
    },
  ],
} satisfies Case;
