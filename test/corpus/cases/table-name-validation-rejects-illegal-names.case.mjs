// Cluster: tables
//
// Real-world scenario: a user creates a worksheet table and gives it a display/definition name Excel
// considers illegal — one with spaces, an apostrophe, other punctuation ("Bob's Accounts"), or a
// leading digit. Excel's table-name rules are strict: the name must start with a letter, underscore,
// or backslash, and every later character must be a letter, digit, period, or underscore; a name
// that collides with a cell reference (like "A1") is also invalid. The library accepts any string and
// writes it straight into the table part, so the workbook opens with a "we found a problem" repair
// prompt and Excel silently rewrites the name. The writer should refuse an illegal table name up
// front with an actionable error rather than emitting a file a consumer flags as corrupt.

/** @typedef {{ name: string, baseline: 'pass'|'fail', expect: (api: any, assert: any) => Promise<void>|void }} Behavior */

const tableSpec = tableName => ({
  sheets: [{name: 'S', tables: [{name: tableName, ref: 'A1', headers: ['H1', 'H2'], rows: [['a', 1]]}]}],
});

export default {
  id: 'table-name-validation-rejects-illegal-names',
  provenance: {source: 'upstream-issue'},
  cluster: 'tables',
  description:
    'A table name that violates Excel’s identifier rules (spaces, apostrophes/punctuation, a leading ' +
    'digit) is rejected at write time with an error, rather than being written into the table part ' +
    'and producing a repair-prompting file; a valid identifier name is accepted and round-trips.',

  /** @type {Behavior[]} */
  behavior: [
    {
      name: 'a table name containing spaces or an apostrophe is rejected',
      baseline: 'fail',
      async expect(api, assert) {
        const result = await api.tryWriteWorkbook(tableSpec("Bob's Accounts"));
        assert.strictEqual(result.ok, false, 'a name with a space and apostrophe must be rejected, not written through');
      },
    },
    {
      name: 'a table name starting with a digit is rejected',
      baseline: 'fail',
      async expect(api, assert) {
        const result = await api.tryWriteWorkbook(tableSpec('1Digit'));
        assert.strictEqual(result.ok, false, 'a name starting with a digit must be rejected');
      },
    },
    {
      // A hyphen is ambiguous with the subtraction operator, so Excel forbids it in a table name and
      // treats a file carrying "test-name" as corrupt — it must be rejected, not written verbatim.
      name: 'a table name containing a hyphen is rejected',
      baseline: 'fail',
      async expect(api, assert) {
        const result = await api.tryWriteWorkbook(tableSpec('test-name'));
        assert.strictEqual(result.ok, false, 'a hyphenated name must be rejected, not emitted into corrupt XML');
      },
    },
    {
      name: 'a valid identifier table name is accepted and survives into the written table part',
      baseline: 'pass',
      async expect(api, assert) {
        const spec = tableSpec('Valid_Name');
        const result = await api.tryWriteWorkbook(spec);
        assert.strictEqual(result.ok, true, `a valid name must write; got ${JSON.stringify(result.error)}`);
        const {tables} = await api.inspectPackage(spec);
        assert.strictEqual(tables[0].name, 'Valid_Name', 'the valid name is written verbatim into the table XML');
      },
    },
  ],
};
