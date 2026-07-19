// Cluster: tables
//
// Real-world scenario: a user builds a table whose header labels wrap across multiple visual lines by
// embedding line-break characters (CR/LF) in a column name — "Test\r\nmultiple\r\nlines" — together
// with wrapText alignment. The column name is emitted as an XML attribute value in the table part.
// Raw CR/LF in an attribute is not preserved by XML attribute-value normalization (a CR becomes a
// space on reparse) and makes the package suspect, so the application reports the file as damaged.
// The control characters must be XML-character-escaped (e.g. &#10;) so the table part stays valid and
// the multi-line name round-trips.

/** @typedef {{ name: string, baseline: 'pass'|'fail', expect: (api: any, assert: any) => Promise<void>|void }} Behavior */

export default {
  id: 'table-column-name-with-line-breaks',
  provenance: {source: 'upstream-issue'},
  cluster: 'tables',
  description:
    'A table column name containing line-break characters is emitted with those characters XML-' +
    'escaped (not raw CR/LF) in the tableColumn name attribute, so the table part stays well-formed ' +
    'and opens without a repair prompt.',

  /** @type {Behavior[]} */
  behavior: [
    {
      name: 'writing a table with a CR/LF column name does not throw',
      baseline: 'pass',
      async expect(api, assert) {
        const {writeOk, writeError} = await api.tableColumnNameControlChars();
        assert.strictEqual(writeOk, true, `writing must not throw; got ${JSON.stringify(writeError)}`);
      },
    },
    {
      name: 'the tableColumn name attribute contains no raw CR/LF control characters',
      baseline: 'pass',
      async expect(api, assert) {
        const {rawControlChars, firstColumnTag} = await api.tableColumnNameControlChars();
        assert.strictEqual(
          rawControlChars,
          false,
          `the name must be XML-escaped, not emit raw control chars; got ${JSON.stringify(firstColumnTag)}`
        );
      },
    },
  ],
};
