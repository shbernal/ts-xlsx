// Cluster: tables
//
// Real-world scenario: a table is built by mapping a header list into column definitions, and that
// list happens to contain repeated names (e.g. every column is called "foo"). OOXML requires every
// tableColumn name to be unique within a table part; a table whose column names collide produces a
// file that Excel flags as corrupt and offers to repair. The library must never emit such a file —
// duplicate incoming names must be disambiguated deterministically (keep the first "foo", then
// "foo1", "foo2", …) so the workbook opens cleanly. The bug: the writer emits the colliding names
// verbatim (name="foo" three times), yielding a corrupt package.

/** @typedef {{ name: string, baseline: 'pass'|'fail', expect: (api: any, assert: any) => Promise<void>|void }} Behavior */

export default {
  id: 'table-duplicate-column-names-disambiguated',
  provenance: {source: 'upstream-issue'},
  cluster: 'tables',
  description:
    'A table whose supplied column names collide must disambiguate them into a unique set in the ' +
    'written tableColumns (OOXML requires unique names), producing a package that reloads cleanly ' +
    'rather than the colliding names being emitted verbatim into a corrupt file.',

  /** @type {Behavior[]} */
  behavior: [
    {
      name: 'distinct column names are emitted unchanged (control)',
      baseline: 'pass',
      async expect(api, assert) {
        const {ok, writtenNames, uniqueNames} = await api.tableDuplicateColumnNamesReport(['a', 'b', 'c']);
        assert.ok(ok, 'the table writes');
        assert.deepStrictEqual(writtenNames, ['a', 'b', 'c'], 'already-distinct names pass through verbatim');
        assert.strictEqual(uniqueNames, true, 'distinct names stay unique');
      },
    },
    {
      name: 'colliding column names are disambiguated into a unique set',
      baseline: 'fail',
      async expect(api, assert) {
        const {ok, writtenNames, uniqueNames} = await api.tableDuplicateColumnNamesReport(['foo', 'foo', 'foo']);
        assert.ok(ok, 'the table writes');
        assert.strictEqual(
          uniqueNames,
          true,
          `each tableColumn name must be unique; got ${JSON.stringify(writtenNames)} — OOXML rejects a ` +
            'table whose column names collide'
        );
      },
    },
  ],
};
