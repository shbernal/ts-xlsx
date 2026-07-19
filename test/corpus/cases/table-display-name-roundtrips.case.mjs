// Cluster: tables
//
// Real-world scenario: an Excel table carries both an internal name and a human-facing display name.
// A caller authors a table with a distinct display name and expects it to appear in the written table
// XML and to survive a round-trip. A serializer that mis-keys the display-name property (e.g. a typo
// like "displyName") silently ignores the supplied value and the table falls back to a default name.
// The display name a caller sets must reach the `displayName` attribute of the table part and read
// back unchanged.

/** @typedef {{ name: string, baseline: 'pass'|'fail', expect: (api: any, assert: any) => Promise<void>|void }} Behavior */

const DISPLAY = 'My Display Name';

export default {
  id: 'table-display-name-roundtrips',
  provenance: {source: 'upstream-issue'},
  cluster: 'tables',
  description:
    "A table's human-facing display name is written to the table XML's displayName attribute and " +
    'survives a read→write round-trip, rather than being dropped in favor of a default.',

  /** @type {Behavior[]} */
  behavior: [
    {
      name: 'the supplied display name reaches the table XML displayName attribute',
      baseline: 'pass',
      async expect(api, assert) {
        const {writtenDisplayName} = await api.tableDisplayNameReport(DISPLAY);
        assert.strictEqual(
          writtenDisplayName,
          DISPLAY,
          'the display name must be serialized, not dropped to a default',
        );
      },
    },
    {
      name: 'the display name survives a reload',
      baseline: 'pass',
      async expect(api, assert) {
        const {reloadedDisplayName} = await api.tableDisplayNameReport(DISPLAY);
        assert.strictEqual(
          reloadedDisplayName,
          DISPLAY,
          'the display name reads back after a round-trip',
        );
      },
    },
    {
      name: 'the internal name is preserved distinctly from the display name',
      baseline: 'pass',
      async expect(api, assert) {
        const {reloadedName} = await api.tableDisplayNameReport(DISPLAY);
        assert.strictEqual(
          reloadedName,
          'MyTable',
          'the internal table name is unaffected by the display name',
        );
      },
    },
  ],
};
