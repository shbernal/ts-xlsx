// Cluster: xlsx-io
//
// Real-world scenario: a worksheet has manual horizontal page breaks (row breaks) so the print
// layout splits at chosen rows. These are persisted in the worksheet XML as a rowBreaks section (a
// set of <brk> elements). On read, the library should surface those manual row page breaks on the
// loaded worksheet; today the reader ignores the rowBreaks section entirely, so a freshly loaded
// worksheet reports none and a load→save round-trip silently drops the print-pagination the user
// configured.
//
// The fixture is a worksheet with a rowBreaks section (breaks at rows 3 and 6) injected into the
// stored XML, since the writer does not currently author rowBreaks.

/** @typedef {{ name: string, baseline: 'pass'|'fail', expect: (api: any, assert: any) => Promise<void>|void }} Behavior */

const FIXTURE = 'read-manual-row-page-breaks/sample.xlsx';

export default {
  id: 'read-manual-row-page-breaks',
  provenance: {source: 'upstream-issue'},
  cluster: 'xlsx-io',
  description:
    'Manual horizontal page breaks (rowBreaks) declared in a worksheet are surfaced on load and ' +
    'preserved through a load→save round-trip, rather than being ignored on read and dropped on write.',

  /** @type {Behavior[]} */
  behavior: [
    {
      name: 'the fixture declares two manual row breaks (precondition)',
      baseline: 'pass',
      async expect(api, assert) {
        const {sourceBreaks} = await api.roundtripFixtureRowBreaks(FIXTURE);
        assert.deepStrictEqual(sourceBreaks, [3, 6], 'the source XML declares row breaks at rows 3 and 6');
      },
    },
    {
      name: 'the loaded worksheet surfaces the manual row breaks',
      baseline: 'fail',
      async expect(api, assert) {
        const {loadedBreaks} = await api.roundtripFixtureRowBreaks(FIXTURE);
        assert.deepStrictEqual(loadedBreaks, [3, 6], `the reader must surface the row breaks; got ${JSON.stringify(loadedBreaks)}`);
      },
    },
    {
      name: 'a load→save round-trip preserves the rowBreaks section',
      baseline: 'fail',
      async expect(api, assert) {
        const {rewrittenBreaks} = await api.roundtripFixtureRowBreaks(FIXTURE);
        assert.deepStrictEqual(rewrittenBreaks, [3, 6], `the row breaks must survive the round-trip; got ${JSON.stringify(rewrittenBreaks)}`);
      },
    },
  ],
};
