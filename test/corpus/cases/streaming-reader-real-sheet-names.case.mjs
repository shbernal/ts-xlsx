// Cluster: streaming
//
// Real-world scenario: a workbook produced by a foreign generator declares worksheets with
// meaningful names (e.g. "Sheet" and "test"). Consumed through the streaming reader — iterating
// worksheet readers one at a time — each reader exposes a generic positional name ("Sheet1",
// "Sheet2") instead of the declared one. The eager (whole-workbook) reader resolves the same
// file's names correctly, so the file is well-formed. The streaming reader must join each incoming
// worksheet part to the workbook-level sheet declaration (which carries the authoritative name),
// deterministically regardless of the order parts arrive.

/** @typedef {{ name: string, baseline: 'pass'|'fail', expect: (api: any, assert: any) => Promise<void>|void }} Behavior */

const FIXTURE = 'streaming-reader-real-sheet-names/source.xlsx';

export default {
  id: 'streaming-reader-real-sheet-names',
  provenance: {source: 'upstream-issue'},
  cluster: 'streaming',
  description:
    'The streaming reader exposes each worksheet\'s real declared name (matching the eager read), ' +
    'not a generic positional placeholder like "Sheet2".',

  /** @type {Behavior[]} */
  behavior: [
    {
      name: 'the eager read resolves the declared sheet names (oracle)',
      baseline: 'pass',
      async expect(api, assert) {
        const {eager} = await api.streamVsEagerSheetNames(FIXTURE);
        assert.deepStrictEqual(eager, ['Sheet', 'test'], 'the eager read is the naming oracle');
      },
    },
    {
      name: 'the streaming reader surfaces the same declared names, not positional placeholders',
      baseline: 'fail',
      async expect(api, assert) {
        const {eager, streaming} = await api.streamVsEagerSheetNames(FIXTURE);
        assert.deepStrictEqual(
          streaming,
          eager,
          `streaming sheet names must equal the declared names, not positional placeholders; got ${JSON.stringify(streaming)}`
        );
      },
    },
  ],
};
