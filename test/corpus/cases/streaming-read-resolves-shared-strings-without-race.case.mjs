// Cluster: streaming
//
// Real-world scenario: a workbook is read in streaming mode, where the zip package is consumed
// entry-by-entry as an async stream. A reported failure was that some zip entries (notably the
// shared-strings part) were occasionally skipped under concurrency — a hand-rolled stream-iteration
// wrapper racing the native async-iterable contract — leaving string cells unresolved or a read
// hanging. A streaming read must resolve every string-typed cell (never skip shared strings), always
// terminate, and yield identical complete results when many independent reads run concurrently.

/** @typedef {{ name: string, baseline: 'pass'|'fail', expect: (api: any, assert: any) => Promise<void>|void }} Behavior */

export default {
  id: 'streaming-read-resolves-shared-strings-without-race',
  provenance: {source: 'upstream-issue'},
  cluster: 'streaming',
  description:
    'A streaming read resolves every shared-string cell to its text (the shared-strings part is never ' +
    'skipped), and running many independent streaming reads concurrently yields complete, identical ' +
    'results for every one — no dropped entries or missing parses.',

  /** @type {Behavior[]} */
  behavior: [
    {
      name: 'a single streaming read resolves all shared-string cells',
      baseline: 'pass',
      async expect(api, assert) {
        const {singleComplete, singleLength} = await api.streamingSharedStringsRead();
        assert.strictEqual(singleComplete, true, `every string cell resolves; got ${singleLength} rows`);
      },
    },
    {
      name: 'many concurrent streaming reads all complete with resolved strings',
      baseline: 'pass',
      async expect(api, assert) {
        const {concurrentAllComplete, concurrentLengths} = await api.streamingSharedStringsRead();
        assert.strictEqual(
          concurrentAllComplete,
          true,
          `all concurrent reads must complete with resolved strings; got lengths ${JSON.stringify(concurrentLengths)}`
        );
      },
    },
  ],
};
