// Cluster: streaming
//
// Real-world scenario: a service streams a spreadsheet directly to a remote sink — object storage,
// a blob store, an outbound HTTP upload — instead of to a file on disk, precisely to avoid holding
// the whole document in memory. It builds a streaming workbook writer over a caller-owned
// writable/pass-through stream, adds a worksheet, commits rows, commits the worksheet, then awaits
// the workbook commit so it can finalize the upload afterward. The reported failure is a deadlock:
// when the sink is a caller-supplied stream (rather than a library-owned file stream), the commit
// promise is claimed never to settle, hanging the upload.
//
// The durable requirement locked here: committing a streaming workbook whose output is a
// caller-supplied writable (a plain PassThrough or a Duplex) must SETTLE within bounded time and the
// sink must receive a complete, valid package — so a caller can deterministically sequence upload
// finalization after commit. (The specific cloud-SDK hangs are downstream stream-implementation
// quirks; the contract the library owes is that its own commit resolves over a standard writable.)

/** @typedef {{ name: string, baseline: 'pass'|'fail', expect: (api: any, assert: any) => Promise<void>|void }} Behavior */

export default {
  id: 'streaming-commit-over-caller-supplied-stream-resolves',
  provenance: {source: 'upstream-issue'},
  cluster: 'streaming',
  description:
    'A streaming workbook commit whose output is a caller-supplied writable settles within bounded ' +
    'time and delivers a complete, valid package to the sink — it must not hang waiting on a finish ' +
    'signal, whether the sink is a plain PassThrough or a Duplex stream.',

  /** @type {Behavior[]} */
  behavior: [
    {
      name: 'committing a streamed workbook over a caller PassThrough resolves within bounded time and yields a valid package',
      baseline: 'pass',
      async expect(api, assert) {
        const {settled, timedOut, bytes, valid} = await api.streamCommitReport({duplex: false});
        assert.strictEqual(timedOut, false, 'the workbook commit must not hang over a caller PassThrough');
        assert.strictEqual(settled, 'resolved', `commit must resolve; got ${settled}`);
        assert.ok(bytes > 0 && valid, 'the sink must receive a complete, re-openable package');
      },
    },
    {
      name: 'the same commit settles over a Duplex sink — completion does not depend on the library owning the stream',
      baseline: 'pass',
      async expect(api, assert) {
        const {settled, timedOut, valid} = await api.streamCommitReport({duplex: true});
        assert.strictEqual(timedOut, false, 'the commit must not hang over a caller Duplex stream');
        assert.strictEqual(settled, 'resolved', `commit must resolve over a Duplex; got ${settled}`);
        assert.ok(valid, 'the Duplex sink must receive a complete, re-openable package');
      },
    },
  ],
};
