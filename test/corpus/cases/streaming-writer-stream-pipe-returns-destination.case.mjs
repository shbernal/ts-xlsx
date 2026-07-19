// Cluster: streaming
//
// Real-world scenario: a caller drives the streaming workbook writer without handing it a destination
// stream, so the writer exposes its own output stream, and routes that stream onward with Node's
// standard idiom — `writer.stream.pipe(out).on('finish', …)` or `stream.pipeline(writer.stream, out,
// cb)`. Node's Readable.pipe(dest) is contractually required to RETURN dest so those forms compose:
// pipe calls chain, and pipeline wires each stage to the next. The writer's stream must obey that
// contract. When it instead returns undefined, `.pipe(out).on('finish', …)` throws ("Cannot read
// properties of undefined") and pipeline mis-wires — the finish handler never attaches and the write
// appears to hang or silently drop, even though the bytes themselves are produced correctly.

/** @typedef {{ name: string, baseline: 'pass'|'fail', expect: (api: any, assert: any) => Promise<void>|void }} Behavior */

export default {
  id: 'streaming-writer-stream-pipe-returns-destination',
  provenance: {source: 'upstream-issue'},
  cluster: 'streaming',
  description:
    "The streaming writer's output stream honors Node's pipe contract: pipe(dest) returns dest so " +
    'the standard `writer.stream.pipe(out).on("finish", …)` and stream.pipeline idioms compose, while ' +
    'the piped payload still reconstitutes a valid workbook.',

  /** @type {Behavior[]} */
  behavior: [
    {
      name: 'writer.stream.pipe(destination) returns the destination stream (Node contract)',
      baseline: 'pass',
      async expect(api, assert) {
        const {pipeReturnsDestination} = await api.streamWriterPipeContract();
        assert.strictEqual(
          pipeReturnsDestination,
          true,
          'pipe(dest) must return dest so .pipe(out).on("finish", …) and stream.pipeline compose',
        );
      },
    },
    {
      // Control: the stream still carries the whole package — proving the defect is the return-value
      // contract, not data delivery, so a baseline flip is unambiguous.
      name: 'the piped destination receives the complete workbook bytes (control)',
      baseline: 'pass',
      async expect(api, assert) {
        const {bytes, valid} = await api.streamWriterPipeContract();
        assert.ok(bytes > 0, 'bytes flow through the piped destination');
        assert.strictEqual(valid, true, 'the piped bytes reload as a valid workbook');
      },
    },
  ],
};
