// Cluster: streaming
//
// Real-world scenario: a large workbook is generated with the streaming writer to keep memory
// bounded — rows are committed as produced, then the worksheet and workbook are committed, piping
// into a file or a pass-through sink. The whole-file writer produces a package that opens cleanly,
// but the streaming writer was reported to sometimes emit an archive that spreadsheet applications
// flag as corrupt and can only open via repair: the auxiliary parts (styles, theme, workbook,
// content-types, rels, docProps) come out zero-byte or carry zip-entry CRC values that do not match
// their bytes, while the primary sheet part is intact. Re-zipping the identical bytes with an
// external tool yields a valid package — the fault is in how the streaming writer assembles the zip
// container, not in the XML.
//
// The invariant: the streaming writer must always produce a well-formed zip — every declared part
// present and non-empty, every entry's stored CRC matching its bytes — so the output re-reads
// without any repair step and matches an equivalent whole-file write. Treating the produced bytes
// as an untrusted archive (CRC-checked extraction) is exactly the hostile-input posture the fork
// takes toward its own output.

/** @typedef {{ name: string, baseline: 'pass'|'fail', expect: (api: any, assert: any) => Promise<void>|void }} Behavior */

export default {
  id: 'streaming-writer-produces-valid-zip-package',
  provenance: {source: 'upstream-issue'},
  cluster: 'streaming',
  description:
    'A package assembled by the streaming writer is a valid zip archive: every declared part is ' +
    'present and non-empty, every entry’s CRC matches its bytes, and the output re-reads to the ' +
    'same sheet names and cell values as a whole-file write — no repair step required.',

  /** @type {Behavior[]} */
  behavior: [
    {
      name: 'every declared part is present and non-empty (no zero-byte auxiliary parts)',
      baseline: 'pass',
      async expect(api, assert) {
        const {partCount, emptyParts} = await api.streamWritePackageReport({rows: 50});
        assert.ok(partCount >= 5, `a real package declares several parts, got ${partCount}`);
        assert.deepStrictEqual(emptyParts, [], 'no declared part may be zero bytes');
      },
    },
    {
      name: 'every zip entry’s stored CRC matches its decompressed bytes',
      baseline: 'pass',
      async expect(api, assert) {
        const {crcValid, crcError} = await api.streamWritePackageReport({rows: 50});
        assert.ok(crcValid, `every entry must extract with a matching CRC-32, got: ${crcError}`);
      },
    },
    {
      name: 'the streamed package re-reads to the same sheet names and cell values',
      baseline: 'pass',
      async expect(api, assert) {
        const {reloadOk, reloadError, sheetNames, firstCol} = await api.streamWritePackageReport({
          rows: 50,
        });
        assert.ok(
          reloadOk,
          `the streamed package must re-read without repair, got: ${reloadError}`,
        );
        assert.deepStrictEqual(sheetNames, ['S'], 'the written sheet name survives the round-trip');
        assert.deepStrictEqual(
          firstCol,
          ['r1', 'r2', 'r3'],
          'the streamed cell values re-read faithfully',
        );
      },
    },
  ],
};
