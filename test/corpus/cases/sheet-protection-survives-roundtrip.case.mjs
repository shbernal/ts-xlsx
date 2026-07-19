// Cluster: security
//
// Real-world scenario: an author protects a worksheet with a password and opts to leave sorting and
// autofilter available to end users. A consumer opens that file, reads it, and saves it back out —
// a plain passthrough. The worksheet-level protection must survive that read→write round-trip, or the
// save silently unlocks a sheet the author locked. The subtlety is the password credential: OOXML
// stores it in finished agile form (algorithm, salted iterated hash, salt, spin count) with no
// recoverable plaintext, so the reader cannot re-hash it — it must carry the exact credential back
// out byte-for-byte. And because OOXML inverts the protection booleans ("1" LOCKS an operation, "0"
// PERMITS it), the permissive flags the author chose must round-trip as "0", not be dropped or flipped.

/** @typedef {{ name: string, baseline: 'pass'|'fail', expect: (api: any, assert: any) => Promise<void>|void }} Behavior */

export default {
  id: 'sheet-protection-survives-roundtrip',
  provenance: {source: 'upstream-issue'},
  cluster: 'security',
  description:
    'Worksheet-level protection present in a loaded workbook survives a read→write round-trip: the ' +
    're-written sheet still declares <sheetProtection sheet="1">, preserves the agile password ' +
    'credential (algorithm/hash/salt/spinCount) verbatim, and keeps the permissive flags the author ' +
    'chose — rather than silently unlocking the sheet on a passthrough save.',

  /** @type {Behavior[]} */
  behavior: [
    {
      name: 'the first write emits a password-guarded sheetProtection with the permissive flags honored',
      baseline: 'pass',
      async expect(api, assert) {
        const {first} = await api.sheetProtectionRoundtrip();
        assert.ok(first, 'protecting the sheet writes a <sheetProtection> element');
        assert.strictEqual(first.sheet, '1', 'the sheet itself is protected (sheet="1")');
        assert.ok(first.hashValue, 'a password-derived hashValue is written');
        assert.ok(first.saltValue, 'a saltValue is written');
        assert.strictEqual(first.sort, '0', 'sorting is left permitted (sort="0")');
        assert.strictEqual(first.autoFilter, '0', 'autofilter is left permitted (autoFilter="0")');
      },
    },
    {
      name: 'protection is not silently dropped by a read→write passthrough — the second write still locks the sheet',
      baseline: 'pass',
      async expect(api, assert) {
        const {second} = await api.sheetProtectionRoundtrip();
        assert.ok(second, 'the re-written sheet must still declare <sheetProtection>, not drop it');
        assert.strictEqual(
          second.sheet,
          '1',
          'the reloaded sheet is still protected after a passthrough save',
        );
      },
    },
    {
      name: 'the agile password credential survives the round-trip byte-for-byte',
      baseline: 'pass',
      async expect(api, assert) {
        const {first, second} = await api.sheetProtectionRoundtrip();
        assert.ok(first && second, 'both writes emit protection');
        assert.strictEqual(
          second.algorithmName,
          first.algorithmName,
          'the hash algorithm is preserved',
        );
        assert.strictEqual(
          second.hashValue,
          first.hashValue,
          'the password hash is preserved verbatim (not re-hashed)',
        );
        assert.strictEqual(second.saltValue, first.saltValue, 'the salt is preserved verbatim');
        assert.strictEqual(second.spinCount, first.spinCount, 'the spin count is preserved');
      },
    },
    {
      name: 'the permissive flags survive the round-trip (not inverted back to a lock)',
      async expect(api, assert) {
        const {second} = await api.sheetProtectionRoundtrip();
        assert.ok(second, 'the re-written sheet declares protection');
        assert.strictEqual(second.sort, '0', 'sorting stays permitted after the round-trip');
        assert.strictEqual(
          second.autoFilter,
          '0',
          'autofilter stays permitted after the round-trip',
        );
      },
      baseline: 'pass',
    },
  ],
};
