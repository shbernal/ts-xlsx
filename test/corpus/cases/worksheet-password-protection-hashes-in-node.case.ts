// Cluster: security
//
// Real-world scenario: on a Node.js server (no browser globals, no window.crypto) a caller protects a
// worksheet with a password. Password protection derives a hash from a randomly-generated salt and
// writes an algorithm/hash/salt/spinCount into the sheet's protection element. If the hashing path
// relies on a browser-oriented randomness shim with no Node fallback, it throws "Secure random number
// generation is not supported by this browser" and the protect call fails on the server. The library
// must obtain secure random bytes from the platform runtime and complete: protect resolves, emits a
// well-formed password-derived protection element, honors the requested options, and — because the
// salt is real randomness, not a stub — two protects with the same password produce different salts.

import type {Assert, Case, CorpusApi} from '../case.ts';

export default {
  id: 'worksheet-password-protection-hashes-in-node',
  provenance: {source: 'upstream-issue'},
  cluster: 'security',
  description:
    'Password-protecting a worksheet in a Node runtime resolves without a secure-random error and ' +
    'emits a valid sheetProtection carrying an algorithm, a password-derived hash, a salt, and a spin ' +
    'count; the requested select-locked/unlocked options are reflected; and the salt is real ' +
    'randomness (two protects with the same password differ), not a constant stub.',

  behavior: [
    {
      name: 'protecting with a password in Node does not throw a secure-random error',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const {threw} = await api.worksheetPasswordProtectionReport();
        assert.strictEqual(threw, null, `password protect must succeed under Node; threw ${threw}`);
      },
    },
    {
      name: 'the emitted protection carries an algorithm, hash, salt, and spin count',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const {algorithm, hasHash, hasSalt, spinCount} =
          await api.worksheetPasswordProtectionReport();
        assert.ok(algorithm, 'a hashing algorithm name is written');
        assert.ok(hasHash, 'a password-derived hashValue is written');
        assert.ok(hasSalt, 'a saltValue is written');
        assert.ok(Number(spinCount) > 0, `a positive spinCount is written; got ${spinCount}`);
      },
    },
    {
      name: 'the requested protection options are reflected in the serialized element',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const {selectLockedCells, selectUnlockedCells} =
          await api.worksheetPasswordProtectionReport();
        assert.strictEqual(
          selectLockedCells,
          '1',
          'disallowing locked-cell selection is serialized',
        );
        assert.strictEqual(
          selectUnlockedCells,
          '1',
          'disallowing unlocked-cell selection is serialized',
        );
      },
    },
    {
      name: 'the salt is real randomness — two protects with the same password differ',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const {saltsDiffer} = await api.worksheetPasswordProtectionReport();
        assert.strictEqual(
          saltsDiffer,
          true,
          'secure randomness must be exercised, not stubbed to a constant salt',
        );
      },
    },
  ],
} satisfies Case;
