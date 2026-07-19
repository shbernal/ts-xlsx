// Cluster: styles
//
// Real-world scenario: a user sets a solid pattern fill and, by habit, supplies the ARGB color as a
// CSS-style '#'-prefixed string (e.g. '#FFBFBFBF') instead of the OOXML bare 8-hex-digit form
// ('FFBFBFBF'). The library passes the value through verbatim into the fill color's rgb attribute,
// producing an invalid 9-character color that Excel and LibreOffice cannot parse and fall back to
// solid black on (with auto-contrast text). A valid bare ARGB must serialize as clean 8 hex digits;
// a malformed '#'-prefixed value must be normalized or rejected, never emitted as-is.

/** @typedef {{ name: string, baseline: 'pass'|'fail', expect: (api: any, assert: any) => Promise<void>|void }} Behavior */

export default {
  id: 'solid-fill-argb-rejects-hash-prefix',
  provenance: {source: 'upstream-issue'},
  cluster: 'styles',
  description:
    'A solid fill ARGB is emitted as a clean 8-hex-digit rgb attribute; a value supplied with a ' +
    'leading "#" is normalized or rejected rather than passed through as a malformed 9-character ' +
    'color that renders black in strict consumers.',

  /** @type {Behavior[]} */
  behavior: [
    {
      name: 'a valid bare ARGB serializes as exactly 8 hex digits',
      baseline: 'pass',
      async expect(api, assert) {
        const {validRgb} = await api.fillArgbHashPrefixReport();
        assert.ok(/^[0-9A-Fa-f]{8}$/.test(String(validRgb)), `a valid ARGB must be 8 hex digits; got ${JSON.stringify(validRgb)}`);
      },
    },
    {
      name: 'a "#"-prefixed ARGB is not emitted as a malformed rgb value',
      baseline: 'pass',
      async expect(api, assert) {
        const {hashRgb} = await api.fillArgbHashPrefixReport();
        assert.ok(
          /^[0-9A-Fa-f]{8}$/.test(String(hashRgb)),
          `a '#'-prefixed color must be normalized to a valid 8-hex ARGB (or rejected), not written as ${JSON.stringify(hashRgb)}`
        );
      },
    },
  ],
};
