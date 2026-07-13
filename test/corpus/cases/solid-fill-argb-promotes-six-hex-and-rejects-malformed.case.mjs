// Cluster: styles
//
// Real-world scenario: a user sets a solid pattern fill and supplies the colour as a 6-hex RGB
// ('00FF00') — the everyday habit of writing a colour without its alpha channel — or, by mistake, as
// a value that is not a colour at all ('12345', 'red'). OOXML's rgb attribute is a bare 8-hex ARGB;
// a 6-char value renders as flat black in Excel, and a malformed value does too, with no error to
// warn the author. A 6-hex RGB must be promoted to an opaque 8-hex ARGB ('FF00FF00'), and a value
// that is neither 6 nor 8 hex digits must be rejected at the API surface rather than written as a
// colour the consumer silently renders black.

/** @typedef {{ name: string, baseline: 'pass'|'fail', expect: (api: any, assert: any) => Promise<void>|void }} Behavior */

export default {
  id: 'solid-fill-argb-promotes-six-hex-and-rejects-malformed',
  provenance: {source: 'upstream-issue'},
  cluster: 'styles',
  description:
    'A solid fill given a 6-hex RGB is promoted to an opaque 8-hex ARGB rather than emitted as a ' +
    '6-char rgb that renders black; a value that is neither 6 nor 8 hex digits is rejected at the ' +
    'API surface rather than passed through as a colour the consumer silently renders black.',

  /** @type {Behavior[]} */
  behavior: [
    {
      name: 'a 6-hex RGB fill colour is promoted to a fully-opaque 8-hex ARGB',
      baseline: 'fail',
      async expect(api, assert) {
        const {sixHexRgb} = await api.argbNormalizationReport();
        assert.strictEqual(
          String(sixHexRgb),
          'FF00FF00',
          `a 6-hex RGB must gain an opaque 'FF' alpha, not be emitted as ${JSON.stringify(sixHexRgb)}`
        );
      },
    },
    {
      name: 'a colour that is neither 6 nor 8 hex digits is rejected, not written',
      baseline: 'fail',
      async expect(api, assert) {
        const {rejectsMalformed} = await api.argbNormalizationReport();
        assert.strictEqual(rejectsMalformed, true, 'a malformed ARGB must be rejected at the API surface, not written as a black-rendering colour');
      },
    },
  ],
};
