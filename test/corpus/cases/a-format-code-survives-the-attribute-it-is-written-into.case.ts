// Cluster: styles
//
// Real-world scenario: a custom number format carries a quoted literal with a tab in it, to line a
// currency symbol up in a column. The format code goes into `formatCode="…"`, an XML attribute, and
// XML 1.0 3.3.3 requires a conforming parser to normalise a raw tab, line feed or carriage return in
// *any* attribute value to a single space. So Excel reads back a format the author did not write.
//
// The escape for this attribute is deliberately not `escapeAttr` -- a lone apostrophe stays bare,
// because Excel writes format codes with bare apostrophes and escaping one would change the code on
// every round trip. That one intended divergence had been written as an independently maintained
// list of four replacements, and a list maintained apart drifts: it had silently lost the three
// whitespace escapes as well. Our own reader preserves the raw character, so a round-trip through
// this library returned the tab unchanged and no test built on one could see it.
//
// The rule this locks, and the reason the assertion is on the emitted bytes rather than on a
// round-trip: the file has to carry what the author wrote, whoever reads it.

import type {Assert, Case, CorpusApi} from '../case.ts';

export default {
  id: 'a-format-code-survives-the-attribute-it-is-written-into',
  provenance: {source: 'write-path-audit'},
  cluster: 'styles',
  description:
    'Every character an XML attribute cannot carry raw is escaped in an emitted formatCode -- tab, ' +
    'line feed and carriage return included -- while the apostrophe stays bare, and the code reads ' +
    'back exactly as authored.',

  behavior: [
    {
      name: 'the whitespace an XML parser would normalise away is escaped',
      expect(api: CorpusApi, assert: Assert) {
        const {attribute} = api.formatCodeAttributeEscaping();
        assert.ok(attribute !== null, 'the custom format was emitted');
        assert.ok(attribute.includes('&#9;'), `tab is escaped (got ${attribute})`);
        assert.ok(attribute.includes('&#10;'), `line feed is escaped (got ${attribute})`);
        assert.ok(attribute.includes('&#13;'), `carriage return is escaped (got ${attribute})`);
      },
    },
    {
      name: 'the markup-significant characters are escaped too',
      expect(api: CorpusApi, assert: Assert) {
        const {attribute} = api.formatCodeAttributeEscaping();
        assert.ok(attribute?.includes('&amp;'), 'ampersand');
        assert.ok(attribute?.includes('&lt;'), 'less-than');
        assert.ok(attribute?.includes('&gt;'), 'greater-than');
        assert.ok(attribute?.includes('&quot;'), 'the quote that delimits the attribute');
      },
    },
    {
      name: 'the apostrophe stays bare: Excel writes format codes with it, and so do we',
      expect(api: CorpusApi, assert: Assert) {
        const {attribute} = api.formatCodeAttributeEscaping();
        assert.ok(attribute?.includes("h'i"), `the apostrophe is untouched (got ${attribute})`);
        assert.ok(!attribute?.includes('&apos;'), 'and is not escaped');
      },
    },
    {
      name: 'the code reads back exactly as it was authored',
      expect(api: CorpusApi, assert: Assert) {
        const {readBack, authored} = api.formatCodeAttributeEscaping();
        assert.equal(readBack, authored);
      },
    },
  ],
} satisfies Case;
