// A package this library did not write carries values it does not model: a preserved part's path and
// content type, the agile-hash credential guarding a sheet, a preserved relationship's target. Every
// one of them is re-emitted verbatim on a round-trip, which makes them the writer's untrusted input.
//
// Two ways to get that wrong, and both were live. Escaping none of them lets an XML special in a
// foreign value close the attribute it sits in: at best the part no longer parses, at worst it injects
// a second `<Override>` for `xl/workbook.xml`. Escaping one of them twice is quieter and just as
// destructive: `s&1.xml` becomes `s&amp;1.xml` in a relationship target, which resolves to a part that
// is not in the package, so the slicer or chart it reached is silently orphaned.

import type {Assert, Case, CorpusApi} from '../case.ts';

export default {
  id: 'foreign-values-are-escaped-once-on-rewrite',
  cluster: 'security',
  provenance: {source: 'write-boundary-audit'},
  description:
    'A part path, content type, relationship target and sheet-protection credential taken off a ' +
    'foreign package and re-emitted on a round-trip are escaped exactly once: the rewritten parts ' +
    'stay parseable and every value still decodes to what the source package held.',
  behavior: [
    {
      name: 'the rewritten content types, worksheet and workbook rels all parse',
      expect(api: CorpusApi, assert: Assert) {
        assert.strictEqual(api.foreignValueRewriteReport().rewritten.wellFormed, true);
      },
    },
    {
      name: 'a preserved part path and content type carrying XML specials survive as themselves',
      expect(api: CorpusApi, assert: Assert) {
        const {source, rewritten} = api.foreignValueRewriteReport();
        assert.strictEqual(rewritten.overridePartName, `/${source.slicerPath}`);
        assert.strictEqual(rewritten.overrideContentType, source.slicerContentType);
      },
    },
    {
      name: 'a preserved relationship target is escaped once, so it still names a part in the package',
      expect(api: CorpusApi, assert: Assert) {
        assert.strictEqual(
          api.foreignValueRewriteReport().rewritten.relTarget,
          'slicerCaches/s&1.xml',
        );
      },
    },
    {
      name: 'a sheet-protection credential round-trips its algorithm name and spin count',
      expect(api: CorpusApi, assert: Assert) {
        const {source, rewritten} = api.foreignValueRewriteReport();
        assert.strictEqual(rewritten.algorithmName, source.algorithmName);
        assert.strictEqual(rewritten.spinCount, '100000');
      },
    },
  ],
} satisfies Case;
