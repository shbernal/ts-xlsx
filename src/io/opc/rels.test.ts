import assert from 'node:assert/strict';
import {test} from 'node:test';

import {AuthoringError} from '../../errors.ts';
import {preservedRelsXml, relationship, relationshipsPart, relsPartXml} from './rels.ts';

// `xml.ts` states that escaping is audited in one place rather than sprinkled through the part
// emitters. `.rels` was the exception: the obligation lived in a comment telling the caller to escape
// a target it did not control, so whether a `.rels` part was well-formed depended on which call site
// built it. These pin the guarantee to the function.

test('a writer-controlled package path comes out byte-identical', () => {
  // The no-op case, and the reason moving the escape inside changes no emitted byte: every generated
  // target is already free of the characters an attribute escape touches.
  assert.equal(
    relationship('rId1', 'http://x/worksheet', '../media/image1.jpeg'),
    '<Relationship Id="rId1" Type="http://x/worksheet" Target="../media/image1.jpeg"/>',
  );
});

test('an ampersand, a quote or an angle bracket in any of the three values is escaped', () => {
  const xml = relationship('rId<1>', 'http://x?a=1&b=2', 'https://example.com/?q="a"&r=1', {
    external: true,
  });
  assert.equal(
    xml,
    '<Relationship Id="rId&lt;1&gt;" Type="http://x?a=1&amp;b=2" ' +
      'Target="https://example.com/?q=&quot;a&quot;&amp;r=1" TargetMode="External"/>',
  );
});

test('escaping happens exactly once, so an already-escaped-looking target is not doubled', () => {
  // The failure the caller-side escape would have caused once the function took the job: `a&b`
  // becoming `a&amp;amp;b`, a target that resolves to different bytes than the one asked for.
  assert.equal(
    relationship('rId1', 't', 'a&b'),
    '<Relationship Id="rId1" Type="t" Target="a&amp;b"/>',
  );
});

test('a character XML cannot carry at all is refused rather than emitted', () => {
  // `escapeAttr` carries `assertRepresentable`, so routing through it buys the refusal too: an
  // unescaped interpolation skipped that guard entirely, not only the substitution.
  assert.throws(() => relationship('rId1', 't', 'bad\u0000path'), AuthoringError);
});

test('TargetMode is a fixed token and appears only for an external target', () => {
  assert.match(relationship('rId1', 't', 'x', {external: true}), / TargetMode="External"\/>$/);
  assert.doesNotMatch(relationship('rId1', 't', 'x'), /TargetMode/);
  assert.doesNotMatch(relationship('rId1', 't', 'x', {external: false}), /TargetMode/);
});

test('relationshipsPart wraps its elements in the OPC envelope', () => {
  const xml = relationshipsPart([relationship('rId1', 't', 'x')]);
  assert.match(xml, /^<\?xml /);
  assert.match(xml, /<Relationships xmlns="[^"]+"><Relationship [^>]+\/><\/Relationships>$/);
});

test('a preserved relationship carrying a foreign target is escaped by the part builder', () => {
  // The caller that used to escape by hand. Its output must be unchanged, and unchanged means
  // escaped once.
  const xml = preservedRelsXml([
    {id: 'rId1', type: 'http://x/hyperlink', target: 'https://e.com/?a=1&b=2', external: true},
  ]);
  assert.match(xml, /Target="https:\/\/e\.com\/\?a=1&amp;b=2" TargetMode="External"/);
});

test('a generated part chain emits the same bytes it always did', () => {
  assert.equal(
    relsPartXml([{id: 'rId1', type: 'http://x/pivotCacheDefinition', target: 'pivotCache/x.xml'}]),
    relationshipsPart([
      '<Relationship Id="rId1" Type="http://x/pivotCacheDefinition" Target="pivotCache/x.xml"/>',
    ]),
  );
});
