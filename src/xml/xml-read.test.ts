import assert from 'node:assert/strict';
import {test} from 'node:test';

import {
  closeEmptyElements,
  decodeEntities,
  decodeSpreadsheetText,
  localName,
  openElements,
  parseXml,
  type XmlAttributes,
  xmlEvents,
} from './xml-read.ts';
import {escapeSpreadsheetText} from './xml.ts';

interface Event {
  readonly kind: 'open' | 'text' | 'close';
  readonly name?: string;
  readonly attrs?: XmlAttributes;
  readonly selfClosing?: boolean;
  readonly text?: string;
}

function events(source: string): Event[] {
  const out: Event[] = [];
  parseXml(source, {
    onOpen(name, attrs, selfClosing) {
      out.push({kind: 'open', name, attrs, selfClosing});
    },
    onText(text) {
      out.push({kind: 'text', text});
    },
    onClose(name) {
      out.push({kind: 'close', name});
    },
  });
  return out;
}

test('decodeEntities resolves the five predefined entities', () => {
  assert.equal(
    decodeEntities('a &amp; b &lt; c &gt; d &quot; e &apos; f'),
    'a & b < c > d " e \' f',
  );
});

test('decodeEntities resolves decimal and hex character references', () => {
  assert.equal(decodeEntities('&#65;&#x42;&#x1F600;'), 'AB\u{1F600}');
});

test('decodeEntities leaves an unknown named entity verbatim — no DTD, nothing to expand', () => {
  // This is the property that makes entity-expansion (billion-laughs) attacks impossible.
  assert.equal(decodeEntities('&lol;&custom;'), '&lol;&custom;');
});

test('decodeEntities leaves an out-of-range character reference verbatim', () => {
  assert.equal(decodeEntities('&#x110000;'), '&#x110000;');
});

test('parseXml reports open/text/close for a simple element', () => {
  assert.deepEqual(events('<a>hi</a>'), [
    {kind: 'open', name: 'a', attrs: {}, selfClosing: false},
    {kind: 'text', text: 'hi'},
    {kind: 'close', name: 'a'},
  ]);
});

test('parseXml runs an onOpen-only handler, ignoring the omitted onText/onClose', () => {
  const opened: string[] = [];
  parseXml('<a>hi<b/></a>', {
    onOpen(name) {
      opened.push(name);
    },
  });
  assert.deepEqual(opened, ['a', 'b']);
});

test('openElements with no filter yields every start, skipping text and close', () => {
  const seen = [...openElements('<a><b>hi</b><c/></a>')].map((e) => e.name);
  assert.deepEqual(seen, ['a', 'b', 'c']);
});

test('openElements restricts to the given local names, ignoring namespace prefixes', () => {
  const source = '<r><a:sheet name="one"/><x/><b:sheet name="two"/></r>';
  const names = [...openElements(source, 'sheet')].map((e) => e.attrs.name);
  assert.deepEqual(names, ['one', 'two']);
});

test('openElements matches any of several local names and reports the matched local name', () => {
  const source = '<Types><Default Extension="xml"/><Override PartName="/x"/><Other/></Types>';
  assert.deepEqual(
    [...openElements(source, 'Override', 'Default')].map((e) => e.local),
    ['Default', 'Override'],
  );
});

test('openElements surfaces decoded attributes and can be stopped early by breaking', () => {
  let first: XmlAttributes | undefined;
  for (const {attrs} of openElements('<r><a v="x &amp; y"/><a v="z"/></r>', 'a')) {
    first = attrs;
    break;
  }
  assert.equal(first?.v, 'x & y');
});

test('closeEmptyElements expands a named self-closing element to open+close, as if <x></x>', () => {
  const events = [...closeEmptyElements(xmlEvents('<r><c/></r>'), new Set(['c']))];
  assert.deepEqual(
    events.map((e) =>
      e.kind === 'open'
        ? `open:${e.name}:${e.selfClosing}`
        : `${e.kind}:${'name' in e ? e.name : ''}`,
    ),
    ['open:r:false', 'open:c:false', 'close:c', 'close:r'],
  );
});

test('closeEmptyElements matches by local name and leaves un-named self-closing tags bare', () => {
  // `a:c` matches on local name `c` and is expanded; the un-named `d` stays a self-closing open
  // with no synthesized close.
  const events = [...closeEmptyElements(xmlEvents('<a:c/><d/>'), new Set(['c']))];
  assert.deepEqual(
    events.map((e) => (e.kind === 'open' ? `${e.name}:${e.selfClosing}` : e.kind)),
    ['a:c:false', 'close', 'd:true'],
  );
});

test('parseXml with closeEmptyElements fires onClose only for the named self-closing element', () => {
  const opens: string[] = [];
  const closes: string[] = [];
  parseXml(
    '<r><c/><v/></r>',
    {
      onOpen: (name) => opens.push(name),
      onClose: (name) => closes.push(name),
    },
    {closeEmptyElements: new Set(['c'])},
  );
  assert.deepEqual(opens, ['r', 'c', 'v']);
  // `<c/>` is expanded (fires a close); the un-named `<v/>` is not; `<r>` closes normally.
  assert.deepEqual(closes, ['c', 'r']);
});

test('parseXml parses attributes in both quote styles and decodes their entities', () => {
  const [open] = events(`<c r="A1" t='inlineStr' note="a &amp; b"/>`);
  assert.deepEqual(open?.attrs, {r: 'A1', t: 'inlineStr', note: 'a & b'});
  assert.equal(open?.selfClosing, true);
});

test('parseXml tolerates a literal ">" inside a quoted attribute value', () => {
  const [open] = events('<f formula="1 > 0"/>');
  assert.equal(open?.attrs?.formula, '1 > 0');
});

test('parseXml normalizes CRLF and lone CR line endings in text to LF (XML §2.11)', () => {
  const evs = events('<t>a\r\nb\rc\nd</t>');
  assert.deepEqual(
    evs.filter((e) => e.kind === 'text').map((e) => e.text),
    ['a\nb\nc\nd'],
  );
});

test('parseXml preserves a carriage return supplied as a character reference', () => {
  // EOL normalization precedes entity decoding, so &#13; survives as a genuine CR — the escape
  // hatch distinguishing an intended carriage return from a producer's newline convention.
  const evs = events('<t>a&#13;b</t>');
  assert.deepEqual(
    evs.filter((e) => e.kind === 'text').map((e) => e.text),
    ['a\rb'],
  );
});

test('parseXml delivers CDATA verbatim, without entity decoding', () => {
  const evs = events('<t><![CDATA[a & b < c]]></t>');
  assert.deepEqual(
    evs.filter((e) => e.kind === 'text').map((e) => e.text),
    ['a & b < c'],
  );
});

test('parseXml skips comments, processing instructions, and the XML declaration', () => {
  const evs = events('<?xml version="1.0"?><!-- note --><a/>');
  assert.deepEqual(evs, [{kind: 'open', name: 'a', attrs: {}, selfClosing: true}]);
});

test('parseXml skips a DOCTYPE with a bracketed internal subset', () => {
  const evs = events('<!DOCTYPE r [ <!ENTITY x "boom"> ]><r>&x;</r>');
  // The entity definition is ignored; the reference stays literal.
  assert.deepEqual(evs, [
    {kind: 'open', name: 'r', attrs: {}, selfClosing: false},
    {kind: 'text', text: '&x;'},
    {kind: 'close', name: 'r'},
  ]);
});

test('parseXml preserves namespace prefixes on tags and attributes; localName strips them', () => {
  const [open] = events('<w:t xml:space="preserve"> hi </w:t>');
  assert.equal(open?.name, 'w:t');
  assert.equal(localName('w:t'), 't');
  assert.equal(open?.attrs?.['xml:space'], 'preserve');
});

test('parseXml throws on an unterminated tag', () => {
  assert.throws(() => events('<a'), /unterminated tag/);
});

test('parseXml throws on an unterminated comment', () => {
  assert.throws(() => events('<!-- oops'), /unterminated comment/);
});

// The `_xHHHH_` convention, in the direction that reads it. Every expectation here was checked
// against Excel Desktop on this host before it was written down; see
// `docs/knowledge/specs/spreadsheetml-xhhhh-escape-is-decoded-on-read.md`.

test('decodeSpreadsheetText resolves an escape to its character', () => {
  assert.equal(decodeSpreadsheetText('_x0041_'), 'A');
  assert.equal(decodeSpreadsheetText('a_x0001_b'), 'a\u0001b');
});

test('decodeSpreadsheetText decodes a character XML could have carried anyway', () => {
  // Excel does, so a decoder restricted to the illegal range would disagree with it on Excel's
  // own files.
  assert.equal(decodeSpreadsheetText('a_x0009_b'), 'a\tb');
});

test('decodeSpreadsheetText leaves an escaped underscore as literal text', () => {
  // The single left-to-right pass is what makes this work: the match at 0 yields `_` and scanning
  // resumes past it, so the `x0041_` that follows never starts an escape.
  assert.equal(decodeSpreadsheetText('_x005F_x0041_'), '_x0041_');
  assert.equal(decodeSpreadsheetText('a_x005F_x0041_b'), 'a_x0041_b');
});

test('decodeSpreadsheetText leaves an escaped underscore alone when nothing follows it', () => {
  assert.equal(decodeSpreadsheetText('_x005F_'), '_');
});

test('decodeSpreadsheetText leaves a lookalike untouched', () => {
  for (const value of ['_xZZZZ_', '_x041_', '_x00041_', '_', '_x', 'x0041_', '_X0041_']) {
    assert.equal(decodeSpreadsheetText(value), value);
  }
});

test('decodeSpreadsheetText accepts either case of hex digit', () => {
  assert.equal(decodeSpreadsheetText('_x00e9_'), 'é');
  assert.equal(decodeSpreadsheetText('_x00E9_'), 'é');
});

test('decodeSpreadsheetText restores a lone surrogate rather than a replacement character', () => {
  assert.equal(decodeSpreadsheetText('a_xD800_b'), 'a\uD800b');
});

test('decodeSpreadsheetText inverts escapeSpreadsheetText', () => {
  // The escape hands XML's own `&<>` to `escapeText`, so the fixed point is checked on a value
  // holding none of those — what matters here is that the `_xHHHH_` round trip is the identity for
  // every awkward shape at once, including ones that only interact when adjacent.
  const awkward =
    'plain _ _x _x0041_ _x005F_ _xZZZZ_ _x041_ __x0041_ \u0001 \uFFFE \uD800 \t\n \u007F';
  assert.equal(decodeSpreadsheetText(escapeSpreadsheetText(awkward)), awkward);
});
