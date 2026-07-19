import assert from 'node:assert/strict';
import {test} from 'node:test';

import {decodeEntities, localName, parseXml, type XmlAttributes} from './xml-read.ts';

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
