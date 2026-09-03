import assert from 'node:assert/strict';
import {test} from 'node:test';

import {decodeSpreadsheetText, numFinite, numInteger} from './xml-attrs.ts';
import {decodeEntities, localName, type XmlAttributes, xmlEvents} from './xml-scan.ts';
import {escapeSpreadsheetText} from './xml.ts';

interface Event {
  readonly kind: 'open' | 'text' | 'close';
  readonly name?: string;
  readonly attrs?: XmlAttributes;
  readonly selfClosing?: boolean;
  readonly text?: string;
}

// The event stream as plain objects. Attributes are spread onto a literal because the parsed map is
// deliberately null-prototype (a file supplies the attribute names), which strict deep equality
// counts as a difference from a literal; the one test that is *about* that prototype reads the
// parser's own map instead.
function events(source: string): Event[] {
  return [...xmlEvents(source)].map((event) =>
    event.kind === 'open'
      ? {
          kind: 'open' as const,
          name: event.name,
          attrs: {...event.attrs},
          selfClosing: event.selfClosing,
        }
      : event.kind === 'text'
        ? {kind: 'text' as const, text: event.text}
        : {kind: 'close' as const, name: event.name},
  );
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

test('decodeEntities leaves an unknown named entity verbatim: no DTD, nothing to expand', () => {
  // This is the property that makes entity-expansion (billion-laughs) attacks impossible.
  assert.equal(decodeEntities('&lol;&custom;'), '&lol;&custom;');
});

test('decodeEntities leaves an out-of-range character reference verbatim', () => {
  assert.equal(decodeEntities('&#x110000;'), '&#x110000;');
});

test('decodeEntities leaves an Object.prototype name verbatim rather than expanding it', () => {
  // The entity table is indexed by a name taken straight out of the file, so a table that
  // inherits Object.prototype hands back a *function* for a dozen attacker-chosen names and
  // stringifies engine source into the model. `&custom;` above proves the miss path; these
  // prove the miss path is reached at all for the names the prototype would otherwise answer.
  for (const name of ['constructor', 'toString', 'hasOwnProperty']) {
    const decoded = decodeEntities(`a&${name};b`);
    assert.equal(typeof decoded, 'string');
    assert.equal(decoded, `a&${name};b`);
  }
});

test('xmlEvents keeps an attribute named for an Object.prototype member out of the prototype chain', () => {
  // Read off the parser's own map rather than through the helper above: the property under test is
  // the map's prototype, and the spread the helper does to compare against literals restores it.
  const [event] = [...xmlEvents('<a constructor="1"/>')];
  const attrs = event?.kind === 'open' ? event.attrs : undefined;
  assert.equal(attrs?.['constructor'], '1');
  assert.equal(attrs?.['toString'], undefined);
});

test('xmlEvents parses attributes in both quote styles and decodes their entities', () => {
  const [open] = events(`<c r="A1" t='inlineStr' note="a &amp; b"/>`);
  assert.deepEqual(open?.attrs, {r: 'A1', t: 'inlineStr', note: 'a & b'});
  assert.equal(open?.selfClosing, true);
});

test('xmlEvents tolerates a literal ">" inside a quoted attribute value', () => {
  const [open] = events('<f formula="1 > 0"/>');
  assert.equal(open?.attrs?.formula, '1 > 0');
});

test('xmlEvents normalizes CRLF and lone CR line endings in text to LF (XML §2.11)', () => {
  const evs = events('<t>a\r\nb\rc\nd</t>');
  assert.deepEqual(
    evs.filter((e) => e.kind === 'text').map((e) => e.text),
    ['a\nb\nc\nd'],
  );
});

test('xmlEvents preserves a carriage return supplied as a character reference', () => {
  // EOL normalization precedes entity decoding, so &#13; survives as a genuine CR: the escape
  // hatch distinguishing an intended carriage return from a producer's newline convention.
  const evs = events('<t>a&#13;b</t>');
  assert.deepEqual(
    evs.filter((e) => e.kind === 'text').map((e) => e.text),
    ['a\rb'],
  );
});

test('xmlEvents delivers CDATA verbatim, without entity decoding', () => {
  const evs = events('<t><![CDATA[a & b < c]]></t>');
  assert.deepEqual(
    evs.filter((e) => e.kind === 'text').map((e) => e.text),
    ['a & b < c'],
  );
});

test('xmlEvents skips comments, processing instructions, and the XML declaration', () => {
  const evs = events('<?xml version="1.0"?><!-- note --><a/>');
  assert.deepEqual(evs, [{kind: 'open', name: 'a', attrs: {}, selfClosing: true}]);
});

test('xmlEvents skips a DOCTYPE with a bracketed internal subset', () => {
  const evs = events('<!DOCTYPE r [ <!ENTITY x "boom"> ]><r>&x;</r>');
  // The entity definition is ignored; the reference stays literal.
  assert.deepEqual(evs, [
    {kind: 'open', name: 'r', attrs: {}, selfClosing: false},
    {kind: 'text', text: '&x;'},
    {kind: 'close', name: 'r'},
  ]);
});

test('xmlEvents preserves namespace prefixes on tags and attributes; localName strips them', () => {
  const [open] = events('<w:t xml:space="preserve"> hi </w:t>');
  assert.equal(open?.name, 'w:t');
  assert.equal(localName('w:t'), 't');
  assert.equal(open?.attrs?.['xml:space'], 'preserve');
});

test('xmlEvents throws on an unterminated tag', () => {
  assert.throws(() => events('<a'), /unterminated tag/);
});

test('xmlEvents throws on an unterminated comment', () => {
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
  // holding none of those. What matters here is that the `_xHHHH_` round trip is the identity for
  // every awkward shape at once, including ones that only interact when adjacent.
  const awkward =
    'plain _ _x _x0041_ _x005F_ _xZZZZ_ _x041_ __x0041_ \u0001 \uFFFE \uD800 \t\n \u007F';
  assert.equal(decodeSpreadsheetText(escapeSpreadsheetText(awkward)), awkward);
});

test('numInteger reads an ordinal, and reads nothing out of anything that is not one', () => {
  for (const [attr, expected] of [
    ['0', 0],
    ['-3', -3],
    ['42', 42],
    [' 7 ', 7],
    ['1e3', 1000],
    [undefined, undefined],
    ['', undefined],
    ['   ', undefined],
    ['abc', undefined],
    ['1.5', undefined],
    ['NaN', undefined],
    ['Infinity', undefined],
    // Past the safe range an "integer" no longer survives its own arithmetic.
    ['9007199254740993', undefined],
  ] as const) {
    assert.equal(numInteger(attr), expected, `numInteger(${JSON.stringify(attr)})`);
  }
});

test('numInteger drops a value below the floor rather than clamping it to it', () => {
  assert.equal(numInteger('0', 0), 0);
  assert.equal(numInteger('-1', 0), undefined);
  assert.equal(numInteger('1', 1), 1);
  assert.equal(numInteger('0', 1), undefined, 'a dropped value leaves the caller its own default');
});

test('numFinite reads a measurement, integer or not, and refuses what is not finite', () => {
  for (const [attr, expected] of [
    ['1.5', 1.5],
    ['0', 0],
    ['-2.25', -2.25],
    ['1e-3', 0.001],
    [undefined, undefined],
    ['', undefined],
    ['pt', undefined],
    ['NaN', undefined],
    ['Infinity', undefined],
    ['-Infinity', undefined],
  ] as const) {
    assert.equal(numFinite(attr), expected, `numFinite(${JSON.stringify(attr)})`);
  }
  assert.equal(numFinite('-0.5', 0), undefined, 'and honours the floor');
  assert.equal(numFinite('9007199254740993', 0), 9007199254740992, 'a measurement may be huge');
});
