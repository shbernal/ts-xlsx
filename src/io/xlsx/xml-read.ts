// A lean, hostile-input-safe SAX pull parser for the read path.
//
// OOXML uses a small, regular subset of XML, so the reader does not need — and must not
// pay for — a general-purpose DOM library (see ADR 0004). This scans the source in a
// single O(n) pass with no recursion, emitting open/text/close events; the OOXML reader
// consumes them and builds only the model, so peak memory tracks real content rather
// than document structure.
//
// Security posture: entities are *decoded, never expanded*. Only the five predefined
// entities and numeric character references are recognised; DTDs and `<!ENTITY>`
// definitions are skipped, so entity-expansion (billion-laughs) and external-entity
// (XXE) attacks are structurally impossible here, not merely mitigated.

export interface XmlAttributes {
  readonly [name: string]: string;
}

export interface SaxHandlers {
  /** An element start. `selfClosing` is true for `<x/>`; no matching {@link onClose} fires for it. */
  onOpen(name: string, attrs: XmlAttributes, selfClosing: boolean): void;
  /** A run of character data (already entity-decoded; CDATA delivered verbatim). */
  onText(text: string): void;
  /** An element end (`</x>`), or the synthetic end of a self-closing element is *not* reported here. */
  onClose(name: string): void;
}

/**
 * One parse event from {@link xmlEvents}. The payloads match {@link SaxHandlers} exactly: `text`
 * is already entity-decoded (or verbatim CDATA), and a `<x/>` yields one `open` with
 * `selfClosing: true` and no matching `close`. The discriminated `kind` lets a *pull* consumer
 * drive the parse — the shape the streaming reader needs, where a push callback cannot `yield`.
 */
export type XmlEvent =
  | {
      readonly kind: 'open';
      readonly name: string;
      readonly attrs: XmlAttributes;
      readonly selfClosing: boolean;
    }
  | {readonly kind: 'text'; readonly text: string}
  | {readonly kind: 'close'; readonly name: string};

const PREDEFINED_ENTITIES: Readonly<Record<string, string>> = {
  amp: '&',
  lt: '<',
  gt: '>',
  quot: '"',
  apos: "'",
};

const ENTITY = /&(#x[0-9a-fA-F]+|#[0-9]+|[a-zA-Z][a-zA-Z0-9]*);/g;

/**
 * Decode XML character references and the five predefined entities. An unrecognised
 * `&name;` is left verbatim rather than expanded — there is no DTD, so there is nothing
 * to expand it to, and refusing to invent one is what makes entity-expansion attacks
 * impossible.
 */
export function decodeEntities(value: string): string {
  if (!value.includes('&')) return value;
  return value.replace(ENTITY, (match, body: string) => {
    if (body.charCodeAt(0) === 0x23 /* # */) {
      const codePoint =
        body.charCodeAt(1) === 0x78 /* x */
          ? parseInt(body.slice(2), 16)
          : parseInt(body.slice(1), 10);
      if (!Number.isInteger(codePoint) || codePoint < 0 || codePoint > 0x10ffff) return match;
      try {
        return String.fromCodePoint(codePoint);
      } catch {
        return match;
      }
    }
    const named = PREDEFINED_ENTITIES[body];
    return named ?? match;
  });
}

// Attribute values cannot contain their own delimiter and cannot contain a literal `<`,
// so a delimiter-respecting scan finds a tag's end even when an attribute value holds a
// `>` (legal but rare). Names may carry a namespace prefix (`r:id`, `xml:space`).
const ATTRIBUTE = /([^\s=/>]+)\s*=\s*(?:"([^"]*)"|'([^']*)')/g;

function parseAttributes(source: string): XmlAttributes {
  const attrs: Record<string, string> = {};
  ATTRIBUTE.lastIndex = 0;
  let match = ATTRIBUTE.exec(source);
  while (match !== null) {
    const value = match[2] ?? match[3] ?? '';
    attrs[match[1] as string] = decodeEntities(value);
    match = ATTRIBUTE.exec(source);
  }
  return attrs;
}

// Scan to the tag's closing `>`, honouring quoted attribute values so a `>` inside a
// value does not end the tag prematurely.
function findTagEnd(source: string, start: number): number {
  let quote = '';
  for (let i = start + 1; i < source.length; i++) {
    const ch = source[i];
    if (quote !== '') {
      if (ch === quote) quote = '';
    } else if (ch === '"' || ch === "'") {
      quote = ch;
    } else if (ch === '>') {
      return i;
    }
  }
  throw new SyntaxError('unterminated tag: missing ">"');
}

// A `<!DOCTYPE …>` may contain a bracketed internal subset with its own `>`; balance the
// brackets so the declaration is skipped whole. We never act on its contents.
function skipDeclaration(source: string, start: number): number {
  let depth = 0;
  for (let i = start; i < source.length; i++) {
    const ch = source[i];
    if (ch === '[') depth++;
    else if (ch === ']') depth--;
    else if (ch === '>' && depth <= 0) return i + 1;
  }
  throw new SyntaxError('unterminated markup declaration: missing ">"');
}

/**
 * Scan an XML document as a *pull* stream of {@link XmlEvent}s in a single O(n) pass with no
 * recursion. This is the parser's core; {@link parseXml} is a thin push adapter over it. A
 * consumer that must produce output incrementally (the streaming row reader) pulls events and
 * yields as it goes, holding only its own running state — a push callback cannot.
 *
 * Throws {@link SyntaxError} on malformed markup.
 */
export function* xmlEvents(source: string): Generator<XmlEvent> {
  const length = source.length;
  let i = 0;

  while (i < length) {
    const lt = source.indexOf('<', i);
    if (lt === -1) {
      const chunk = source.slice(i);
      if (chunk.length > 0) yield {kind: 'text', text: decodeEntities(normalizeLineEndings(chunk))};
      return;
    }
    if (lt > i) {
      const chunk = source.slice(i, lt);
      if (chunk.length > 0) yield {kind: 'text', text: decodeEntities(normalizeLineEndings(chunk))};
    }

    if (source.startsWith('<!--', lt)) {
      const end = source.indexOf('-->', lt + 4);
      if (end === -1) throw new SyntaxError('unterminated comment');
      i = end + 3;
      continue;
    }
    if (source.startsWith('<![CDATA[', lt)) {
      const end = source.indexOf(']]>', lt + 9);
      if (end === -1) throw new SyntaxError('unterminated CDATA section');
      yield {kind: 'text', text: source.slice(lt + 9, end)};
      i = end + 3;
      continue;
    }
    if (source.startsWith('<?', lt)) {
      const end = source.indexOf('?>', lt + 2);
      if (end === -1) throw new SyntaxError('unterminated processing instruction');
      i = end + 2;
      continue;
    }
    if (source.startsWith('<!', lt)) {
      i = skipDeclaration(source, lt);
      continue;
    }

    const gt = findTagEnd(source, lt);
    const raw = source.slice(lt + 1, gt);
    if (raw.charCodeAt(0) === 0x2f /* / */) {
      yield {kind: 'close', name: raw.slice(1).trim()};
    } else {
      const selfClosing = raw.charCodeAt(raw.length - 1) === 0x2f;
      const body = selfClosing ? raw.slice(0, -1) : raw;
      const nameEnd = firstWhitespace(body);
      const name = nameEnd === -1 ? body : body.slice(0, nameEnd);
      const attrs = nameEnd === -1 ? {} : parseAttributes(body.slice(nameEnd));
      yield {kind: 'open', name, attrs, selfClosing};
    }
    i = gt + 1;
  }
}

/**
 * Parse an XML document, dispatching SAX events to `handlers`. A thin push adapter over
 * {@link xmlEvents} — one scanning core serves both the callback and the pull consumers.
 * Throws {@link SyntaxError} on malformed markup.
 */
export function parseXml(source: string, handlers: SaxHandlers): void {
  for (const event of xmlEvents(source)) {
    switch (event.kind) {
      case 'open':
        handlers.onOpen(event.name, event.attrs, event.selfClosing);
        break;
      case 'text':
        handlers.onText(event.text);
        break;
      case 'close':
        handlers.onClose(event.name);
        break;
    }
  }
}

// XML end-of-line handling (spec §2.11): a literal CRLF or lone CR in character data is
// normalized to a single LF, so a value's in-cell line breaks read back identically whatever
// newline convention the producer wrote. Normalization precedes entity decoding, so a
// deliberately-encoded carriage return (&#13;) survives it — the escape hatch for a real CR.
// CDATA is delivered verbatim (it bypasses this), matching the reader's CDATA contract.
function normalizeLineEndings(chunk: string): string {
  if (!chunk.includes('\r')) return chunk;
  return chunk.replace(/\r\n?/g, '\n');
}

function firstWhitespace(source: string): number {
  for (let i = 0; i < source.length; i++) {
    const code = source.charCodeAt(i);
    if (code === 0x20 || code === 0x09 || code === 0x0a || code === 0x0d) return i;
  }
  return -1;
}

/** Strip a namespace prefix from a qualified name (`r:id` → `id`, `sheet` → `sheet`). */
export function localName(qualified: string): string {
  const colon = qualified.indexOf(':');
  return colon === -1 ? qualified : qualified.slice(colon + 1);
}

// OOXML spells booleans three ways, and the reader needs all three. A `<b/>`-style font flag
// defaults to on when present with no value, so its absence is meaningful (`boolPresent`). Most
// attributes are plain xsd:booleans that are off unless an explicit "1"/"true" turns them on
// (`boolStrict`). An optional attribute that must round-trip byte-clean has to distinguish absent
// from present-and-false and drop an unrecognised token rather than coerce it (`boolTristate`).

/** An OOXML boolean that is on when present with no value (`<b/>` is bold) and off only on an
 * explicit `"0"`/`"false"`; absence reads as on. */
export function boolPresent(val: string | undefined): boolean {
  return val === undefined || (val !== '0' && val !== 'false');
}

/** An OOXML boolean that is on only when explicitly `"1"`/`"true"`; anything else — including
 * absence and a truthy-looking `"0"` — is off. */
export function boolStrict(val: string | undefined): boolean {
  return val === '1' || val === 'true';
}

/** An optional OOXML boolean: `undefined` when the attribute is absent or carries an unrecognised
 * token, otherwise its `"1"`/`"true"` vs `"0"`/`"false"` value. Lets a caller store only the
 * attributes the source actually carried, so a re-write stays byte-clean. */
export function boolTristate(val: string | undefined): boolean | undefined {
  if (val === '1' || val === 'true') return true;
  if (val === '0' || val === 'false') return false;
  return undefined;
}
