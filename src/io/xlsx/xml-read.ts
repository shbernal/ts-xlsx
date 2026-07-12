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
      const codePoint = body.charCodeAt(1) === 0x78 /* x */ ? parseInt(body.slice(2), 16) : parseInt(body.slice(1), 10);
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
  let match: RegExpExecArray | null;
  while ((match = ATTRIBUTE.exec(source)) !== null) {
    const value = match[2] ?? match[3] ?? '';
    attrs[match[1] as string] = decodeEntities(value);
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

/** Parse an XML document, dispatching SAX events. Throws {@link SyntaxError} on malformed markup. */
export function parseXml(source: string, handlers: SaxHandlers): void {
  const length = source.length;
  let i = 0;

  while (i < length) {
    const lt = source.indexOf('<', i);
    if (lt === -1) {
      emitText(source.slice(i), handlers);
      return;
    }
    if (lt > i) emitText(source.slice(i, lt), handlers);

    if (source.startsWith('<!--', lt)) {
      const end = source.indexOf('-->', lt + 4);
      if (end === -1) throw new SyntaxError('unterminated comment');
      i = end + 3;
      continue;
    }
    if (source.startsWith('<![CDATA[', lt)) {
      const end = source.indexOf(']]>', lt + 9);
      if (end === -1) throw new SyntaxError('unterminated CDATA section');
      handlers.onText(source.slice(lt + 9, end));
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
      handlers.onClose(raw.slice(1).trim());
    } else {
      const selfClosing = raw.charCodeAt(raw.length - 1) === 0x2f;
      const body = selfClosing ? raw.slice(0, -1) : raw;
      const nameEnd = firstWhitespace(body);
      const name = nameEnd === -1 ? body : body.slice(0, nameEnd);
      const attrs = nameEnd === -1 ? {} : parseAttributes(body.slice(nameEnd));
      handlers.onOpen(name, attrs, selfClosing);
    }
    i = gt + 1;
  }
}

function emitText(chunk: string, handlers: SaxHandlers): void {
  if (chunk.length > 0) handlers.onText(decodeEntities(normalizeLineEndings(chunk)));
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
