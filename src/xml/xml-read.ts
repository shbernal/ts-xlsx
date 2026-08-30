// A lean, hostile-input-safe SAX pull parser for the read path.
//
// OOXML uses a small, regular subset of XML, so the reader does not need, and must not
// pay for, a general-purpose DOM library (see ADR 0004). This scans the source in a
// single O(n) pass with no recursion, emitting open/text/close events; the OOXML reader
// consumes them and builds only the model, so peak memory tracks real content rather
// than document structure.
//
// Security posture: entities are *decoded, never expanded*. Only the five predefined
// entities and numeric character references are recognised; DTDs and `<!ENTITY>`
// definitions are skipped, so entity-expansion (billion-laughs) and external-entity
// (XXE) attacks are structurally impossible here, not merely mitigated.

import {XmlParseError} from './errors.ts';

export interface XmlAttributes {
  readonly [name: string]: string;
}

export interface SaxHandlers {
  /** An element start. `selfClosing` is true for `<x/>`; no matching {@link onClose} fires for it. */
  onOpen(name: string, attrs: XmlAttributes, selfClosing: boolean): void;
  /** A run of character data (already entity-decoded; CDATA delivered verbatim). Omit to ignore text. */
  onText?(text: string): void;
  /** An element end (`</x>`); the synthetic end of a self-closing element is *not* reported here.
   * Omit to ignore closes. */
  onClose?(name: string): void;
}

/**
 * One parse event from {@link xmlEvents}. The payloads match {@link SaxHandlers} exactly: `text`
 * is already entity-decoded (or verbatim CDATA), and a `<x/>` yields one `open` with
 * `selfClosing: true` and no matching `close`. The discriminated `kind` lets a *pull* consumer
 * drive the parse: the shape the streaming reader needs, where a push callback cannot `yield`.
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
 * `&name;` is left verbatim rather than expanded: there is no DTD, so there is nothing
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

/**
 * The SpreadsheetML `_xHHHH_` escape, in the only place it may appear: a complete cell-text value.
 *
 * The mirror of `escapeSpreadsheetText` in `./xml.ts`, and it sits here rather than beside it for
 * the same reason `decodeEntities` sits apart from `escapeText`: the write helpers carry an
 * `AuthoringError` and a whole serialisation vocabulary the reader has no business importing.
 *
 * **One left-to-right pass, and that is load-bearing.** `005F` maps to `_` like any other code
 * point, with no special case, because a single pass already gives the underscore escape its
 * meaning: in `_x005F_x0041_` the match at 0 yields `_` and scanning resumes at `x0041_`, which has
 * no leading underscore left to start an escape. So the value reads back as the literal seven
 * characters `_x0041_` the author wrote. Decoding `_x005F_` in a pass of its own, before or after
 * the rest, collapses that to `A` and loses the distinction the encoder went to trouble to keep.
 * Excel agrees: it reads that cell as `_x0041_`.
 *
 * The decode is unconditional, not a repair of characters XML cannot carry. Excel reads
 * `a_x0009_b` as a tab even though a literal tab would have been perfectly legal there, so a
 * decoder that only handled the illegal range would disagree with Excel on files Excel wrote.
 */
export function decodeSpreadsheetText(value: string): string {
  if (!value.includes('_')) return value;
  return value.replace(/_x([0-9A-Fa-f]{4})_/g, (_match, hex: string) =>
    String.fromCharCode(parseInt(hex, 16)),
  );
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
  throw new XmlParseError('unterminated tag: missing ">"');
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
  throw new XmlParseError('unterminated markup declaration: missing ">"');
}

/**
 * Scan an XML document as a *pull* stream of {@link XmlEvent}s in a single O(n) pass with no
 * recursion. This is the parser's core; {@link parseXml} is a thin push adapter over it. A
 * consumer that must produce output incrementally (the streaming row reader) pulls events and
 * yields as it goes, holding only its own running state; a push callback cannot.
 *
 * Throws {@link XmlParseError} on malformed markup.
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
      if (end === -1) throw new XmlParseError('unterminated comment');
      i = end + 3;
      continue;
    }
    if (source.startsWith('<![CDATA[', lt)) {
      const end = source.indexOf(']]>', lt + 9);
      if (end === -1) throw new XmlParseError('unterminated CDATA section');
      yield {kind: 'text', text: source.slice(lt + 9, end)};
      i = end + 3;
      continue;
    }
    if (source.startsWith('<?', lt)) {
      const end = source.indexOf('?>', lt + 2);
      if (end === -1) throw new XmlParseError('unterminated processing instruction');
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
 * What {@link elementSubtrees} is to capture: for each container element's local name, the local name
 * of the children to take verbatim inside it (`'dxfs' -> 'dxf'`). Scoping the child to a container is
 * what keeps a `<color>` in `<mruColors>` from being confused with the many other `<color>` elements
 * a stylesheet carries.
 */
export type SubtreeSelection = ReadonlyMap<string, string>;

/** What {@link elementSubtrees} captured: the verbatim source of each selected child, keyed by its
 * container's local name, and each container's own attributes as written. */
export interface SubtreeCapture {
  readonly fragments: ReadonlyMap<string, readonly string[]>;
  readonly attributes: ReadonlyMap<string, XmlAttributes>;
}

/**
 * Capture the verbatim source text of selected elements, in one scan.
 *
 * Some content is re-emitted byte for byte rather than modelled: a differential style, a custom
 * indexed palette, an author's recent-colour swatches, a table-style definition. Preserving the raw
 * text is what keeps a foreign `<dxf>`'s number format a real format code across a re-write instead
 * of a coerced `"[object Object]"`, so the reader needs a subtree's *source*, which an event stream
 * by definition cannot hand back.
 *
 * That gap is why four callers each grew a `<container>([\s\S]*?)</container>` scanner of their own,
 * which is a regular expression parsing XML, over untrusted input, in the same directory as the
 * reader written specifically to avoid that (ADR-0004). This is the same capability done properly:
 * one linear scan with comments, CDATA, processing instructions and declarations skipped as markup
 * rather than matched as text, and the nesting depth counted so a same-named descendant does not end
 * a capture early.
 *
 * Only a container's *first* occurrence is read, matching the single block these documents declare;
 * a second is ignored rather than merged. A captured element that never closes throws
 * {@link XmlParseError}, like the reader's other truncation cases: a partial subtree re-emitted
 * verbatim is broken markup handed on as though it were content.
 */
export function elementSubtrees(source: string, selection: SubtreeSelection): SubtreeCapture {
  const fragments = new Map<string, string[]>();
  const attributes = new Map<string, XmlAttributes>();
  // A container is read once: closed containers are recorded so a later one of the same name is
  // skipped rather than appended to.
  const finished = new Set<string>();
  // The container currently open, the child name to capture inside it, and how deep the same
  // container name is nested within itself.
  let container: {local: string; child: string; depth: number} | undefined;
  // The child element being captured: where its `<` sits, and how deep its own name is nested inside
  // it, so `</name>` for a descendant does not end the capture.
  let capture: {local: string; start: number; depth: number} | undefined;

  const length = source.length;
  let i = 0;
  while (i < length) {
    const lt = source.indexOf('<', i);
    if (lt === -1) break;
    if (source.startsWith('<!--', lt)) {
      const end = source.indexOf('-->', lt + 4);
      if (end === -1) throw new XmlParseError('unterminated comment');
      i = end + 3;
      continue;
    }
    if (source.startsWith('<![CDATA[', lt)) {
      const end = source.indexOf(']]>', lt + 9);
      if (end === -1) throw new XmlParseError('unterminated CDATA section');
      i = end + 3;
      continue;
    }
    if (source.startsWith('<?', lt)) {
      const end = source.indexOf('?>', lt + 2);
      if (end === -1) throw new XmlParseError('unterminated processing instruction');
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
      const local = localName(raw.slice(1).trim());
      if (capture !== undefined && local === capture.local) {
        if (capture.depth > 0) capture.depth -= 1;
        else {
          fragments.get(container?.local ?? '')?.push(source.slice(capture.start, gt + 1));
          capture = undefined;
        }
      } else if (capture === undefined && container !== undefined && local === container.local) {
        if (container.depth > 0) container.depth -= 1;
        else {
          finished.add(container.local);
          container = undefined;
        }
      }
      i = gt + 1;
      continue;
    }

    const selfClosing = raw.charCodeAt(raw.length - 1) === 0x2f;
    const body = selfClosing ? raw.slice(0, -1) : raw;
    const nameEnd = firstWhitespace(body);
    const name = nameEnd === -1 ? body : body.slice(0, nameEnd);
    const local = localName(name);
    if (capture !== undefined) {
      if (!selfClosing && local === capture.local) capture.depth += 1;
    } else if (container !== undefined) {
      if (!selfClosing && local === container.local) container.depth += 1;
      else if (local === container.child) {
        if (selfClosing) fragments.get(container.local)?.push(source.slice(lt, gt + 1));
        else capture = {local, start: lt, depth: 0};
      }
    } else {
      const child = selection.get(local);
      if (child !== undefined && !finished.has(local)) {
        attributes.set(local, nameEnd === -1 ? {} : parseAttributes(body.slice(nameEnd)));
        if (selfClosing) finished.add(local);
        else {
          fragments.set(local, fragments.get(local) ?? []);
          container = {local, child, depth: 0};
        }
      }
    }
    i = gt + 1;
  }

  if (capture !== undefined) {
    throw new XmlParseError(`unterminated <${capture.local}> element`);
  }
  return {fragments, attributes};
}

/** An element start surfaced by {@link openElements}: its qualified `name`, the namespace-stripped
 * `local` name the filter matched on, and its already-decoded `attrs`. */
export interface OpenElement {
  readonly name: string;
  readonly local: string;
  readonly attrs: XmlAttributes;
}

/**
 * Yield each element start in `source` as an {@link OpenElement}, optionally restricted to the given
 * local names. This is the pull shape for the ubiquitous "scan opens, read attributes" pass: a caller
 * writes a plain `for..of` and reads `attrs` directly, instead of threading a mutable accumulator out
 * through a {@link parseXml} `onOpen` closure. With no names every start is yielded; with one or more,
 * only starts whose namespace-stripped name matches one. Text and close events are skipped.
 *
 * Throws {@link XmlParseError} on malformed markup.
 */
export function* openElements(source: string, ...localNames: string[]): Generator<OpenElement> {
  const filter = localNames.length > 0 ? new Set(localNames) : undefined;
  for (const event of xmlEvents(source)) {
    if (event.kind !== 'open') continue;
    const local = localName(event.name);
    if (filter !== undefined && !filter.has(local)) continue;
    yield {name: event.name, local, attrs: event.attrs};
  }
}

/**
 * Wrap an {@link XmlEvent} stream so a self-closing `<x/>` whose local name is in `names` is
 * presented as an open (with `selfClosing: false`) immediately followed by a close: the exact
 * event shape of `<x></x>`. This lets a consumer commit such an element from its close handling
 * alone, instead of hand-coding a parallel self-closing branch: {@link xmlEvents} fires no close
 * for `<x/>`, and forgetting that branch silently drops the empty element. Names not in the set
 * pass through untouched, so an element whose close would wrongly act on absent content (an empty
 * `<v/>`/`<f/>` that must not commit captured text) is left as a bare self-closing open.
 */
export function* closeEmptyElements(
  events: Iterable<XmlEvent>,
  names: ReadonlySet<string>,
): Generator<XmlEvent> {
  for (const event of events) {
    if (event.kind === 'open' && event.selfClosing && names.has(localName(event.name))) {
      yield {kind: 'open', name: event.name, attrs: event.attrs, selfClosing: false};
      yield {kind: 'close', name: event.name};
    } else {
      yield event;
    }
  }
}

/** Options for {@link parseXml}. */
export interface ParseXmlOptions {
  /**
   * Local names whose self-closing form should also fire {@link SaxHandlers.onClose}. For each name
   * listed, `<x/>` is delivered as an open followed by a close (via {@link closeEmptyElements}), so a
   * handler commits the element once in `onClose` rather than duplicating that logic in a self-closing
   * branch. Only name elements the handler is safe to run on close when empty.
   */
  readonly closeEmptyElements?: ReadonlySet<string>;
}

/**
 * One reader's share of a parse: the handlers it wants the events delivered to, and the self-closing
 * elements it needs expanded. Named separately from {@link ParseXmlOptions} because several readers
 * of the same part run over a single parse of it, and each has to bring its own requirements rather
 * than have the caller remember them; see {@link parseXmlPasses}.
 */
export interface SaxPass {
  readonly handlers: SaxHandlers;
  /** As {@link ParseXmlOptions.closeEmptyElements}, for this reader's elements. */
  readonly closeEmptyElements?: ReadonlySet<string>;
}

/** A {@link SaxPass} that gathers something during the parse rather than committing as it goes. */
export interface CollectingPass<T> extends SaxPass {
  /** What the pass collected. Meaningful only once the parse driving it has finished. */
  result(): T;
}

/**
 * Parse `source` once, delivering every event to each pass in turn. The alternative, a parse per
 * reader, costs a full scan of the document per reader and finds nothing in most of them: the
 * worksheet part is the largest in a package, and reading it five times over spent 45% of a large
 * file's read on four scans that matched no element.
 *
 * The expansions the passes ask for are unioned, so *every* pass sees `<x/>` as an open plus a close
 * for any name *any* of them named. That is the one way a pass can observe that it is sharing a
 * parse, and it is why the option is a set of element names rather than a flag: a pass sees an extra
 * close only for elements another pass had to name, and reaching a close for an element a reader
 * does not handle is already the ordinary case.
 */
export function parseXmlPasses(source: string, passes: readonly SaxPass[]): void {
  const expanded = new Set<string>();
  for (const pass of passes) {
    for (const name of pass.closeEmptyElements ?? []) expanded.add(name);
  }
  const handlers = passes.map((pass) => pass.handlers);
  parseXml(
    source,
    {
      onOpen(name, attrs, selfClosing) {
        for (const handler of handlers) handler.onOpen(name, attrs, selfClosing);
      },
      onText(text) {
        for (const handler of handlers) handler.onText?.(text);
      },
      onClose(name) {
        for (const handler of handlers) handler.onClose?.(name);
      },
    },
    expanded.size > 0 ? {closeEmptyElements: expanded} : undefined,
  );
}

/**
 * Parse an XML document, dispatching SAX events to `handlers`. A thin push adapter over
 * {@link xmlEvents}: one scanning core serves both the callback and the pull consumers.
 * Throws {@link XmlParseError} on malformed markup.
 */
export function parseXml(source: string, handlers: SaxHandlers, options?: ParseXmlOptions): void {
  const events = options?.closeEmptyElements
    ? closeEmptyElements(xmlEvents(source), options.closeEmptyElements)
    : xmlEvents(source);
  for (const event of events) {
    switch (event.kind) {
      case 'open':
        handlers.onOpen(event.name, event.attrs, event.selfClosing);
        break;
      case 'text':
        handlers.onText?.(event.text);
        break;
      case 'close':
        handlers.onClose?.(event.name);
        break;
    }
  }
}

// XML end-of-line handling (spec §2.11): a literal CRLF or lone CR in character data is
// normalized to a single LF, so a value's in-cell line breaks read back identically whatever
// newline convention the producer wrote. Normalization precedes entity decoding, so a
// deliberately-encoded carriage return (&#13;) survives it: the escape hatch for a real CR.
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
// There is no fourth variant: a hand-rolled `attr !== '0'` reads `"false"` as true, because
// xsd:boolean spells false both ways and only Excel's own files consistently pick the digit.

/** An OOXML boolean that is on when present with no value (`<b/>` is bold) and off only on an
 * explicit `"0"`/`"false"`; absence reads as on. */
export function boolPresent(val: string | undefined): boolean {
  return val === undefined || (val !== '0' && val !== 'false');
}

/** An OOXML boolean that is on only when explicitly `"1"`/`"true"`; anything else, including
 * absence and a truthy-looking `"0"`, is off. */
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

// The numeric attributes need the same treatment, and for the same reason: a bare `Number(attr)`
// turns a token it cannot parse into `NaN`, which is a number, so every guard downstream passes it
// along until something far from the file throws about a value the caller never wrote. Two readings
// cover the format. An ordinal - a count, an index, a row or a column - is an integer or it is
// nothing (`numInteger`). A measurement - a width, a height, a tint, a margin - is any finite
// number (`numFinite`). Both take the floor the attribute's kind implies, because nearly every call
// site wants "at least 0" or "at least 1" and would otherwise spell it inline and sometimes forget.
// Absent, unparseable and out-of-floor all read as `undefined`, so a caller stores only what the
// source carried and a re-write stays byte-clean.

/** An OOXML integer attribute at or above `min` (default: unbounded below); `undefined` when the
 * attribute is absent, blank, fractional, not a number, or below the floor. Integers past
 * `Number.MAX_SAFE_INTEGER` read as `undefined` too: no index or count is usable out there, and
 * arithmetic on one silently lies. */
export function numInteger(
  val: string | undefined,
  min = -Number.MAX_SAFE_INTEGER,
): number | undefined {
  const n = parseAttrNumber(val);
  if (n === undefined || !Number.isSafeInteger(n) || n < min) return undefined;
  return n;
}

/** An OOXML decimal attribute at or above `min` (default: unbounded below); `undefined` when the
 * attribute is absent, blank, not a number, or below the floor. Infinities are not finite numbers
 * and read as `undefined`. */
export function numFinite(val: string | undefined, min = -Infinity): number | undefined {
  const n = parseAttrNumber(val);
  if (n === undefined || n < min) return undefined;
  return n;
}

// `Number("")` and `Number(" ")` are both 0, which would turn an empty attribute into a real value.
function parseAttrNumber(val: string | undefined): number | undefined {
  if (val === undefined || val.trim() === '') return undefined;
  const n = Number(val);
  return Number.isFinite(n) ? n : undefined;
}

/** Read an operand's text as a number only when it is a canonical decimal literal (optional sign,
 * digits, optional fraction). A cell reference, defined name, expression, or exotically-spelled
 * number (`1E5`, hex) keeps its verbatim text, so it is neither coerced to `NaN` and lost nor
 * re-spelled into a number that would not re-write byte-clean. Callers layer their own type rules
 * (a data-validation `list`/`custom` operand stays a string regardless of what it looks like). */
export function coerceNumericLiteral(text: string): string | number {
  const trimmed = text.trim();
  return /^-?\d+(?:\.\d+)?$/.test(trimmed) ? Number(trimmed) : text;
}

// The third kind of attribute, after the booleans and the numbers, and dropped on the same terms: a
// token from a closed OOXML enumeration. `checkedToken` in `./xml.ts` is the write-side half of the
// same grammar, and the pair is deliberately asymmetric in what it does when the token is foreign.
// The reader drops it, because a file it did not write is allowed to be wrong and losing one
// attribute beats losing the sheet. The writer throws, because a value an author supplied is a
// mistake at the call and the file it would produce is one Excel refuses to open.

/** Narrow an enumerated attribute through its guard; `undefined` when absent or not a member. */
export function enumToken<T extends string>(
  val: string | undefined,
  isMember: (candidate: string) => candidate is T,
): T | undefined {
  return val !== undefined && isMember(val) ? val : undefined;
}

/**
 * Gathers one element's character data across the open/text/close events a SAX parse delivers it in.
 *
 * Nine parsers used to open-code this: latch a flag and clear a buffer on the open, append every
 * chunk while latched, consume the buffer and unlatch on the close. Six spellings of one idea, and
 * none of them honoured the one thing {@link SaxHandlers.onOpen} warns about. A self-closing `<x/>`
 * fires no matching close, so `<t/>`, `<text/>`, `<xm:f/>` and `<totalsRowFormula/>`, all legal and
 * all written by real files, latched a capture that nothing would ever close. What kept that from
 * corrupting anything was the order the next open happened to reset things in, which is an accident
 * rather than a property anyone chose, on a path that reads untrusted input. Taking `selfClosing`
 * here makes it structural, once.
 *
 * The other thing the open-coded versions disagreed on is what an unrelated element opening
 * mid-capture should do. Ending the capture is never what a caller wants: the text belongs to the
 * element that opened it, and a nested or sibling element is not that element. So an open that is
 * not for a captured name leaves an capture in progress alone, and {@link close} answers only for
 * the element that started it.
 *
 * Decoding stays outside. A `<t>` needs `decodeSpreadsheetText` over the whole element and never
 * over a chunk, an `<xm:f>` needs nothing, and a coordinate needs a number: the caller knows which.
 */
export class TextCapture {
  readonly #names: ReadonlySet<string>;
  #capturing: string | undefined;
  #text = '';

  /** @param names the element local name, or the set of names this instance may capture. */
  constructor(names: string | Iterable<string>) {
    this.#names = new Set(typeof names === 'string' ? [names] : names);
  }

  /** Whether a capture is currently open. */
  get capturing(): boolean {
    return this.#capturing !== undefined;
  }

  /** Begin capturing `local` if it is one of this instance's names and is not self-closing. */
  open(local: string, selfClosing: boolean): void {
    if (selfClosing || !this.#names.has(local)) return;
    this.#capturing = local;
    this.#text = '';
  }

  /** Feed a chunk of character data; ignored when no capture is open. */
  text(chunk: string): void {
    if (this.#capturing !== undefined) this.#text += chunk;
  }

  /** The gathered text when `local` closes the captured element, else `undefined`. Unlatches. */
  close(local: string): string | undefined {
    if (this.#capturing !== local) return undefined;
    this.#capturing = undefined;
    return this.#text;
  }
}

/**
 * Yield each named element's text as that element closes, as `{local, text}`.
 *
 * The third member of the pull-shaped family beside {@link openElements} ("scan opens, read
 * attributes") and {@link closeEmptyElements}: this one is "capture these elements' text, tell me
 * each as it closes". A parser whose whole job is reading a handful of text elements out of a part
 * writes a `for..of` over it instead of a {@link parseXml} handler triple whose open and text arms
 * are the same three lines every time.
 *
 * It is deliberately not for every {@link TextCapture} caller. A parser that interleaves capture
 * with per-element state of its own (a `<dataValidation>` gathering formulae, a `<tableColumn>`
 * attaching a totals formula to the column it is inside) needs the open and attribute events too,
 * and stays bespoke; forcing it through here would trade a handler triple for a second pass.
 *
 * A self-closing `<x/>` carries no text and fires no close, so it yields nothing, which is the
 * behaviour {@link TextCapture} exists to make structural rather than a branch each caller
 * remembers.
 *
 * Throws {@link XmlParseError} on malformed markup.
 */
export function* capturedText(
  source: string,
  names: string | Iterable<string>,
): Generator<{local: string; text: string}> {
  const capture = new TextCapture(names);
  for (const event of xmlEvents(source)) {
    if (event.kind === 'open') {
      capture.open(localName(event.name), event.selfClosing);
    } else if (event.kind === 'text') {
      capture.text(event.text);
    } else {
      const local = localName(event.name);
      const text = capture.close(local);
      if (text !== undefined) yield {local, text};
    }
  }
}
