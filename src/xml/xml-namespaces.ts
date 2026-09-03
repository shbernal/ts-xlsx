// Which namespace an element or attribute is actually in, rather than which prefix it happens to use.
//
// The reader is prefix-agnostic almost everywhere: `localName()` strips whatever prefix a file chose
// and matching happens on the bare name. That is what makes the places it is *not* dangerous. A
// handful of paths test a prefix string instead of a namespace, and a file that binds the same
// namespace to a different prefix, which is legal and which real toolchains emit, loses a whole
// feature with no error at all:
//
//   - `attrs['r:id']` finds nothing in a workbook that binds the relationships namespace to `rel:`,
//     so every sheet loads permanently empty;
//   - `name.includes(':')` used to mean "this element is in the x14 extension namespace" is true of
//     *every* element in a worksheet that binds the main SpreadsheetML namespace to a prefix, so
//     every data validation and conditional format is discarded as an unknown extension.
//
// A separate module from `xml-scan.ts` deliberately. `/customui` is a ribbon reader that imports the
// scanner and needs none of this, and while the traversals lived in the scanner's module every helper
// added to it was charged to that entry's bundle budget, which is how it once drifted over. The same
// reasoning applies to the same file, so this lives beside it rather than in it.
//
// It is driven by the reader's own `onOpen`/`onClose`, not by a change to the scanner: the handlers
// already receive every element's attributes, which is where `xmlns` declarations are.

import type {XmlAttributes} from './xml-scan.ts';

const XMLNS = 'xmlns';
const XMLNS_PREFIX = 'xmlns:';

/**
 * The prefix-to-namespace bindings in force at the current point of a scan.
 *
 * Properly scoped: a declaration binds only within the element that carries it, so the bindings are
 * pushed and popped with the element stack. In practice OOXML puts them all on the part's root, but
 * a reader that assumed so would be making the same kind of assumption this class exists to remove.
 */
export class NamespaceScope {
  // The bindings in force, prefix → URI, with `''` as the key for the default (unprefixed) namespace.
  readonly #bindings = new Map<string, string>();
  // What each open element added, so closing it can put back exactly what it shadowed. The entry for
  // an element declaring nothing is `undefined` rather than an empty array, and the pairs are only
  // allocated when there is something to record: this runs once per element on the worksheet body,
  // which is the largest part in a package, and every element there declares nothing.
  readonly #undo: ([string, string | undefined][] | undefined)[] = [];

  /** Enter an element, taking any `xmlns` declarations it carries. */
  open(attrs: XmlAttributes): void {
    let undo: [string, string | undefined][] | undefined;
    for (const name in attrs) {
      // `xmlns` and `xmlns:…` are the only two spellings, and the second is far the commoner, so the
      // cheap prefix test comes first and the exact-match one only runs when it fails.
      const prefix = name.startsWith(XMLNS_PREFIX)
        ? name.slice(XMLNS_PREFIX.length)
        : name === XMLNS
          ? ''
          : null;
      if (prefix === null) continue;
      (undo ??= []).push([prefix, this.#bindings.get(prefix)]);
      this.#bindings.set(prefix, attrs[name] ?? '');
    }
    this.#undo.push(undo);
  }

  /** Leave an element, restoring whatever its declarations shadowed. */
  close(): void {
    const undo = this.#undo.pop();
    if (undo === undefined) return;
    for (const [prefix, previous] of undo) {
      if (previous === undefined) this.#bindings.delete(prefix);
      else this.#bindings.set(prefix, previous);
    }
  }

  /**
   * The namespace URI a qualified *element* name resolves to, or `undefined` when its prefix is
   * unbound. An unprefixed element takes the default namespace, which is what makes
   * `<worksheet xmlns="…main">` and `<x:worksheet xmlns:x="…main">` the same document.
   */
  elementNamespace(qualified: string): string | undefined {
    const colon = qualified.indexOf(':');
    return this.#bindings.get(colon === -1 ? '' : qualified.slice(0, colon));
  }

  /** Whether a qualified element name is in `uri`. */
  isElementIn(qualified: string, uri: string): boolean {
    return this.elementNamespace(qualified) === uri;
  }

  /**
   * An attribute's value looked up by namespace and local name rather than by qualified name.
   *
   * An *unprefixed* attribute is in no namespace at all, never the default one, which is why this
   * only ever matches a prefixed spelling. Every namespaced attribute OOXML uses (`r:id`, `r:embed`,
   * `xml:space`) is written with a prefix for exactly that reason.
   */
  attr(attrs: XmlAttributes, uri: string, local: string): string | undefined {
    for (const [prefix, bound] of this.#bindings) {
      if (prefix === '' || bound !== uri) continue;
      const value = attrs[`${prefix}:${local}`];
      if (value !== undefined) return value;
    }
    return undefined;
  }
}
