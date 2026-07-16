// The shared-strings table: `xl/sharedStrings.xml`, the workbook-wide pool a `t="s"` cell indexes.
//
// When the writer runs with `useSharedStrings`, a plain string cell value is interned here — an
// identical string is stored once and every cell holding it references the single `<si>` entry by
// index — rather than repeating the text inline in each cell. This trades a second part and an
// indirection for a smaller package when strings repeat, the storage Excel itself prefers.
//
// Only plain string values are pooled. A rich-text value stays inline (`t="inlineStr"`) even under
// the option, so its per-run formatting round-trips through the existing inline path with no reader
// change; moving rich text into the pool as rich `<si>` runs is a later slice.

import {textElement, XML_DECLARATION} from './xml.ts';

const NS_MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';

/**
 * Interns plain string cell values into the shared-strings pool. {@link intern} returns the index a
 * `t="s"` cell writes as its `<v>`; identical strings collapse to one entry, so `count` (total
 * references) and `uniqueCount` (distinct entries) diverge exactly as Excel records them.
 */
export class SharedStringTable {
  readonly #indexByString = new Map<string, number>();
  readonly #entries: string[] = [];
  #references = 0;

  /** Intern a string and return its `<si>` index, reusing the entry when the string repeats. */
  intern(value: string): number {
    this.#references += 1;
    const existing = this.#indexByString.get(value);
    if (existing !== undefined) return existing;
    const index = this.#entries.length;
    this.#indexByString.set(value, index);
    this.#entries.push(value);
    return index;
  }

  /** Whether no string has been interned — the writer omits the part entirely when so. */
  get isEmpty(): boolean {
    return this.#entries.length === 0;
  }

  /** Serialise the pool as the `xl/sharedStrings.xml` part. */
  toXml(): string {
    const items = this.#entries.map(text => `<si>${textElement(text)}</si>`).join('');
    return (
      XML_DECLARATION +
      `<sst xmlns="${NS_MAIN}" count="${this.#references}" uniqueCount="${this.#entries.length}">` +
      items +
      '</sst>'
    );
  }
}
