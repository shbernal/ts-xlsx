// The shared-strings table: `xl/sharedStrings.xml`, the workbook-wide pool a `t="s"` cell indexes.
//
// When the writer runs with `useSharedStrings`, a plain string cell value is interned here — an
// identical string is stored once and every cell holding it references the single `<si>` entry by
// index — rather than repeating the text inline in each cell. This trades a second part and an
// indirection for a smaller package when strings repeat, the storage Excel itself prefers.
//
// Both plain strings and rich text are pooled: a plain value becomes a `<si><t>…</t></si>` entry, a
// rich value a `<si><r>…</r>…</si>` entry carrying its per-run formatting — the rich `<si>` runs Excel
// itself writes. A `t="s"` cell then indexes either, and the reader reconstructs the runs, so pooled
// rich text round-trips its formatting rather than flattening to text.

import type {RichTextValue} from '../../core/value.ts';
import {richTextRunsXml} from './rich-text.ts';
import {textElement, XML_DECLARATION} from './xml.ts';

const NS_MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';

/**
 * Interns cell string values into the shared-strings pool. {@link intern} returns the index a
 * `t="s"` cell writes as its `<v>`; identical entries collapse to one, so `count` (total references)
 * and `uniqueCount` (distinct entries) diverge exactly as Excel records them. Each entry is stored as
 * its rendered `<si>` inner XML, which is also its dedup key — a plain string (`<t>…`) and rich runs
 * (`<r>…`) render to distinct markup, so the two kinds never collide in the pool.
 */
export class SharedStringTable {
  readonly #indexByEntry = new Map<string, number>();
  readonly #entries: string[] = [];
  #references = 0;

  /** Intern a plain or rich string and return its `<si>` index, reusing the entry when it repeats. */
  intern(value: string | RichTextValue): number {
    this.#references += 1;
    const inner = typeof value === 'string' ? textElement(value) : richTextRunsXml(value.richText);
    const existing = this.#indexByEntry.get(inner);
    if (existing !== undefined) return existing;
    const index = this.#entries.length;
    this.#indexByEntry.set(inner, index);
    this.#entries.push(inner);
    return index;
  }

  /** Whether no string has been interned — the writer omits the part entirely when so. */
  get isEmpty(): boolean {
    return this.#entries.length === 0;
  }

  /** Serialise the pool as the `xl/sharedStrings.xml` part. */
  toXml(): string {
    const items = this.#entries.map(inner => `<si>${inner}</si>`).join('');
    return (
      XML_DECLARATION +
      `<sst xmlns="${NS_MAIN}" count="${this.#references}" uniqueCount="${this.#entries.length}">` +
      items +
      '</sst>'
    );
  }
}
