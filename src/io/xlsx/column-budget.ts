// The ceiling on how much work one sheet's `<col>` elements may cost between them.
//
// Both worksheet readers clamp a single `<col min max>` span to `MAX_COLUMN`, which bounds one
// element's loop and nothing else: the *number* of `<col>` elements is unbounded, and each may span
// the whole grid, so the cost is the product. 154 KB of worksheet XML holding 2,000 full-grid spans
// cost the buffered reader 12.3 s; extrapolated against the 512 MiB inflate ceiling that is hours of
// CPU from a package compressing to a few hundred KB. The zip-bomb guard cannot see it, because after
// inflation the payload really is small: this is a CPU bomb, not an allocation one.
//
// Shared rather than written twice, because the two readers reached the same shape independently and
// the number they agree on is the contract.

import {MAX_COLUMN} from '../../core/address.ts';
import {numInteger} from '../../xml/xml-attrs.ts';
import {type XmlAttributes} from '../../xml/xml-scan.ts';

// Four times the whole grid. Excel writes disjoint spans covering at most `MAX_COLUMN` columns in
// total, so no legitimate file comes near this, while the worst case stays a fraction of a second.
const MAX_COLUMN_RECORD_TOUCHES = MAX_COLUMN * 4;

/**
 * A per-sheet allowance of column touches, spent by each `<col>` span the reader applies.
 *
 * A span past the allowance is truncated, and one starting past it is dropped, rather than the read
 * being refused: that is the stance both readers already take on an out-of-grid span, and it means a
 * hostile file loses formatting it could not have meant while a real one is untouched.
 */
export class ColumnRecordBudget {
  #remaining = MAX_COLUMN_RECORD_TOUCHES;

  /** The last column of `[first, last]` this sheet can still afford, or `undefined` when none is. */
  take(first: number, last: number): number | undefined {
    if (this.#remaining <= 0 || first > last) return undefined;
    const affordable = Math.min(last, first + this.#remaining - 1);
    this.#remaining -= affordable - first + 1;
    return affordable;
  }
}

/**
 * The columns a `<col min max>` element actually applies to, or `undefined` for one that applies to
 * none: unreadable bounds, a span starting past the grid, or a budget already spent.
 *
 * Beside the budget because the two decisions are one decision. Both readers had their own copy, and
 * the copies had already drifted: one tested `min > MAX_COLUMN` and the other did not, which happened
 * not to matter only because `budget.take` returns `undefined` when `first > last` - an accident of
 * this class's contract rather than an agreement between the readers.
 */
export function takeColumnSpan(
  attrs: XmlAttributes,
  budget: ColumnRecordBudget,
): {first: number; last: number} | undefined {
  const first = numInteger(attrs.min, 1);
  const declared = numInteger(attrs.max, 1);
  if (first === undefined || declared === undefined || first > MAX_COLUMN) return undefined;
  // Clamped to the format's ceiling rather than refused: `<col max="99999999">` is a file Excel
  // opens, and an unclamped loop would materialise 16.7 million column records before dying.
  const last = budget.take(first, Math.min(declared, MAX_COLUMN));
  return last === undefined ? undefined : {first, last};
}
