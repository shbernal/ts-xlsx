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
