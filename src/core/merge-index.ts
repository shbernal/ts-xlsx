// The index behind the two questions a sheet asks about its merged regions: does a candidate
// rectangle overlap one, and which region covers a given cell.
//
// Both were answered by walking every region on the sheet, which put a linear scan on two paths that
// run once per item. `mergeCells` overlap-checks each new region, and the reader calls it once per
// `<mergeCell>` in the part, so a file's merges cost O(n^2) to load: 20,000 of them took 1.4 seconds
// where 2,000 took 58 ms. Merged regions are cheap to write and capped only by the grid itself, so
// the input that makes that hurt is a few megabytes of XML, and `CLAUDE.md` does not let a parser
// path carry work that grows like that. `getCell` pays the same scan on every call, resolving a
// covered address to its region's master.
//
// The index is one bucket per band of rows, holding the regions whose *top* row falls in it, plus the
// height of the tallest region on the sheet. A region overlapping a query must begin no earlier than
// `query.top - tallest + 1` and no later than `query.bottom`, so only the bands spanning that window
// are visited.
//
// One entry per region and no more, which is the reason the top row alone is bucketed rather than
// every row a region covers. A sheet of whole-column merges is legal, disjoint, and cheap to write,
// and bucketing by covered row would let it claim a bucket entry per row per region: the unbounded
// allocation a hostile-input path must not have. The cost of that choice is the tallest-region
// window, which widens as regions get taller until the scan is the whole sheet again. That is where
// this started, so a sheet of tall regions is never worse off than before, and every other sheet
// (regions in real files are small rectangles clustered in bands) visits one or two buckets.

import {type GridRect, rectsOverlap} from './address.ts';
import type {MergeRect} from './merge.ts';

// Rows per bucket. Wide enough that a table's worth of merged headers shares one bucket, so the
// common query reads a single band; narrow enough that regions spread down a large sheet do not all
// land in the same list.
const BAND_ROWS = 64;

function bandOf(row: number): number {
  return Math.floor(row / BAND_ROWS);
}

/**
 * A row-banded index over a sheet's merged regions. Holds the region list by reference and never
 * mutates it: the sheet stays the owner, reports what it adds, and says when it has rewritten the
 * list wholesale (a splice, an unmerge, a model assignment), which the index answers by rebuilding
 * on the next query rather than by tracking the edit.
 */
export class MergeIndex {
  readonly #rects: readonly MergeRect[];
  #bands = new Map<number, MergeRect[]>();
  #tallest = 0;
  // Starts stale so a list that already holds regions is picked up on the first query.
  #stale = true;

  constructor(rects: readonly MergeRect[]) {
    this.#rects = rects;
  }

  /** Note a region added to the list. Ignored while stale: the rebuild reads the list itself. */
  note(rect: MergeRect): void {
    if (!this.#stale) this.#place(rect);
  }

  /** Report that the region list has been rewritten from outside this index. */
  invalidate(): void {
    this.#stale = true;
  }

  /**
   * The region overlapping `rect`, or `undefined` when the rectangle is free. Which of several
   * overlapping regions is returned is unspecified beyond being one of them; a sheet admits no
   * overlap, so at most one ever applies.
   */
  overlapping(rect: GridRect): MergeRect | undefined {
    this.#refresh();
    const first = bandOf(Math.max(1, rect.top - this.#tallest + 1));
    const last = bandOf(rect.bottom);
    for (let band = first; band <= last; band++) {
      const bucket = this.#bands.get(band);
      if (bucket === undefined) continue;
      for (const candidate of bucket) {
        if (rectsOverlap(candidate, rect)) return candidate;
      }
    }
    return undefined;
  }

  /**
   * Resolve a position to the top-left of the merged region covering it, or to itself when no region
   * does. Only fully-bounded regions participate: an unbounded whole-row/column merge carries no
   * rectangle and so resolves nothing.
   */
  masterOf(row: number, col: number): {row: number; col: number} {
    const covering = this.overlapping({top: row, left: col, bottom: row, right: col});
    return covering === undefined ? {row, col} : {row: covering.top, col: covering.left};
  }

  #place(rect: MergeRect): void {
    const height = rect.bottom - rect.top + 1;
    if (height > this.#tallest) this.#tallest = height;
    const band = bandOf(rect.top);
    const bucket = this.#bands.get(band);
    if (bucket === undefined) this.#bands.set(band, [rect]);
    else bucket.push(rect);
  }

  #refresh(): void {
    if (!this.#stale) return;
    this.#stale = false;
    this.#bands = new Map();
    this.#tallest = 0;
    for (const rect of this.#rects) this.#place(rect);
  }
}
