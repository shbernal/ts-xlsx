// Inclusive grid rectangles and the geometry a worksheet uses to reason about merged regions and the
// `sqref` ranges that overlays (data validations, conditional formats) apply to: overlap detection,
// decoding an OOXML `sqref` into containment rectangles, resolving a covered position to its region's
// master, and collapsing the values a new merge covers.
//
// The last two used to be private methods on `Worksheet`, which split merge geometry across two files:
// the rectangle type and the overlap test lived here, while the two operations that consume them lived
// there. Reasoning about "what does a merge do to the grid" meant reading both. The storage arrives as
// a parameter, so these stay pure functions of the rects and rows handed in.

import {decodeRange} from './address.ts';
import type {Cell} from './cell.ts';

/** A merged region as inclusive 1-based grid bounds. */
export interface MergeRect {
  readonly top: number;
  readonly left: number;
  readonly bottom: number;
  readonly right: number;
}

/** Whether two inclusive grid rectangles share at least one cell. */
export function rectsOverlap(a: MergeRect, b: MergeRect): boolean {
  return a.left <= b.right && b.left <= a.right && a.top <= b.bottom && b.top <= a.bottom;
}

/**
 * Resolve a position to the master (top-left) of the merged region covering it, or to itself when no
 * region does. First covering region wins; `Worksheet.mergeCells` rejects overlaps, so at most one
 * region ever applies. Only fully-bounded rects participate — an unbounded whole-row/column merge
 * carries no rect and so resolves nothing.
 */
export function masterOf(
  rects: readonly MergeRect[],
  row: number,
  col: number,
): {row: number; col: number} {
  for (const rect of rects) {
    if (row >= rect.top && row <= rect.bottom && col >= rect.left && col <= rect.right) {
      return {row: rect.top, col: rect.left};
    }
  }
  return {row, col};
}

/**
 * Drop any value already sitting in a merge's covered non-anchor cells, keeping only the top-left
 * anchor — the collapse Excel performs on merge. A leftover covered value would serialise as a
 * populated `<c>` under the range's `<mergeCell>` ref, the geometry that trips Excel's repair prompt.
 * Styles are untouched: a border spanning the merged region rides the covered cells.
 */
export function clearCoveredValues(rows: Map<number, Map<number, Cell>>, rect: MergeRect): void {
  for (let row = rect.top; row <= rect.bottom; row++) {
    const cols = rows.get(row);
    if (cols === undefined) continue;
    for (let col = rect.left; col <= rect.right; col++) {
      if (row === rect.top && col === rect.left) continue;
      const covered = cols.get(col);
      if (covered !== undefined) covered.value = null;
    }
  }
}

/** Decode an OOXML `sqref` (one or more space-separated ranges) into containment rectangles. A whole
 * column or row leaves one axis unbounded, so its missing edges open to `Infinity` rather than
 * clamping — a cell anywhere down the column still resolves inside it. */
export function decodeSqrefRects(sqref: string): MergeRect[] {
  const rects: MergeRect[] = [];
  for (const part of sqref.split(/\s+/)) {
    if (part === '') continue;
    const {top, left, bottom, right} = decodeRange(part);
    rects.push({
      top: top ?? 1,
      left: left ?? 1,
      bottom: bottom ?? Infinity,
      right: right ?? Infinity,
    });
  }
  return rects;
}
