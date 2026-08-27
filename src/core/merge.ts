// Inclusive grid rectangles and the geometry a worksheet uses to reason about merged regions and the
// `sqref` ranges that overlays (data validations, conditional formats) apply to: overlap detection,
// decoding an OOXML `sqref` into containment rectangles, resolving a covered position to its region's
// master, and collapsing the values a new merge covers.
//
// The last two used to be private methods on `Worksheet`, which split merge geometry across two files:
// the rectangle type and the overlap test lived here, while the two operations that consume them lived
// there. Reasoning about "what does a merge do to the grid" meant reading both. The storage arrives as
// a parameter, so these stay pure functions of the rects and rows handed in.

import {encodeCornerRef, type GridRect, tryDecodeRange} from './address.ts';
import type {Cell} from './cell.ts';
import {isDeletedSpan, shiftIndex} from './grid-shift.ts';

/** A merged region, as the {@link GridRect} every range-shaped thing in the library is. */
export type MergeRect = GridRect;

/**
 * Resolve a position to the master (top-left) of the merged region covering it, or to itself when no
 * region does. First covering region wins; `Worksheet.mergeCells` rejects overlaps, so at most one
 * region ever applies. Only fully-bounded rects participate: an unbounded whole-row/column merge
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
 * anchor, the collapse Excel performs on merge. A leftover covered value would serialise as a
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
 * clamping: a cell anywhere down the column still resolves inside it. An area that names no region
 * that can exist contributes no rectangle, the reader's rule for every other foreign attribute. */
export function decodeSqrefRects(sqref: string): MergeRect[] {
  const rects: MergeRect[] = [];
  for (const part of sqref.split(/\s+/)) {
    if (part === '') continue;
    const decoded = tryDecodeRange(part);
    if (decoded === undefined) continue;
    const {top, left, bottom, right} = decoded;
    rects.push({
      top: top ?? 1,
      left: left ?? 1,
      bottom: bottom ?? Infinity,
      right: right ?? Infinity,
    });
  }
  return rects;
}

/**
 * Re-anchor an OOXML `sqref` through a splice of `count` lines at `start` on `axis`, the inverse of
 * {@link decodeSqrefRects}. Returns `undefined` when the splice deleted every area the `sqref` named:
 * an empty `sqref` is not writable, so the entry holding it goes too.
 *
 * Each space-separated area shifts on its own, and one the splice does not move is returned as the
 * *original text*. That matters for a file the library did not author: `B:B` and `B1:B1048576` decode
 * identically, so a re-encode would rewrite a foreign spelling and cost the byte-clean round trip.
 */
export function shiftSqref(
  sqref: string,
  axis: 'row' | 'col',
  start: number,
  count: number,
  delta: number,
): string | undefined {
  const areas: string[] = [];
  for (const area of sqref.split(/\s+/)) {
    if (area === '') continue;
    const shifted = shiftSqrefArea(area, axis, start, count, delta);
    if (shifted !== undefined) areas.push(shifted);
  }
  return areas.length > 0 ? areas.join(' ') : undefined;
}

function shiftSqrefArea(
  area: string,
  axis: 'row' | 'col',
  start: number,
  count: number,
  delta: number,
): string | undefined {
  const decoded = tryDecodeRange(area);
  // An area this library cannot read is one it cannot move either, and the `sqref` it came from is
  // still the file's own text: return it untouched rather than dropping a region on a guess.
  if (decoded === undefined) return area;
  const {top, left, bottom, right} = decoded;
  const [lo, hi] = axis === 'row' ? [top, bottom] : [left, right];
  // An area unbounded on the spliced axis covers every line of it, so the splice cannot move it: `B:B`
  // after a row insert is still `B:B`, never `B2:B1048577`.
  if (lo === undefined || hi === undefined) return area;
  if (isDeletedSpan(lo, hi, start, count)) return undefined;
  const movedLo = shiftIndex(lo, start, count, delta);
  const movedHi = shiftIndex(hi, start, count, delta);
  if (movedLo === lo && movedHi === hi) return area;
  const [tl, br] =
    axis === 'row'
      ? [encodeCornerRef(left, movedLo), encodeCornerRef(right, movedHi)]
      : [encodeCornerRef(movedLo, top), encodeCornerRef(movedHi, bottom)];
  // A single-cell area stays a single cell: `B5` must not come back as `B6:B6`.
  return area.includes(':') ? `${tl}:${br}` : tl;
}
