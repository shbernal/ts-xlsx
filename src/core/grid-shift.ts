// How a grid coordinate moves through a splice, in one place.
//
// Every re-anchoring pass a structural edit runs (the cell grid, line metadata, merges, tables,
// anchored images, shared-formula anchors, the range-bound overlays, the autofilter) answers the same
// two questions about a coordinate: where does it land, and did the delete swallow it outright? Three
// copies of that arithmetic are three places it can drift, and a pass that drifts re-points a rule at
// cells the user never chose. So it lives here and nowhere else.
//
// The grid's last line is part of that arithmetic, which is why the axis is a parameter. An insert
// above a region whose bottom edge is already the last row would otherwise push that edge off the
// grid, and `B1:B1048576` (the bounded spelling of a whole column, and what Excel itself writes)
// becoming `B1:B1048578` is a `sqref` naming rows that cannot exist. Excel does not repair such a
// file quietly: it opens with the "we found a problem with some content" prompt, where the same
// workbook with the edge left inside the grid opens clean. See
// `docs/knowledge/specs/a-splice-must-not-push-geometry-off-the-grid.md` for that measurement and
// for Excel's own answer to the same edit, which is the clamp applied here.

import {MAX_COLUMN, MAX_ROW} from './address.ts';

/**
 * Where a 1-based coordinate on `axis` lands after `count` lines are deleted at `start` and the
 * replacement lines shift what follows by `delta`: before the edit it stays put, at or after the
 * edited span it shifts, and inside a deleted span it clamps to the cut line. The clamp is the best
 * effort for a geometry straddling the cut; a caller that must instead *drop* what a delete swallowed
 * whole tests {@link isDeletedSpan} first.
 *
 * A shift never leaves the grid: the result is clamped to the axis's last line, so a region already
 * touching the bottom (or the right edge) keeps its edge there instead of naming a line the format
 * has no room for. That shrinks such a region by what it could not move, which is the lesser of the
 * two evils and is what Excel does to the same region on the same edit.
 */
export function shiftIndex(
  v: number,
  start: number,
  count: number,
  delta: number,
  axis: 'row' | 'col',
): number {
  const moved = v < start ? v : v >= start + count ? v + delta : start;
  return Math.min(moved, axis === 'row' ? MAX_ROW : MAX_COLUMN);
}

/**
 * Whether the inclusive span `lo..hi` lies entirely within the `count` lines deleted at `start`: the
 * test that separates "this moved" from "this is gone". A single coordinate is the degenerate span
 * `lo === hi`.
 */
export function isDeletedSpan(lo: number, hi: number, start: number, count: number): boolean {
  return lo >= start && hi < start + count;
}
