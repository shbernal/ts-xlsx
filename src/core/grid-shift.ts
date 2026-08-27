// How a grid coordinate moves through a splice, in one place.
//
// Every re-anchoring pass a structural edit runs (the cell grid, line metadata, merges, tables,
// anchored images, shared-formula anchors, the range-bound overlays, the autofilter) answers the same
// two questions about a coordinate: where does it land, and did the delete swallow it outright? Three
// copies of that arithmetic are three places it can drift, and a pass that drifts re-points a rule at
// cells the user never chose. So it lives here and nowhere else.

/**
 * Where a 1-based coordinate lands after `count` lines are deleted at `start` and the replacement
 * lines shift what follows by `delta`: before the edit it stays put, at or after the edited span it
 * shifts, and inside a deleted span it clamps to the cut line. The clamp is the best effort for a
 * geometry straddling the cut; a caller that must instead *drop* what a delete swallowed whole tests
 * {@link isDeletedSpan} first.
 */
export function shiftIndex(v: number, start: number, count: number, delta: number): number {
  return v < start ? v : v >= start + count ? v + delta : start;
}

/**
 * Whether the inclusive span `lo..hi` lies entirely within the `count` lines deleted at `start`: the
 * test that separates "this moved" from "this is gone". A single coordinate is the degenerate span
 * `lo === hi`.
 */
export function isDeletedSpan(lo: number, hi: number, start: number, count: number): boolean {
  return lo >= start && hi < start + count;
}
