// How a grid coordinate moves through a splice, in one place.
//
// Every re-anchoring pass a structural edit runs (the cell grid, line metadata, merges, tables,
// anchored images, shared-formula anchors, the range-bound overlays, the autofilter) answers the same
// two questions about a coordinate: where does it land, and did the delete swallow it outright? Three
// copies of that arithmetic are three places it can drift, and a pass that drifts re-points a rule at
// cells the user never chose. So it lives here and nowhere else.
//
// The splice itself is one value for the same reason, and it is the reason `SheetReferences` and
// `ContentTypeInputs` exist a layer up: `start`, `count` and `delta` are three `number`s with no
// order the reader of a call site can check, and the file a transposed pair produces is perfectly
// well-formed -- it just applies the author's rules to somebody else's cells.
//
// The grid's last line is part of that arithmetic, which is why the axis rides along with the rest
// of the splice. An insert above a region whose bottom edge is already the last row would otherwise
// push that edge off the grid, and `B1:B1048576` (the bounded spelling of a whole column, and what
// Excel itself writes) becoming `B1:B1048578` is a `sqref` naming rows that cannot exist. Excel does
// not repair such a file quietly: it opens with the "we found a problem with some content" prompt,
// where the same workbook with the edge left inside the grid opens clean. See
// `docs/knowledge/specs/a-splice-must-not-push-geometry-off-the-grid.md` for that measurement and
// for Excel's own answer to the same edit, which is the clamp applied here.

import {type CellPosition, type GridRect, MAX_COLUMN, MAX_ROW} from './address.ts';

/**
 * One structural edit to an axis: `count` lines removed at the 1-based `start`, with everything
 * after the removed span moved by `delta`. `delta` is the *net* movement, so a pure insert of two
 * lines is `{count: 0, delta: 2}` and a pure delete of three is `{count: 3, delta: -3}`.
 *
 * The four travel as one value because they describe one edit, and because three of them are
 * `number`: passed positionally, two adjacent ones transposed compiles, typechecks, lints, and moves
 * a merge, a dropdown, a highlight or a comment anchor onto cells the author never chose -- silently,
 * since nothing about the resulting file is malformed. A named field cannot be transposed.
 */
export interface AxisSplice {
  /** The axis the lines were spliced on. */
  readonly axis: 'row' | 'col';
  /** 1-based first line of the removed span. */
  readonly start: number;
  /** How many lines were removed. */
  readonly count: number;
  /** Net movement of every line after the removed span. */
  readonly delta: number;
}

/**
 * Where a 1-based coordinate on the spliced axis lands: before the edit it stays put, at or after
 * the edited span it shifts by `delta`, and inside a deleted span it clamps to the cut line. The
 * clamp is the best effort for a geometry straddling the cut; a caller that must instead *drop*
 * what a delete swallowed whole tests {@link isDeletedSpan} first.
 *
 * A shift never leaves the grid: the result is clamped to the axis's last line, so a region already
 * touching the bottom (or the right edge) keeps its edge there instead of naming a line the format
 * has no room for. That shrinks such a region by what it could not move, which is the lesser of the
 * two evils and is what Excel does to the same region on the same edit.
 */
export function shiftIndex(v: number, splice: AxisSplice): number {
  const {start, count, delta} = splice;
  const moved = v < start ? v : v >= start + count ? v + delta : start;
  return Math.min(moved, splice.axis === 'row' ? MAX_ROW : MAX_COLUMN);
}

/**
 * Whether the inclusive span `lo..hi` lies entirely within the deleted lines: the test that
 * separates "this moved" from "this is gone". A single coordinate is the degenerate span
 * `lo === hi`.
 */
export function isDeletedSpan(lo: number, hi: number, splice: AxisSplice): boolean {
  return lo >= splice.start && hi < splice.start + splice.count;
}

/**
 * Both edges of a region's span on the spliced axis, moved together, or `undefined` when the delete
 * swallowed the span whole. The two answers are exclusive on purpose: a span straddling the cut
 * clamps and survives, one wholly inside it is dropped, and a caller must not clamp its way out of
 * a drop.
 *
 * The one-axis form, for a region whose *other* axis is not a number the caller holds -- a `sqref`
 * area may be unbounded across the axis it is not being spliced on, and re-encoding it as bounded
 * would rewrite the file's own spelling. Callers holding a real rectangle want {@link shiftRect}.
 */
export function shiftSpan(
  lo: number,
  hi: number,
  splice: AxisSplice,
): {lo: number; hi: number} | undefined {
  if (isDeletedSpan(lo, hi, splice)) return undefined;
  return {lo: shiftIndex(lo, splice), hi: shiftIndex(hi, splice)};
}

/**
 * A rectangle re-anchored through the splice, or `undefined` when the delete swallowed it whole.
 * The unspliced axis passes through untouched; both edges of the spliced axis move and clamp.
 *
 * This projection -- pick the spliced axis's two edges out of the rectangle, ask whether they
 * survived, move them, put the rectangle back together -- is the shape every range-bound thing in
 * the model re-anchors by, and writing it per caller is how the axis ternary gets inverted in one
 * of them.
 */
export function shiftRect(rect: GridRect, splice: AxisSplice): GridRect | undefined {
  const rowAxis = splice.axis === 'row';
  const moved = shiftSpan(
    rowAxis ? rect.top : rect.left,
    rowAxis ? rect.bottom : rect.right,
    splice,
  );
  if (moved === undefined) return undefined;
  return rowAxis
    ? {...rect, top: moved.lo, bottom: moved.hi}
    : {...rect, left: moved.lo, right: moved.hi};
}

/**
 * The degenerate one-coordinate form of {@link shiftRect}: a cell that moves, or `undefined` when
 * the delete took the line it sat on. The coordinate off the spliced axis passes through.
 */
export function shiftPoint(point: CellPosition, splice: AxisSplice): CellPosition | undefined {
  const line = splice.axis === 'row' ? point.row : point.col;
  if (isDeletedSpan(line, line, splice)) return undefined;
  const moved = shiftIndex(line, splice);
  return splice.axis === 'row' ? {col: point.col, row: moved} : {col: moved, row: point.row};
}
