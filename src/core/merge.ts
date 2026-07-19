// Inclusive grid rectangles and the geometry a worksheet uses to reason about merged regions and the
// `sqref` ranges that overlays (data validations, conditional formats) apply to: overlap detection and
// decoding an OOXML `sqref` into containment rectangles.

import {decodeRange} from './address.ts';

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
