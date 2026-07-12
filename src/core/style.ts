// Cell style primitives.
//
// Styles are a large surface in OOXML (fonts, fills, borders, alignment, number
// formats, protection). The rewrite grows them corpus-first; this module models the
// facets landed so far — colours, fills, borders, fonts, alignment, and protection.

/** Underline can be a plain flag or one of Excel's named underline styles. */
export type UnderlineStyle = boolean | 'none' | 'single' | 'double' | 'singleAccounting' | 'doubleAccounting';

/** A colour, expressed as an ARGB hex string (`"FF0000FF"`) or an indexed theme colour. */
export interface Color {
  /** 8-digit ARGB hex, uppercase, no leading `#`. */
  readonly argb?: string;
  /** Index into the workbook theme's colour scheme. */
  readonly theme?: number;
  /** Tint applied to the theme colour, in `[-1, 1]`. */
  readonly tint?: number;
  /** Index into the legacy indexed colour palette. In a solid fill, `bgColor` is the
   *  automatic placeholder `indexed="64"`; the visible colour lives on `fgColor`. */
  readonly indexed?: number;
}

/**
 * Fill pattern kinds, as OOXML's `ST_PatternType` enumerates them. `none` is the
 * absence of a fill; `solid` paints the whole cell with the foreground colour (the
 * common case). The remaining hatch patterns are carried for fidelity on read.
 */
export type FillPatternType =
  | 'none'
  | 'solid'
  | 'gray125'
  | 'darkGray'
  | 'mediumGray'
  | 'lightGray'
  | 'gray0625'
  | 'darkHorizontal'
  | 'darkVertical'
  | 'darkDown'
  | 'darkUp'
  | 'darkGrid'
  | 'darkTrellis'
  | 'lightHorizontal'
  | 'lightVertical'
  | 'lightDown'
  | 'lightUp'
  | 'lightGrid'
  | 'lightTrellis';

/**
 * A pattern fill. For a `solid` fill the visible colour is the pattern *foreground*
 * (`fgColor`) — OOXML's counter-intuitive rule — while `bgColor` is the automatic
 * indexed placeholder. Gradient fills land with a later slice.
 */
export interface PatternFill {
  readonly type: 'pattern';
  readonly pattern: FillPatternType;
  readonly fgColor?: Color;
  readonly bgColor?: Color;
}

/** A cell/row background fill. Only pattern fills exist today; gradients extend this union later. */
export type Fill = PatternFill;

/**
 * Line styles a cell border edge can take, as OOXML's `ST_BorderStyle` enumerates them.
 * `none` is the absence of an edge and is expressed by omitting the edge, not by this value.
 */
export type BorderStyle =
  | 'thin'
  | 'medium'
  | 'thick'
  | 'dashed'
  | 'dotted'
  | 'double'
  | 'hair'
  | 'mediumDashed'
  | 'dashDot'
  | 'mediumDashDot'
  | 'dashDotDot'
  | 'mediumDashDotDot'
  | 'slantDashDot';

/** One edge of a cell border: its line style, and optionally the line colour. */
export interface BorderEdge {
  readonly style: BorderStyle;
  readonly color?: Color;
}

/**
 * A cell's border. Each of the four sides plus the diagonal is an independent edge; an
 * absent edge means that side has no border (it is not rendered), so a cell bordered on
 * one side never implies the other three. `diagonalUp`/`diagonalDown` select which way a
 * present diagonal edge runs.
 */
export interface Border {
  readonly left?: BorderEdge;
  readonly right?: BorderEdge;
  readonly top?: BorderEdge;
  readonly bottom?: BorderEdge;
  readonly diagonal?: BorderEdge;
  readonly diagonalUp?: boolean;
  readonly diagonalDown?: boolean;
}

/** Vertical alignment of a font relative to the baseline (super/subscript). */
export type FontVerticalAlignment = 'superscript' | 'subscript';

/** A font, as it applies to a cell or a single rich-text run. */
export interface Font {
  readonly name: string;
  readonly size: number;
  readonly family: number;
  readonly scheme: 'minor' | 'major' | 'none';
  readonly charset: number;
  readonly color: Color;
  readonly bold: boolean;
  readonly italic: boolean;
  readonly underline: UnderlineStyle;
  readonly strike: boolean;
  readonly outline: boolean;
  readonly vertAlign: FontVerticalAlignment;
}

/** How a cell's content sits horizontally within its bounds, as OOXML's `ST_HorizontalAlignment`
 *  enumerates it. `general` is the type-dependent default (text left, numbers right) and reads
 *  back as no explicit horizontal alignment. */
export type HorizontalAlignment =
  | 'general'
  | 'left'
  | 'center'
  | 'right'
  | 'fill'
  | 'justify'
  | 'centerContinuous'
  | 'distributed';

/** How a cell's content sits vertically within its bounds, as OOXML's `ST_VerticalAlignment`
 *  enumerates it. */
export type VerticalAlignment = 'top' | 'center' | 'bottom' | 'justify' | 'distributed';

/**
 * A cell's alignment. Every facet is optional and independent; an absent facet means the cell
 * takes Excel's default for it. The boolean flags default to off — a cell that never enabled
 * `wrapText`/`shrinkToFit` must never read back with them on. `textRotation` is in degrees
 * (0–180, where 91–180 encodes -1° to -90°); `indent` is a non-negative indent level.
 */
export interface Alignment {
  readonly horizontal?: HorizontalAlignment;
  readonly vertical?: VerticalAlignment;
  readonly textRotation?: number;
  readonly wrapText?: boolean;
  readonly indent?: number;
  readonly shrinkToFit?: boolean;
  readonly readingOrder?: number;
}

/**
 * A cell's protection state, enforced only when the worksheet itself is protected — the flags
 * do nothing on an unprotected sheet. `locked` defaults to TRUE in OOXML (every cell is locked
 * unless told otherwise), so the meaningful, information-carrying state is an explicitly
 * *unlocked* cell (`locked: false`); marking a cell locked merely restates the default and
 * records nothing. `hidden` (defaults false) hides the cell's formula from the formula bar of a
 * protected sheet. A cell that never set either flag reads back with neither.
 */
export interface Protection {
  readonly locked?: boolean;
  readonly hidden?: boolean;
}
