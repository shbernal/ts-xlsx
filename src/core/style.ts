// Cell style primitives.
//
// Styles are a large surface in OOXML (fonts, fills, borders, alignment, number
// formats, protection). The rewrite grows them corpus-first; this module currently
// defines only `Font`, which the value model needs because a rich-text run carries
// its own font. Fills, borders, and alignment land with the styles slice.

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

/** Vertical alignment of a font relative to the baseline (super/subscript). */
export type VerticalAlignment = 'superscript' | 'subscript';

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
  readonly vertAlign: VerticalAlignment;
}
