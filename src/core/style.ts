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
}

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
