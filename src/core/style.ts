// Cell style primitives.
//
// Styles are a large surface in OOXML (fonts, fills, borders, alignment, number
// formats, protection). The rewrite grows them corpus-first; this module models the
// facets landed so far — colours, fills, borders, fonts, alignment, and protection.

/** Underline can be a plain flag or one of Excel's named underline styles. */
export type UnderlineStyle =
  | boolean
  | 'none'
  | 'single'
  | 'double'
  | 'singleAccounting'
  | 'doubleAccounting';

const NAMED_UNDERLINE_STYLES: ReadonlySet<string> = new Set([
  'none',
  'single',
  'double',
  'singleAccounting',
  'doubleAccounting',
]);

/** Narrow a raw `<u val>` token to a named {@link UnderlineStyle} (the non-boolean members). */
export function isNamedUnderlineStyle(value: string): value is Exclude<UnderlineStyle, boolean> {
  return NAMED_UNDERLINE_STYLES.has(value);
}

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
 * indexed placeholder.
 */
export interface PatternFill {
  readonly type: 'pattern';
  readonly pattern: FillPatternType;
  readonly fgColor?: Color;
  readonly bgColor?: Color;
}

/** A colour stop in a gradient, at a fractional `position` in `[0, 1]` along the gradient axis. */
export interface GradientStop {
  readonly position: number;
  readonly color: Color;
}

/**
 * A gradient fill, as OOXML's `CT_GradientFill`. A `linear` gradient runs at `degree`
 * degrees across the cell; a `path` gradient radiates from an inner rectangle whose
 * insets are `left`/`right`/`top`/`bottom` (each a fraction in `[0, 1]`). The `stops`
 * place colours along the axis; a well-formed gradient names at least two.
 */
export interface GradientFill {
  readonly type: 'gradient';
  readonly gradient: 'linear' | 'path';
  /** Rotation of a `linear` gradient, in degrees. Absent (and meaningless) for `path`. */
  readonly degree?: number;
  /** Inner-rectangle insets of a `path` gradient, each a fraction in `[0, 1]`. */
  readonly left?: number;
  readonly right?: number;
  readonly top?: number;
  readonly bottom?: number;
  readonly stops: readonly GradientStop[];
}

/** A cell/row background fill: a flat pattern or a colour gradient. */
export type Fill = PatternFill | GradientFill;

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

const BORDER_STYLES: ReadonlySet<string> = new Set<BorderStyle>([
  'thin',
  'medium',
  'thick',
  'dashed',
  'dotted',
  'double',
  'hair',
  'mediumDashed',
  'dashDot',
  'mediumDashDot',
  'dashDotDot',
  'mediumDashDotDot',
  'slantDashDot',
]);

/** Narrow a raw border-edge `style` attribute to a known {@link BorderStyle}. */
export function isBorderStyle(value: string): value is BorderStyle {
  return BORDER_STYLES.has(value);
}

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

/** Narrow a raw `<vertAlign val>` token to a known {@link FontVerticalAlignment}. */
export function isFontVerticalAlignment(value: string): value is FontVerticalAlignment {
  return value === 'superscript' || value === 'subscript';
}

/** Narrow a raw `<scheme val>` token to a known font scheme ({@link Font}'s `scheme` member). */
export function isFontScheme(value: string): value is Font['scheme'] {
  return value === 'minor' || value === 'major' || value === 'none';
}

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

const HORIZONTAL_ALIGNMENTS: ReadonlySet<string> = new Set<HorizontalAlignment>([
  'general',
  'left',
  'center',
  'right',
  'fill',
  'justify',
  'centerContinuous',
  'distributed',
]);

/** Narrow a raw `<alignment horizontal>` token to a known {@link HorizontalAlignment}. */
export function isHorizontalAlignment(value: string): value is HorizontalAlignment {
  return HORIZONTAL_ALIGNMENTS.has(value);
}

/** How a cell's content sits vertically within its bounds, as OOXML's `ST_VerticalAlignment`
 *  enumerates it. */
export type VerticalAlignment = 'top' | 'center' | 'bottom' | 'justify' | 'distributed';

const VERTICAL_ALIGNMENTS: ReadonlySet<string> = new Set<VerticalAlignment>([
  'top',
  'center',
  'bottom',
  'justify',
  'distributed',
]);

/** Narrow a raw `<alignment vertical>` token to a known {@link VerticalAlignment}. */
export function isVerticalAlignment(value: string): value is VerticalAlignment {
  return VERTICAL_ALIGNMENTS.has(value);
}

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

/**
 * The six direct-format facets a cell can carry — its fill, number format, font, border, alignment,
 * and protection. Every facet is optional and independent: a cell sets only the facets it overrides
 * and inherits the rest. This one tuple is the unit of style throughout the library, so the
 * interfaces that carry a cell's formatting compose it rather than re-listing the fields — a column,
 * table column, or named style whose facets *default* the cells that leave them unset (see
 * {@link ColumnProperties}, {@link NamedCellStyle}), and a cell's own resolved format. Because they
 * share this type, "add a facet" is a single edit here and the compiler enforces that no read/write
 * path silently drops one — the round-trip symmetry the merge-loss contract depends on.
 */
export interface CellStyle {
  fill?: Fill | undefined;
  numFmt?: string | undefined;
  font?: Partial<Font> | undefined;
  border?: Border | undefined;
  alignment?: Alignment | undefined;
  protection?: Protection | undefined;
}

// One entry per facet, as a Record rather than a bare list so the compiler rejects it the moment a
// facet is added to CellStyle and forgotten here — which would otherwise let the facet-list-driven
// copies below (assignStyleFacets) silently skip the new facet, a merge-loss.
const CELL_STYLE_FACET_KEYS: Record<keyof CellStyle, true> = {
  fill: true,
  numFmt: true,
  font: true,
  border: true,
  alignment: true,
  protection: true,
};

/** The names of the {@link CellStyle} facets, for helpers that copy the tuple facet-by-facet. */
export const CELL_STYLE_FACETS = Object.keys(CELL_STYLE_FACET_KEYS) as (keyof CellStyle)[];

/**
 * Copy each present facet of `source` onto `target`, leaving facets `source` omits untouched — the
 * plain-record counterpart to a cell's `applyCellStyle`, for the {@link CellStyle}-shaped targets a
 * `Cell`'s setters don't reach (a column's cell-defaults, a named style being assembled on read).
 * Driven by {@link CELL_STYLE_FACETS}, so a facet added to the tuple reaches these paths the moment
 * it joins — the same single-point-of-change the cell path gets.
 */
export function assignStyleFacets(target: CellStyle, source: Readonly<CellStyle>): void {
  for (const facet of CELL_STYLE_FACETS) copyFacet(target, source, facet);
}

// A single facet key at a time, so the write's key type is one member (not the whole union) and
// `target[key] = source[key]` typechecks without a cast — the correlated-key access TS can't verify
// when the key is a union.
function copyFacet<K extends keyof CellStyle>(
  target: CellStyle,
  source: Readonly<CellStyle>,
  key: K,
): void {
  const value = source[key];
  if (value !== undefined) target[key] = value;
}

/**
 * A named cell style — the OOXML `cellStyleXfs`/`cellStyles` layer. A spreadsheet applies a built-in
 * or custom style (e.g. "Normal", "Accent1") whose visual facets live in this shared, named layer
 * rather than on each cell's direct format; a cell links to it and inherits any facet the direct
 * format leaves unset. The facets are a cell's own (see {@link CellStyle}); `name` is the style's
 * display name and `builtinId` its Excel gallery index when it is a built-in style.
 */
export type NamedCellStyle = Readonly<CellStyle> & {
  readonly name?: string;
  readonly builtinId?: number;
};
