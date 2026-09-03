// Cell style primitives.
//
// Styles are a large surface in OOXML (fonts, fills, borders, alignment, number
// formats, protection). The rewrite grows them corpus-first; this module models the
// facets landed so far: colours, fills, borders, fonts, alignment, and protection.

import {tokenSet, tokenSetOf} from '../token-set.ts';
import {type ClonePlan, cloneWith} from './clone.ts';
import {type AssertNever, NAMED_STYLE_ID} from './internal.ts';

/** Underline can be a plain flag or one of Excel's named underline styles. */
export type UnderlineStyle =
  | boolean
  | 'none'
  | 'single'
  | 'double'
  | 'singleAccounting'
  | 'doubleAccounting';

/** Narrow a raw `<u val>` token to a named {@link UnderlineStyle} (the non-boolean members). */
export const isNamedUnderlineStyle = tokenSet<Exclude<UnderlineStyle, boolean>>({
  none: true,
  single: true,
  double: true,
  singleAccounting: true,
  doubleAccounting: true,
});

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
 * Parse a colour written in any of the shapes the API accepts into the bare 8-hex ARGB OOXML wants,
 * or `undefined` if it is not one of them.
 *
 * Two conveniences are accepted, and nothing else: a leading `#` is a CSS habit and is stripped
 * (`'#FFBFBFBF'` → `'FFBFBFBF'`), and a 6-hex RGB is promoted with a fully-opaque alpha (`'00FF00'` →
 * `'FF00FF00'`), the common case of a colour written without its alpha channel. Casing is preserved,
 * so a foreign file's lowercase value round-trips as it arrived.
 *
 * This states the grammar once for both directions. What a malformed value *means* differs by
 * direction and is decided by the caller: on read it is foreign data and resolves to nothing, on
 * write it is a caller's bug and throws (see `normalizeArgb` in `io/xlsx/color-xml.ts`). Neither can
 * be a silently half-parsed value, because Excel does not report a malformed `rgb` at all; it
 * renders flat black.
 *
 * `normalizeThemeColor` in `core/theme.ts` asks a similar question and stays separate: a theme
 * slot is `<a:srgbClr val>`, which DrawingML gives no alpha channel, so the two differ in exactly the
 * thing this one exists to add.
 */
export function parseArgb(value: string): string | undefined {
  const hex = value.startsWith('#') ? value.slice(1) : value;
  const argb = hex.length === 6 ? `FF${hex}` : hex;
  return /^[0-9a-fA-F]{8}$/.test(argb) ? argb : undefined;
}

/**
 * Fill pattern kinds, in the order OOXML's `ST_PatternType` enumerates them. `none` is the
 * absence of a fill; `solid` paints the whole cell with the foreground colour (the
 * common case). The remaining hatch patterns are carried for fidelity on read.
 *
 * The order is load-bearing rather than cosmetic, which is why it is stated twice over -- here, and
 * in {@link FILL_PATTERNS_IN_SCHEMA_ORDER}, tied together by a proof. BIFF12 stores a pattern as its
 * *index* into this enumeration, so the binary codec needs the sequence as a value; the doc above
 * this type used to claim schema order while the union put `gray125` and `gray0625` at positions 2
 * and 6, where the schema puts them last, and the binary codec carried a third copy that was right.
 */
export type FillPatternType =
  | 'none'
  | 'solid'
  | 'mediumGray'
  | 'darkGray'
  | 'lightGray'
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
  | 'lightTrellis'
  | 'gray125'
  | 'gray0625';

/**
 * Every {@link FillPatternType}, in `ST_PatternType` order, which is also the order BIFF12's `fls`
 * field indexes: `fls` 1 is `solid`, 2 is `mediumGray`, and so on. Index 0 is `none`, which the
 * model spells as no fill at all rather than as a pattern.
 *
 * One list, consulted by the guard below and by `io/xlsb/read-styles.ts`. It replaced three
 * hand-maintained copies of one enumeration -- the union, the guard's own table, and the binary
 * codec's index array -- of which only the first two were checked against each other, so adding a
 * pattern forced the guard to be updated and let the binary codec silently drop it.
 */
export const FILL_PATTERNS_IN_SCHEMA_ORDER = [
  'none',
  'solid',
  'mediumGray',
  'darkGray',
  'lightGray',
  'darkHorizontal',
  'darkVertical',
  'darkDown',
  'darkUp',
  'darkGrid',
  'darkTrellis',
  'lightHorizontal',
  'lightVertical',
  'lightDown',
  'lightUp',
  'lightGrid',
  'lightTrellis',
  'gray125',
  'gray0625',
] as const satisfies readonly FillPatternType[];

/**
 * The two halves of the proof {@link FILL_PATTERNS_IN_SCHEMA_ORDER} owes, since a list -- unlike the
 * `Record` shape every other token guard here is built from -- can name fewer members than its union
 * without the compiler minding. `satisfies` above covers "no invented name"; this covers "no
 * omission".
 */
export type EveryFillPatternIsOrdered = AssertNever<
  Exclude<FillPatternType, (typeof FILL_PATTERNS_IN_SCHEMA_ORDER)[number]>
>;

/** Narrow a raw `<patternFill patternType>` token to a known {@link FillPatternType}. */
export const isFillPatternType = tokenSetOf<FillPatternType>(FILL_PATTERNS_IN_SCHEMA_ORDER);

/**
 * A pattern fill. For a `solid` fill the visible colour is the pattern *foreground*
 * (`fgColor`), OOXML's counter-intuitive rule, while `bgColor` is the automatic
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

// The three style primitives that reach a defensive copy (here, and through `DifferentialStyle` in a
// conditional-formatting rule) are none of them flat, so none of them can be copied by a spread. A
// gradient carries an array of stops and a font and a border each carry a nested colour; copying one
// with `{...v}` shared exactly those with the caller, so a later mutation of theirs reached into the
// stored rule. Each declares a plan instead, with the proof beside it that the plan names every
// field, so a facet added to any of these types does not compile until its copy is stated.

const GRADIENT_STOP_CLONE: ClonePlan<GradientStop> = {position: 'value', color: 'record'};

/** The proof that {@link GRADIENT_STOP_CLONE} names every field of a gradient stop. */
export type EveryGradientStopFieldIsCloned = AssertNever<
  Exclude<keyof Required<GradientStop>, keyof typeof GRADIENT_STOP_CLONE>
>;

const PATTERN_FILL_CLONE: ClonePlan<PatternFill> = {
  type: 'value',
  pattern: 'value',
  fgColor: 'record',
  bgColor: 'record',
};

/** The proof that {@link PATTERN_FILL_CLONE} names every field of a pattern fill. */
export type EveryPatternFillFieldIsCloned = AssertNever<
  Exclude<keyof Required<PatternFill>, keyof typeof PATTERN_FILL_CLONE>
>;

const GRADIENT_FILL_CLONE: ClonePlan<GradientFill> = {
  type: 'value',
  gradient: 'value',
  degree: 'value',
  left: 'value',
  right: 'value',
  top: 'value',
  bottom: 'value',
  stops: (stops) => stops.map((stop) => cloneWith(stop, GRADIENT_STOP_CLONE)),
};

/** The proof that {@link GRADIENT_FILL_CLONE} names every field of a gradient fill. */
export type EveryGradientFillFieldIsCloned = AssertNever<
  Exclude<keyof Required<GradientFill>, keyof typeof GRADIENT_FILL_CLONE>
>;

/** A defensive deep copy of a fill, whichever of the two shapes it is. The union is dispatched on
 * `type` because the two halves share no field beyond it, so one plan could not describe both. */
export function cloneFill(fill: Fill): Fill {
  return fill.type === 'gradient'
    ? cloneWith(fill, GRADIENT_FILL_CLONE)
    : cloneWith(fill, PATTERN_FILL_CLONE);
}

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

/** Narrow a raw border-edge `style` attribute to a known {@link BorderStyle}. */
export const isBorderStyle = tokenSet<BorderStyle>({
  thin: true,
  medium: true,
  thick: true,
  dashed: true,
  dotted: true,
  double: true,
  hair: true,
  mediumDashed: true,
  dashDot: true,
  mediumDashDot: true,
  dashDotDot: true,
  mediumDashDotDot: true,
  slantDashDot: true,
});

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

const BORDER_EDGE_CLONE: ClonePlan<BorderEdge> = {style: 'value', color: 'record'};

/** The proof that {@link BORDER_EDGE_CLONE} names every field of a border edge. */
export type EveryBorderEdgeFieldIsCloned = AssertNever<
  Exclude<keyof Required<BorderEdge>, keyof typeof BORDER_EDGE_CLONE>
>;

const cloneBorderEdge = (edge: BorderEdge): BorderEdge => cloneWith(edge, BORDER_EDGE_CLONE);

const BORDER_CLONE: ClonePlan<Border> = {
  left: cloneBorderEdge,
  right: cloneBorderEdge,
  top: cloneBorderEdge,
  bottom: cloneBorderEdge,
  diagonal: cloneBorderEdge,
  diagonalUp: 'value',
  diagonalDown: 'value',
};

/** The proof that {@link BORDER_CLONE} names every side of a border. */
export type EveryBorderFieldIsCloned = AssertNever<
  Exclude<keyof Required<Border>, keyof typeof BORDER_CLONE>
>;

/** A defensive deep copy of a border, each present edge and its colour included. */
export function cloneBorder(border: Border): Border {
  return cloneWith(border, BORDER_CLONE);
}

/** Vertical alignment of a font relative to the baseline (super/subscript). */
export type FontVerticalAlignment = 'superscript' | 'subscript';

/** Narrow a raw `<vertAlign val>` token to a known {@link FontVerticalAlignment}. */
export const isFontVerticalAlignment = tokenSet<FontVerticalAlignment>({
  superscript: true,
  subscript: true,
});

/** The theme-font role a `<scheme val>` names: `"minor"`/`"major"` bind the font to whichever
 * face the workbook theme assigns that role, `"none"` leaves it a literal, unbound face. */
export type FontScheme = 'minor' | 'major' | 'none';

/** Narrow a raw `<scheme val>` token to a known {@link FontScheme}. */
export const isFontScheme = tokenSet<FontScheme>({minor: true, major: true, none: true});

/** A font, as it applies to a cell or a single rich-text run. Every facet is optional and
 * independent, like {@link Border}/{@link Alignment}/{@link Protection}: a font sets only the
 * facets it overrides (Excel's own default font backs the rest), so no consumer ever holds every
 * field populated at once. */
export interface Font {
  readonly name?: string;
  readonly size?: number;
  readonly family?: number;
  readonly scheme?: FontScheme;
  readonly charset?: number;
  readonly color?: Color;
  readonly bold?: boolean;
  readonly italic?: boolean;
  readonly underline?: UnderlineStyle;
  readonly strike?: boolean;
  readonly outline?: boolean;
  readonly vertAlign?: FontVerticalAlignment;
}

const FONT_CLONE: ClonePlan<Font> = {
  name: 'value',
  size: 'value',
  family: 'value',
  scheme: 'value',
  charset: 'value',
  color: 'record',
  bold: 'value',
  italic: 'value',
  underline: 'value',
  strike: 'value',
  outline: 'value',
  vertAlign: 'value',
};

/** The proof that {@link FONT_CLONE} names every facet of a font. */
export type EveryFontFieldIsCloned = AssertNever<
  Exclude<keyof Required<Font>, keyof typeof FONT_CLONE>
>;

/** A defensive deep copy of a font, its nested colour included. */
export function cloneFont(font: Font): Font {
  return cloneWith(font, FONT_CLONE);
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

/** Narrow a raw `<alignment horizontal>` token to a known {@link HorizontalAlignment}. */
export const isHorizontalAlignment = tokenSet<HorizontalAlignment>({
  general: true,
  left: true,
  center: true,
  right: true,
  fill: true,
  justify: true,
  centerContinuous: true,
  distributed: true,
});

/** How a cell's content sits vertically within its bounds, as OOXML's `ST_VerticalAlignment`
 *  enumerates it. */
export type VerticalAlignment = 'top' | 'center' | 'bottom' | 'justify' | 'distributed';

/** Narrow a raw `<alignment vertical>` token to a known {@link VerticalAlignment}. */
export const isVerticalAlignment = tokenSet<VerticalAlignment>({
  top: true,
  center: true,
  bottom: true,
  justify: true,
  distributed: true,
});

/**
 * A cell's alignment. Every facet is optional and independent; an absent facet means the cell
 * takes Excel's default for it. The boolean flags default to off, so a cell that never enabled
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
 * How one `<alignment>` facet encodes: which model key it is, what kind of value it carries, and the
 * value that means "default", which the writer omits and the reader drops.
 *
 * Format-blind on purpose. `Alignment` is a core type and the layering gate forbids core importing a
 * serialisation, so the table states what a facet *is* and each codec supplies the reading and the
 * writing off the `kind`. That is where {@link SHEET_PROTECTION_FLAGS} sits and how it is consumed,
 * and it is the shape the BIFF12 codec drives off too: `io/xlsb/read-styles.ts` walks this list and
 * looks each facet up in a `Record` keyed by {@link Alignment}, so the bit layout stays a BIFF12 fact
 * in the BIFF12 codec while the *set* of facets is decided here, once, for both. (That sentence used
 * to read "if alignment ever reaches it". Alignment had reached it; the comment had not noticed, and
 * the binary reader was restating all seven facets and their default-omission rules by hand.)
 *
 * The OOXML attribute name is the model key for all seven facets, so it is not restated here: a
 * second list that is always identical is a second list to keep in step. A future facet whose
 * attribute differs from its key is the one that would have to add the field.
 */
export type AlignmentFacet =
  | {
      readonly key: 'horizontal' | 'vertical';
      readonly kind: 'token';
      /** The enumeration guard, and what to call it in the error when a value fails it. */
      readonly isValid: (value: string) => boolean;
      readonly label: string;
      /** The token that means the OOXML default, expressed by omitting the attribute. */
      readonly omit?: string;
    }
  | {readonly key: 'textRotation' | 'indent' | 'readingOrder'; readonly kind: 'number'}
  | {readonly key: 'wrapText' | 'shrinkToFit'; readonly kind: 'flag'};

/**
 * The seven `<alignment>` facets, declared once. Both directions key off this list, so a facet
 * written but not read (it survives a re-write and vanishes on load) or read but not written (the
 * reverse) is no longer something a reviewer has to notice: {@link EveryAlignmentFacetIsDeclared}
 * makes a facet added to {@link Alignment} and forgotten here a compile error.
 *
 * In ECMA-376 CT_CellAlignment order, which is the order the writer emits.
 */
export const ALIGNMENT_FACETS: readonly AlignmentFacet[] = [
  {
    key: 'horizontal',
    kind: 'token',
    isValid: isHorizontalAlignment,
    label: 'horizontal alignment',
    // `general` is the type-dependent default and is expressed by omitting the attribute.
    omit: 'general',
  },
  {key: 'vertical', kind: 'token', isValid: isVerticalAlignment, label: 'vertical alignment'},
  // ST_TextRotation is an integer, 0 to 180 plus the sentinel 255, so a fractional rotation is not
  // schema-legal. The reader still takes one, deliberately: the model accepts any number an author
  // assigns and the writer emits it, so a reader stricter than the writer would make an authored or
  // foreign 45.5 write out and then vanish on reload. Losing a facet on a round trip is the worse of
  // the two failures, and tightening only this end would be the one that causes it.
  {key: 'textRotation', kind: 'number'},
  {key: 'wrapText', kind: 'flag'},
  {key: 'indent', kind: 'number'},
  {key: 'shrinkToFit', kind: 'flag'},
  {key: 'readingOrder', kind: 'number'},
];

/**
 * Compile-time proof that {@link ALIGNMENT_FACETS} covers every {@link Alignment} facet. A facet
 * added to the type without a table entry resolves this to that facet's name, which does not satisfy
 * `never`, so the error names what is missing.
 */
export type EveryAlignmentFacetIsDeclared = AssertNever<
  Exclude<keyof Alignment, (typeof ALIGNMENT_FACETS)[number]['key']>
>;

/**
 * A cell's protection state, enforced only when the worksheet itself is protected. The flags
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
 * The six direct-format facets a cell can carry: its fill, number format, font, border, alignment,
 * and protection. Every facet is optional and independent: a cell sets only the facets it overrides
 * and inherits the rest. This one tuple is the unit of style throughout the library, so the
 * interfaces that carry a cell's formatting compose it rather than re-listing the fields: a column,
 * table column, or named style whose facets *default* the cells that leave them unset (see
 * {@link ColumnProperties}, {@link NamedCellStyle}), and a cell's own resolved format. Because they
 * share this type, "add a facet" is a single edit here and the compiler enforces that no read/write
 * path silently drops one, the round-trip symmetry the merge-loss contract depends on.
 */
export interface CellStyle {
  fill?: Fill | undefined;
  numFmt?: string | undefined;
  font?: Font | undefined;
  border?: Border | undefined;
  alignment?: Alignment | undefined;
  protection?: Protection | undefined;
}

// One entry per facet, as a Record rather than a bare list so the compiler rejects it the moment a
// facet is added to CellStyle and forgotten here, which would otherwise let the facet-list-driven
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
 * Everything a cell's *formatting* is, which is the six {@link CellStyle} facets plus two that only a
 * cell can carry: the quote-prefix flag and the link to a named cell style.
 *
 * The two extras are on the cell's `xf` record exactly as the six are, and are written and read back
 * exactly as the six are, but they sat outside the tuple, so every copy path driven by the tuple
 * dropped them: a splice, a `duplicateRow`, or a `dst.model = src.model` turned a leading-apostrophe
 * text cell back into an unprefixed one. That is the merge-loss the tuple exists to make impossible,
 * so they join it. Callers that mean "the six shared facets" (a column default, a named style, a
 * `<dxf>`) still say {@link CellStyle}; callers copying a *cell* say this.
 */
export type CellContent = CellStyle & {
  quotePrefix?: boolean | undefined;
  [NAMED_STYLE_ID]?: number | undefined;
};

// One entry per cell-only facet, on the same terms as CELL_STYLE_FACET_KEYS: the Record shape is what
// makes the compiler reject this the moment CellContent gains a facet it does not list.
const CELL_CONTENT_ONLY_KEYS: Record<Exclude<keyof CellContent, keyof CellStyle>, true> = {
  quotePrefix: true,
  [NAMED_STYLE_ID]: true,
};

/**
 * The names of every {@link CellContent} facet, the six shared ones first. Helpers that copy a cell's
 * whole formatting drive off this, so a facet added to either half reaches all of them at once.
 */
export const CELL_CONTENT_FACETS = [
  ...CELL_STYLE_FACETS,
  ...(Reflect.ownKeys(CELL_CONTENT_ONLY_KEYS) as Exclude<keyof CellContent, keyof CellStyle>[]),
] as (keyof CellContent)[];

/**
 * Copy every present facet of a cell's formatting from `source` onto `target`, leaving facets
 * `source` omits untouched: {@link assignStyleFacets} widened to the two cell-only facets. This is
 * what a copy of a *cell* uses, so no structural edit can drop one.
 */
export function assignContentFacets(target: CellContent, source: Readonly<CellContent>): void {
  for (const facet of CELL_CONTENT_FACETS) copyContentFacet(target, source, facet);
}

// One key at a time, for the reason copyFacet is: a correlated-key write the compiler cannot verify
// when the key is the whole union.
function copyContentFacet<K extends keyof CellContent>(
  target: CellContent,
  source: Readonly<CellContent>,
  key: K,
): void {
  const value = source[key];
  if (value !== undefined) target[key] = value;
}

/**
 * Copy each present facet of `source` onto `target`, leaving facets `source` omits untouched: the
 * plain-record counterpart to a cell's `applyCellStyle`, for the {@link CellStyle}-shaped targets a
 * `Cell`'s setters don't reach (a column's cell-defaults, a named style being assembled on read).
 * Driven by {@link CELL_STYLE_FACETS}, so a facet added to the tuple reaches these paths the moment
 * it joins, the same single-point-of-change the cell path gets.
 */
export function assignStyleFacets(target: CellStyle, source: Readonly<CellStyle>): void {
  for (const facet of CELL_STYLE_FACETS) copyFacet(target, source, facet);
}

/**
 * The {@link CellStyle} facets of `source` as a plain tuple of their own, for a source that carries
 * more than the facets (a column's properties also hold width, hidden and outline state). The
 * projection counterpart to {@link assignStyleFacets}, driven by the same list, so a facet added to
 * the tuple reaches a `<col>` style without anyone remembering to widen a literal.
 */
export function pickStyleFacets(source: Readonly<CellStyle>): CellStyle {
  const facets: CellStyle = {};
  assignStyleFacets(facets, source);
  return facets;
}

// A single facet key at a time, so the write's key type is one member (not the whole union) and
// `target[key] = source[key]` typechecks without a cast: the correlated-key access TS can't verify
// when the key is a union.
function copyFacet<K extends keyof CellStyle>(
  target: CellStyle,
  source: Readonly<CellStyle>,
  key: K,
): void {
  const value = source[key];
  if (value !== undefined) target[key] = value;
}
