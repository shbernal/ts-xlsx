// A custom table style: the named, reusable look a table (or a pivot) applies to itself by name.
//
// A table style is a list of *elements*, each naming one region of the table — the whole table, the
// header row, the first row stripe — and the differential formatting to lay over it. Excel's built-in
// gallery ("TableStyleMedium2" and its sixty siblings) is exactly this shape; a workbook that declares
// its own joins the gallery for that file, and a table reaches it by name through
// `TableStyleInfo.name`.

import type {DifferentialStyle} from './style.ts';

/**
 * The regions a table style can format (`ST_TableStyleType`).
 *
 * The first thirteen apply to a **table**; the rest style a **pivot table**, which has regions a
 * table does not have (subtotal rows, page-field labels, subheadings). Both live in the same
 * enumeration and the same `<tableStyle>` element — what decides which regions a consumer honours is
 * the style's own `table`/`pivot` flags, not the element names — so the type carries all of them
 * rather than splitting into two enumerations that a caller would have to choose between up front.
 */
export const TABLE_STYLE_ELEMENT_TYPES = [
  'wholeTable',
  'headerRow',
  'totalRow',
  'firstColumn',
  'lastColumn',
  'firstRowStripe',
  'secondRowStripe',
  'firstColumnStripe',
  'secondColumnStripe',
  'firstHeaderCell',
  'lastHeaderCell',
  'firstTotalCell',
  'lastTotalCell',
  'firstSubtotalColumn',
  'secondSubtotalColumn',
  'thirdSubtotalColumn',
  'firstSubtotalRow',
  'secondSubtotalRow',
  'thirdSubtotalRow',
  'blankRow',
  'firstColumnSubheading',
  'secondColumnSubheading',
  'thirdColumnSubheading',
  'firstRowSubheading',
  'secondRowSubheading',
  'thirdRowSubheading',
  'pageFieldLabels',
  'pageFieldValues',
] as const;

/** One region of a table or pivot that a table style can format. */
export type TableStyleElementType = (typeof TABLE_STYLE_ELEMENT_TYPES)[number];

/** The four element types banded across several rows or columns — the only ones {@link TableStyleElement.size} means anything on. */
export const STRIPE_ELEMENT_TYPES: ReadonlySet<TableStyleElementType> = new Set([
  'firstRowStripe',
  'secondRowStripe',
  'firstColumnStripe',
  'secondColumnStripe',
]);

export function isTableStyleElementType(value: string): value is TableStyleElementType {
  return (TABLE_STYLE_ELEMENT_TYPES as readonly string[]).includes(value);
}

/**
 * How one region of a table is formatted: a {@link DifferentialStyle} laid over whatever the cells
 * already carry, plus — for a stripe — how many rows or columns wide one band is.
 *
 * A `numFmt` here is carried faithfully but has no visible effect: Excel's own table-style element
 * exposes a font, an interior and borders, and nothing for a number format. See
 * {@link DifferentialStyle}.
 */
export interface TableStyleElement extends DifferentialStyle {
  /**
   * The band width, in rows or columns, for a striped element — `2` makes each band two rows deep.
   * Defaults to 1.
   *
   * Meaningful **only** on the four stripe types ({@link STRIPE_ELEMENT_TYPES}); ECMA-376 says so and
   * Excel ignores it elsewhere. Setting it on any other element is rejected rather than silently
   * dropped: a caller who wrote it meant something by it, and a value that vanishes into a file that
   * still opens cleanly is the kind of bug nobody finds.
   */
  readonly size?: number | undefined;
}

/**
 * A custom table style, ready to be registered on a workbook and named by a table's
 * {@link TableStyleInfo.name}.
 *
 * Elements are applied in the order ECMA-376 fixes, not the order they are written here: whole table,
 * then the column stripes, then the row stripes, then last/first column, header row, total row, and
 * the four corner cells. So a row stripe wins over a column stripe, and both win over the whole-table
 * formatting — worth knowing when a stripe colour appears not to take.
 */
export interface TableStyle {
  /** The name a table references, and the name Excel shows in its style gallery. */
  readonly name: string;
  /** The regions this style formats. An element left out is not styled by it. */
  readonly elements: Readonly<Partial<Record<TableStyleElementType, TableStyleElement>>>;
  /** Whether the style is offered for tables. Defaults to true. */
  readonly table?: boolean | undefined;
  /** Whether the style is offered for pivot tables. Defaults to true. */
  readonly pivot?: boolean | undefined;
}

/**
 * Reject a table style the writer would otherwise emit as valid-but-inert XML.
 *
 * Both failures here are of the same kind: Excel accepts the file and quietly does nothing with the
 * part the caller cared about. An empty name means no table can ever reference the style, and a
 * `size` outside a stripe is ignored — neither shows up as a repair prompt or a schema error, so the
 * only place to catch them is the call that made them.
 *
 * @throws {Error} if the name is empty, or a non-stripe element carries a `size`, or a `size` is not
 *   a positive integer.
 */
export function checkTableStyle(style: TableStyle): void {
  if (style.name === '') {
    throw new Error('a table style needs a name: a table references its style by name');
  }
  for (const [type, element] of Object.entries(style.elements)) {
    if (element === undefined) continue;
    if (!isTableStyleElementType(type)) {
      throw new Error(
        `Invalid table style element ${JSON.stringify(type)}: not a value the OOXML enumeration allows`,
      );
    }
    const {size} = element;
    if (size === undefined) continue;
    if (!STRIPE_ELEMENT_TYPES.has(type)) {
      throw new Error(
        `table style element "${type}" cannot carry a size: band width applies only to ` +
          `${[...STRIPE_ELEMENT_TYPES].join(', ')}`,
      );
    }
    if (!Number.isInteger(size) || size < 1) {
      throw new Error(`Invalid table style band size ${size}: expected a positive integer`);
    }
  }
}
