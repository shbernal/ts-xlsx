import {decodeRange} from './address.ts';

/**
 * A worksheet's autofilter: the filtered region plus any per-column criteria narrowing it. A bare
 * range (no columns) is just the header-row dropdowns Excel draws; adding {@link FilterColumn}s
 * records the criteria a column is actively filtered by.
 */
export interface AutoFilter {
  /** The filtered region in canonical `A1:C10` form; its top row is the header the dropdowns sit on. */
  readonly ref: string;
  /** The columns actively narrowed, each addressed by its offset from the range's left edge. Empty
   *  when the filter only draws dropdowns without hiding any row. */
  readonly columns: readonly FilterColumn[];
}

/** One filtered column, addressed by its 0-based offset (`colId`) from the filter range's left edge. */
export interface FilterColumn {
  readonly colId: number;
  readonly criteria: FilterCriteria;
}

/** The two criteria kinds this library models: a discrete value set, or operator predicates. */
export type FilterCriteria = ValuesFilter | CustomFilter;

/**
 * A column narrowed to a discrete set of allowed values — the checkbox list in Excel's dropdown.
 * A row survives when its cell in this column matches one of {@link values} (or is blank, when
 * {@link blank} is set).
 */
export interface ValuesFilter {
  readonly kind: 'values';
  readonly values: readonly string[];
  readonly blank: boolean;
}

/**
 * A column narrowed to one or two operator predicates (`> 6`, `<> "draft"`). Two predicates are
 * AND-combined when {@link and} is set, else OR-combined; Excel permits at most two.
 */
export interface CustomFilter {
  readonly kind: 'custom';
  readonly and: boolean;
  readonly predicates: readonly CustomFilterPredicate[];
}

export interface CustomFilterPredicate {
  readonly operator: CustomFilterOperator;
  /** The comparison operand, kept as its raw string form (a number, or wildcard text like `a*`). */
  readonly val: string;
}

export type CustomFilterOperator =
  | 'equal'
  | 'notEqual'
  | 'lessThan'
  | 'lessThanOrEqual'
  | 'greaterThan'
  | 'greaterThanOrEqual';

const CUSTOM_FILTER_OPERATORS: ReadonlySet<string> = new Set<CustomFilterOperator>([
  'equal',
  'notEqual',
  'lessThan',
  'lessThanOrEqual',
  'greaterThan',
  'greaterThanOrEqual',
]);

/** Narrow a raw `operator` attribute to a known {@link CustomFilterOperator}. */
export function isCustomFilterOperator(value: string): value is CustomFilterOperator {
  return CUSTOM_FILTER_OPERATORS.has(value);
}

/**
 * Validate and normalise a settable autofilter into its canonical stored form. A bare range string
 * is the common case (dropdowns, no criteria); an {@link AutoFilter} object carries per-column
 * criteria too. Throws when the range is not a bounded rectangle, or a column's `colId` falls
 * outside it, or a custom filter does not carry one or two predicates — a filter authored wrong is
 * a bug to surface, not to silently repair. (The reader sanitises hostile input before it reaches
 * here, so load never trips these guards.)
 */
export function canonicalizeAutoFilter(input: string | AutoFilter): AutoFilter {
  const ref = typeof input === 'string' ? input : input.ref;
  const {top, left, bottom, right, dimensions} = decodeRange(ref);
  if (top === undefined || left === undefined || bottom === undefined || right === undefined) {
    throw new Error(`autofilter range "${ref}" must be a bounded rectangle`);
  }
  if (typeof input === 'string') return {ref: dimensions, columns: []};
  const width = right - left + 1;
  return {
    ref: dimensions,
    columns: input.columns.map((column) => canonicalizeColumn(column, width)),
  };
}

function canonicalizeColumn(column: FilterColumn, width: number): FilterColumn {
  if (!Number.isInteger(column.colId) || column.colId < 0 || column.colId >= width) {
    throw new Error(`autofilter colId ${column.colId} is outside the filter range`);
  }
  if (column.criteria.kind === 'custom') {
    const count = column.criteria.predicates.length;
    if (count < 1 || count > 2) {
      throw new Error(`a custom filter needs one or two predicates, got ${count}`);
    }
  }
  return column;
}
