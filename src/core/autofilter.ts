import {AuthoringError} from '../errors.ts';
import {decodeRange, encodeAddress} from './address.ts';
import {isDeletedSpan, shiftIndex} from './grid-shift.ts';

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
 * A column narrowed to a discrete set of allowed values: the checkbox list in Excel's dropdown.
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

const CUSTOM_FILTER_OPERATORS: Record<CustomFilterOperator, true> = {
  equal: true,
  notEqual: true,
  lessThan: true,
  lessThanOrEqual: true,
  greaterThan: true,
  greaterThanOrEqual: true,
};

/** Narrow a raw `operator` attribute to a known {@link CustomFilterOperator}. */
export function isCustomFilterOperator(value: string): value is CustomFilterOperator {
  return Object.hasOwn(CUSTOM_FILTER_OPERATORS, value);
}

/**
 * Validate and normalise a settable autofilter into its canonical stored form. A bare range string
 * is the common case (dropdowns, no criteria); an {@link AutoFilter} object carries per-column
 * criteria too. Throws when the range is not a bounded rectangle, or a column's `colId` falls
 * outside it, or a custom filter does not carry one or two predicates: a filter authored wrong is
 * a bug to surface, not to silently repair. (The reader sanitises hostile input before it reaches
 * here, so load never trips these guards.)
 */
export function canonicalizeAutoFilter(input: string | AutoFilter): AutoFilter {
  const ref = typeof input === 'string' ? input : input.ref;
  const {top, left, bottom, right, dimensions} = decodeRange(ref);
  if (top === undefined || left === undefined || bottom === undefined || right === undefined) {
    throw new AuthoringError(`autofilter range "${ref}" must be a bounded rectangle`);
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
    throw new AuthoringError(`autofilter colId ${column.colId} is outside the filter range`);
  }
  if (column.criteria.kind === 'custom') {
    const count = column.criteria.predicates.length;
    if (count < 1 || count > 2) {
      throw new AuthoringError(`a custom filter needs one or two predicates, got ${count}`);
    }
  }
  return column;
}

/**
 * Re-anchor a filter through a splice of `count` lines at `start` on `axis`, or drop it (`undefined`)
 * when the splice deleted every line it covered.
 *
 * A row splice only moves the range. A column splice moves its left edge too, and a criterion is
 * addressed by its offset from that edge rather than by an absolute column, so every offset is
 * re-measured against the new edge and a criterion whose column was deleted goes with the column.
 * Left alone, those offsets would keep their old numbers and silently re-point each filter at a
 * neighbouring column.
 */
export function shiftAutoFilter(
  filter: AutoFilter,
  axis: 'row' | 'col',
  start: number,
  count: number,
  delta: number,
): AutoFilter | undefined {
  const {top, left, bottom, right} = decodeRange(filter.ref);
  // Unreachable for a stored filter: canonicalizeAutoFilter refuses anything but a bounded rectangle.
  if (top === undefined || left === undefined || bottom === undefined || right === undefined) {
    return filter;
  }
  const [lo, hi] = axis === 'row' ? [top, bottom] : [left, right];
  if (isDeletedSpan(lo, hi, start, count)) return undefined;
  const movedLo = shiftIndex(lo, start, count, delta);
  const movedHi = shiftIndex(hi, start, count, delta);
  if (axis === 'row') {
    const ref = `${encodeAddress(left, movedLo)}:${encodeAddress(right, movedHi)}`;
    return {ref, columns: filter.columns};
  }
  const ref = `${encodeAddress(movedLo, top)}:${encodeAddress(movedHi, bottom)}`;
  const columns: FilterColumn[] = [];
  for (const column of filter.columns) {
    const absolute = left + column.colId;
    if (isDeletedSpan(absolute, absolute, start, count)) continue;
    columns.push({...column, colId: shiftIndex(absolute, start, count, delta) - movedLo});
  }
  return {ref, columns};
}
