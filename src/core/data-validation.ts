// Data validation: the input constraints Excel enforces on a cell (a dropdown list, a numeric
// range, a length limit, a custom formula). Unlike a style facet, a validation is not owned by one
// cell: it is a worksheet-level overlay keyed by a target range (`sqref`), and a cell inherits
// whichever rule's range contains it. Keying by range is what keeps a whole-column dropdown a single
// entry rather than a million per-cell copies.

// Each union below is paired with a `Record` keyed by it, and the guard tests that record. The
// pairing is what keeps the two honest: the compiler refuses a key the union does not name *and* a
// union member the record omits, so a token cannot be added to one side and forgotten on the other.
// The keys are `ST_DataValidationType`, `ST_DataValidationOperator` and `ST_DataValidationErrorStyle`
// verbatim; a token outside them is not a value this model can hold, and a reader narrows through
// the guard rather than asserting past it.

/** The kind of constraint a validation enforces. `list` is a dropdown; `custom` is an arbitrary
 * boolean formula; `none` constrains nothing and exists only to carry the rule's messages; the rest
 * bound a typed value (`whole`/`decimal`/`date`/`time`/`textLength`). */
export type DataValidationType =
  | 'none'
  | 'list'
  | 'whole'
  | 'decimal'
  | 'date'
  | 'time'
  | 'textLength'
  | 'custom';

const DATA_VALIDATION_TYPES: Record<DataValidationType, true> = {
  none: true,
  list: true,
  whole: true,
  decimal: true,
  date: true,
  time: true,
  textLength: true,
  custom: true,
};

/** Narrow a raw `<dataValidation type>` token to a known {@link DataValidationType}. */
export function isDataValidationType(value: string): value is DataValidationType {
  return Object.hasOwn(DATA_VALIDATION_TYPES, value);
}

/** How a typed validation compares its operand(s). Absent on a `list`/`custom` rule; defaults to
 * `between` on a typed rule (the value Excel omits from the XML). */
export type DataValidationOperator =
  | 'between'
  | 'notBetween'
  | 'equal'
  | 'notEqual'
  | 'greaterThan'
  | 'lessThan'
  | 'greaterThanOrEqual'
  | 'lessThanOrEqual';

const DATA_VALIDATION_OPERATORS: Record<DataValidationOperator, true> = {
  between: true,
  notBetween: true,
  equal: true,
  notEqual: true,
  greaterThan: true,
  lessThan: true,
  greaterThanOrEqual: true,
  lessThanOrEqual: true,
};

/** Narrow a raw `<dataValidation operator>` token to a known {@link DataValidationOperator}. */
export function isDataValidationOperator(value: string): value is DataValidationOperator {
  return Object.hasOwn(DATA_VALIDATION_OPERATORS, value);
}

/** How Excel reacts to input that fails the rule. */
export type DataValidationErrorStyle = 'stop' | 'warning' | 'information';

const DATA_VALIDATION_ERROR_STYLES: Record<DataValidationErrorStyle, true> = {
  stop: true,
  warning: true,
  information: true,
};

/** Narrow a raw `<dataValidation errorStyle>` token to a known {@link DataValidationErrorStyle}. */
export function isDataValidationErrorStyle(value: string): value is DataValidationErrorStyle {
  return Object.hasOwn(DATA_VALIDATION_ERROR_STYLES, value);
}

/** One validation rule. `formulae` holds the operand(s), `formula1` then optional `formula2`: a
 * numeric literal is stored as a number, while a cell reference, defined name, or list source keeps
 * its verbatim string. */
export interface DataValidation {
  type: DataValidationType;
  operator?: DataValidationOperator;
  formulae?: (string | number)[];
  allowBlank?: boolean;
  showInputMessage?: boolean;
  showErrorMessage?: boolean;
  errorStyle?: DataValidationErrorStyle;
  error?: string;
  errorTitle?: string;
  prompt?: string;
  promptTitle?: string;
}

/** A validation bound to the range(s) it covers. `sqref` is an OOXML `sqref`: one or more
 * space-separated ranges. `extended` marks a rule stored in the 2009 extension form
 * (`<x14:dataValidation>` inside the worksheet `<extLst>`), Excel's carrier for validations a
 * legacy `<dataValidation>` cannot express, such as a list source on another sheet. The flag is how
 * a rule read from that form remembers to be written back to it, rather than downgraded to the
 * standard element (which would corrupt a cross-sheet reference). */
export interface DataValidationEntry {
  sqref: string;
  rule: DataValidation;
  extended?: boolean;
}

/** A defensive copy of a rule, so a stored validation never aliases the caller's object (nor its
 * `formulae` array). */
export function cloneDataValidation(rule: DataValidation): DataValidation {
  return {
    ...rule,
    ...(rule.formulae !== undefined ? {formulae: [...rule.formulae]} : {}),
  };
}
