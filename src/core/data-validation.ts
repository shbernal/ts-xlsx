// Data validation — the input constraints Excel enforces on a cell (a dropdown list, a numeric
// range, a length limit, a custom formula). Unlike a style facet, a validation is not owned by one
// cell: it is a worksheet-level overlay keyed by a target range (`sqref`), and a cell inherits
// whichever rule's range contains it. Keying by range is what keeps a whole-column dropdown a single
// entry rather than a million per-cell copies.

/** The kind of constraint a validation enforces. `list` is a dropdown; `custom` is an arbitrary
 * boolean formula; the rest bound a typed value (`whole`/`decimal`/`date`/`time`/`textLength`). */
export type DataValidationType =
  | 'list'
  | 'whole'
  | 'decimal'
  | 'date'
  | 'time'
  | 'textLength'
  | 'custom';

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

/** How Excel reacts to input that fails the rule. */
export type DataValidationErrorStyle = 'stop' | 'warning' | 'information';

/** One validation rule. `formulae` holds the operand(s) — `formula1` then optional `formula2`: a
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

/** A validation bound to the range(s) it covers. `sqref` is an OOXML `sqref` — one or more
 * space-separated ranges. `extended` marks a rule stored in the 2009 extension form
 * (`<x14:dataValidation>` inside the worksheet `<extLst>`) — Excel's carrier for validations a
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
