// The spreadsheet-flavoured readings of one attribute's value: the two numbers, the numeric-literal
// coercion, the enumerated token, and the `_xHHHH_` cell-text escape.
//
// Apart from `xml-scan.ts` for the reason ADR-0004's split states and stopped one module short of
// finishing: `/customui` is a ribbon reader that wants the scan and one boolean, and while these
// lived beside the scan every helper added here was charged to that entry's bundle budget. That is
// how it drifted 3 KB over one; it had 0.7 KB of headroom left when these moved.
//
// The three OOXML booleans stay with the scanner, and that is where the line falls rather than at
// "everything that reads a value". A `<button enabled="0">` in a ribbon part is the same boolean a
// `<sheetPr>` attribute is, so the booleans are not a spreadsheet's question; a `numInteger` with a
// floor, a `_xHHHH_` escape and a `coerceNumericLiteral` are.
//
// What every reading here has in common is what it does with a value it cannot read: it returns
// `undefined` and lets the caller store nothing. A file this library did not write is allowed to be
// wrong, losing one attribute beats losing the sheet, and a caller recording only what the source
// carried is what keeps a re-write byte-clean. `./xml.ts` holds the write-side halves, which throw
// instead, because a value an author supplied is a mistake at the call.

/**
 * The SpreadsheetML `_xHHHH_` escape, in the only place it may appear: a complete cell-text value.
 *
 * The mirror of `escapeSpreadsheetText` in `./xml.ts`, and it sits here rather than beside it for
 * the same reason `decodeEntities` sits apart from `escapeText`: the write helpers carry an
 * `AuthoringError` and a whole serialisation vocabulary the reader has no business importing.
 *
 * **One left-to-right pass, and that is load-bearing.** `005F` maps to `_` like any other code
 * point, with no special case, because a single pass already gives the underscore escape its
 * meaning: in `_x005F_x0041_` the match at 0 yields `_` and scanning resumes at `x0041_`, which has
 * no leading underscore left to start an escape. So the value reads back as the literal seven
 * characters `_x0041_` the author wrote. Decoding `_x005F_` in a pass of its own, before or after
 * the rest, collapses that to `A` and loses the distinction the encoder went to trouble to keep.
 * Excel agrees: it reads that cell as `_x0041_`.
 *
 * The decode is unconditional, not a repair of characters XML cannot carry. Excel reads
 * `a_x0009_b` as a tab even though a literal tab would have been perfectly legal there, so a
 * decoder that only handled the illegal range would disagree with Excel on files Excel wrote.
 */
export function decodeSpreadsheetText(value: string): string {
  if (!value.includes('_')) return value;
  return value.replace(/_x([0-9A-Fa-f]{4})_/g, (_match, hex: string) =>
    String.fromCharCode(Number.parseInt(hex, 16)),
  );
}

// The numeric attributes need the same treatment, and for the same reason: a bare `Number(attr)`
// turns a token it cannot parse into `NaN`, which is a number, so every guard downstream passes it
// along until something far from the file throws about a value the caller never wrote. Two readings
// cover the format. An ordinal - a count, an index, a row or a column - is an integer or it is
// nothing (`numInteger`). A measurement - a width, a height, a tint, a margin - is any finite
// number (`numFinite`). Both take the floor the attribute's kind implies, because nearly every call
// site wants "at least 0" or "at least 1" and would otherwise spell it inline and sometimes forget.
// Absent, unparseable and out-of-floor all read as `undefined`, so a caller stores only what the
// source carried and a re-write stays byte-clean.

/** An OOXML integer attribute at or above `min` (default: unbounded below); `undefined` when the
 * attribute is absent, blank, fractional, not a number, or below the floor. Integers past
 * `Number.MAX_SAFE_INTEGER` read as `undefined` too: no index or count is usable out there, and
 * arithmetic on one silently lies. */
export function numInteger(
  val: string | undefined,
  min = -Number.MAX_SAFE_INTEGER,
): number | undefined {
  const n = parseAttrNumber(val);
  if (n === undefined || !Number.isSafeInteger(n) || n < min) return undefined;
  return n;
}

/** An OOXML decimal attribute at or above `min` (default: unbounded below); `undefined` when the
 * attribute is absent, blank, not a number, or below the floor. Infinities are not finite numbers
 * and read as `undefined`. */
export function numFinite(val: string | undefined, min = -Infinity): number | undefined {
  const n = parseAttrNumber(val);
  if (n === undefined || n < min) return undefined;
  return n;
}

// `Number("")` and `Number(" ")` are both 0, which would turn an empty attribute into a real value.
function parseAttrNumber(val: string | undefined): number | undefined {
  if (val === undefined || val.trim() === '') return undefined;
  const n = Number(val);
  return Number.isFinite(n) ? n : undefined;
}

/** Read an operand's text as a number only when it is a canonical decimal literal (optional sign,
 * digits, optional fraction). A cell reference, defined name, expression, or exotically-spelled
 * number (`1E5`, hex) keeps its verbatim text, so it is neither coerced to `NaN` and lost nor
 * re-spelled into a number that would not re-write byte-clean. Callers layer their own type rules
 * (a data-validation `list`/`custom` operand stays a string regardless of what it looks like). */
export function coerceNumericLiteral(text: string): string | number {
  const trimmed = text.trim();
  return /^-?\d+(?:\.\d+)?$/.test(trimmed) ? Number(trimmed) : text;
}

// The third kind of attribute, after the booleans and the numbers, and dropped on the same terms: a
// token from a closed OOXML enumeration. `checkedToken` in `./xml.ts` is the write-side half of the
// same grammar, and the pair is deliberately asymmetric in what it does when the token is foreign.
// The reader drops it, because a file it did not write is allowed to be wrong and losing one
// attribute beats losing the sheet. The writer throws, because a value an author supplied is a
// mistake at the call and the file it would produce is one Excel refuses to open.

/** Narrow an enumerated attribute through its guard; `undefined` when absent or not a member. */
export function enumToken<T extends string>(
  val: string | undefined,
  isMember: (candidate: string) => candidate is T,
): T | undefined {
  return val !== undefined && isMember(val) ? val : undefined;
}
