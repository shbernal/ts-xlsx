// OOXML froze its formula-function grammar around Excel 2007. Every function Microsoft has added
// since — the dynamic-array family, LAMBDA and its helpers, the newer text and logical functions —
// is persisted in the sheet XML under an `_xlfn.` name-mangling prefix. The prefix is purely an
// on-disk convention: the model only ever holds the plain, readable name, the writer applies the
// prefix on the way out, and the reader strips it back on the way in. A writer that omits it emits
// a formula current Excel silently drops, because the function is unknown under its bare name. This
// module is the single place that knows the mangling — shared by the xlsx writer and reader like
// address.ts and date.ts own their domains.
//
// It also owns formula *translation*: a spreadsheet fills a formula down or across a range by storing
// it once on a master cell and marking the rest as shared clones. Reading a clone means recovering the
// master's formula shifted to the clone's position — relative references move by the row/column
// offset, absolute (`$`-anchored) parts stay put. That relative-reference arithmetic lives here too.

import {columnToNumber, numberToColumn} from './address.ts';

// The post-2007 functions Excel persists with an `_xlfn.` prefix, keyed by their uppercased name.
// The tokenizer below treats '.' as part of a function name, so both the plain modern functions and
// the dotted 2010 statistical rename family are matched as whole names and prefixed.
const MODERN_FUNCTIONS: ReadonlySet<string> = new Set([
  // Dynamic arrays (Excel 365)
  'FILTER', 'SORT', 'SORTBY', 'UNIQUE', 'SEQUENCE', 'RANDARRAY', 'XLOOKUP', 'XMATCH',
  // LAMBDA and its helpers
  'LAMBDA', 'LET', 'BYROW', 'BYCOL', 'MAKEARRAY', 'MAP', 'REDUCE', 'SCAN', 'ISOMITTED',
  // Array shaping (Excel 365)
  'VSTACK', 'HSTACK', 'TOROW', 'TOCOL', 'WRAPROWS', 'WRAPCOLS', 'TAKE', 'DROP', 'EXPAND',
  'CHOOSEROWS', 'CHOOSECOLS',
  // Text (Excel 2019 / 365)
  'TEXTJOIN', 'CONCAT', 'TEXTBEFORE', 'TEXTAFTER', 'TEXTSPLIT', 'ARRAYTOTEXT', 'VALUETOTEXT',
  // Logical and conditional aggregation (Excel 2016 / 2019)
  'IFS', 'SWITCH', 'MAXIFS', 'MINIFS',

  // Other bare-name functions added after the frozen grammar (Excel 2010 / 2013) — trigonometric,
  // bitwise, engineering, information, and math/financial additions. Their names carry no '.', so
  // they need no tokenizer work; they simply have to be recognised as modern to earn the prefix.
  'AGGREGATE',
  'ACOT', 'ACOTH', 'COT', 'COTH', 'CSC', 'CSCH', 'SEC', 'SECH',
  'ARABIC', 'BASE', 'DECIMAL', 'COMBINA', 'PERMUTATIONA', 'GAMMA', 'GAUSS', 'PHI', 'MUNIT',
  'BITAND', 'BITOR', 'BITXOR', 'BITLSHIFT', 'BITRSHIFT',
  'IMCOSH', 'IMCOT', 'IMCSC', 'IMCSCH', 'IMSEC', 'IMSECH', 'IMSINH', 'IMTAN',
  'DAYS', 'ISOWEEKNUM', 'IFNA', 'NUMBERVALUE', 'SHEET', 'SHEETS',
  'FORMULATEXT', 'ISFORMULA', 'ENCODEURL', 'WEBSERVICE', 'FILTERXML',
  'UNICHAR', 'UNICODE', 'XOR', 'PDURATION', 'RRI',

  // The Excel 2010 statistical-consistency rename family and the handful of other post-2007
  // functions whose canonical names contain a '.'. They carry the same `_xlfn.` prefix; the whole
  // dotted name is stored, e.g. `_xlfn.NORM.DIST`, `_xlfn.T.DIST.2T`.
  'BETA.DIST', 'BETA.INV', 'BINOM.DIST', 'BINOM.DIST.RANGE', 'BINOM.INV',
  'CHISQ.DIST', 'CHISQ.DIST.RT', 'CHISQ.INV', 'CHISQ.INV.RT', 'CHISQ.TEST',
  'CONFIDENCE.NORM', 'CONFIDENCE.T', 'COVARIANCE.P', 'COVARIANCE.S', 'EXPON.DIST',
  'F.DIST', 'F.DIST.RT', 'F.INV', 'F.INV.RT', 'F.TEST',
  'GAMMA.DIST', 'GAMMA.INV', 'GAMMALN.PRECISE', 'HYPGEOM.DIST',
  'LOGNORM.DIST', 'LOGNORM.INV', 'MODE.MULT', 'MODE.SNGL', 'NEGBINOM.DIST',
  'NORM.DIST', 'NORM.INV', 'NORM.S.DIST', 'NORM.S.INV',
  'PERCENTILE.EXC', 'PERCENTILE.INC', 'PERCENTRANK.EXC', 'PERCENTRANK.INC', 'POISSON.DIST',
  'QUARTILE.EXC', 'QUARTILE.INC', 'RANK.AVG', 'RANK.EQ', 'SKEW.P',
  'STDEV.P', 'STDEV.S', 'T.DIST', 'T.DIST.2T', 'T.DIST.RT', 'T.INV', 'T.INV.2T', 'T.TEST',
  'VAR.P', 'VAR.S', 'WEIBULL.DIST', 'Z.TEST',
  'CEILING.PRECISE', 'FLOOR.PRECISE', 'ISO.CEILING', 'ERF.PRECISE', 'ERFC.PRECISE',
]);

const XLFN = '_xlfn.';
const XLPM = '_xlpm.';

// LET and LAMBDA are the only functions that bind names. Their parameter identifiers are persisted
// under an `_xlpm.` prefix — at the declaration site and at every in-body reference — exactly as the
// modern functions themselves carry `_xlfn.`. The prefix is scoped: a name bound by one LET/LAMBDA
// is only prefixed inside that call, so a same-named defined-name reference elsewhere is untouched.
const SCOPING_FUNCTIONS: ReadonlySet<string> = new Set(['LET', 'LAMBDA']);

// A function call is an identifier — dots included, so a dotted name like NORM.DIST is matched whole
// rather than by its tail — immediately followed by '('. The negative lookbehind rejects a name
// preceded by an identifier character or '.', so an already-qualified name (`_xlfn.XLOOKUP`) is
// consumed as a single token whose uppercased form is absent from the set, and is therefore never
// double-prefixed. Lookbehind rather than a consumed boundary char so adjacent calls
// (SUM(FILTER(…))) both match.
const FUNCTION_CALL = /(?<![A-Za-z0-9_.])([A-Za-z_][A-Za-z0-9_.]*)(\s*\()/g;
const PREFIX = /_xlfn\.|_xlpm\./g;

/**
 * Apply `transform` to a formula's code while copying its string literals verbatim. OOXML formula
 * strings are double-quoted, with a doubled `""` standing for an embedded quote; keeping the
 * transform out of them means a literal like `"FILTER("` is never mistaken for a function call.
 */
function outsideStrings(formula: string, transform: (code: string) => string): string {
  let out = '';
  let i = 0;
  const n = formula.length;
  while (i < n) {
    const quote = formula.indexOf('"', i);
    if (quote === -1) {
      out += transform(formula.slice(i));
      break;
    }
    out += transform(formula.slice(i, quote));
    let j = quote + 1;
    while (j < n) {
      if (formula[j] === '"') {
        if (formula[j + 1] === '"') {
          j += 2;
          continue;
        }
        j += 1;
        break;
      }
      j += 1;
    }
    out += formula.slice(quote, j);
    i = j;
  }
  return out;
}

/**
 * Prefix every modern function called by its plain name with `_xlfn.` so Excel accepts the stored
 * formula. Names already prefixed are left alone (never doubled), unknown/legacy functions pass
 * through untouched, and string literals are preserved verbatim. No other rewriting occurs — in
 * particular no `@` implicit-intersection operator is ever introduced.
 */
export function mangleFunctions(formula: string): string {
  return outsideStrings(formula, (code) =>
    code.replace(FUNCTION_CALL, (whole, name: string, open: string) =>
      MODERN_FUNCTIONS.has(name.toUpperCase()) ? `${XLFN}${name}${open}` : whole,
    ),
  );
}

/**
 * Strip the `_xlfn.` function prefix and the `_xlpm.` LET-parameter prefix back to the plain names,
 * so the model holds the readable form regardless of how a file stored it. String literals are left
 * untouched.
 */
export function unmangleFunctions(formula: string): string {
  return outsideStrings(formula, (code) => code.replace(PREFIX, ''));
}

const NAME_START = /[A-Za-z_]/;
const NAME_CHAR = /[A-Za-z0-9_.]/;
const WHITESPACE = /\s/;

// Advance past an identifier — dots included, matching FUNCTION_CALL — starting at `i`, or return `i`
// unchanged when no identifier begins there.
function readName(formula: string, i: number): number {
  if (!NAME_START.test(formula[i] ?? '')) return i;
  let j = i + 1;
  while (j < formula.length && NAME_CHAR.test(formula[j] ?? '')) j += 1;
  return j;
}

// Advance past a double-quoted string literal (opening quote at `i`), honouring the doubled-quote
// escape, so a comma or paren inside a literal never registers as syntax.
function skipString(formula: string, i: number): number {
  let j = i + 1;
  const n = formula.length;
  while (j < n) {
    if (formula[j] === '"') {
      if (formula[j + 1] === '"') {
        j += 2;
        continue;
      }
      return j + 1;
    }
    j += 1;
  }
  return n;
}

// Advance past a bracketed structured reference (opening bracket at `i`), which may nest, so a comma
// inside `Table[[#Data],[Col]]` never registers as an argument separator.
function skipBrackets(formula: string, i: number): number {
  let depth = 0;
  let j = i;
  const n = formula.length;
  while (j < n) {
    const ch = formula[j];
    if (ch === '[') depth += 1;
    else if (ch === ']') {
      depth -= 1;
      if (depth === 0) return j + 1;
    }
    j += 1;
  }
  return n;
}

/**
 * From the index of a call's opening paren, find the matching close and the `[start, end)` ranges of
 * its top-level, comma-separated arguments. Nested parens, string literals, and bracketed structured
 * references are skipped so their commas do not split an argument.
 */
function parseCall(formula: string, open: number): {close: number; args: [number, number][]} {
  const args: [number, number][] = [];
  const n = formula.length;
  let depth = 0;
  let argStart = open + 1;
  let i = open;
  while (i < n) {
    const ch = formula[i];
    if (ch === '"') {
      i = skipString(formula, i);
    } else if (ch === '[') {
      i = skipBrackets(formula, i);
    } else if (ch === '(') {
      depth += 1;
      i += 1;
    } else if (ch === ')') {
      depth -= 1;
      if (depth === 0) {
        args.push([argStart, i]);
        return {close: i, args};
      }
      i += 1;
    } else if (ch === ',' && depth === 1) {
      args.push([argStart, i]);
      argStart = i + 1;
      i += 1;
    } else {
      i += 1;
    }
  }
  args.push([argStart, n]);
  return {close: n, args};
}

// Extract the single, unprefixed identifier occupying an argument range, or `undefined` when the
// range is not one clean name (whitespace-trimmed) — a malformed binding we decline to touch.
function boundName(formula: string, [start, end]: [number, number]): string | undefined {
  let s = start;
  let e = end;
  while (s < e && WHITESPACE.test(formula[s] ?? '')) s += 1;
  while (e > s && WHITESPACE.test(formula[e - 1] ?? '')) e -= 1;
  if (s >= e || readName(formula, s) !== e) return undefined;
  const name = formula.slice(s, e);
  return name.startsWith(XLPM) ? undefined : name;
}

// The parameter names a LET/LAMBDA call binds. LAMBDA binds every argument but its last (the body);
// LET binds the even-indexed arguments up to but excluding its last (the calculation).
function parameterNames(formula: string, keyword: string, args: [number, number][]): ReadonlySet<string> {
  const names = new Set<string>();
  const isLambda = keyword === 'LAMBDA';
  for (let a = 0; a < args.length - 1; a += 1) {
    if (isLambda || a % 2 === 0) {
      const name = boundName(formula, args[a] as [number, number]);
      if (name !== undefined) names.add(name);
    }
  }
  return names;
}

/**
 * Prefix every LET/LAMBDA parameter identifier with `_xlpm.` — at its declaration and at each
 * reference within the binding call's parentheses — so Excel accepts the stored formula. The prefix
 * is lexically scoped: a name is only rewritten inside the call that binds it, string literals are
 * copied verbatim, and a lambda-valued parameter used as a call (`f(…)`) is prefixed too. Formulas
 * with no LET/LAMBDA pass through unchanged.
 */
export function mangleParams(formula: string): string {
  let out = '';
  let i = 0;
  const n = formula.length;
  // A stack of active bindings, each expiring exactly at its owner call's close paren. Nested
  // LET/LAMBDA push inner frames that pop first, so shadowing resolves to the same prefix anyway.
  const frames: {end: number; names: ReadonlySet<string>}[] = [];
  const inScope = (name: string): boolean => frames.some((frame) => frame.names.has(name));

  while (i < n) {
    const top = frames[frames.length - 1];
    if (top !== undefined && i >= top.end) {
      frames.pop();
      continue;
    }

    const ch = formula[i] as string;
    if (ch === '"') {
      const j = skipString(formula, i);
      out += formula.slice(i, j);
      i = j;
      continue;
    }
    if (ch === '[') {
      const j = skipBrackets(formula, i);
      out += formula.slice(i, j);
      i = j;
      continue;
    }

    const nameEnd = readName(formula, i);
    if (nameEnd === i) {
      out += ch;
      i += 1;
      continue;
    }

    const name = formula.slice(i, nameEnd);
    let k = nameEnd;
    while (k < n && WHITESPACE.test(formula[k] ?? '')) k += 1;
    const heads = formula[k] === '(';

    if (heads && SCOPING_FUNCTIONS.has(name.toUpperCase()) && !inScope(name)) {
      const {close, args} = parseCall(formula, k);
      // The keyword stays at the outer scope; its parameters take effect inside the parens.
      out += formula.slice(i, k + 1);
      frames.push({end: close, names: parameterNames(formula, name.toUpperCase(), args)});
      i = k + 1;
      continue;
    }

    // Any other identifier — a bare reference, an ordinary call, or a lambda-valued parameter call.
    // In-scope names (declaration sites included, as they lie inside their own binding's parens) take
    // the prefix; the rest pass through. Call arguments are covered by the continuing scan, so a
    // nested LET/LAMBDA within them is still seen.
    out += inScope(name) ? `${XLPM}${name}` : name;
    i = nameEnd;
  }
  return out;
}

/**
 * Mangle a model formula into its on-disk form: LET/LAMBDA parameter names first (`_xlpm.`), then the
 * modern-function prefix (`_xlfn.`). Ordering matters — parameter mangling reads the plain LET/LAMBDA
 * names before the function pass qualifies them. The inverse for both prefixes is unmangleFunctions.
 */
export function mangleFormula(formula: string): string {
  return mangleFunctions(mangleParams(formula));
}

// A cell reference at the start of a slice: an optional `$` then 1–3 uppercase column letters, an
// optional `$`, then the row digits. The column is uppercase-only because Excel stores it that way and
// so that a lowercase identifier (a defined name) is never mistaken for a reference. The row is capped
// at seven digits (Excel's last row is 1048576).
const CELL_REFERENCE = /^(\$?)([A-Z]{1,3})(\$?)([0-9]{1,7})/;

// A reference token must sit at an identifier boundary: not glued to a preceding name character (so
// the `A1` inside `_xlfn.A1` or a defined name `FOO_A1` is left alone) and not continued by one, nor
// opening a call `(` nor separating a sheet name `!` — a token followed by `!` is the sheet name, not
// a cell, and its reference sits after the `!`.
const NAME_CHAR_OR_DOT = /[A-Za-z0-9_.]/;
const REFERENCE_TAIL = /[A-Za-z0-9_.!(]/;

// Advance past a single-quoted sheet name (opening quote at `i`), honouring the doubled-quote escape,
// so a token inside `'My Sheet'` is never read as a cell reference.
function skipSingleQuoted(formula: string, i: number): number {
  let j = i + 1;
  const n = formula.length;
  while (j < n) {
    if (formula[j] === "'") {
      if (formula[j + 1] === "'") {
        j += 2;
        continue;
      }
      return j + 1;
    }
    j += 1;
  }
  return n;
}

/**
 * Shift every relative cell reference in a formula by `colDelta` columns and `rowDelta` rows, leaving
 * absolute (`$`-anchored) axes fixed. This is how a shared-formula clone recovers its own formula from
 * the master's: a master `A1*2` shared one row down reads back as `A2*2`, and `$A$1*B1` shared one row
 * and one column across as `$A$1*C2`. String literals, single-quoted sheet names, and bracketed
 * structured references are copied verbatim, and a sheet-qualified reference shifts the cell while its
 * sheet name is untouched. Function names and defined names carry no row digits, so they pass through.
 */
export function translateFormula(formula: string, colDelta: number, rowDelta: number): string {
  if (colDelta === 0 && rowDelta === 0) return formula;
  let out = '';
  let i = 0;
  const n = formula.length;
  while (i < n) {
    const ch = formula[i] as string;
    if (ch === '"') {
      const j = skipString(formula, i);
      out += formula.slice(i, j);
      i = j;
      continue;
    }
    if (ch === "'") {
      const j = skipSingleQuoted(formula, i);
      out += formula.slice(i, j);
      i = j;
      continue;
    }
    if (ch === '[') {
      const j = skipBrackets(formula, i);
      out += formula.slice(i, j);
      i = j;
      continue;
    }

    const prev = formula[i - 1];
    const atBoundary = prev === undefined || !NAME_CHAR_OR_DOT.test(prev);
    if (atBoundary && (ch === '$' || (ch >= 'A' && ch <= 'Z'))) {
      const match = CELL_REFERENCE.exec(formula.slice(i));
      if (match !== null) {
        const after = formula[i + match[0].length];
        if (after === undefined || !REFERENCE_TAIL.test(after)) {
          const [, colAbs, colLetters, rowAbs, rowDigits] = match as unknown as [
            string, string, string, string, string,
          ];
          const col = colAbs === '$' ? colLetters : numberToColumn(columnToNumber(colLetters) + colDelta);
          const row = rowAbs === '$' ? rowDigits : String(Number(rowDigits) + rowDelta);
          out += `${colAbs}${col}${rowAbs}${row}`;
          i += match[0].length;
          continue;
        }
      }
    }

    // Not a reference: consume a whole identifier at once (so its interior is never re-examined) or
    // copy a single non-name character.
    const nameEnd = readName(formula, i);
    if (nameEnd > i) {
      out += formula.slice(i, nameEnd);
      i = nameEnd;
      continue;
    }
    out += ch;
    i += 1;
  }
  return out;
}
