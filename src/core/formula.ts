// OOXML froze its formula-function grammar around Excel 2007. Every function Microsoft has added
// since — the dynamic-array family, LAMBDA and its helpers, the newer text and logical functions —
// is persisted in the sheet XML under an `_xlfn.` name-mangling prefix. The prefix is purely an
// on-disk convention: the model only ever holds the plain, readable name, the writer applies the
// prefix on the way out, and the reader strips it back on the way in. A writer that omits it emits
// a formula current Excel silently drops, because the function is unknown under its bare name. This
// module is the single place that knows the mangling — shared by the xlsx writer and reader like
// address.ts and date.ts own their domains.

// The post-2007 functions Excel persists with an `_xlfn.` prefix. Restricted to names with no
// internal '.', so the single-pass tokenizer below never mistakes a dotted segment for a call: the
// renamed 2010 statistical family (NORM.DIST, BETA.INV, …) also carries the prefix but needs the
// tokenizer to treat '.' as part of a function name, so it is a deliberate follow-up.
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
]);

const XLFN = '_xlfn.';

// A function call is an identifier immediately followed by '('. The negative lookbehind rejects a
// name preceded by an identifier character or '.', so a name already qualified by a prefix
// (_xlfn.XLOOKUP) or a dotted segment is never matched — this is what stops an already-prefixed
// formula from being double-prefixed. Lookbehind rather than a consumed boundary char so adjacent
// calls (SUM(FILTER(…))) both match.
const FUNCTION_CALL = /(?<![A-Za-z0-9_.])([A-Za-z_][A-Za-z0-9_]*)(\s*\()/g;
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
