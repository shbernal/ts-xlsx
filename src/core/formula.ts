// OOXML froze its formula-function grammar around Excel 2007. Every function Microsoft has added
// since — the dynamic-array family, LAMBDA and its helpers, the newer text and logical functions —
// is persisted in the sheet XML under an `_xlfn.` name-mangling prefix. The prefix is purely an
// on-disk convention: the model only ever holds the plain, readable name, the writer applies the
// prefix on the way out, and the reader strips it back on the way in. A writer that omits it emits
// a formula current Excel silently drops, because the function is unknown under its bare name. This
// module is the single place that knows the mangling — shared by the xlsx writer and reader like
// address.ts and date.ts own their domains.

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
