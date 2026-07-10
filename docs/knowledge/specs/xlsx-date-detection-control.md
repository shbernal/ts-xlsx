# Control automatic date detection when reading xlsx

Cluster: types

## Scenario

OOXML stores dates as a numeric serial value plus a date-shaped number-format code; there is no
intrinsic "date type" on the cell. To return a JS `Date`, the library applies a heuristic on the
number format. A user reading files where that heuristic guesses wrong — a numeric code that looks
date-ish but is not, or a genuine date they would rather handle as a raw serial — needs a way to make
the policy predictable and to opt out. The older API also conflated this xlsx read option with the
CSV-only date-parsing list, causing confusion when a moment-style format string was passed to the
wrong path.

> Spec note, not a corpus case: several corpus cases already lock specific date-detection outcomes
> (`date-value-written-as-serial-not-text`, `builtin-cjk-date-numfmt-ids-resolve-to-date-format`,
> `numfmt-date-detection-literal-m-scaling`, `strict-mode-iso8601-date-parses-correctly`). This note
> is the umbrella policy + opt-out those cases operate under.

## Desired behavior

- **A documented, deterministic rule** for which number-format codes trigger date coercion on read:
  the built-in date/time format ids, plus custom formats containing date tokens (with the literal-`m`
  minutes-vs-months disambiguation already locked separately), and correct handling of the workbook's
  **`date1904` epoch flag**.
- **An opt-out**: a consumer can disable automatic date coercion and receive the **raw numeric serial**
  for every cell, applying their own interpretation. (A per-cell or per-read granularity to decide.)
- **Independence from CSV date parsing**: this xlsx read policy is separate from the CSV-only
  `dateFormats` parsing list — the two must not share an option name, so passing a moment-style format
  string to one does not silently affect the other.

## Open questions

- Granularity of the opt-out: a global read flag (`{ dates: 'raw' | 'auto' }`), a predicate the caller
  supplies (`isDateFormat(numFmtCode) => boolean`), or both?
- Does opting out still expose the number-format code so the caller can run the same heuristic
  themselves?
- Should the default remain auto-detect (ecosystem expectation) with opt-out, or flip to raw-by-
  default for predictability? Auto-detect-by-default with a documented rule + opt-out is the
  least-surprise path.
- Interaction with the streaming reader, which must apply the same policy (see
  `streaming-read-applies-date-format`).

Related: `date-value-written-as-serial-not-text`, `numfmt-date-detection-literal-m-scaling`,
`builtin-cjk-date-numfmt-ids-resolve-to-date-format`, `strict-mode-iso8601-date-parses-correctly`,
`date-value-timezone-conversion`, `streaming-read-applies-date-format`.
