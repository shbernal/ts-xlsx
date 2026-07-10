# Literal letter runs in a number format: quote them so a consumer cannot mis-tokenize them

Cluster: formats

## Scenario

A user sets a currency number format whose symbol is a multi-letter code, e.g. `CHF #,##0.00`, and
reports the resulting file opens with a repair prompt in Excel. In an OOXML number-format code,
unquoted letters are format *tokens* (in a date/time context `h`/`m`/`s`/`d`/`y` mean hours/minutes/
seconds/day/year, `e` era, etc.), so a bare alphabetic run like `CHF` can be interpreted as tokens
rather than literal text and make the format code ill-formed for a strict consumer. The safe form
quotes the literal: `"CHF" #,##0.00`.

> Spec note, not a corpus case: probing shows the library round-trips *both* forms faithfully — it
> writes `CHF #,##0.00` verbatim and re-reads it verbatim, and the quoted `"CHF" #,##0.00` likewise
> (the quoted-literal round-trip is already locked by `custom-numfmt-string-roundtrips-verbatim`,
> whose format contains quoted currency literals). There is no library round-trip bug to assert; the
> corruption is a downstream-consumer interpretation of an unquoted literal, and whether the writer
> should *auto-quote* bare letter runs is an undecided design choice. The durable value is that
> decision and its constraints.

## Desired behavior / the decision to make

- **A literal alphabetic run in a number format must not be mis-read as format tokens by a strict
  consumer.** The library must ensure a format like a multi-letter currency symbol produces a file
  that opens without repair.
- **Two candidate contracts, pick one and document it:**
  1. **Author-supplies-quotes (minimal).** The library stores and round-trips the exact format code
     given; the caller is responsible for quoting literal text (`"CHF"`). Documented clearly, with
     examples, so the failure mode is a known caller responsibility rather than a surprise.
  2. **Writer auto-quotes bare letter runs (safer, more magic).** On write, the library detects an
     alphabetic run that is not a recognized format token and wraps it in quotes, so `CHF #,##0.00`
     is emitted as `"CHF" #,##0.00`. This prevents the corruption without caller effort, but must be
     careful not to quote legitimate tokens (`General`, `AM/PM`, date/time letters, `E+`/`E-`
     scientific, `#`/`0`/`?`/`,`/`.`/`%`/`@` and the `[$-…]` locale/`[Red]`/`[>0]` bracket
     constructs), and must be idempotent (never double-quote an already-quoted literal).
- **Round-trip fidelity holds either way.** Whatever is written must re-read as the same format code,
  and a re-write must not drift (no gradual re-quoting on each save).

## Open questions

- Auto-quote or not? Auto-quoting is friendlier but risks mangling an edge-case token; the
  token-vs-literal classifier must be precise and conservative (quote only runs that are provably not
  tokens).
- Locale currency: is the `[$CHF]`/`[$USD-…]` bracketed-currency syntax the preferred canonical form
  to steer callers toward, rather than a bare or quoted literal?
- Validation mode: offer a strict/validate path that rejects (or warns about) a format code with an
  unquoted ambiguous letter run, so a caller can catch it before shipping the file?

Related: `custom-numfmt-string-roundtrips-verbatim`, `numfmt-date-detection-literal-m-scaling`,
`builtin-locale-date-format-code-reporting`.
