# Writing a JS Date must follow a documented, configurable timezone projection

Cluster: types / dates

## Scenario

A developer assigns a JavaScript `Date` to a cell and expects the wall-clock date/time they
constructed to appear verbatim in the spreadsheet. Excel/OOXML stores dates as timezone-naive
serial numbers (fractional days since an epoch), while a JS `Date` is an absolute UTC instant — so
the library must choose a rule for projecting the instant onto a naive serial number. When it uses
the instant's UTC components, a caller in a non-UTC zone who wrote a local-midnight date (e.g.
`new Date(y, m, d)`) sees the value land on the previous calendar day at a nonzero time (the prior
day at 22:00 for a +2h zone). A caller who built a UTC-based date is happy. There is no single rule
that satisfies everyone: the same absolute instant legitimately maps to different naive wall-clock
values depending on the intended timezone, and the library cannot infer intent from a bare `Date`.

## Desired behavior

- Writing a `Date` produces a serial number by a **documented, deterministic** rule, and that rule
  is **configurable** because the correct projection depends on the caller's intended timezone,
  which is unrecoverable from a bare instant.
- Offer an explicit conversion policy at write time:
  - a default **UTC projection** (serial = the instant's UTC components) — interoperable with
    foreign generators that also serialize naive dates, and matching historical behavior;
  - an opt-in **local / host-timezone projection** (serial = local wall-clock components), covering
    the most common surprise;
  - ideally an explicit **named-zone or fixed-offset projection** so the same input + declared zone
    yields identical bytes regardless of the host machine's zone (a UTC server and a +2h workstation
    must not diverge — the "works on my laptop, wrong on CI" trap).
- The **reverse direction** (serial → `Date`) must round-trip consistently with whatever write
  policy was chosen, or values drift on read-modify-write cycles.
- Centralize the serial math (epoch base `1899-12-30`, the 1900 leap-year quirk, the optional 1904
  date system) in the same layer so the projection and the epoch handling cannot disagree.

## Prior art / workarounds

Users converge on hand-rolled offset math the library should absorb:

- Pre-shift the instant so its UTC components equal the desired wall-clock:
  `new Date(d.getTime() - d.getTimezoneOffset()*60000)` — makes local wall-clock survive, at the
  cost of the `Date` no longer being the true instant.
- Construct in UTC up front: `new Date('2017-03-15')` (date-only ISO parses as UTC midnight) or
  `Date.UTC(...)`.
- Reinterpret local time as UTC without moving the timestamp (Luxon
  `setZone('utc', {keepLocalTime:true})`, moment/dayjs `utcOffset`).

The recurring, years-long request is that the library manage this itself rather than forcing every
caller to pre-shift.

## Open questions

- Is the policy per-workbook, per-write, or per-cell/column? Per-write with a workbook default is
  the pragmatic middle.
- Is a fixed IANA zone / offset a first-class option so output is independent of the host TZ?
- Accept only `Date`, or also epoch-millis and ISO strings — and how does each interact with the
  chosen policy?
- Safest default given both interop (favors UTC) and the volume of confused users (favors local)?
  Leaning: keep UTC as the interoperable default but make local/explicit-zone a trivially
  discoverable one-liner so callers stop hand-rolling offset math.

Related: `date-serial-1900-epoch-leap-year`, `date-value-written-as-serial-not-text`,
`strict-mode-iso8601-date-parses-correctly`.
