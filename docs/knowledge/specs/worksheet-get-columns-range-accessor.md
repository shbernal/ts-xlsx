# A range-of-columns accessor, symmetric with the row-range accessor

Cluster: address-decoding

## Scenario

The worksheet API exposes a single-column accessor and a range-of-rows accessor (get one row,
get a contiguous block of rows), but no symmetric range-of-columns accessor. Callers who want to
operate on several adjacent columns at once must loop over single-column lookups and assemble the
array themselves. A range-of-columns accessor would return a contiguous block of column objects
in one call, mirroring the row-range accessor for symmetry and ergonomics.

## Desired behavior

Provide a worksheet method that returns a contiguous run of column objects, symmetric with the
existing method that returns a contiguous run of row objects. Given a start column and a count
(or a start and end column), return the corresponding array of column accessors, each usable for
setting width/hidden/numFmt and per-cell access.

Prior art: the library already ships a single-column accessor (by index or letter) and a
multi-row accessor that takes a start and a length and returns an array of row objects. The
column accessor should follow the same shape so the two axes are consistent.

## Open questions

- ~~**Signature:** `(start, count)` vs `(firstCol, lastCol)`~~ **Settled by precedent: first and
  last, inclusive.** `Worksheet.getRange` takes `(top, left, bottom, right)` as inclusive corners in
  either order, and states that convention as binding on every range-shaped accessor here — so a
  columns accessor must be `(firstCol, lastCol)`, and a rows one `(firstRow, lastRow)`. One reading
  of a pair of numbers across all three axes; no start-and-count overload beside it.
- **Column identity:** accept 1-based numeric indices and/or letter references consistently with
  the single-column accessor.
- **Out-of-range / non-positive count:** define whether an empty array, `undefined`, or an error
  is returned; state it explicitly rather than inheriting the row accessor's behavior implicitly.
- **Laziness:** returned column objects should be the same lazily-materialized column objects the
  single accessor returns, so reading a range does not force allocation of empty columns.

Related note: `worksheet-columns-mutable-array-ergonomics` — both concern making the column axis
as ergonomic as the row axis.
