// The plumbing behind a row's and a column's formatting accessors.
//
// `Row` and `Column` are one handle on two axes, and what they share is not a face but a mechanism:
// look the position's format record up in the worksheet's sparse store, read a field off it, write a
// field into it, and treat a written `undefined` as a clear. That is identical on both axes apart
// from which pair of `WorksheetInternals` methods it reaches, and the clear rule in particular is
// subtle enough that stating it twice is how it drifts.
//
// The accessors themselves deliberately stay spelled out, one per class: they are the public API,
// their doc comments differ where the axes genuinely do, and `EveryRowPropertyIsMirrored` /
// `EveryColumnPropertyIsMirrored` are compile-time proofs that each one exists. Generating them
// would defeat the check that exists to stop this very mirror from rotting.
//
// `values` is likewise left duplicated on both classes. Sharing it would take an abstract
// cell-by-position accessor and an abstract position-of-cell reader to bridge `cell.col` against
// `cell.row`, which is more indirection than the six lines it would replace. The rules it obeys are
// stated on `Row.values` and pointed at from `Column.values`.
export abstract class AxisHandle<P extends object> {
  /** This position's format record, or `undefined` when it has none. Never fabricates one. */
  protected abstract propertiesOf(): P | undefined;

  /** This position's format record, created if it does not exist yet. */
  protected abstract ensureProperties(): P;

  /** Remove this position's format record entirely, so nothing is left declaring the line. */
  protected abstract dropProperties(): void;

  protected read<K extends keyof P>(key: K): P[K] | undefined {
    return this.propertiesOf()?.[key];
  }

  // `undefined` clears rather than stores: both property records declare their fields optional under
  // `exactOptionalPropertyTypes`, so a present-but-undefined key is not the same shape as an absent
  // one, and it would make an unformatted row or column look formatted to anything reading
  // `properties`. Clearing a position that has no record at all is a no-op, so a write of `undefined`
  // never materialises one.
  //
  // Clearing the *last* field takes the record with it. An emptied record is not the same thing as no
  // record: the used extent derives its bounds from which positions have one, so a row formatted at
  // 500 and then unformatted kept `rowCount` at 500 forever; `properties` promised a read that never
  // fabricates and answered `{}`; and a model round-trip was not idempotent, because the importer
  // assigns an empty record through setters that never fire, so the destination never recreated it.
  // One rule, stated here, fixes all three.
  protected write<K extends keyof P>(key: K, value: P[K]): void {
    if (value === undefined) {
      const properties = this.propertiesOf();
      if (properties === undefined) return;
      delete properties[key];
      if (Object.keys(properties).length === 0) this.dropProperties();
      return;
    }
    this.ensureProperties()[key] = value;
  }
}
