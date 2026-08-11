// The timestamp every package entry is stamped with, and why it is not the clock.
//
// A zip entry carries a modification time in both its local header and the central directory. Left
// to fflate that is `Date.now()`, which makes a package a function of *when* it was written: two
// writes of an unchanged workbook differ in a few bytes per entry and in nothing else. Excel cannot
// see the difference; everything around the file can. A committed `.xlsx` deliverable churns in
// `git diff` on every regeneration, a cache keyed on output bytes never hits, and a byte-comparison
// gate cannot separate "the writer changed" from "the clock moved" — so it stops being able to fail
// for the right reason, which is worse than not having one.
//
// Pinning it makes writing a pure function of the workbook. Two constraints picked the value:
//
//   - Zip stores DOS dates, which cover 1980-2099 only, so the Unix epoch is not expressible.
//   - fflate encodes the stamp through the *local-time* getters (`getFullYear`, `getHours`, …), so
//     a fixed UTC instant would still land on different DOS bytes in different timezones. The
//     constant is therefore built from local components, which read back identically everywhere,
//     and midday is used so that no DST transition can shift the date under it.
//
// The sibling `ts-pptx` writer pins its own stamp for the same reason; the two need not agree on
// the instant, only on refusing the clock.
//
// A caller who genuinely wants wall-clock times can re-stamp the archive afterwards; a caller who
// wants reproducibility could not get it back. That asymmetry is why this is the behaviour rather
// than an option.

/**
 * Modification time stamped on every zip entry, so package bytes depend on the workbook alone.
 * Encodes as 2001-01-01 12:00:00 in the DOS date field on any machine, in any timezone.
 */
export const FIXED_ENTRY_MTIME = new Date(2001, 0, 1, 12, 0, 0, 0).getTime();
