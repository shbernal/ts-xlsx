// Excel date serials — the bridge between a JS `Date` and OOXML's numeric storage.
//
// A spreadsheet never stores a date as such: it stores a number (the "serial") and a
// number format that tells the viewer to render that number as a date. Serial 1 is
// 1900-01-01. The catch is a bug Excel has carried since 1985 and every consumer must
// reproduce for compatibility: the 1900 date system counts a phantom 1900-02-29 that
// never existed (serial 60), so serials at or below 59 sit one calendar day later than
// a naive "days since the epoch" offset would place them. Round-tripping a date means
// reproducing that quirk faithfully in both directions, or imported pre-March-1900
// dates land a day early.
//
// Dates here are treated as UTC wall-clock: a serial carries no timezone, so we convert
// against a UTC epoch in both directions. A `Date` written and read back through this
// module is exact; the caller owns any timezone interpretation before it reaches here.

const MS_PER_DAY = 86_400_000;

// Serial 0 in the 1900 date system is nominally 1899-12-30. The phantom leap day lives
// at serial 60, so the offset correction pivots there.
const EPOCH_1900_UTC = Date.UTC(1899, 11, 30);
const PHANTOM_SERIAL = 60;

/**
 * The number format applied to a `Date` cell that carries no explicit format of its own,
 * so the value renders — and reads back — as a date rather than a bare serial number.
 */
export const DEFAULT_DATE_NUMFMT = 'yyyy-mm-dd';

/**
 * Convert a JS `Date` to its 1900-system Excel serial, reproducing the phantom-leap-day
 * quirk so the value renders on the calendar date Excel would show. Fractional serials
 * carry the time of day. The caller must reject a non-finite (invalid) date before here.
 */
export function dateToSerial(date: Date): number {
  const days = (date.getTime() - EPOCH_1900_UTC) / MS_PER_DAY;
  // Days at or below the phantom (1900-02-28 is day 60 from the nominal epoch) are shifted
  // one earlier to skip the fake 1900-02-29 that Excel counts at serial 60.
  return days <= PHANTOM_SERIAL ? days - 1 : days;
}

/**
 * Convert a 1900-system Excel serial back to a UTC `Date`, accounting for the phantom
 * 1900-02-29 so serial 1 reads as 1900-01-01 (not 1899-12-31) and consecutive serials
 * map to consecutive days.
 */
export function serialToDate(serial: number): Date {
  const dayOffset = serial < PHANTOM_SERIAL ? serial + 1 : serial;
  return new Date(EPOCH_1900_UTC + dayOffset * MS_PER_DAY);
}

/**
 * Whether a number-format code renders its value as a date or time. A format is a date
 * format when, once its non-formatting sections are removed — bracketed color/locale/
 * condition directives, quoted literals, and escaped characters — any of the date/time
 * placeholder letters (`y m d h s`) remain. So `"$"#,##0.00` and `0.00%` are not dates
 * while `yyyy-mm-dd`, `dd/mm/yyyy`, and `[$-409]mmmm d, yyyy` are.
 */
export function isDateFormat(code: string): boolean {
  const stripped = code
    .replace(/\[[^\]]*\]/g, '')
    .replace(/"[^"]*"/g, '')
    .replace(/\\./g, '');
  return /[ymdhs]/i.test(stripped);
}
