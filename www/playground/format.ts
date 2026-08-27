/**
 * The three quantities the playground puts on screen, formatted once.
 *
 * They live in their own module because a number a reader is asked to believe should be
 * formatted by something with a test beside it, not by a template expression nobody looks
 * at twice. Nothing here touches a DOM.
 */

const KIB = 1024;
const UNITS = ['B', 'KiB', 'MiB'] as const;

/**
 * Bytes, in the largest unit that keeps the number readable.
 *
 * Binary units, not decimal, because the number beside it on the page is a zip entry's
 * uncompressed size, and every tool a reader would check it against reports that in KiB.
 */
export function formatBytes(bytes: number): string {
  if (!Number.isFinite(bytes) || bytes < 0) return '-';
  let value = bytes;
  let unit = 0;
  while (value >= KIB && unit < UNITS.length - 1) {
    value /= KIB;
    unit += 1;
  }
  const digits = unit === 0 ? 0 : value < 10 ? 2 : 1;
  return `${value.toFixed(digits)} ${UNITS[unit] ?? 'B'}`;
}

/**
 * A duration, at the precision the measurement actually has.
 *
 * A sub-millisecond write is reported as such rather than rounded to `0 ms`, which reads as
 * "did not run". `performance.now()` in a browser is deliberately coarsened against timing
 * attacks, so anything below a tenth of a millisecond is reported as a bound and not as a
 * figure: a number with more digits than the clock has is a number that invites arithmetic
 * nobody can reproduce.
 */
export function formatMillis(ms: number): string {
  if (!Number.isFinite(ms) || ms < 0) return '-';
  if (ms < 0.1) return '<0.1 ms';
  if (ms < 10) return `${ms.toFixed(1)} ms`;
  if (ms < 1000) return `${Math.round(ms)} ms`;
  return `${(ms / 1000).toFixed(2)} s`;
}

/** A count, grouped, so four thousand rows does not read as forty thousand at a glance. */
export function formatCount(count: number): string {
  if (!Number.isFinite(count)) return '-';
  return new Intl.NumberFormat('en-GB').format(count);
}
