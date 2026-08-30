// The [MS-CFB] numbers the reader and the writer both answer to.
//
// `vbaProject.bin` is the one place in this tree where a hand-written binary reader and a
// hand-written binary writer are expected to round-trip each other's output byte for byte, and the
// whole edit-in-place path rests on that. Two independent transcriptions of the same five numbers
// is the cheapest way for that to break, and it would break quietly: a wrong `TYPE_*` in one file
// yields a container this library still reads and Excel rejects.
//
// Only what both directions are bound by lives here. The v3 sector layout the writer chooses to
// emit does not, because the reader deliberately takes every one of those off the header it was
// handed: a file may legally say otherwise, and a shared constant would invite a reader to trust
// the layout over the header field.

/** Sector chain markers ([MS-CFB] 2.2). */
export const FREESECT = 0xffffffff;
export const ENDOFCHAIN = 0xfffffffe;
export const FATSECT = 0xfffffffd;
export const DIFSECT = 0xfffffffc;

/**
 * Sector values 0xFFFFFFFA..0xFFFFFFFF are the reserved markers above, not data-sector indices, so
 * any value at or above this ceiling is chain-terminal. A reader walking a hostile chain tests the
 * ceiling rather than the four markers by name: an unassigned reserved value is terminal too.
 */
export const MAX_REGULAR_SECTOR = 0xfffffffa;

/** Directory-tree links reuse the sector convention: NOSTREAM, and any value at or above
 * {@link MAX_REGULAR_SECTOR}, mean "no such sibling or child". */
export const NOSTREAM = 0xffffffff;

/** Object types ([MS-CFB] 2.6.1). */
export const TYPE_EMPTY = 0;
export const TYPE_STORAGE = 1;
export const TYPE_STREAM = 2;
export const TYPE_ROOT = 5;

/** A directory entry is a fixed 128 bytes ([MS-CFB] 2.6.1), whatever the sector size. */
export const DIR_ENTRY_SIZE = 128;

/** 32 UTF-16 code units including the NUL terminator. This is also why a VBA module name is capped
 * at 31: the module is a stream, and the stream's name is the module's. */
export const MAX_NAME_CHARS = 31;
