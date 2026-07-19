// Worksheet-level protection: the `<sheetProtection>` state that makes the per-cell
// `locked`/`hidden` flags enforceable. A cell's protection does nothing until the sheet
// itself is protected; this module models that sheet-level switch and the optional
// password credential that guards lifting it.
//
// The option surface is stated in the AUTHOR's terms — each flag answers "may a user do
// this while the sheet is protected?" (`sort: true` = sorting stays available). OOXML
// encodes the inverse ("1" LOCKS an operation, "0"/omission PERMITS it) and its per-
// attribute defaults differ; that encoding table is {@link SHEET_PROTECTION_FLAGS} below,
// shared by the writer and reader, while the translation that consumes it lives in the io layer.

import {createHash, randomBytes} from 'node:crypto';

/**
 * Whether each protected-sheet operation stays available to a user. Every flag is an
 * *allow* flag: `true` keeps the operation permitted, `false` forbids it, and an absent
 * flag falls to Excel's default for that operation (most editing operations default to
 * forbidden once a sheet is protected; selecting cells defaults to permitted).
 */
export interface SheetProtectionFlags {
  /** Select locked cells (Excel permits this by default). */
  readonly selectLockedCells?: boolean;
  /** Select unlocked cells (permitted by default). */
  readonly selectUnlockedCells?: boolean;
  readonly formatCells?: boolean;
  readonly formatColumns?: boolean;
  readonly formatRows?: boolean;
  readonly insertColumns?: boolean;
  readonly insertRows?: boolean;
  readonly insertHyperlinks?: boolean;
  readonly deleteColumns?: boolean;
  readonly deleteRows?: boolean;
  readonly sort?: boolean;
  readonly autoFilter?: boolean;
  readonly pivotTables?: boolean;
  readonly objects?: boolean;
  readonly scenarios?: boolean;
}

/** {@link SheetProtectionFlags} plus the password-hardening knob accepted by `protect`. */
export interface SheetProtectionOptions extends SheetProtectionFlags {
  /**
   * Iteration count for the password hash. Higher is slower to brute-force; Excel writes
   * 100000 by default. Ignored when no password is given.
   */
  readonly spinCount?: number;
}

/**
 * A password-derived credential, in OOXML's agile form: the hash algorithm, the salted
 * iterated hash of the password, the salt, and the iteration count — everything a consumer
 * needs to verify a supplied password without the password ever being stored.
 */
export interface SheetProtectionCredential {
  readonly algorithmName: string;
  readonly hashValue: string;
  readonly saltValue: string;
  readonly spinCount: number;
}

/** A sheet's protection: which operations stay allowed, and the optional password guard. */
export interface SheetProtection {
  readonly flags: SheetProtectionFlags;
  readonly credential?: SheetProtectionCredential;
}

/**
 * The OOXML encoding table for the protection flags: each `<sheetProtection>` attribute paired
 * with whether that operation is *forbidden by default* once a sheet is protected. Both directions
 * key off this one list — the writer turns an author allow-flag into an attribute (omitting values
 * equal to the default), the reader turns an attribute back into an allow-flag — so serialization
 * and deserialization can never fall out of step. Most editing operations default to forbidden
 * under protection; selecting cells and the object/scenario operations default to permitted.
 */
export const SHEET_PROTECTION_FLAGS: readonly {
  readonly key: keyof SheetProtectionFlags;
  readonly defaultForbidden: boolean;
}[] = [
  {key: 'formatCells', defaultForbidden: true},
  {key: 'formatColumns', defaultForbidden: true},
  {key: 'formatRows', defaultForbidden: true},
  {key: 'insertColumns', defaultForbidden: true},
  {key: 'insertRows', defaultForbidden: true},
  {key: 'insertHyperlinks', defaultForbidden: true},
  {key: 'deleteColumns', defaultForbidden: true},
  {key: 'deleteRows', defaultForbidden: true},
  {key: 'sort', defaultForbidden: true},
  {key: 'autoFilter', defaultForbidden: true},
  {key: 'pivotTables', defaultForbidden: true},
  {key: 'objects', defaultForbidden: false},
  {key: 'scenarios', defaultForbidden: false},
  {key: 'selectLockedCells', defaultForbidden: false},
  {key: 'selectUnlockedCells', defaultForbidden: false},
];

// OOXML's agile hashing (ECMA-376 / MS-OFFCRYPTO): the password is UTF-16LE, prefixed with
// the salt for the first hash, then re-hashed `spinCount` times with a little-endian uint32
// iteration counter mixed in. SHA-512 is the modern choice Excel writes.
const ALGORITHM_NAME = 'SHA-512';
const HASH = 'sha512';
const DEFAULT_SPIN_COUNT = 100000;
const SALT_BYTES = 16;

/**
 * Derive a fresh {@link SheetProtectionCredential} for a password. Each call generates a new
 * random salt, so protecting two sheets with the same password yields different credentials —
 * the salt is real randomness, not a stub.
 */
export function deriveCredential(
  password: string,
  spinCount: number = DEFAULT_SPIN_COUNT,
): SheetProtectionCredential {
  const salt = randomBytes(SALT_BYTES);
  const secret = Buffer.from(password, 'utf16le');
  let hash = createHash(HASH)
    .update(Buffer.concat([salt, secret]))
    .digest();
  const iteration = Buffer.alloc(4);
  for (let i = 0; i < spinCount; i++) {
    iteration.writeUInt32LE(i, 0);
    hash = createHash(HASH)
      .update(Buffer.concat([hash, iteration]))
      .digest();
  }
  return {
    algorithmName: ALGORITHM_NAME,
    hashValue: hash.toString('base64'),
    saltValue: salt.toString('base64'),
    spinCount,
  };
}
