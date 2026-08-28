/**
 * Node built-ins, absent, with an error that says so.
 *
 * Three of them are statically reachable from the library's entry point today. `node:crypto`
 * arrives through `src/core/protection.ts`, which hashes a sheet-protection password;
 * `node:fs` and `node:stream` arrive through `src/io/xlsx/write-stream.ts`, the streaming
 * writer. Nothing this site calls reaches any of the three, but a bundler resolves imports
 * and not call graphs, so the build stops at the first one regardless.
 *
 * `docs/knowledge/specs/browser-safe-io-boundary.md` asks for the opposite arrangement: no
 * Node built-in statically reachable from a browser entry, with the filesystem conveniences
 * behind an export condition a browser never resolves. That is a change to the library, not
 * to its website, and it is on the follow-up list. Until it lands, this module is the
 * boundary: aliased in the site's Vite config, so a browser bundle contains no Node import
 * and the one path that would have needed one fails by name rather than as
 * `undefined is not a function` from inside a minified chunk.
 *
 * The error text is the one the spec asks for, which makes this the site demonstrating the
 * boundary rather than papering over it.
 */

const WHY =
  'This build of ts-xlsx runs in a browser, where Node built-ins do not exist. ' +
  'Sheet-protection passwords and the streaming writer are Node-only; use the buffer API ' +
  '(readXlsx and writeXlsx), which needs neither.';

function absent(name: string): never {
  throw new Error(`ts-xlsx: ${name} is not available in this environment. ${WHY}`);
}

export function createHash(): never {
  return absent('createHash (node:crypto)');
}

export function randomBytes(): never {
  return absent('randomBytes (node:crypto)');
}

export function createWriteStream(): never {
  return absent('createWriteStream (node:fs)');
}

export class PassThrough {
  constructor() {
    absent('PassThrough (node:stream)');
  }
}
