// Where the repo is, and where the tools it runs live.
//
// Eleven scripts each opened with `resolve(dirname(fileURLToPath(import.meta.url)), '..')`, which
// is not a spelling of "the repo root" but of "one directory above this file". The two agree only
// as long as every script sits directly in `scripts/`; the day one moves into a subdirectory it
// resolves to `scripts/` and every path built from it lands somewhere plausible and wrong. One
// definition makes that a compile-time question rather than a runtime surprise.
//
// The tool entrypoints are here for a different reason: `node_modules/typescript/bin/tsc` was
// written out in two scripts, and the reasoning for *why* it is the package entry rather than the
// `.bin` shim was written out in four places plus `lefthook.yml`. That reasoning is a property of
// this repo's platform, not of any one script, so it is stated once, below.

import {mkdirSync, mkdtempSync, readFileSync} from 'node:fs';
import {dirname, join, resolve} from 'node:path';
import {fileURLToPath} from 'node:url';

/** The repository root, resolved from this module's own location. */
export const ROOT = resolve(dirname(fileURLToPath(import.meta.url)), '..');

/** A repo-relative path made absolute. */
export function fromRoot(...segments: readonly string[]): string {
  return resolve(ROOT, ...segments);
}

/**
 * The fields of `package.json` anything in this repo reads, as one declaration.
 *
 * Three readers had three private `PackageJson` interfaces over the same file, which is three
 * chances for one of them to describe a field the file no longer has. Widened rather than exact: a
 * consumer names what it needs and the compiler still checks the shape of what it names.
 */
export interface PackageJson {
  readonly name: string;
  readonly description: string;
  readonly author: string;
  readonly license: string;
  readonly dependencies: Readonly<Record<string, string>>;
  readonly repository: {readonly url: string};
  readonly exports: Readonly<Record<string, string | {readonly default?: string}>>;
}

export function readPackageJson(): PackageJson {
  return JSON.parse(readFileSync(fromRoot('package.json'), 'utf8')) as PackageJson;
}

/**
 * A fresh scratch directory under the repo's own `.tmp/`, and never under the system temp.
 *
 * CLAUDE.md puts scratch in `.tmp/` so it is inspectable and already git-ignored, and three callers
 * reached for `os.tmpdir()` instead. `$TMPDIR` is pointed at `.tmp/` only in an agent's environment,
 * so on a CI runner or a plain shell that constraint silently did not apply at all. One of them named
 * its file by process id with no `mkdtemp`, which is a stray `.xlsx` left in system temp forever
 * whenever the test throws before its cleanup, and a collision between two concurrent checkouts.
 */
export function scratchDir(prefix: string): string {
  const base = fromRoot('.tmp');
  mkdirSync(base, {recursive: true});
  return mkdtempSync(join(base, `${prefix}-`));
}

/**
 * A tool's own package entrypoint - the file `node_modules/.bin` generates its shim from, and the
 * one to hand `node` directly.
 *
 * Never the `.bin` entry itself. On Windows that is a `.cmd`, which Node will not spawn without
 * `shell: true`, and a shell brings quoting rules that differ per platform into a path where every
 * argument is a file name this repo controls. Going to the entrypoint also skips a process: the
 * shim costs ~0.3 s per invocation over the direct call, measured on a `--version` that does no
 * work at all, which on a pre-commit hook dwarfs the lint it is wrapping.
 */
export const NODE = process.execPath;
export const TSC = fromRoot('node_modules/typescript/bin/tsc');
export const OXLINT = fromRoot('node_modules/oxlint/bin/oxlint');
export const OXFMT = fromRoot('node_modules/oxfmt/bin/oxfmt');
export const CHARCHECK = fromRoot('node_modules/charcheck/dist/cli.js');
