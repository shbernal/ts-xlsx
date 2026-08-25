#!/usr/bin/env node
// The one place the formatter's file set is defined.
//
// Usage:
//   node scripts/format.ts --write    rewrite what is not already formatted
//   node scripts/format.ts --check    exit non-zero if anything would change
//
// The glob list exists because bare `oxfmt` considers a *wider* set than this repo
// formats — markdown, JSON, YAML — and a formatter that silently reaches further than
// you think is how a generated file gets rewritten. `.oxfmtrc.jsonc`'s `ignorePatterns`
// is the second, overlapping defence.
//
// It lives in a script rather than inline in package.json because three callers need it
// (`format`, `format:check`, and the format gate in verify.ts) and a glob list copied
// three times is a glob list that drifts twice.
//
// oxfmt expands the globs itself; they are never handed to a shell, so `**` means the
// same thing on every platform. Invoked as `node <entrypoint>` rather than through
// `pnpm exec` or a .bin shim for the reason lefthook.yml records — the wrapper costs
// more than the work.

import {spawn} from 'node:child_process';
import {dirname, resolve} from 'node:path';
import {fileURLToPath} from 'node:url';

const ROOT = resolve(dirname(fileURLToPath(import.meta.url)), '..');
const OXFMT = resolve(ROOT, 'node_modules/oxfmt/bin/oxfmt');

/** Must stay in step with the `include` list tsconfig.json and tsconfig.test.json span. */
const TARGETS = [
  'src/**/*.ts',
  'scripts/**/*.ts',
  'test/**/*.ts',
  'tools/**/*.ts',
  'charcheck.config.ts',
];

const mode = process.argv[2];
if (mode !== '--write' && mode !== '--check') {
  console.error('usage: node scripts/format.ts --write | --check');
  process.exit(2);
}

const child = spawn(process.execPath, [OXFMT, mode, ...TARGETS], {
  cwd: ROOT,
  stdio: 'inherit',
});
child.on('exit', (code, signal) => process.exit(signal !== null ? 1 : (code ?? 1)));
