#!/usr/bin/env node
// `pnpm run lint`, spawning oxlint over the one target list.
//
// A thin wrapper for the reason `format.ts` and `build.ts` are: the arguments were spelled out in
// `package.json` and again in `scripts/verify.ts`, whose comment asked the next editor to keep the
// two in step. A JSON file cannot import a constant, so the way to have one list is for the script
// to be TypeScript. `--fix` and any other flag pass through.
//
// Invoked as `node <entrypoint>` rather than through a `.bin` shim, for the reason `repo.ts` records:
// on Windows the shim is a `.cmd` that needs a shell, and the shim costs ~0.3 s per invocation.
//
//   node scripts/lint.ts [--fix]

import {spawn} from 'node:child_process';

import {NODE, OXLINT, ROOT} from './repo.ts';
import {LINT_STRICT, LINT_TARGETS, LINT_TYPE_AWARE, LINT_UNUSED_DIRECTIVES} from './targets.ts';

const child = spawn(
  NODE,
  [
    OXLINT,
    ...LINT_TARGETS,
    LINT_TYPE_AWARE,
    LINT_STRICT,
    LINT_UNUSED_DIRECTIVES,
    ...process.argv.slice(2),
  ],
  {cwd: ROOT, stdio: 'inherit'},
);
child.on('exit', (code, signal) => process.exit(signal !== null ? 1 : (code ?? 1)));
