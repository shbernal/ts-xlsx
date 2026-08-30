#!/usr/bin/env node
// The publishable build: clear `dist/`, emit the runtime JS, then emit the declarations.
//
// A script rather than a chain in `package.json` because `verify --full` runs the build too (the
// size budgets can only be measured on emitted JS, and a budget that only CI can check is one a
// push discovers rather than a commit). Verify spawns its tools as `node <entrypoint>`, never
// through the package manager: a `.bin` entry on Windows is a `.cmd` shim, which would force
// `shell: true` and with it per-platform quoting. So the two callers can only share *one*
// definition of what a build is if that definition is a file both can run.
//
// Two `tsc` passes, and the split is deliberate: the JS pass strips comments (ADR 0001, amended),
// the declarations pass puts the JSDoc back for the audience that reads types. See the
// `//removeComments` note in `tsconfig.build.json`.
//
//   node scripts/build.ts

import {spawnSync} from 'node:child_process';
import {rmSync} from 'node:fs';
import {dirname, join, resolve} from 'node:path';
import {fileURLToPath} from 'node:url';

const ROOT = resolve(dirname(fileURLToPath(import.meta.url)), '..');
const TSC = resolve(ROOT, 'node_modules/typescript/bin/tsc');

rmSync(join(ROOT, 'dist'), {recursive: true, force: true});

for (const project of ['tsconfig.build.json', 'tsconfig.build.dts.json']) {
  const {status} = spawnSync(process.execPath, [TSC, '-p', project], {
    cwd: ROOT,
    stdio: 'inherit',
  });
  if (status !== 0) {
    console.error(`build: tsc -p ${project} exited ${status ?? 'on a signal'}`);
    process.exit(status ?? 1);
  }
}
