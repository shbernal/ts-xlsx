#!/usr/bin/env node
// `pnpm run test:src`, over the one definition of what the unit suite is.
//
// The glob was written out three times: here, in `verify.ts`'s gate list, and in `coverage.ts`'s
// suite list. Two of those decide what a gate runs and the third decides what the coverage number
// is *of*, so a suffix added to one and not the others produces a green gate over a smaller suite
// and a coverage figure that agrees with neither.
//
// The glob is passed to `node --test` and never to a shell, so `**` means the same thing on every
// platform. Extra arguments pass through, which is how `--test-name-pattern` reaches it.
//
//   node scripts/test-src.ts [--test-name-pattern …]

import {spawn} from 'node:child_process';

import {NODE, ROOT} from './repo.ts';
import {UNIT_SUITE_GLOB} from './targets.ts';

const child = spawn(NODE, ['--test', ...process.argv.slice(2), UNIT_SUITE_GLOB], {
  cwd: ROOT,
  stdio: 'inherit',
});
child.on('exit', (code, signal) => process.exit(signal !== null ? 1 : (code ?? 1)));
