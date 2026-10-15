#!/usr/bin/env node
// Every tree the harness gate spans has a `tsconfig.json` claiming it.
//
// The type-aware lint rules resolve compiler options from the nearest `tsconfig.json` that
// *includes* the file being linted. `tsconfig.test.json` is a non-standard name and is invisible to
// that search, and the root `tsconfig.json` includes only `src/**`, so a tree with no
// `tsconfig.json` of its own is linted under TypeScript's defaults: no `noUncheckedIndexedAccess`,
// no `exactOptionalPropertyTypes`. A type-aware verdict decided by something other than this repo's
// settings is not a verdict about this repo, and it is silent: the rules still run, and still pass.
//
// `test/`, `scripts/` and `tools/` each carry one, and each explains itself at length. What none of
// them could do is notice a fourth tree being added to `tsconfig.test.json`'s `include` without one.
// The playbook's answer was a manual plant-a-control procedure. This is the automatic half.
//
// The `incremental: false` / `tsBuildInfoFile: null` pair is checked too: nothing runs `tsc` against
// these files, so a build-info cache written from here could only race the gate's.
//
//   node scripts/check-tsconfig-coverage.ts

import {existsSync, readFileSync} from 'node:fs';
import {join} from 'node:path';

import {ROOT} from './repo.ts';
import {HARNESS_TREES} from './targets.ts';
import {reportCrash, verdict} from './verdict.ts';

const HARNESS_CONFIG = 'tsconfig.test.json';

interface TsConfig {
  readonly extends?: string;
  readonly include?: readonly string[];
  readonly compilerOptions?: {readonly incremental?: boolean; readonly tsBuildInfoFile?: unknown};
}

function readConfig(path: string): TsConfig {
  return JSON.parse(readFileSync(path, 'utf8')) as TsConfig;
}

function main(): void {
  const problems: string[] = [];

  // The list this gate is about is `targets.ts`'s, and the harness config's `include` is what it has
  // to agree with: a tree in one and not the other is the drift, whichever way round.
  const spanned = new Set(
    (readConfig(join(ROOT, HARNESS_CONFIG)).include ?? [])
      .map((pattern) => pattern.split('/')[0] ?? '')
      .filter((segment) => segment !== '' && !segment.includes('.')),
  );
  for (const tree of HARNESS_TREES) {
    if (!spanned.has(tree)) {
      problems.push(
        `  ${tree}\n    is named in targets.ts's HARNESS_TREES and not in ${HARNESS_CONFIG}'s include`,
      );
    }
  }
  for (const tree of spanned) {
    if (!(HARNESS_TREES as readonly string[]).includes(tree)) {
      problems.push(
        `  ${tree}\n    is in ${HARNESS_CONFIG}'s include and not in targets.ts's HARNESS_TREES:\n` +
          '    add it there, so this gate and the formatter and the linter all see it',
      );
    }
  }

  for (const tree of HARNESS_TREES) {
    const path = join(ROOT, tree, 'tsconfig.json');
    if (!existsSync(path)) {
      problems.push(
        `  ${tree}/tsconfig.json\n    does not exist: the type-aware linter would judge ${tree}/ under\n` +
          "    TypeScript's defaults rather than this repo's settings, silently",
      );
      continue;
    }
    const config = readConfig(path);
    if (config.extends !== `../${HARNESS_CONFIG}`) {
      problems.push(
        `  ${tree}/tsconfig.json\n    extends ${JSON.stringify(config.extends)} rather than "../${HARNESS_CONFIG}":\n` +
          '    the options must come from the gate, or the linter and tsc disagree about this tree',
      );
    }
    if (config.compilerOptions?.incremental !== false) {
      problems.push(
        `  ${tree}/tsconfig.json\n    does not set incremental: false: nothing runs tsc against it, so a\n` +
          "    build-info cache written from here could only race the gate's",
      );
    }
    if (config.compilerOptions?.tsBuildInfoFile !== null) {
      problems.push(
        `  ${tree}/tsconfig.json\n    does not set tsBuildInfoFile: null, which is the other half of that`,
      );
    }
  }

  verdict({
    gate: 'tsconfig coverage',
    problems,
    ok: `${HARNESS_TREES.length} harness trees each claimed by their own tsconfig.json`,
    failure: "problem(s) in the harness's promises about itself",
  });
}

try {
  main();
} catch (error) {
  reportCrash('tsconfig coverage', error);
}
