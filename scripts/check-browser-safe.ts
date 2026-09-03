#!/usr/bin/env node
// The browser boundary, enforced on the module graph.
//
// A bundler resolves imports, not call graphs. One `import {createHash} from 'node:crypto'` in a
// module nothing on the buffered path ever calls is still enough to put Node's crypto into every
// browser build that touches the package root. That is exactly what happened here: three Node
// built-ins were reachable from `src/index.ts` (a password hash and the streaming writer), and a
// consumer bundling for a tab got warnings, a broken chunk, or a build failure depending on which
// bundler they used, with nothing in the package to tell them why.
//
// `docs/knowledge/specs/browser-safe-io-boundary.md` asks for the opposite arrangement, and ADR
// 0038 records how it is built: everything reachable from the root specifier is environment-neutral,
// and the one Node-only face lives behind `@shbernal/ts-xlsx/node`, which a browser condition
// resolves to a module that throws by name. That arrangement is one careless import away from
// being untrue again, and the import in question would typecheck, pass every test, and look
// entirely reasonable where it sits. This is the mechanism that says so instead.
//
// Two things are checked, over the closure of every browser-facing entry:
//
//   1. No `node:` specifier, and none of the bare names those built-ins also answer to.
//   2. No Node-only *global*: `process`, `Buffer`, `__dirname`, `global`. Those pass a bundler
//      silently and fail in the tab, which is worse: upstream shipped a `process.version` read on
//      an ordinary in-memory path and made the library unusable in a browser on its first call.
//
//   node scripts/check-browser-safe.ts

import {readFileSync} from 'node:fs';
import {join} from 'node:path';

import {
  closure,
  resolveSpecifier,
  sourceFiles,
  specifiers,
  withoutComments,
} from './module-graph.ts';
import {ROOT} from './repo.ts';
import {verdict} from './verdict.ts';

// The root barrel is the entry most consumers name, and each browser-facing subpath is checked in
// its own right: a subpath's closure is a subset of the root's today, and nothing guarantees that
// stays true as entries move.
const BROWSER_ENTRIES = [
  'src/index.ts',
  'src/entries/core.ts',
  'src/entries/csv.ts',
  'src/entries/customui.ts',
  'src/entries/errors.ts',
  'src/entries/vba.ts',
  'src/entries/xlsb.ts',
  'src/entries/xlsx.ts',
];

/** The Node-only entry, whose whole reason for existing is that this check would fail on it. */
const NODE_ENTRY = 'src/entries/node.ts';

// `node:fs` is the modern spelling; `fs` is the same module and resolves identically in Node, so a
// rule that matched only the prefixed form would be blind to half the ways in.
const BUILTINS = [
  'assert',
  'buffer',
  'child_process',
  'crypto',
  'events',
  'fs',
  'http',
  'https',
  'module',
  'net',
  'os',
  'path',
  'process',
  'querystring',
  'readline',
  'stream',
  'string_decoder',
  'timers',
  'tls',
  'tty',
  'url',
  'util',
  'v8',
  'vm',
  'worker_threads',
  'zlib',
];

// Only these: a name that also exists in a browser (`crypto`, `TextEncoder`, `URL`) is not evidence
// of anything, and listing it would make the check lie. Each is matched as a whole identifier, so
// `ArrayBuffer` and `globalThis` do not trip the two that would otherwise look like prefixes.
const NODE_GLOBALS: readonly RegExp[] = [
  /(?<![\w$.])process\s*\./,
  /(?<![\w$.])Buffer(?![\w$])/,
  /(?<![\w$.])__dirname(?![\w$])/,
  /(?<![\w$.])global(?![\w$])/,
];

/**
 * Every specifier the module imports or re-exports from: relative ones resolved, the rest as written.
 * A bare specifier is what this gate is *for*, so both halves are wanted here where the other
 * consumers of the shared walker only ever need the relative one.
 */
function imports(file: string): {relative: string[]; bare: string[]} {
  const relative: string[] = [];
  const bare: string[] = [];
  for (const specifier of specifiers(readFileSync(join(ROOT, file), 'utf8'))) {
    if (specifier.startsWith('.')) relative.push(resolveSpecifier(file, specifier));
    else bare.push(specifier);
  }
  return {relative, bare};
}

/** Every module that has to be present for `entry` to evaluate, itself included. */
function reachableFrom(entry: string): string[] {
  return [...closure(entry, (file) => imports(file).relative)].sort();
}

const problems: string[] = [];
// file -> the first browser entry that reaches it, so a violation is reported once and still names
// a specifier a consumer would recognise.
const reachable = new Map<string, string>();

for (const entry of BROWSER_ENTRIES) {
  for (const file of reachableFrom(entry)) if (!reachable.has(file)) reachable.set(file, entry);
}

for (const [file, entry] of [...reachable].sort(([a], [b]) => a.localeCompare(b))) {
  for (const specifier of imports(file).bare) {
    const name = specifier.replace(/^node:/, '').split('/')[0] as string;
    if (specifier.startsWith('node:') || BUILTINS.includes(name)) {
      problems.push(
        `  ${file}\n    imports ${specifier}\n` +
          `    it is reachable from ${entry}, which a browser bundles: publish it from ${NODE_ENTRY} instead`,
      );
    }
  }
}

for (const file of [...reachable.keys()].sort()) {
  const source = withoutComments(readFileSync(join(ROOT, file), 'utf8'));
  for (const pattern of NODE_GLOBALS) {
    const match = pattern.exec(source);
    if (match === null) continue;
    const line = source.slice(0, match.index).split('\n').length;
    problems.push(
      `  ${file}:${line}\n    reads the Node-only global \`${match[0].trim()}\`\n` +
        `    a bundler will not warn and it is undefined in a tab: use the platform's own API`,
    );
  }
}

// The Node-only entry is checked from the other side: it exists to carry the imports the browser
// entries may not, so an empty one would mean the boundary had quietly moved rather than held.
const nodeOnly = reachableFrom(NODE_ENTRY).filter((file) => !reachable.has(file));
const carried = nodeOnly.flatMap((file) => imports(file).bare.filter((s) => s.startsWith('node:')));
if (carried.length === 0) {
  problems.push(
    `  ${NODE_ENTRY}\n    reaches no Node built-in that the browser entries do not\n` +
      `    it exists only to hold that boundary; fold it back into /xlsx if there is nothing left to hold`,
  );
}

// A directory read, purely so the entry list above cannot silently fall behind the files on disk.
const unclassified = sourceFiles(`${ROOT}/src/entries`, '.ts')
  .map((path) => `src/entries/${path.slice(path.lastIndexOf('/') + 1)}`)
  .filter(
    (file) =>
      !BROWSER_ENTRIES.includes(file) && file !== NODE_ENTRY && !file.endsWith('-unavailable.ts'),
  );
for (const file of unclassified) {
  problems.push(
    `  ${file}\n    is an entry this check has no verdict on\n` +
      `    add it to BROWSER_ENTRIES, or publish it from ${NODE_ENTRY} if it is Node-only`,
  );
}

verdict({
  gate: 'browser-safe',
  problems,
  ok:
    `${reachable.size} modules reachable from ${BROWSER_ENTRIES.length} entries, no Node built-in ` +
    `and no Node global; ${NODE_ENTRY} carries ${[...new Set(carried)].sort().join(', ')}`,
  failure: 'break(s) in the browser boundary',
});
