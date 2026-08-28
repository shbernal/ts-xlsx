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

import {readdirSync, readFileSync} from 'node:fs';
import {dirname, join, resolve} from 'node:path';
import {fileURLToPath} from 'node:url';

const ROOT = resolve(dirname(fileURLToPath(import.meta.url)), '..');

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

const SPECIFIER = /\b(?:from|import)\s+'([^']*)'/g;

/** Every specifier the module imports or re-exports from: relative ones resolved, the rest as written. */
function imports(file: string): {relative: string[]; bare: string[]} {
  const source = readFileSync(join(ROOT, file), 'utf8');
  const relative: string[] = [];
  const bare: string[] = [];
  const dir = file.slice(0, file.lastIndexOf('/'));
  for (const match of source.matchAll(SPECIFIER)) {
    const specifier = match[1] as string;
    if (!specifier.startsWith('.')) {
      bare.push(specifier);
      continue;
    }
    const out: string[] = [];
    for (const segment of `${dir}/${specifier}`.split('/')) {
      if (segment === '' || segment === '.') continue;
      if (segment === '..') out.pop();
      else out.push(segment);
    }
    relative.push(out.join('/'));
  }
  return {relative, bare};
}

/** Every module that has to be present for `entry` to evaluate, itself included. */
function closure(entry: string): string[] {
  const reached = new Set<string>();
  const pending = [entry];
  while (pending.length > 0) {
    const file = pending.pop() as string;
    if (reached.has(file)) continue;
    reached.add(file);
    pending.push(...imports(file).relative);
  }
  return [...reached].sort();
}

/**
 * The module's code with its comments blanked out, newlines kept so a match still reports its own
 * line. Prose is where these identifiers legitimately appear (this file's own header names three
 * of them), so a scan that read comments would report nothing but itself. Strings are walked
 * rather than skipped so that a `//` inside one is not mistaken for the start of a comment.
 */
function withoutComments(source: string): string {
  let out = '';
  for (let i = 0; i < source.length; i++) {
    const two = source.slice(i, i + 2);
    if (two === '//') {
      const end = source.indexOf('\n', i);
      const stop = end === -1 ? source.length : end;
      out += ' '.repeat(stop - i);
      i = stop - 1;
      continue;
    }
    if (two === '/*') {
      const end = source.indexOf('*/', i + 2);
      const stop = end === -1 ? source.length : end + 2;
      out += source.slice(i, stop).replace(/[^\n]/g, ' ');
      i = stop - 1;
      continue;
    }
    const char = source[i];
    out += char;
    if (char !== "'" && char !== '"' && char !== '`') continue;
    // Inside a string literal: copy to its close, honouring backslash escapes.
    for (i += 1; i < source.length; i++) {
      const inner = source[i] as string;
      out += inner;
      if (inner === '\\') {
        out += source[i + 1] ?? '';
        i += 1;
        continue;
      }
      if (inner === char) break;
    }
  }
  return out;
}

const problems: string[] = [];
// file -> the first browser entry that reaches it, so a violation is reported once and still names
// a specifier a consumer would recognise.
const reachable = new Map<string, string>();

for (const entry of BROWSER_ENTRIES) {
  for (const file of closure(entry)) if (!reachable.has(file)) reachable.set(file, entry);
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
const nodeOnly = closure(NODE_ENTRY).filter((file) => !reachable.has(file));
const carried = nodeOnly.flatMap((file) => imports(file).bare.filter((s) => s.startsWith('node:')));
if (carried.length === 0) {
  problems.push(
    `  ${NODE_ENTRY}\n    reaches no Node built-in that the browser entries do not\n` +
      `    it exists only to hold that boundary; fold it back into /xlsx if there is nothing left to hold`,
  );
}

// A directory read, purely so the entry list above cannot silently fall behind the files on disk.
const unclassified = readdirSync(join(ROOT, 'src/entries'))
  .filter((name) => name.endsWith('.ts') && !name.endsWith('.test.ts'))
  .map((name) => `src/entries/${name}`)
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

if (problems.length === 0) {
  console.log(
    `browser-safe: ${reachable.size} modules reachable from ${BROWSER_ENTRIES.length} entries, ` +
      `no Node built-in and no Node global; ${NODE_ENTRY} carries ${[...new Set(carried)].sort().join(', ')}`,
  );
} else {
  console.error(`\nbrowser-safe: ${problems.length} break(s) in the browser boundary.\n`);
  console.error(`${problems.join('\n\n')}\n`);
  process.exit(1);
}
