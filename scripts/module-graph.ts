// What a module imports, and everything a module reaches, for the four gates that ask.
//
// `check-layering`, `check-browser-safe`, `size-budget` and `smoke-dist` each walk the module graph,
// and each had grown its own copy of the walk: four file listers, four specifier regexes, two
// byte-identical path-segment collapse loops and two byte-identical closure functions.
//
// The four regexes did not agree, and the disagreement was live rather than theoretical.
// `size-budget` matched either quote style and recorded *why* in a comment: the emitter's quoting
// flipped between TypeScript 6 and 7. `check-layering` and `check-browser-safe` matched single quotes
// only, so the same flip in the *source* formatter would have made both gates see zero imports and
// report a clean graph. A gate that silently stops checking is worse than no gate, and these two are
// named in CLAUDE.md as part of what makes the machine-checkable net the primary guarantor of
// correctness. Four answers to "what does this module import" is four definitions of the graph that
// net is guaranteeing.
//
// None of the four stripped comments before matching either, so a commented-out import read as a real
// one: a false layering violation, from prose.
//
// Paths here are always `/`-separated, whatever the platform, because a graph keyed by path needs one
// spelling per file: half these callers build repo-relative strings by hand and half use `node:path`,
// and on Windows the two disagree on the separator, which would count one module twice.

import {readdirSync, readFileSync} from 'node:fs';

/** A path with `/` separators, whatever `node:path` produced on this platform. */
export function toPosix(path: string): string {
  return path.replaceAll('\\', '/');
}

/**
 * Every file under `dir` (recursively) whose name ends with `suffix`, as `/`-separated paths built
 * from `dir` as given. Tests are not source: a `.test.ts` is excluded from a `.ts` listing, because
 * every caller here asks about what ships.
 */
export function sourceFiles(dir: string, suffix: string): string[] {
  return readdirSync(dir, {withFileTypes: true}).flatMap((entry) => {
    const path = `${toPosix(dir)}/${entry.name}`;
    if (entry.isDirectory()) return sourceFiles(path, suffix);
    if (!entry.name.endsWith(suffix)) return [];
    return suffix === '.ts' && entry.name.endsWith('.test.ts') ? [] : [path];
  });
}

// Both spellings, and both quote styles. `… from '…'` is the ordinary form; the bare `import '…'` has
// no `from` to anchor on and should never appear here, since the package declares `"sideEffects":
// false` and an import kept only for its effect is a lie to every bundler. But a rule that cannot see
// it would report a clean graph while a layer was being crossed by the one form it was blind to.
const SPECIFIER = /\b(?:from|import)\s+["']([^"']*)["']/g;

/** Every specifier a module imports or re-exports from, in source order, comments excluded. */
export function specifiers(source: string): string[] {
  return [...withoutComments(source).matchAll(SPECIFIER)].map((match) => match[1] ?? '');
}

/** Only the relative specifiers: a bare one names a dependency, which is not part of this graph. */
export function relativeSpecifiers(source: string): string[] {
  return specifiers(source).filter((specifier) => specifier.startsWith('.'));
}

/**
 * Resolve a relative specifier against the module that wrote it, collapsing `.` and `..`.
 *
 * Deliberately not `node:path`: these paths are graph keys rather than filesystem arguments, and
 * `join`/`normalize` would hand back `\` on Windows for callers that built their side with `/`.
 */
export function resolveSpecifier(fromFile: string, specifier: string): string {
  const from = toPosix(fromFile);
  const dir = from.slice(0, from.lastIndexOf('/'));
  const out: string[] = [];
  for (const segment of `${dir}/${specifier}`.split('/')) {
    if (segment === '' || segment === '.') continue;
    if (segment === '..') out.pop();
    else out.push(segment);
  }
  // A leading `/` (or a drive-relative Windows path) survives the split as an empty first segment,
  // which the loop drops; put it back so an absolute path stays absolute.
  return (from.startsWith('/') ? '/' : '') + out.join('/');
}

/** Every relative import of `file`, read off disk and resolved. */
export function importedPaths(file: string): string[] {
  return relativeSpecifiers(readFileSync(file, 'utf8')).map((specifier) =>
    resolveSpecifier(file, specifier),
  );
}

/**
 * Every module that has to be present for `entry` to evaluate, `entry` included.
 *
 * `imports` is a parameter because the two source gates read repo-relative paths through a root and
 * the two `dist` gates read real ones; what they share is the walk, not the lookup.
 */
export function closure(entry: string, imports: (file: string) => string[]): Set<string> {
  const reached = new Set<string>();
  const pending = [toPosix(entry)];
  while (pending.length > 0) {
    const file = pending.pop() as string;
    if (reached.has(file)) continue;
    reached.add(file);
    pending.push(...imports(file));
  }
  return reached;
}

/**
 * A module's code with its comments blanked out, newlines kept so a match still reports its own line.
 *
 * Strings are walked rather than skipped so a `//` inside one is not mistaken for the start of a
 * comment. Exported because `check-browser-safe` scans the same blanked source for the Node globals
 * it forbids, for the same reason: prose is where those identifiers legitimately appear.
 */
export function withoutComments(source: string): string {
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
