#!/usr/bin/env node
// The public entry points, held to their contract.
//
// `src/entries/*.ts` are the package's public faces. `package.json`'s `exports` map publishes
// them, and `src/index.ts` unions them with `export *` so the root specifier keeps carrying
// everything. Five things can go wrong silently, and each is checked here:
//
//   1. An entry file exists but nothing publishes it: dead code that reads like API.
//   2. `exports` names a subpath whose entry file is gone, so a consumer's import resolves to
//      a missing file at runtime, long after CI was green.
//   3. Two entries export the same name. This is the dangerous one: `export *` does not error on
//      an ambiguous re-export, it *drops the name*, so the symbol would vanish from the root
//      specifier with no diagnostic anywhere. Disjointness is what makes `src/index.ts` a
//      faithful union, and it is why the whole failure taxonomy is exported from `/errors`
//      alone: `UnsupportedFormatError` belongs to no single codec.
//   4. `/node` gets unioned into the root barrel after all. That entry is the one the root
//      specifier must NOT carry: it reaches `node:fs` and `node:stream`, and re-exporting it here
//      would put them back on every browser consumer's graph (ADR 0040). Absence is the contract,
//      so absence is asserted rather than merely tolerated.
//   5. The browser stub drifts. `package.json` resolves `/node` to `entries/node-unavailable.ts`
//      under a bundler's `browser` condition, and a name exported by the entry but missing from
//      the stub is a browser build that fails on an import it cannot find. The two value-export
//      lists are held equal.
//
// Reading the syntax, not the types: an entry is a pure re-export list, so every question here is
// answered by the parse tree alone. TypeScript 7 publishes no standalone parser, though. A
// SourceFile is only reachable through a project, so the tree comes from one built over
// `tsconfig.json`, which already includes `src/**/*.ts`. That costs about as much as loading the
// TypeScript 6 module used to, and no diagnostics are requested, so a tree that does not typecheck
// still gets a verdict.
//
//   node scripts/check-entries.ts

import {readdirSync, readFileSync} from 'node:fs';
import {dirname, join, resolve} from 'node:path';
import {fileURLToPath} from 'node:url';

import * as ast from 'typescript/unstable/ast';
import {API, type Project} from 'typescript/unstable/sync';

const ROOT = resolve(dirname(fileURLToPath(import.meta.url)), '..');
const CONFIG = join(ROOT, 'tsconfig.json');
const ENTRY_DIR = 'src/entries';
const BARREL = 'src/index.ts';
// The Node-only entry, and the module a browser condition resolves it to instead. Neither is an
// ordinary barrel: the first is published but not unioned, the second is not published at all.
const NODE_ENTRY = 'src/entries/node.ts';
const NODE_STUB = 'src/entries/node-unavailable.ts';

/** The names an entry barrel re-exports. Entries hold nothing but `export {…} from '…'`. */
function exportedNames(source: ast.SourceFile): string[] {
  const names: string[] = [];
  for (const statement of source.statements) {
    if (!ast.isExportDeclaration(statement)) continue;
    const clause = statement.exportClause;
    if (clause === undefined || !ast.isNamedExports(clause)) continue;
    for (const element of clause.elements) names.push(element.name.text);
  }
  return names;
}

/**
 * The names a module exports as *values*, which are the only ones a stub has to stand in for: the
 * `types` condition is never switched, so a type-only export resolves to the real declarations in
 * either environment and needs no runtime counterpart.
 */
function exportedValueNames(source: ast.SourceFile): Set<string> {
  const names = new Set<string>();
  for (const statement of source.statements) {
    if (ast.isClassDeclaration(statement) || ast.isFunctionDeclaration(statement)) {
      const exported = statement.modifiers?.some((m) => m.kind === ast.SyntaxKind.ExportKeyword);
      if (exported === true && statement.name) names.add(statement.name.text);
      continue;
    }
    if (!ast.isExportDeclaration(statement) || statement.isTypeOnly) continue;
    const clause = statement.exportClause;
    if (clause === undefined || !ast.isNamedExports(clause)) continue;
    for (const element of clause.elements) {
      if (!element.isTypeOnly) names.add(element.name.text);
    }
  }
  return names;
}

/** The entry modules `src/index.ts` unions, as repo-relative paths. */
function starExportedEntries(source: ast.SourceFile): string[] {
  const paths: string[] = [];
  for (const statement of source.statements) {
    if (!ast.isExportDeclaration(statement) || statement.exportClause !== undefined) continue;
    const specifier = statement.moduleSpecifier;
    if (specifier === undefined || !ast.isStringLiteral(specifier)) continue;
    paths.push(`src/${specifier.text.replace(/^\.\//, '')}`);
  }
  return paths;
}

/** A subpath's target: a bare file (`./package.json` is one) or the conditions object entries use. */
type ExportTarget = string | {readonly default?: string};

interface PackageJson {
  readonly exports: Readonly<Record<string, ExportTarget>>;
}

/** The subpath → entry-source mapping `package.json` publishes, with `dist/*.js` read back to `src/*.ts`. */
function publishedEntries(): Map<string, string> {
  const pkg = JSON.parse(readFileSync(join(ROOT, 'package.json'), 'utf8')) as PackageJson;
  const map = new Map<string, string>();
  for (const [subpath, target] of Object.entries(pkg.exports)) {
    const emitted = typeof target === 'string' ? undefined : target.default;
    if (emitted === undefined || !emitted.endsWith('.js')) continue;
    map.set(subpath, emitted.replace(/^\.\/dist\//, 'src/').replace(/\.js$/, '.ts'));
  }
  return map;
}

function main(project: Project): void {
  const parse = (file: string): ast.SourceFile => {
    const source = project.program.getSourceFile(join(ROOT, file));
    if (!source) throw new Error(`${file} is not in the program; is it covered by tsconfig.json?`);
    return source;
  };

  const problems: string[] = [];

  const onDisk = readdirSync(join(ROOT, ENTRY_DIR))
    .filter((name) => name.endsWith('.ts') && !name.endsWith('.test.ts'))
    .map((name) => `${ENTRY_DIR}/${name}`)
    .filter((file) => file !== NODE_STUB)
    .sort();

  const published = publishedEntries();
  const publishedSources = new Set(published.values());
  for (const file of onDisk) {
    if (!publishedSources.has(file)) {
      problems.push(`  ${file} is an entry barrel that package.json "exports" does not publish`);
    }
  }
  for (const [subpath, source] of published) {
    if (source === BARREL) continue;
    if (!onDisk.includes(source)) {
      problems.push(`  package.json publishes "${subpath}" but ${source} does not exist`);
    }
  }

  const unioned = starExportedEntries(parse(BARREL));
  for (const file of onDisk) {
    if (file === NODE_ENTRY) {
      if (unioned.includes(file)) {
        problems.push(
          `  ${BARREL} unions ${file}
` +
            `    that entry reaches node:fs and node:stream, so the root specifier must not carry it (ADR 0040)`,
        );
      }
      continue;
    }
    if (!unioned.includes(file)) {
      problems.push(`  ${BARREL} does not \`export *\` from ${file}; the root specifier loses it`);
    }
  }

  const stubbed = exportedValueNames(parse(NODE_STUB));
  for (const name of exportedValueNames(parse(NODE_ENTRY))) {
    if (!stubbed.has(name)) {
      problems.push(
        `  ${NODE_ENTRY} exports the value "${name}" and ${NODE_STUB} does not
` + `    a browser build resolves the subpath to the stub, so the import would fail to link`,
      );
    }
  }

  const owner = new Map<string, string>();
  let total = 0;
  for (const file of onDisk) {
    for (const name of exportedNames(parse(file))) {
      total += 1;
      const first = owner.get(name);
      if (first === undefined) owner.set(name, file);
      else {
        problems.push(
          `  "${name}" is exported by both ${first} and ${file}\n` +
            `    \`export *\` resolves that ambiguity by dropping the name, so it would disappear from ${BARREL}`,
        );
      }
    }
  }

  if (problems.length === 0) {
    console.log(
      `entries: ${onDisk.length} public faces, ${total} disjoint exports, all published; ` +
        `${onDisk.length - 1} unioned into the root barrel and /node held out of it`,
    );
  } else {
    console.error(`\nentries: ${problems.length} problem(s) with the public entry points.\n`);
    console.error(`${problems.join('\n\n')}\n`);
    process.exitCode = 1;
  }
}

// The compiler is a separate process; the snapshot and the server both have to be handed back, or a
// failing check would leave a tsgo behind and this process hanging on its pipe. `process.exitCode`
// rather than `process.exit` above, so the close still runs before the exit is taken.
const api = new API({cwd: ROOT});
try {
  const snapshot = api.updateSnapshot({openProjects: [CONFIG]});
  try {
    const project = snapshot.getProject(CONFIG);
    if (!project) throw new Error(`cannot open project ${CONFIG}`);
    main(project);
  } finally {
    snapshot.dispose();
  }
} finally {
  api.close();
}
