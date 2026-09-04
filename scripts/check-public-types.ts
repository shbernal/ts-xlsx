#!/usr/bin/env node
// Every type a consumer can reach, reachable by name.
//
// `check-entries.ts` gates *disjointness*, because `export *` drops an ambiguous name with no
// diagnostic. This gates the dual, which is just as silent: a type named in a published signature
// that no entry barrel exports. `TableColumn` is published and its `totalsRowFunction` is a
// `TotalsRowFunction`, so a consumer could write `const c: TableColumn` and had no way to name the
// type of `c.totalsRowFunction`. Nothing reported it: the emitted `.d.ts` typechecks either way,
// because the declarations import each other by relative path regardless of what `exports` publishes,
// and `docs:check` regenerates from the barrel, so it sees only what the barrel already lists.
//
// The walk starts at each entry's exports, follows every type reference out of their declarations,
// and reports any it reaches that is exported from `src/` and from no entry. It is transitive: a
// type reached only through another unpublished one is reported too, so one run names the whole
// closure rather than one layer of it per fix.
//
// Types *not* exported from their own module are out of scope here. A public signature naming one is
// also a defect, but a different one, and `typecheck:dist` is what sees it: the emitted declaration
// cannot import a name its module does not export.
//
// A declaration may decline to be published by carrying `@unpublished` in its doc comment with the
// reason after it. That is a decision on the record, which is the difference between a shape held
// back while its surface settles and one nobody noticed; the run reports how many were declined, so
// the count is a number someone reads rather than an absence. A declined type is not walked into
// either: what is reachable only through something a consumer cannot name is not reachable.
//
//   node scripts/check-public-types.ts

import {readdirSync} from 'node:fs';
import {join} from 'node:path';

import * as ast from 'typescript/unstable/ast';
import {API, type Project} from 'typescript/unstable/sync';

import {sourceFiles, toPosix} from './module-graph.ts';
import {ROOT} from './repo.ts';
import {verdict} from './verdict.ts';

const CONFIG = join(ROOT, 'tsconfig.json');
const ENTRY_DIR = 'src/entries';

/** A declaration a module exports under a name, with the file it lives in for the report. */
interface Declared {
  readonly file: string;
  readonly node: ast.Node;
  /** Whether its doc comment declines publication with `@unpublished`. */
  readonly declined: boolean;
}

const DECLINED = /@unpublished\b/;

function declinesPublication(statement: ast.Statement): boolean {
  return DECLINED.test(statement.getFullText());
}

// The declaration kinds a type reference can resolve to. A value export (a function, a `const`) is
// indexed too: its own signature is part of the public graph and is walked the same way.
function declaredName(statement: ast.Statement): string | undefined {
  if (
    ast.isInterfaceDeclaration(statement) ||
    ast.isTypeAliasDeclaration(statement) ||
    ast.isClassDeclaration(statement) ||
    ast.isEnumDeclaration(statement) ||
    ast.isFunctionDeclaration(statement)
  ) {
    return statement.name?.text;
  }
  return undefined;
}

function isExported(statement: ast.Statement): boolean {
  return statement.getText().startsWith('export');
}

/** Every name `src/` exports, outside the entry barrels, indexed to where it is declared. */
function declarationIndex(project: Project): Map<string, Declared> {
  const index = new Map<string, Declared>();
  for (const absolute of sourceFiles(join(ROOT, 'src'), '.ts')) {
    const file = toPosix(absolute).slice(toPosix(ROOT).length + 1);
    if (file.startsWith(`${ENTRY_DIR}/`) || file.endsWith('.test.ts')) continue;
    const source = project.program.getSourceFile(absolute);
    if (!source) continue;
    for (const statement of source.statements) {
      if (!isExported(statement)) continue;
      const name = declaredName(statement);
      // First declaration wins. A name declared in two modules is `check-entries.ts`'s problem when
      // both are published, and cannot be this one's: either declaration answers "is it nameable".
      const declined = declinesPublication(statement);
      if (name !== undefined && !index.has(name)) {
        index.set(name, {file, node: statement, declined});
      }
      if (ast.isVariableStatement(statement)) {
        for (const declaration of statement.declarationList.declarations) {
          if (ast.isIdentifier(declaration.name) && !index.has(declaration.name.text)) {
            index.set(declaration.name.text, {file, node: statement, declined});
          }
        }
      }
    }
  }
  return index;
}

/** Every name the entry barrels re-export, which is the package's nameable surface. */
function publishedNames(project: Project): Set<string> {
  const names = new Set<string>();
  for (const entry of readdirSync(join(ROOT, ENTRY_DIR))) {
    if (!entry.endsWith('.ts') || entry.endsWith('.test.ts')) continue;
    const source = project.program.getSourceFile(join(ROOT, ENTRY_DIR, entry));
    if (!source) continue;
    for (const statement of source.statements) {
      if (!ast.isExportDeclaration(statement)) continue;
      const clause = statement.exportClause;
      if (clause === undefined || !ast.isNamedExports(clause)) continue;
      for (const element of clause.elements) names.add(element.name.text);
    }
  }
  return names;
}

/**
 * Whether a class member is part of what a consumer can see.
 *
 * A `#field`, a `private` one, and a member keyed by a computed name are all unreachable from
 * outside: the last is how the codec-only surface is hung off the `INTERNAL` symbol. Walking into
 * them would report every slice `Worksheet` delegates to as an unnameable public type, which is the
 * opposite of true -- they are unnameable *because* they are not public.
 */
function isVisibleMember(member: ast.Node): boolean {
  const named = member as {name?: ast.Node; modifiers?: readonly ast.Node[]};
  const name = named.name;
  if (name !== undefined && (ast.isPrivateIdentifier(name) || ast.isComputedPropertyName(name))) {
    return false;
  }
  return (
    named.modifiers?.some(
      (m) => m.kind === ast.SyntaxKind.PrivateKeyword || m.kind === ast.SyntaxKind.ProtectedKeyword,
    ) !== true
  );
}

/** The type names one declaration mentions, in source order, deduplicated. */
function referencedTypes(node: ast.Node): string[] {
  const found = new Set<string>();
  const walk = (child: ast.Node): void => {
    if (!isVisibleMember(child)) return;
    // A statement body is not a signature. A `const flushed = new Map<Worksheet, FlushedSheet>()`
    // inside a method says nothing about what a consumer can reach: the annotations, casts and type
    // arguments in there are the implementation talking to the compiler. Walking them reported the
    // streaming writer's own plumbing as an unnameable public type the moment that plumbing stopped
    // declining publication, which is the opposite of what had just been fixed. A type a body
    // genuinely leaks does so through the declaration's *inferred* return type, and that is a
    // reference node this walk never sees anyway.
    if (ast.isBlock(child)) return;
    if (ast.isTypeReferenceNode(child)) {
      const name = child.typeName;
      // A qualified `A.B` names a namespace member, which this package has none of; the head is
      // still the name a consumer would have to reach, so it is what gets recorded.
      found.add(ast.isIdentifier(name) ? name.text : name.getText());
    } else if (ast.isExpressionWithTypeArguments(child) && ast.isIdentifier(child.expression)) {
      // A heritage clause: `interface CellModel extends CellContent`.
      found.add(child.expression.text);
    }
    child.forEachChild(walk);
  };
  walk(node);
  return [...found];
}

function main(project: Project): void {
  const index = declarationIndex(project);
  const published = publishedNames(project);

  const problems: string[] = [];
  const seen = new Set<string>();
  // Breadth-first from the published surface, so a type is reported at the shallowest place it is
  // reachable from and each is reported once however many signatures name it.
  const queue = [...published].filter((name) => index.has(name));
  const reachedFrom = new Map<string, string>();
  for (let next = 0; next < queue.length; next++) {
    const name = queue[next];
    if (name === undefined || seen.has(name)) continue;
    seen.add(name);
    const declared = index.get(name);
    if (declared === undefined || declared.declined) continue;
    for (const referenced of referencedTypes(declared.node)) {
      if (seen.has(referenced) || !index.has(referenced)) continue;
      if (!reachedFrom.has(referenced)) reachedFrom.set(referenced, name);
      queue.push(referenced);
    }
  }

  let declined = 0;
  for (const name of seen) {
    if (published.has(name)) continue;
    const declared = index.get(name);
    if (declared?.declined === true) {
      declined++;
      continue;
    }
    const via = reachedFrom.get(name) ?? '(an entry)';
    problems.push(
      `  ${name} (${declared?.file ?? 'unknown'})\n` +
        `    is named by the published ${via}, and no entry barrel exports it:\n` +
        `    a consumer can hold the value and cannot write its type`,
    );
  }

  verdict({
    gate: 'public types',
    problems,
    ok:
      `${seen.size} type(s) reachable from the published surface, all nameable` +
      (declined > 0 ? `; ${declined} decline publication with @unpublished` : ''),
    failure: 'type(s) reachable from a published signature that no entry exports',
  });
}

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
