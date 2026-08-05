// Generates the public API reference under `docs/api/` straight from the types.
//
// The public barrel (`src/index.ts`) is the single source of truth: this walks the
// symbols it re-exports via the TypeScript compiler API, renders each one's JSDoc
// summary + tags + a body-stripped TypeScript signature, and writes one Markdown
// page per originating module plus an index. No new dependency — `typescript` is
// already the toolchain — so the docs cannot describe a shape the compiler wouldn't
// accept. Run `pnpm run docs`; `pnpm run docs:check` fails if the committed pages have
// drifted from a fresh generation (the docs are gated like any other artifact).
//
// See docs/decisions/0006-docs-from-types.md.

import {mkdirSync, rmSync, writeFileSync} from 'node:fs';
import {dirname, join, resolve} from 'node:path';
import {fileURLToPath} from 'node:url';
import * as ast from 'typescript/unstable/ast';
import {
  API,
  type Checker,
  type Project,
  SymbolFlags,
  type Symbol as TypeSymbol,
} from 'typescript/unstable/sync';

const ROOT = resolve(dirname(fileURLToPath(import.meta.url)), '..');
const ENTRY = join(ROOT, 'src/index.ts');
const CONFIG = join(ROOT, 'tsconfig.json');
const OUT_DIR = join(ROOT, 'docs/api');

/**
 * `TypeFormatFlags.NoTruncation`. TypeScript 7 does not export that enum, but the checker
 * still reads its bits — and without this a wide union renders as `… 18 more …`.
 */
const NO_TRUNCATION = 1;

/** The modifiers that keep a class member out of the reference. */
const HIDDEN = ast.ModifierFlags.Private | ast.ModifierFlags.Protected;

// Human-facing page titles + ordering, keyed by the source module basename an export
// resolves to. A module not listed here still renders (alphabetically, titled from its
// filename) so a new public module can never silently vanish from the docs.
const GROUPS: ReadonlyArray<readonly [string, string]> = [
  ['address', 'Addresses & ranges'],
  ['value', 'Cell values'],
  ['cell', 'Cell'],
  ['workbook', 'Workbook'],
  ['worksheet', 'Worksheet'],
  ['style', 'Styles'],
  ['table', 'Tables'],
  ['pivot-table', 'Pivot tables'],
  ['image', 'Images'],
  ['protection', 'Protection'],
  ['workbook-protection', 'Protection'],
  ['xlsx/read', 'Reading .xlsx'],
  ['xlsx/read-rows', 'Streaming reads'],
  ['xlsx/write', 'Writing .xlsx'],
  ['xlsx/write-stream', 'Streaming writes'],
  ['csv/read', 'CSV'],
  ['csv/write', 'CSV'],
];

/** The body-bearing declarations — the ones whose signature ends where their body begins. */
function bodyOf(node: ast.Node): ast.Node | undefined {
  if (
    ast.isFunctionDeclaration(node) ||
    ast.isMethodDeclaration(node) ||
    ast.isGetAccessorDeclaration(node) ||
    ast.isSetAccessorDeclaration(node)
  ) {
    return node.body;
  }
  return undefined;
}

/**
 * Render a declaration as a signature by slicing its own source text, cut at the body when it
 * has one.
 *
 * Source text rather than a compiler re-print: the reference then shows the shape the author
 * wrote — parameter line breaks and all — instead of the printer's normalization, and the
 * generator needs no emit machinery. Starting at `getStart` drops leading trivia, so the JSDoc
 * above a declaration stays out of the code block that renders it.
 */
function printSignature(node: ast.Node, sourceFile: ast.SourceFile): string {
  const body = bodyOf(node);
  const text = sourceFile.text
    .slice(node.getStart(sourceFile), body ? body.getStart(sourceFile) : node.getEnd())
    .trimEnd()
    .replace(/^export (?:default )?/, '');
  // Cutting the body leaves the signature unterminated; `;` restores the `.d.ts` shape a reader
  // expects. Declarations that keep their text (interfaces, type aliases) already carry their own.
  return body ? `${text};` : text;
}

/** Where a `{@link}` target lives in the reference, or `undefined` if the reference does not document it. */
type LinkResolver = (target: string) => string | undefined;

/** The resolver for prose rendered before the link index exists — the index's own entries. */
const NO_LINKS: LinkResolver = () => undefined;

/**
 * A link target's URL.
 *
 * `{@link addRow}` written inside `Worksheet`'s own prose means `Worksheet.addRow` — the author had
 * the class in scope and wrote what they would say aloud — so a scoped lookup runs first. A
 * qualified target with no block of its own falls back to its container: `DefinedName.scope` is an
 * interface property, which renders inside the interface's signature rather than under a heading,
 * and the interface is where a reader finds it.
 */
function linkResolver(targets: ReadonlyMap<string, string>, scope?: string): LinkResolver {
  return (target) => {
    const scoped = scope === undefined ? undefined : targets.get(`${scope}.${target}`);
    if (scoped !== undefined) return scoped;
    const direct = targets.get(target);
    if (direct !== undefined) return direct;
    const dot = target.indexOf('.');
    return dot > 0 ? targets.get(target.slice(0, dot)) : undefined;
  };
}

/**
 * Turn TSDoc `{@link Target}` / `{@link Target | label}` into a code span, linked when the reference
 * documents the target and left bare when it does not.
 *
 * Bare rather than dropped for an unresolved target: plenty of prose links to something deliberately
 * internal, and a reader is better served by the name than by a link to nothing.
 *
 * Only the pipe form carries a label. TSDoc also permits `{@link Target label}` (space, no pipe), which
 * falls through here and renders as the whole string — so write the pipe, or write a bare target.
 */
function resolveLinks(text: string, href: LinkResolver = NO_LINKS): string {
  return text.replace(
    /\{@link(?:code|plain)?\s+([^}|]+?)(?:\s*\|\s*([^}]+))?\}/g,
    (_m: string, target: string, label: string | undefined) => {
      const span = `\`${(label ?? target).trim()}\``;
      const url = href(target.trim());
      return url === undefined ? span : `[${span}](${url})`;
    },
  );
}

/**
 * The JSDoc block a declaration is documented by.
 *
 * A `const` carries its block on the enclosing statement rather than on the declarator, so that
 * one walks out through the declaration list; every other kind we render is annotated directly.
 */
function jsDocOf(node: ast.Node): ast.JSDoc | undefined {
  const documented = ast.isVariableDeclaration(node) ? node.parent.parent : node;
  return documented.jsDoc?.at(-1) as ast.JSDoc | undefined;
}

/**
 * A declaration's summary prose, read from its JSDoc block.
 *
 * The block rather than `Symbol.getDocumentationComment`, because the checker hands back a string
 * in which `{@link Target}` has already been flattened to a bare `Target` — losing the one piece of
 * markup {@link resolveLinks} exists to render. The AST still carries the tag intact.
 *
 * A bare `@word` mid-sentence silently truncates the summary: TypeScript's JSDoc parser reads it as
 * the start of an unknown tag and drops everything after it. Prose about, say, an `@mention` has to
 * wrap it in backticks — the parse happens upstream of here, so this cannot be fixed downstream.
 */
function docText(node: ast.Node, href: LinkResolver = NO_LINKS): string {
  return resolveLinks((ast.getTextOfJSDocComment(jsDocOf(node)?.comment) ?? '').trim(), href);
}

/** A bare type name, possibly qualified — what the brace slot of a `@throws` is allowed to hold. */
const ERROR_TYPE = /^[A-Za-z_$][\w$.]*$/;

/**
 * The error type a `@throws` names, and the prose after it.
 *
 * `@throws {ErrorType}` is the only spelling that survives the parse. TypeScript reads the braces
 * after `@throws` as a *type expression*, and `{@link Target}` is not one: the parse runs past the
 * close brace and consumes the rest of the comment, so the tag arrives here as the bare string `{`
 * with its prose already gone — from editor hovers too, not just from this generator. Nothing
 * downstream can recover it, so a slot that is not a type name fails the run rather than
 * publishing the truncation as if it were the whole sentence.
 */
function splitThrows(raw: string, symbolName: string): {errorType: string; prose: string} {
  const close = raw.startsWith('{') ? raw.indexOf('}') : -1;
  const errorType = close === -1 ? '' : raw.slice(1, close).trim();
  if (!ERROR_TYPE.test(errorType)) {
    throw new Error(
      `${symbolName}: write \`@throws {ErrorType} prose\`, not \`@throws ${raw}\` — ` +
        "a `{@link …}` in the brace slot is swallowed by TypeScript's JSDoc type-expression " +
        'parse, and takes the description with it. Links mid-sentence are fine.',
    );
  }
  return {errorType, prose: raw.slice(close + 1).trim()};
}

/** Render `@throws`, `@param`, `@returns`, `@example` tags into Markdown, in a stable order. */
function docTags(symbol: TypeSymbol, checker: Checker, href: LinkResolver = NO_LINKS): string[] {
  const order: Record<string, number> = {param: 0, returns: 1, throws: 2, example: 3};
  const tags = symbol
    .getJsDocTags(checker)
    .filter((t) => t.name in order)
    .sort((a, b) => (order[a.name] ?? 0) - (order[b.name] ?? 0));
  const lines: string[] = [];
  for (const tag of tags) {
    const raw = (tag.text ?? '').trim();
    const text = resolveLinks(raw, href);
    if (tag.name === 'example') {
      lines.push('', '```ts', text, '```');
    } else if (tag.name === 'param') {
      const [name, ...rest] = text.split(/\s+/);
      lines.push(`- \`${name}\` — ${rest.join(' ')}`);
    } else if (tag.name === 'returns') {
      lines.push(`**Returns** — ${text}`);
    } else if (tag.name === 'throws') {
      // The type slot is the tag's subject — which error — so it leads the line. Dropping it, as
      // this once did, left the reader told that a throw happens but never told what is thrown.
      const {errorType, prose} = splitThrows(raw, symbol.name);
      const named = href(errorType);
      const subject = named === undefined ? `\`${errorType}\`` : `[\`${errorType}\`](${named})`;
      lines.push(`**Throws** — ${subject}${prose ? ` ${resolveLinks(prose, href)}` : ''}`);
    }
  }
  return lines;
}

function kindLabel(node: ast.Node): string {
  if (ast.isInterfaceDeclaration(node)) return 'interface';
  if (ast.isTypeAliasDeclaration(node)) return 'type';
  if (ast.isClassDeclaration(node)) return 'class';
  if (ast.isFunctionDeclaration(node)) return 'function';
  if (ast.isEnumDeclaration(node)) return 'enum';
  if (ast.isVariableDeclaration(node)) return 'const';
  return 'value';
}

/**
 * The class members the reference can render: the ones carrying a name to key them by and
 * modifiers to judge their visibility.
 *
 * `ClassElement` itself declares neither — TypeScript 7 models it as a bare brand — so the four
 * kinds that do are narrowed to explicitly. The kinds left out have no name to render under
 * anyway: a constructor, an index signature, a static block, a stray semicolon.
 */
type NamedMember = ast.MethodDeclaration | ast.PropertyDeclaration | ast.AccessorDeclaration;

function asNamedMember(member: ast.ClassElement): NamedMember | undefined {
  return ast.isMethodDeclaration(member) ||
    ast.isPropertyDeclaration(member) ||
    ast.isGetAccessorDeclaration(member) ||
    ast.isSetAccessorDeclaration(member)
    ? member
    : undefined;
}

function isPublicMember(member: NamedMember): boolean {
  if (ast.isPrivateIdentifier(member.name)) return false;
  // A computed name is the codec's back channel (`src/core/internal.ts`), keyed by a symbol that
  // never leaves the package. Reachable only by a holder of that symbol, so it is no more public
  // than a `#private` field — and rendering it would dump the whole channel, initializer included,
  // into the reference a caller reads.
  if (ast.isComputedPropertyName(member.name)) return false;
  return (member.modifierFlags & HIDDEN) === 0;
}

/**
 * The declarations that share a member name: an overload set, a get/set pair, or a lone member.
 *
 * Grouped because the reference documents a *name*, not a declaration — the two halves of an
 * accessor and every overload of a method are one thing to a caller, and the JSDoc sits on
 * whichever declaration the author chose.
 */
function groupMembersByName(
  node: ast.ClassDeclaration,
  sourceFile: ast.SourceFile,
): {sigs: string[]; groups: Map<string, NamedMember[]>} {
  const sigs: string[] = [];
  const groups = new Map<string, NamedMember[]>();
  for (const element of node.members) {
    const member = asNamedMember(element);
    if (!member || !isPublicMember(member)) continue;
    sigs.push(`  ${printSignature(member, sourceFile).trim()}`);
    const name = member.name.getText(sourceFile);
    let group = groups.get(name);
    if (!group) {
      group = [];
      groups.set(name, group);
    }
    group.push(member);
  }
  return {sigs, groups};
}

/**
 * The signatures worth showing for one member name.
 *
 * A method declared more than once is an overload set whose last declaration is the
 * implementation — an artefact of how TypeScript spells overloading, and not a signature any
 * caller may pass. Accessors are excluded from that collapse: a get/set pair is also two
 * body-bearing declarations of one name, but *both* are the caller's surface.
 */
function shownSignatures(group: readonly NamedMember[]): readonly NamedMember[] {
  if (group.length < 2 || !group.every(ast.isMethodDeclaration)) return group;
  return group.filter((member) => !bodyOf(member));
}

/** One member name that earns a block: which declarations to show, and which one is documented. */
type DocumentedMember = {
  name: string;
  signatures: readonly NamedMember[];
  symbol: TypeSymbol | undefined;
  documenting: NamedMember;
};

/**
 * The members that get a block of their own: those with prose, tags, or both. A member carrying
 * neither is left to the signature overview, which already lists it — a heading over a bare
 * signature says nothing the overview did not.
 *
 * The link index and the renderer both go through here rather than each deciding for itself. When
 * they decided separately the index advertised anchors for members the renderer had skipped, and
 * the reference shipped links to headings that were never written.
 */
function documentedMembers(
  node: ast.ClassDeclaration,
  sourceFile: ast.SourceFile,
  checker: Checker,
): {sigs: string[]; members: DocumentedMember[]} {
  const {sigs, groups} = groupMembersByName(node, sourceFile);
  const members: DocumentedMember[] = [];
  for (const [name, group] of groups) {
    const documenting = group.find((member) => docText(member) !== '') ?? group[0];
    if (!documenting) continue;
    const symbol = checker.getSymbolAtLocation(documenting.name);
    const documented =
      docText(documenting) !== '' || (symbol && docTags(symbol, checker).length > 0);
    if (!documented) continue;
    members.push({name, signatures: shownSignatures(group), symbol, documenting});
  }
  return {sigs, members};
}

/**
 * A class renders as a signature overview (every public member) followed by one block per
 * documented member — signature, summary, and the same `@throws`/`@param`/`@returns` rendering a
 * top-level symbol gets. Members went un-tagged for a long time, which left ~40 documented throws
 * reaching the reference never; routing them through {@link docTags} is what fixed that, and is
 * why a member is a block rather than the one-line bullet it used to be — a tag list does not fit
 * on a bullet.
 *
 * Returns the overview lines and the member blocks separately so the caller can put the compact
 * signature block above the prose.
 */
function renderClassMembers(
  node: ast.ClassDeclaration,
  sourceFile: ast.SourceFile,
  className: string,
  checker: Checker,
  href: LinkResolver,
): {sigs: string[]; docs: string[]} {
  const {sigs, members} = documentedMembers(node, sourceFile, checker);
  const docs: string[] = [];
  for (const {name, signatures, symbol, documenting} of members) {
    const summary = docText(documenting, href);
    const tagLines = symbol ? docTags(symbol, checker, href) : [];
    const printed = signatures.map((member) => printSignature(member, sourceFile).trim());
    docs.push(`#### \`${className}.${name}\``, '', '```ts', ...printed, '```', '');
    if (summary) docs.push(summary, '');
    if (tagLines.length > 0) docs.push(...tagLines, '');
  }
  while (docs.at(-1) === '') docs.pop();
  return {sigs, docs};
}

/** Every heading a page emits, as the anchor a reader's browser will resolve. */
function anchorsIn(body: string): Set<string> {
  const found = new Set<string>();
  let fenced = false;
  for (const line of body.split('\n')) {
    if (line.startsWith('```')) fenced = !fenced;
    // A fence holds source text, and a signature's own JSDoc rides along in it — `#` there starts
    // no heading, and `{@link}` there is never rewritten into a link.
    if (fenced) continue;
    const heading = /^#{1,6}\s+(.*)$/.exec(line);
    if (heading?.[1] !== undefined) found.add(anchor(heading[1]));
  }
  return found;
}

/**
 * Refuse to write a reference that links to a heading it does not contain.
 *
 * The link index and the renderer derive their member lists from one function so they cannot
 * disagree — but they did disagree once, when the index advertised every member and the renderer
 * gave a block only to the documented ones, and the result was four links to headings that were
 * never written. Nothing about the output looked wrong; the anchors simply went nowhere. This is
 * the check that makes that failure loud, and it costs one pass over text already in memory.
 */
function checkLinks(bodies: ReadonlyMap<string, string>): void {
  const anchors = new Map([...bodies].map(([file, body]) => [file, anchorsIn(body)]));
  const dangling: string[] = [];
  for (const [file, body] of bodies) {
    let fenced = false;
    for (const [i, line] of body.split('\n').entries()) {
      if (line.startsWith('```')) fenced = !fenced;
      if (fenced) continue;
      for (const [, target, frag] of line.matchAll(/\]\(\.\/([a-z0-9-]+\.md)#([a-z0-9]+)\)/g)) {
        if (target !== undefined && frag !== undefined && !anchors.get(target)?.has(frag)) {
          dangling.push(`  ${file}:${i + 1} → ${target}#${frag}`);
        }
      }
    }
  }
  if (dangling.length > 0) {
    throw new Error(
      `the reference links to ${dangling.length} missing heading(s):\n${dangling.join('\n')}`,
    );
  }
}

function main(project: Project) {
  const {program, checker} = project;
  const entrySf = program.getSourceFile(ENTRY);
  if (!entrySf) throw new Error(`cannot load entry ${ENTRY}`);
  const moduleSymbol = checker.getSymbolAtLocation(entrySf);
  if (!moduleSymbol) throw new Error('entry has no module symbol — is src/index.ts a module?');

  const groupTitle = new Map(GROUPS);
  const groupOrder = new Map<string, number>(GROUPS.map(([key], i) => [key, i]));
  type Entry = {name: string; block: string};
  type Page = {title: string; key: string; entries: Entry[]};
  const pages = new Map<string, Page>();

  type Exported = {
    symbol: TypeSymbol;
    decl: ast.Node;
    sourceFile: ast.SourceFile;
    name: string;
    page: Page;
  };
  const walked: Exported[] = [];

  for (const exported of checker.getExportsOfModule(moduleSymbol)) {
    const symbol =
      exported.flags & SymbolFlags.Alias ? checker.getAliasedSymbol(exported) : exported;
    // A declaration crosses the API boundary as a handle, not a node — resolving it is a
    // round-trip to the compiler server that holds the tree.
    const decl = symbol.declarations[0]?.resolve(project);
    if (!decl) continue;
    const sourceFile = decl.getSourceFile();
    const rel = sourceFile.fileName
      .replace(/^.*\/src\//, '')
      .replace(/\.ts$/, '')
      .replace(/\/index$/, '');
    const groupKey = rel.replace(/^core\//, '').replace(/^io\//, '');
    const title = groupTitle.get(groupKey) ?? titleize(groupKey);
    const pageId = title;

    let page = pages.get(pageId);
    if (!page) {
      page = {title, key: groupKey, entries: []};
      pages.set(pageId, page);
    }
    walked.push({symbol, decl, sourceFile, name: exported.name, page});
  }

  // Every target has to be known before the first block renders: prose links forward as freely as
  // it links back, and a symbol's page is only settled once the walk that assigns pages has ended.
  // Hence the split — walk, index, then render.
  const targets = new Map<string, string>();
  for (const {decl, sourceFile, name, page} of walked) {
    const slug = slugify(page.title);
    targets.set(name, `./${slug}.md#${anchor(name)}`);
    if (!ast.isClassDeclaration(decl)) continue;
    for (const member of documentedMembers(decl, sourceFile, checker).members) {
      const qualified = `${name}.${member.name}`;
      targets.set(qualified, `./${slug}.md#${anchor(qualified)}`);
    }
  }

  for (const {symbol, decl, sourceFile, name, page} of walked) {
    const kind = kindLabel(decl);
    // A class scopes its own prose: inside `Worksheet`, a bare `{@link addRow}` is `Worksheet.addRow`.
    const href = linkResolver(targets, ast.isClassDeclaration(decl) ? name : undefined);
    const summary = docText(decl, href);
    const tagLines = docTags(symbol, checker, href);

    const block = [`### \`${name}\``, '', `<sub>${kind}</sub>`, ''];
    if (summary) block.push(summary, '');

    let signature: string;
    let memberDocs: string[] = [];
    if (ast.isClassDeclaration(decl)) {
      const {sigs, docs} = renderClassMembers(decl, sourceFile, name, checker, href);
      const heritage = decl.heritageClauses?.map((h) => h.getText(sourceFile)).join(' ');
      signature = `class ${name}${heritage ? ` ${heritage}` : ''} {\n${sigs.join('\n')}\n}`;
      memberDocs = docs;
    } else if (ast.isVariableDeclaration(decl)) {
      const type = checker.typeToString(
        checker.getTypeOfSymbolAtLocation(symbol, decl),
        decl,
        NO_TRUNCATION,
      );
      signature = `const ${name}: ${type}`;
    } else if (ast.isFunctionDeclaration(decl)) {
      // Print every overload declaration (those without a body); skip the impl signature.
      const declarations = symbol.declarations
        .map((handle) => handle.resolve(project))
        .filter((d) => d !== undefined);
      const overloads = declarations.filter(
        (d) => ast.isFunctionDeclaration(d) && (!d.body || declarations.length === 1),
      );
      signature = overloads.map((d) => printSignature(d, d.getSourceFile())).join('\n');
    } else {
      signature = printSignature(decl, sourceFile);
    }
    block.push('```ts', signature, '```');
    if (tagLines.length > 0) block.push('', ...tagLines);
    if (memberDocs.length > 0) block.push('', '**Members**', '', ...memberDocs);

    page.entries.push({name, block: block.join('\n')});
  }

  // Deterministic output regardless of `getExportsOfModule` iteration order — the CI
  // drift check compares committed pages against a fresh generation byte for byte.
  for (const page of pages.values()) {
    page.entries.sort((a: Entry, b: Entry) => a.name.localeCompare(b.name));
  }

  const ordered = [...pages.values()].sort((a, b) => {
    const oa = groupOrder.get(a.key) ?? Number.MAX_SAFE_INTEGER;
    const ob = groupOrder.get(b.key) ?? Number.MAX_SAFE_INTEGER;
    return oa - ob || a.title.localeCompare(b.title);
  });

  const bodies = new Map<string, string>();
  for (const page of ordered) {
    const body = [
      `# ${page.title}`,
      '',
      '<!-- Generated from the public types by `pnpm run docs`. Do not edit by hand. -->',
      '',
      page.entries.map((e) => e.block).join('\n\n---\n\n'),
      '',
    ].join('\n');
    bodies.set(`${slugify(page.title)}.md`, body);
  }
  checkLinks(bodies);

  rmSync(OUT_DIR, {recursive: true, force: true});
  mkdirSync(OUT_DIR, {recursive: true});
  for (const [file, body] of bodies) writeFileSync(join(OUT_DIR, file), body);

  const index = [
    '# API reference',
    '',
    '<!-- Generated from the public types by `pnpm run docs`. Do not edit by hand. -->',
    '',
    'Every symbol below is re-exported from the package root and generated straight from',
    'its TypeScript declaration — the types are the contract.',
    '',
    ...ordered.map((page) => {
      const slug = slugify(page.title);
      const symbols = page.entries
        .map((e) => `[\`${e.name}\`](./${slug}.md#${anchor(e.name)})`)
        .join(', ');
      return `- **[${page.title}](./${slug}.md)** — ${symbols}`;
    }),
    '',
  ].join('\n');
  writeFileSync(join(OUT_DIR, 'README.md'), index);

  const count = ordered.reduce((n, p) => n + p.entries.length, 0);
  process.stdout.write(`docs: ${count} symbols across ${ordered.length} pages → docs/api/\n`);
}

function titleize(key: string): string {
  return key.replace(/[-/]/g, ' ').replace(/\b\w/g, (c) => c.toUpperCase());
}
function slugify(title: string): string {
  return title
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, '-')
    .replace(/^-|-$/g, '');
}
function anchor(name: string): string {
  return name.toLowerCase().replace(/[^a-z0-9]+/g, '');
}

// The compiler runs out-of-process, so the snapshot and the server it lives in are both
// resources: without the `finally` a throw mid-generation would leave a tsgo process behind
// and this one hanging on its open pipe.
//
// The project comes from `tsconfig.json` rather than a hand-written option set. There is no
// inline-options door in this API — and shutting it removed a real hazard, since the options
// gen-docs used to pass were its own and had already drifted from the gate's.
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
