/**
 * Reads `docs/`, the tracked documentation tree, as the site's input.
 *
 * `docs/` is written for someone standing in a checkout: it carries no frontmatter, it
 * links to its neighbours by relative path, and it links out to files the site does not
 * publish. This module translates that into what VitePress needs, and **fails rather than
 * guesses** at every step. A page the manifest does not reach, a manifest entry with no
 * file, a link that resolves to nothing: each one throws. The value of a generated mirror
 * is precisely that `docs/` cannot rot behind the site, and it can only deliver that if
 * drift is an error instead of a warning.
 */

import {readdirSync, readFileSync, statSync} from 'node:fs';
import {posix, resolve} from 'node:path';

import {blobUrl, repoRoot, treeUrl} from './repo.ts';

const DOCS_DIR = resolve(repoRoot, 'docs');
const MANIFEST = 'docs/docs.json';

/** Where the mirror is mounted on the site. Every route below is built from it. */
const MOUNT = 'docs';

export interface DocPage {
  /** Path under `docs/`, extension included: `api/workbook.md`. The file a reader edits. */
  readonly sourcePath: string;
  /** How the manifest names this page: `api/workbook`, or `api/index` for a directory index. */
  readonly slug: string;
  /** The site route: `/docs/api/workbook`, or `/docs/api/` for an index. */
  readonly route: string;
  readonly title: string;
  /** The lead paragraph, if the page opens with prose. Absent is allowed; invented is not. */
  readonly description: string | undefined;
  /** The page as the mirror should write it: source text with every link rewritten. */
  readonly body: string;
}

export interface DocsGroup {
  readonly text: string;
  readonly collapsed: boolean;
  readonly pages: readonly DocPage[];
}

export interface DocsSource {
  readonly groups: readonly DocsGroup[];
  readonly pages: readonly DocPage[];
  /**
   * Mirror path to route path, for VitePress's `rewrites`. It exists so a directory's
   * `README.md` can be served at the directory root while the mirror keeps the source
   * filename: with the mirror path and the repository path identical, `editLink`'s `:path`,
   * which is the file's own path, points at the tracked file with no mapping in the theme.
   */
  readonly rewrites: Readonly<Record<string, string>>;
}

interface ManifestGroup {
  readonly group: string;
  /** Pages a reader should meet in this order. */
  readonly pages?: readonly string[];
  /** A directory whose remaining pages follow, sorted. What `pages` names is not repeated. */
  readonly tree?: string;
  readonly collapsed?: boolean;
}

interface Manifest {
  readonly navigation: readonly ManifestGroup[];
}

export function readDocsSource(): DocsSource {
  const manifest = JSON.parse(readFileSync(resolve(DOCS_DIR, 'docs.json'), 'utf8')) as Manifest;
  const files = markdownFiles();
  const bySlug = new Map(files.map((sourcePath) => [slugFor(sourcePath), sourcePath]));

  const claimed = new Map<string, string>();
  const claim = (slug: string, by: string): void => {
    if (!bySlug.has(slug)) {
      throw new Error(`${MANIFEST}: ${by} names "${slug}", and no file under docs/ has that slug.`);
    }
    const first = claimed.get(slug);
    if (first !== undefined) {
      throw new Error(`${MANIFEST}: "${slug}" is claimed by both ${first} and ${by}.`);
    }
    claimed.set(slug, by);
  };

  // Two passes, because a `tree` is defined as the rest of its directory: every explicit
  // page has to be known before any tree can say what it does not already cover.
  for (const group of manifest.navigation) {
    for (const slug of group.pages ?? []) claim(slug, `group "${group.group}"`);
  }

  const ordered = manifest.navigation.map((group) => {
    const rest = group.tree === undefined ? [] : treeSlugs(group.tree, bySlug, claimed);
    for (const slug of rest) claim(slug, `tree "${group.tree ?? ''}"`);
    return {
      text: group.group,
      collapsed: group.collapsed ?? false,
      slugs: [...(group.pages ?? []), ...rest],
    };
  });

  const unreached = [...bySlug.keys()].filter((slug) => !claimed.has(slug)).sort();
  if (unreached.length > 0) {
    throw new Error(
      `docs/ holds ${unreached.length} page(s) no group in ${MANIFEST} reaches: ` +
        `${unreached.slice(0, 10).join(', ')}${unreached.length > 10 ? ', ...' : ''}. ` +
        'List them, or give their directory a tree; an unreachable page is one nobody edits.',
    );
  }

  const pages = new Map(
    [...bySlug].map(([slug, sourcePath]) => [slug, readPage(slug, sourcePath, bySlug)] as const),
  );

  return {
    groups: ordered.map((group) => ({
      text: group.text,
      collapsed: group.collapsed,
      pages: group.slugs.map((slug) => found(pages, slug)),
    })),
    pages: [...pages.values()],
    rewrites: Object.fromEntries(
      [...pages.values()]
        .filter((page) => page.slug.endsWith('/index'))
        .map((page) => [`${MOUNT}/${page.sourcePath}`, `${MOUNT}/${page.slug}.md`]),
    ),
  };
}

function found(pages: ReadonlyMap<string, DocPage>, slug: string): DocPage {
  const page = pages.get(slug);
  if (page === undefined) throw new Error(`docs-source.ts validated "${slug}" and then lost it.`);
  return page;
}

/** Every `.md` under a directory that no explicit page already claims, sorted. */
function treeSlugs(
  dir: string,
  bySlug: ReadonlyMap<string, string>,
  claimed: ReadonlyMap<string, string>,
): string[] {
  const under = [...bySlug.keys()].filter((slug) => slug.startsWith(`${dir}/`));
  if (under.length === 0)
    throw new Error(`${MANIFEST}: tree "${dir}" matches no page under docs/.`);
  return under.filter((slug) => !claimed.has(slug)).sort();
}

/** Paths under `docs/`, posix, `.md` only. */
function markdownFiles(): string[] {
  const found: string[] = [];
  const walk = (relative: string): void => {
    for (const entry of readdirSync(resolve(DOCS_DIR, relative)).sort()) {
      const child = relative === '' ? entry : posix.join(relative, entry);
      if (statSync(resolve(DOCS_DIR, child)).isDirectory()) walk(child);
      else if (entry.endsWith('.md')) found.push(child);
    }
  };
  walk('');
  return found;
}

/** `api/README.md` becomes `api/index`; `architecture.md` becomes `architecture`. */
function slugFor(sourcePath: string): string {
  const stem = sourcePath.slice(0, -'.md'.length);
  return stem.endsWith('/README') ? `${stem.slice(0, -'README'.length)}index` : stem;
}

/** A directory index is served at the directory root, so its route keeps the slash. */
function routeFor(slug: string): string {
  return slug.endsWith('/index')
    ? `/${MOUNT}/${slug.slice(0, -'index'.length)}`
    : `/${MOUNT}/${slug}`;
}

function readPage(slug: string, sourcePath: string, bySlug: ReadonlyMap<string, string>): DocPage {
  const source = readFileSync(resolve(DOCS_DIR, sourcePath), 'utf8');
  const blocks = splitBlocks(source);
  assertRenderable(source, `docs/${sourcePath}`);
  return {
    sourcePath,
    slug,
    route: routeFor(slug),
    title: headingOf(blocks, `docs/${sourcePath}`),
    description: descriptionOf(blocks),
    body: rewriteLinks(source, sourcePath, bySlug),
  };
}

// ---------------------------------------------------------------------------
// Titles and descriptions
//
// Derived from the prose, never from frontmatter. None of the pages under docs/ carries
// any, and adding a schema to all of them to serve a `<title>` tag would be a migration
// paid for by every page written afterwards. The derivation is checkable, which is what
// the schema was going to buy.
// ---------------------------------------------------------------------------

const FENCE = /^\s*(?:```|~~~)/;
const HEADING = /^(#{1,6})\s+(.+?)\s*#*\s*$/;

interface Block {
  readonly lines: readonly string[];
}

/** Paragraph-sized blocks, with fenced code kept whole so a `#` inside one is not a heading. */
function splitBlocks(source: string): Block[] {
  const blocks: Block[] = [];
  let current: string[] = [];
  let fenced = false;
  const flush = (): void => {
    if (current.length > 0) blocks.push({lines: current});
    current = [];
  };
  for (const line of source.split('\n')) {
    if (FENCE.test(line)) {
      fenced = !fenced;
      current.push(line);
      continue;
    }
    if (!fenced && line.trim() === '') flush();
    else current.push(line);
  }
  flush();
  return blocks;
}

/**
 * Markdown the site can actually compile.
 *
 * VitePress renders every page as a Vue template, so an angle bracket in prose is an
 * element and an element that never closes fails the build. It fails it badly: the reported
 * position is wherever the parser finally gave up, which was 50 lines below the cause in
 * both of the cases here. This check names the line instead, and runs in a gate that takes
 * a second rather than in a build that takes ten.
 *
 * Two shapes are refused. A tag whose name is not an HTML element is prose someone forgot
 * to wrap in backticks, and it is invisible on GitHub today for the same reason. A
 * paragraph line starting with `<` is read as an element even when markdown would have kept
 * it inside an inline code span opened on the line above, so a span wrapped across a
 * newline breaks the page; putting the span on one line fixes it. A tag alone on its own
 * line, like the `<sub>` markers on the generated API pages, opens its own block and is fine.
 */
function assertRenderable(source: string, where: string): void {
  // Two views of the page, because the two failures live on opposite sides of a code span.
  // A tag is only a problem where markdown would have rendered it as markup; a continuation
  // line is a problem even where markdown would have kept it as code, which is the whole
  // reason the second shape is surprising.
  const prose = withoutFences(source).split('\n');
  const markup = withoutCode(prose.join('\n')).split('\n');
  for (const [index, line] of prose.entries()) {
    const at = `${where}:${index + 1}`;
    const tag = NON_HTML_TAG.exec(markup[index] ?? '');
    if (tag !== null) {
      throw new Error(
        `${at}: "${tag[0]}>" is not an HTML element, so the site reads it as a component that ` +
          'never closes and the page fails to compile. Wrap it in backticks; it is already ' +
          'invisible when GitHub renders this page.',
      );
    }
    if (line.startsWith('<') && (prose[index - 1] ?? '').trim() !== '') {
      throw new Error(
        `${at}: a paragraph line starts with "${line.slice(0, 24)}". The site reads that as an ` +
          'element even inside an inline code span opened on the line above, and the page ' +
          'fails to compile. Put the span on one line.',
      );
    }
  }
}

/** The elements docs/ may use. Anything else in angle brackets is prose that needs backticks. */
const HTML_ELEMENTS = new Set([
  'a',
  'b',
  'br',
  'code',
  'details',
  'em',
  'i',
  'img',
  'kbd',
  'p',
  'span',
  'strong',
  'sub',
  'summary',
  'sup',
]);
const TAG = /<\/?([A-Za-z][A-Za-z0-9-]*)/g;
const NON_HTML_TAG = new RegExp(TAG.source);
const CODE_SPAN = /(`+)[\s\S]*?\1/g;
const COMMENT = /<!--[\s\S]*?-->/g;
// `<https://example.com>` and `<name@example.com>`: markdown autolinks, which render as links
// and never reach the Vue compiler as markup. They look exactly like an unclosed tag to the
// scan below, and one in `docs/architecture.md` was the first thing this check reported.
const AUTOLINK = /<[A-Za-z][A-Za-z0-9+.-]*:[^>\s]*>|<[^>\s@]+@[^>\s]+>/g;

const blank = (text: string): string => text.replace(/[^\n]/g, ' ');

/** Fenced code emptied, line count preserved so a finding can still name its line. */
function withoutFences(source: string): string {
  let fenced = false;
  return source
    .split('\n')
    .map((line) => {
      if (FENCE.test(line)) {
        fenced = !fenced;
        return '';
      }
      return fenced ? '' : line;
    })
    .join('\n');
}

/** What is left once comments, inline code and known elements are blanked out: markup. */
function withoutCode(source: string): string {
  return source
    .replace(COMMENT, blank)
    .replace(CODE_SPAN, blank)
    .replace(AUTOLINK, blank)
    .replace(TAG, (whole, name: string) =>
      HTML_ELEMENTS.has(name.toLowerCase()) ? blank(whole) : whole,
    );
}

function headingOf(blocks: readonly Block[], where: string): string {
  for (const block of blocks) {
    const first = block.lines[0] ?? '';
    if (FENCE.test(first)) continue;
    const match = HEADING.exec(first);
    if (match?.[1] === '#' && match[2] !== undefined) return flattenInline(match[2]);
  }
  throw new Error(
    `${where}: no top-level heading. The site takes every page title from its first "# " line, ` +
      'so a page without one has no name to be listed under.',
  );
}

const DESCRIPTION_LIMIT = 160;

/**
 * The lead paragraph, or nothing.
 *
 * A paragraph counts as prose only if it contains a sentence terminator. That single test
 * is what separates an opening sentence from the three things that reliably precede one
 * here: a `Cluster: images` metadata line, a `<sub>interface</sub>` marker on a generated
 * API page, and a status line of dates and separators. None of them ends a sentence, and
 * none of them is worth putting in a `<meta name="description">`. A page that never reaches
 * prose gets none, and the site's own description stands in; inventing one from the
 * filename would put a sentence nobody wrote in front of a search engine.
 */
function descriptionOf(blocks: readonly Block[]): string | undefined {
  let seenHeading = false;
  for (const block of blocks) {
    const first = block.lines[0]?.trimStart() ?? '';
    if (FENCE.test(first)) continue;
    if (HEADING.test(first)) {
      seenHeading = true;
      continue;
    }
    if (!seenHeading) continue;
    if (first.startsWith('<!--') || first.startsWith('|')) continue;
    if (/^(?:[-*+]\s|\d+\.\s)/.test(first)) continue;
    const text = flattenInline(block.lines.join(' ').replace(/^\s*>\s?/gm, ''));
    if (text === '' || !/[.?!]/.test(text)) continue;
    return truncate(text, DESCRIPTION_LIMIT);
  }
  return undefined;
}

/** Markdown inline syntax removed, whitespace collapsed. What a search result would show. */
function flattenInline(text: string): string {
  return text
    .replace(/<!--[\s\S]*?-->/g, '')
    .replace(/!?\[([^\]]*)\]\([^)]*\)/g, '$1')
    .replace(/<[^>]+>/g, '')
    .replace(/[`*_]/g, '')
    .replace(/\s+/g, ' ')
    .trim();
}

function truncate(text: string, limit: number): string {
  if (text.length <= limit) return text;
  const cut = text.slice(0, limit);
  const lastSpace = cut.lastIndexOf(' ');
  return `${(lastSpace > 0 ? cut.slice(0, lastSpace) : cut).replace(/[,;:.]$/, '')}...`;
}

// ---------------------------------------------------------------------------
// Links
// ---------------------------------------------------------------------------

const LINK = /\]\(([^()\s]+)((?:\s+"[^"]*")?)\)/g;
const ABSOLUTE = /^(?:[a-z][a-z0-9+.-]*:|\/\/)/i;

/**
 * Rewrites the links `docs/` writes for a reader of the repository into links a reader of
 * the site can follow. A link inside a code fence is sample text and not navigation, so
 * fences are left alone.
 */
function rewriteLinks(
  source: string,
  sourcePath: string,
  bySlug: ReadonlyMap<string, string>,
): string {
  let fenced = false;
  return source
    .split('\n')
    .map((line, index) => {
      if (FENCE.test(line)) {
        fenced = !fenced;
        return line;
      }
      if (fenced) return line;
      const where = `docs/${sourcePath}:${index + 1}`;
      return line.replace(
        LINK,
        (_whole, target: string, title: string) =>
          `](${rewriteTarget(target, sourcePath, bySlug, where)}${title})`,
      );
    })
    .join('\n');
}

function rewriteTarget(
  target: string,
  sourcePath: string,
  bySlug: ReadonlyMap<string, string>,
  where: string,
): string {
  if (target.startsWith('#') || ABSOLUTE.test(target)) return target;

  const hashAt = target.indexOf('#');
  const path = hashAt === -1 ? target : target.slice(0, hashAt);
  const hash = hashAt === -1 ? '' : target.slice(hashAt);
  const isDirectory = path.endsWith('/');

  // Resolved against the linking page, the way the reader's editor resolves it.
  const resolved = posix.normalize(posix.join(posix.dirname(sourcePath), path));

  if (resolved.startsWith('../')) {
    const outside = resolved.slice('../'.length);
    if (!exists(resolve(DOCS_DIR, '..', outside))) {
      throw new Error(`${where}: link "${target}" leaves docs/ and points at nothing: ${outside}`);
    }
    // A directory has no blob. GitHub serves one under /tree/, and saying so keeps the link
    // off a redirect.
    return `${isDirectory ? treeUrl(outside) : blobUrl(outside)}${hash}`;
  }

  const slug = isDirectory ? `${resolved.replace(/\/$/, '')}/index` : slugFor(resolved);
  if (!bySlug.has(slug)) {
    throw new Error(
      `${where}: link "${target}" resolves to docs/${resolved}, which the site does not publish.` +
        (isDirectory ? ' A directory is published through its README.md, and there is none.' : ''),
    );
  }
  return `${routeFor(slug)}${hash}`;
}

function exists(path: string): boolean {
  try {
    statSync(path);
    return true;
  } catch {
    return false;
  }
}
