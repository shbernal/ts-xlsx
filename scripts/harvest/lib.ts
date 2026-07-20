// Shared core for the harvest toolchain.
//
// The harvest is a *one-time* operation: we pull the entire upstream backlog
// into a local queue once, then drain that queue by distilling each thread into
// durable product (a corpus case and/or spec note) and deleting the raw record.
// We never re-harvest to pick up new upstream activity.
//
// This module holds the pieces the CLI entry points share (fetch-issue,
// harvest-all, list-backlog, status) so the fetch logic lives in exactly one
// place. Auth comes from the `gh` CLI (already logged in); we shell out to
// `gh api` rather than embedding a token so credentials never touch this process.

import {execFile} from 'node:child_process';
import {access, mkdir, writeFile} from 'node:fs/promises';
import {dirname, resolve} from 'node:path';
import {fileURLToPath} from 'node:url';
import {promisify} from 'node:util';

const execFileAsync = promisify(execFile);

const HERE = dirname(fileURLToPath(import.meta.url));

export const REPO_ROOT = resolve(HERE, '..', '..');
export const BACKLOG_DIR = resolve(REPO_ROOT, 'docs', 'knowledge', 'backlog');
export const ISSUES_DIR = resolve(BACKLOG_DIR, 'issues');
export const ATTACHMENTS_DIR = resolve(BACKLOG_DIR, 'attachments');
export const MANIFEST_PATH = resolve(BACKLOG_DIR, 'manifest.json');
export const DEFAULT_REPO = 'exceljs/exceljs';

// A spreadsheet library ingests hostile input; the harvester ingests hostile
// URLs. Cap attachment downloads so a malicious/huge asset can't exhaust disk.
const MAX_ATTACHMENT_BYTES = 25 * 1024 * 1024;
const SPREADSHEET_EXTENSIONS = new Set(['xlsx', 'xlsm', 'xlsb', 'xls', 'csv', 'zip', 'ods']);

// --- External JSON shapes ----------------------------------------------------
// The `gh api` responses are untrusted JSON; these describe only the fields we
// read. Every property is optional because we defensively default each access.

interface GhLabel {
  name?: string;
}

interface GhReactions {
  total_count?: number;
}

interface GhUser {
  login?: string;
}

interface GhComment {
  user?: GhUser;
  author_association?: string;
  created_at?: string;
  reactions?: GhReactions;
  body?: string;
}

interface GhIssue {
  number?: number;
  title?: string;
  state?: string;
  html_url?: string;
  body?: string;
  comments?: number;
  user?: GhUser;
  created_at?: string;
  closed_at?: string;
  labels?: Array<string | GhLabel>;
  reactions?: GhReactions;
  pull_request?: unknown;
}

interface GhPullDetail {
  merged?: boolean;
  mergeable?: boolean | null;
  base?: {ref?: string};
  additions?: number;
  deletions?: number;
}

interface GhPullFile {
  filename?: string;
  status?: string;
  additions?: number;
  deletions?: number;
}

// --- Durable record shapes ---------------------------------------------------
// What we write to disk and pass between the CLI entry points.

export type ItemType = 'issue' | 'pull_request';

export interface ManifestItem {
  number: number;
  type: ItemType;
  title: string;
  labels: Array<string | undefined>;
  reactions: number;
  comments: number;
  url: string | undefined;
}

export interface Manifest {
  schema: string;
  repo: string;
  generatedAt: string;
  total: number;
  issues: number;
  pullRequests: number;
  harvestComplete: boolean;
  items: ManifestItem[];
}

interface NormalizedComment {
  author: string | null;
  authorAssociation: string | null;
  createdAt: string | null;
  reactions: number;
  body: string;
}

interface DiscoveredAttachment {
  url: string;
  ext: string | null;
  isFixture: boolean;
}

interface ChangedFile {
  path: string | undefined;
  status: string | undefined;
  additions: number | undefined;
  deletions: number | undefined;
}

interface PrSummary {
  merged: boolean;
  mergeable: boolean | null;
  baseRef: string | null;
  additions: number | null;
  deletions: number | null;
  changedFiles: ChangedFile[];
}

interface DownloadResult {
  name: string;
  bytes?: number;
  error?: string;
}

export interface HarvestSummary {
  number: number;
  recordPath: string;
  skipped: boolean;
  type?: ItemType;
  commentCount?: number;
  attachmentCount?: number;
  fixtureCount?: number;
  downloaded?: DownloadResult[];
}

export interface HarvestOptions {
  number: number;
  repo?: string;
  skipIfExists?: boolean;
}

export function padNumber(number: number): string {
  return String(number).padStart(4, '0');
}

export function recordPathFor(number: number): string {
  return resolve(ISSUES_DIR, `${padNumber(number)}.json`);
}

export function isValidRepo(repo: string): boolean {
  return /^[\w.-]+\/[\w.-]+$/.test(repo);
}

export async function fileExists(path: string): Promise<boolean> {
  try {
    await access(path);
    return true;
  } catch {
    return false;
  }
}

export async function gh(path: string, {paginate = false} = {}): Promise<unknown> {
  const argv = ['api', path];
  if (paginate) argv.push('--paginate');
  const {stdout} = await execFileAsync('gh', argv, {maxBuffer: 128 * 1024 * 1024});
  if (paginate) {
    // `gh --paginate` concatenates JSON arrays as `][` at page seams; splice them.
    return JSON.parse(stdout.replace(/\]\s*\[/g, ','));
  }
  return JSON.parse(stdout);
}

function labelName(label: string | GhLabel): string | undefined {
  return typeof label === 'string' ? label : label.name;
}

// GitHub's issues endpoint returns PRs too (they carry a `pull_request` key), so
// one paginated call enumerates the whole open universe. Used once by
// list-backlog to snapshot the manifest.
export async function listOpenItems(repo = DEFAULT_REPO): Promise<ManifestItem[]> {
  const raw = (await gh(`repos/${repo}/issues?state=open&per_page=100`, {
    paginate: true,
  })) as GhIssue[];
  return raw.map((item) => ({
    number: item.number ?? 0,
    type: item.pull_request ? 'pull_request' : 'issue',
    title: item.title ?? '',
    labels: (item.labels ?? []).map(labelName),
    reactions: item.reactions?.total_count ?? 0,
    comments: item.comments ?? 0,
    url: item.html_url,
  }));
}

// Pull spreadsheet-shaped asset links out of markdown/text bodies. GitHub serves
// user uploads from a handful of hosts; we record every match and download only
// the ones that look like real fixtures (by extension).
const URL_RE = /https?:\/\/[^\s)"'<>\]]+/g;
function discoverAttachments(text: string): DiscoveredAttachment[] {
  if (!text) return [];
  const out: DiscoveredAttachment[] = [];
  for (const url of text.match(URL_RE) ?? []) {
    const clean = url.replace(/[.,);]+$/, '');
    const lastSegment = clean.split('?')[0]?.split('#')[0]?.split('/').pop() ?? '';
    // Only treat a trailing `.foo` on the final path segment as an extension —
    // a bare URL like `.../pull/636` has no filename and no fixture.
    const match = lastSegment.match(/\.([a-z0-9]{1,5})$/i);
    const ext = match?.[1]?.toLowerCase() ?? null;
    out.push({url: clean, ext, isFixture: ext !== null && SPREADSHEET_EXTENSIONS.has(ext)});
  }
  return out;
}

async function downloadFixture(url: string, destPath: string): Promise<number> {
  const res = await fetch(url, {redirect: 'follow'});
  if (!res.ok) throw new Error(`HTTP ${res.status} for ${url}`);
  const length = Number(res.headers.get('content-length') ?? '0');
  if (length > MAX_ATTACHMENT_BYTES) {
    throw new Error(`Refusing ${length} bytes (> ${MAX_ATTACHMENT_BYTES} cap): ${url}`);
  }
  const buf = Buffer.from(await res.arrayBuffer());
  if (buf.byteLength > MAX_ATTACHMENT_BYTES) {
    throw new Error(`Refusing ${buf.byteLength} bytes (> ${MAX_ATTACHMENT_BYTES} cap): ${url}`);
  }
  await mkdir(dirname(destPath), {recursive: true});
  await writeFile(destPath, buf);
  return buf.byteLength;
}

function normalizeComment(c: GhComment): NormalizedComment {
  return {
    author: c.user?.login ?? null,
    authorAssociation: c.author_association ?? null,
    createdAt: c.created_at ?? null,
    reactions: c.reactions?.total_count ?? 0,
    body: c.body ?? '',
  };
}

// Harvest a single thread into a durable JSON record plus any spreadsheet-shaped
// attachments. This is the atom Phase 1 fans out across the whole backlog.
// Returns a summary; with `skipIfExists`, a thread already on disk is left
// untouched so a bulk fill is cheaply resumable.
export async function harvestOne({
  number,
  repo = DEFAULT_REPO,
  skipIfExists = false,
}: HarvestOptions): Promise<HarvestSummary> {
  if (!Number.isInteger(number) || number <= 0) {
    throw new Error(`Expected a positive issue/PR number, got: ${number}`);
  }
  if (!isValidRepo(repo)) {
    throw new Error(`Invalid repo: ${repo}`);
  }

  const padded = padNumber(number);
  const recordPath = recordPathFor(number);
  if (skipIfExists && (await fileExists(recordPath))) {
    return {number, recordPath, skipped: true};
  }

  const issue = (await gh(`repos/${repo}/issues/${number}`)) as GhIssue;
  const isPR = Boolean(issue.pull_request);
  const type: ItemType = isPR ? 'pull_request' : 'issue';
  const comments =
    (issue.comments ?? 0) > 0
      ? ((await gh(`repos/${repo}/issues/${number}/comments`, {paginate: true})) as GhComment[])
      : [];

  const attachmentSources = [issue.body, ...comments.map((c) => c.body)].join('\n');
  const attachments = discoverAttachments(attachmentSources);

  // PRs carry their real value in the reproduction and the changed-file map, not
  // the diff itself (we are deleting the code the diff targets). Capture the shape.
  let pr: PrSummary | undefined;
  if (isPR) {
    const [detail, files] = (await Promise.all([
      gh(`repos/${repo}/pulls/${number}`),
      gh(`repos/${repo}/pulls/${number}/files`, {paginate: true}),
    ])) as [GhPullDetail, GhPullFile[]];
    pr = {
      merged: detail.merged ?? false,
      mergeable: detail.mergeable ?? null,
      baseRef: detail.base?.ref ?? null,
      additions: detail.additions ?? null,
      deletions: detail.deletions ?? null,
      changedFiles: files.map((f) => ({
        path: f.filename,
        status: f.status,
        additions: f.additions,
        deletions: f.deletions,
      })),
    };
  }

  const record = {
    schema: 'ts-xlsx/backlog-item@1',
    repo,
    number,
    type,
    url: issue.html_url,
    title: issue.title,
    state: issue.state,
    author: issue.user?.login ?? null,
    createdAt: issue.created_at ?? null,
    closedAt: issue.closed_at ?? null,
    labels: (issue.labels ?? []).map(labelName),
    reactions: issue.reactions?.total_count ?? 0,
    commentCount: issue.comments ?? 0,
    body: issue.body ?? '',
    comments: comments.map(normalizeComment),
    attachments,
    pr,
  };

  await mkdir(ISSUES_DIR, {recursive: true});
  await writeFile(recordPath, `${JSON.stringify(record, null, 2)}\n`);

  const fixtures = attachments.filter((a) => a.isFixture);
  const downloaded: DownloadResult[] = [];
  for (const [i, att] of fixtures.entries()) {
    const name = `${padded}-${i}.${att.ext}`;
    const dest = resolve(ATTACHMENTS_DIR, padded, name);
    try {
      const bytes = await downloadFixture(att.url, dest);
      downloaded.push({name, bytes});
    } catch (err) {
      downloaded.push({name, error: errorMessage(err)});
    }
  }

  return {
    number,
    recordPath,
    skipped: false,
    type: record.type,
    commentCount: comments.length,
    attachmentCount: attachments.length,
    fixtureCount: fixtures.length,
    downloaded,
  };
}

// Coerce an unknown thrown value into a printable message without assuming Error.
export function errorMessage(err: unknown): string {
  if (err instanceof Error) return err.message;
  return String(err);
}
