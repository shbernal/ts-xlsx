// The OPC (Open Packaging Conventions) layer of the reader: resolving relationship targets to part
// paths, reading a part's `.rels`, resolving a part's declared content type, and walking the transitive
// closure of parts a preserved reference reaches. Every helper here is pure over the inflated package —
// it takes part text/bytes accessors and returns paths or records, touching no Workbook model.

import type {PreservedPart, PreservedRelationship} from '../../core/preserved.ts';
import {openElements} from './xml-read.ts';

// The extension of a part path (`xl/media/image1.png` → `png`), or '' when it carries none.
export function extensionOf(partPath: string): string {
  const dot = partPath.lastIndexOf('.');
  return dot === -1 ? '' : partPath.slice(dot + 1);
}

// The relationships for `dir/name.ext` live at `dir/_rels/name.ext.rels`.
export function relsPathFor(partPath: string): string {
  const slash = partPath.lastIndexOf('/');
  const dir = slash === -1 ? '' : partPath.slice(0, slash + 1);
  const base = slash === -1 ? partPath : partPath.slice(slash + 1);
  return `${dir}_rels/${base}.rels`;
}

// The Target of the first relationship whose Type ends with `/<suffix>` (local-name match, so a
// namespaced or oddly-cased type still resolves), or undefined when none is declared.
export function relationshipTargetByType(xml: string, suffix: string): string | undefined {
  for (const {attrs} of openElements(xml, 'Relationship')) {
    if (
      attrs.Type !== undefined &&
      attrs.Target !== undefined &&
      attrs.Type.endsWith(`/${suffix}`)
    ) {
      return attrs.Target;
    }
  }
  return undefined;
}

// Every Target whose Type ends with `/<suffix>`, in declaration order — for a part class a sheet may
// reference more than once (a sheet can own several tables), where the singular helper's first-match
// would miss all but one.
export function relationshipTargetsByType(xml: string, suffix: string): string[] {
  const targets: string[] = [];
  for (const {attrs} of openElements(xml, 'Relationship')) {
    if (
      attrs.Type !== undefined &&
      attrs.Target !== undefined &&
      attrs.Type.endsWith(`/${suffix}`)
    ) {
      targets.push(attrs.Target);
    }
  }
  return targets;
}

// Resolve a relationship target (relative to the referencing part's directory, or absolute from the
// package root) into a package part path, collapsing `.`/`..` segments.
export function resolveRelativePart(basePart: string, target: string): string {
  if (target.startsWith('/')) return target.slice(1);
  const baseDir = basePart.slice(0, basePart.lastIndexOf('/') + 1);
  const out: string[] = [];
  for (const segment of `${baseDir}${target}`.split('/')) {
    if (segment === '' || segment === '.') continue;
    if (segment === '..') out.pop();
    else out.push(segment);
  }
  return out.join('/');
}

// A workbook relationship target is relative to the `xl/` directory (`worksheets/sheet1.xml`)
// or absolute from the package root (`/xl/worksheets/sheet1.xml`); normalise both to a part path.
export function resolveWorkbookPart(target: string): string {
  if (target.startsWith('/')) return target.slice(1);
  return `xl/${target.replace(/^\.\//, '')}`;
}

export function parseRelationships(xml: string): Map<string, string> {
  const rels = new Map<string, string>();
  for (const {attrs} of openElements(xml, 'Relationship')) {
    if (attrs.Id !== undefined && attrs.Target !== undefined) {
      rels.set(attrs.Id, attrs.Target);
    }
  }
  return rels;
}

// A relationship as declared, with the fields a preserved-part closure needs: its id, Type URI,
// Target, and whether the target lies outside the package (`TargetMode="External"`). A fuller record
// than {@link parseRelationships}'s id→target map, which the closure walk uses to skip external
// targets and carry each relationship's type through a re-write.
export interface RelationshipRecord {
  readonly id: string;
  readonly type: string;
  readonly target: string;
  readonly external: boolean;
}

export function parseRelationshipRecords(xml: string): RelationshipRecord[] {
  const records: RelationshipRecord[] = [];
  for (const {attrs} of openElements(xml, 'Relationship')) {
    if (attrs.Id !== undefined && attrs.Type !== undefined && attrs.Target !== undefined) {
      records.push({
        id: attrs.Id,
        type: attrs.Type,
        target: attrs.Target,
        external: attrs.TargetMode === 'External',
      });
    }
  }
  return records;
}

// Resolve a package part path to its declared content type the way OPC does: an `<Override>` naming
// the exact part wins, else the `<Default>` registered for the part's extension. An unknown part
// falls back to the generic binary type so re-declaring it never emits an empty content type.
export function contentTypeResolver(contentTypesXml: string): (path: string) => string {
  const overrides = new Map<string, string>();
  const defaults = new Map<string, string>();
  for (const {local, attrs} of openElements(contentTypesXml, 'Override', 'Default')) {
    if (local === 'Override' && attrs.PartName !== undefined && attrs.ContentType !== undefined) {
      overrides.set(attrs.PartName, attrs.ContentType);
    } else if (
      local === 'Default' &&
      attrs.Extension !== undefined &&
      attrs.ContentType !== undefined
    ) {
      defaults.set(attrs.Extension.toLowerCase(), attrs.ContentType);
    }
  }
  return (path: string): string =>
    overrides.get(`/${path}`) ??
    defaults.get(extensionOf(path).toLowerCase()) ??
    'application/octet-stream';
}

// Gather the transitive closure of package parts reachable from an entry part — the part itself, then
// every internal part its relationships target, breadth-first — each with its raw bytes, content type,
// and (internal) relationships. Returns undefined when the entry part is absent (a dangling reference
// preserves nothing). A `visited` set dedupes shared parts and bounds the walk to the (finite,
// inflate-capped) package, so a maliciously self-referential rels graph cannot loop.
export function capturePartClosure(
  entryPath: string,
  partText: (path: string) => string | undefined,
  partBytes: (path: string) => Uint8Array | undefined,
  contentTypeOf: (path: string) => string,
): readonly PreservedPart[] | undefined {
  const parts: PreservedPart[] = [];
  const visited = new Set<string>();
  const queue: string[] = [entryPath];
  while (queue.length > 0) {
    const path = queue.shift();
    if (path === undefined || visited.has(path)) continue;
    visited.add(path);
    const bytes = partBytes(path);
    if (bytes === undefined) continue;
    const relsXml = partText(relsPathFor(path));
    const rels: PreservedRelationship[] = [];
    if (relsXml !== undefined) {
      for (const rel of parseRelationshipRecords(relsXml)) {
        if (rel.external) continue;
        const targetPath = resolveRelativePart(path, rel.target);
        rels.push({id: rel.id, type: rel.type, targetPath});
        queue.push(targetPath);
      }
    }
    parts.push({path, contentType: contentTypeOf(path), bytes, rels});
  }
  return parts.some((part) => part.path === entryPath) ? parts : undefined;
}
