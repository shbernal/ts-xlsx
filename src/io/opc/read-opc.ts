// The OPC (Open Packaging Conventions) layer of the reader: reading a part's `.rels`, resolving a
// part's declared content type, and walking the transitive closure of parts a preserved reference
// reaches. Every helper here is pure over the inflated package: it takes part text/bytes accessors
// and returns paths or records, touching no Workbook model. The path arithmetic those answers are
// expressed in is `part-paths.ts`, which holds both directions of it.

import {strFromU8} from 'fflate';

import type {PreservedPart, PreservedRelationship} from '../../core/preserved.ts';
import {openElements} from '../../xml/xml-read.ts';
import {extensionOf, relsPathFor, resolveRelativePart} from './part-paths.ts';
import {DEFAULT_MAX_UNCOMPRESSED} from './read-options.ts';
import {inflateSpreadsheetPackage} from './sniff-format.ts';

// The two ways a reader reaches into an inflated package: a part's UTF-8-decoded text, or its raw
// bytes. Built once per read (see {@link packageAccessors}) so the buffered and streaming readers
// share one implementation of the lookup rather than each re-declaring the same closures.
export interface PackageAccessors {
  /** A part's decoded text, or undefined when the package holds no such part. Decodes lazily, so a
   * part the reader never asks for is never stringified. */
  partText: (path: string) => string | undefined;
  /** A part's raw bytes, or undefined when the package holds no such part. */
  partBytes: (path: string) => Uint8Array | undefined;
}

/** A spreadsheet package opened for reading: inflated under the read bound, its parts bound to
 * accessors, and its XML office document read if it has one. */
export interface OpenedSpreadsheet {
  /** The inflated parts, for a reader that hands the whole package on to another codec. */
  readonly files: Record<string, Uint8Array>;
  readonly pkg: PackageAccessors;
  /** `xl/workbook.xml`, or undefined when the package carries no XML office document (a `.xlsb`
   * carries `xl/workbook.bin` instead). Each entry point answers that case for itself: they
   * genuinely want different answers, and the difference is documented where they diverge. */
  readonly workbookXml: string | undefined;
}

/**
 * Open a spreadsheet package: the six-step preamble every reader shares, and in particular the two
 * decisions worth having exactly one of. The inflate bound is a security decision (an unbounded
 * inflate is a zip bomb) and "does this package carry an XML office document" is the dispatch
 * decision between the two codecs, so neither should be restated once per entry point.
 *
 * It deliberately stops short of the shared strings and the style table: those are codec-specific
 * (`xl/sharedStrings.xml` against BIFF12's own record stream), and lifting them here would put
 * SpreadsheetML knowledge into the container layer.
 */
export function openSpreadsheetPackage(
  data: Uint8Array,
  maxUncompressedBytes: number | undefined,
): OpenedSpreadsheet {
  const files = inflateSpreadsheetPackage(data, maxUncompressedBytes ?? DEFAULT_MAX_UNCOMPRESSED);
  const pkg = packageAccessors(files);
  return {files, pkg, workbookXml: pkg.partText('xl/workbook.xml')};
}

// Bind the part-lookup accessors over an inflated package (a part-path → bytes map).
//
// Both members are declared as function-typed properties rather than with method syntax, and both
// are written as arrows here, because every reader destructures them off the returned object:
// `const {partText, partBytes} = packageAccessors(files)`. Method syntax would say these values
// may read `this`, which they never do (they close over `files`), and would make each of those
// fifteen destructurings report as an unbound method. Property syntax is also the stricter
// declaration: method-syntax parameters are checked bivariantly even under `strictFunctionTypes`.
export function packageAccessors(files: Record<string, Uint8Array>): PackageAccessors {
  return {
    partText: (path: string): string | undefined => {
      const bytes = files[path];
      return bytes === undefined ? undefined : strFromU8(bytes);
    },
    partBytes: (path: string): Uint8Array | undefined => files[path],
  };
}

// The Target of the first relationship whose Type ends with `/<suffix>`, or undefined when none is
// declared. For a single expected reference, where the plural form below would over-gather.
export function relationshipTargetByType(xml: string, suffix: string): string | undefined {
  return relationshipTargetsByType(xml, suffix)[0];
}

// Every Target whose Type ends with `/<suffix>`, in declaration order. The type is matched on its
// final segment (a local-name match, so a namespaced or oddly-cased type still resolves). For a part
// class a sheet may reference more than once (a sheet can own several tables), where the singular
// helper's first-match would miss all but one.
export function relationshipTargetsByType(xml: string, suffix: string): string[] {
  return parseRelationshipRecords(xml)
    .filter((record) => record.type.endsWith(`/${suffix}`))
    .map((record) => record.target);
}

// A relationship as declared, with the fields a preserved-part closure needs: its id, Type URI,
// Target, and whether the target lies outside the package (`TargetMode="External"`). This is the one
// shape a `.rels` part is read into; the narrower views above and below are projections of it, so the
// element is scanned once and there is one answer to what counts as a relationship.
export interface RelationshipRecord {
  readonly id: string;
  readonly type: string;
  readonly target: string;
  readonly external: boolean;
}

// OPC makes Id, Type and Target all mandatory, so an element missing any of the three is not a
// relationship and is skipped rather than half-read.
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

// One part's `.rels`, parsed once and then queried many times, with the owning part's path bound in so
// every answer comes back as a package path rather than a target still needing resolution.
//
// A worksheet is the reason this exists. Its rels part is the index to nearly everything hanging off the
// sheet (notes, threads, printer settings, the drawing, the background image, tables, pivots, and the
// preserved-reference closure) and each of those lookups used to re-read and re-parse the same XML,
// eight times per sheet on a workbook of any size.
export interface PartRelationships {
  /** Every relationship the part declares, in declaration order. */
  readonly records: readonly RelationshipRecord[];
  /** The relationship with this id, or undefined. Its `target` is raw: an external one is a URL, not a
   * package path, so a caller that may see `TargetMode="External"` must read it before resolving. */
  byId(id: string): RelationshipRecord | undefined;
  /** Resolve one of this part's targets against the part's own directory. */
  pathOf(target: string): string;
  /** The package part reached through the first relationship whose Type ends with `/<suffix>`, or
   * undefined when the part declares none: the single-part lookup (notes, printer settings, drawing,
   * background) in one call. */
  targetPath(suffix: string): string | undefined;
  /** Every package part reached through a relationship of this type, in declaration order. For a part
   * class one sheet may reference more than once (tables, pivot tables). */
  targetPaths(suffix: string): string[];
  /**
   * The text of the single part reached through a relationship of this type, or `undefined` when the
   * part declares no such relationship *or* names one the package does not contain.
   *
   * Those two are deliberately one answer. A relationship pointing at an absent part is a damaged
   * package, and every reader here answers a damaged package the same way Excel does: the feature is
   * simply not there, rather than the load failing over it. Collapsing `targetPath` then `partText`
   * into one call is what lets that contract be stated once instead of at each of the seven places
   * that used to spell it out.
   */
  relatedText(suffix: string): string | undefined;
  /** As {@link relatedText}, for a part whose content is opaque bytes (a printer-settings blob, an
   * image) rather than XML. */
  relatedBytes(suffix: string): Uint8Array | undefined;
}

// Read and parse a part's `.rels`. A part with no rels part yields an empty set rather than undefined,
// so a caller never has to distinguish "no relationships" from "no rels part": nothing downstream
// treats those two differently.
export function readPartRelationships(
  partPath: string,
  partText: (path: string) => string | undefined,
  partBytes?: (path: string) => Uint8Array | undefined,
): PartRelationships {
  const records = parseRelationshipRecords(partText(relsPathFor(partPath)) ?? '');
  const byId = new Map(records.map((record) => [record.id, record]));
  const pathOf = (target: string): string => resolveRelativePart(partPath, target);
  const targetsOf = (suffix: string): string[] =>
    records.filter((record) => record.type.endsWith(`/${suffix}`)).map((record) => record.target);
  const targetPath = (suffix: string): string | undefined => targetsOf(suffix).map(pathOf)[0];
  return {
    records,
    byId: (id) => byId.get(id),
    pathOf,
    targetPath,
    targetPaths: (suffix) => targetsOf(suffix).map(pathOf),
    relatedText: (suffix) => {
      const path = targetPath(suffix);
      return path === undefined ? undefined : partText(path);
    },
    // `partBytes` is optional because the two readers that only ever ask for XML (the streaming row
    // reader, the xlsb workbook part) have no bytes accessor to hand over; asking one of them for a
    // binary part is a bug in that reader rather than a package that lacks it.
    relatedBytes: (suffix) => {
      const path = targetPath(suffix);
      return path === undefined ? undefined : partBytes?.(path);
    },
  };
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
    overrides.get(`/${path}`) ?? defaults.get(extensionOf(path)) ?? 'application/octet-stream';
}

// Gather the transitive closure of package parts reachable from an entry part: the part itself, then
// every internal part its relationships target, breadth-first, each with its raw bytes, content type,
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
        if (rel.external) {
          // A linked workbook lives outside the package: keep the wiring verbatim (an externalLink
          // part's pointer to its source), but do not walk into it: there is no package part to visit.
          rels.push({id: rel.id, type: rel.type, targetPath: rel.target, external: true});
          continue;
        }
        const targetPath = resolveRelativePart(path, rel.target);
        rels.push({id: rel.id, type: rel.type, targetPath});
        queue.push(targetPath);
      }
    }
    parts.push({path, contentType: contentTypeOf(path), bytes, rels});
  }
  return parts.some((part) => part.path === entryPath) ? parts : undefined;
}
