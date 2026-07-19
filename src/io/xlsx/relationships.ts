// The OOXML/OPC namespace URIs, relationship-type URIs, and `.rels` part envelopes shared by the
// writer's serialisers. The canonical namespace set lives in `namespaces.ts`; the writer-local
// groupings here (the props vocabularies, the content-types namespace, and the rel-type URIs derived
// from them) are plumbing every part of the writer references.

import {PKG_RELS_NS, RELATIONSHIPS_NS, SPREADSHEETML_NS} from './namespaces.ts';
import {escapeAttr, XML_DECLARATION} from './xml.ts';

export const NS = {
  contentTypes: 'http://schemas.openxmlformats.org/package/2006/content-types',
  packageRels: PKG_RELS_NS,
  main: SPREADSHEETML_NS,
  docRels: RELATIONSHIPS_NS,
  coreProps: 'http://schemas.openxmlformats.org/package/2006/metadata/core-properties',
  extProps: 'http://schemas.openxmlformats.org/officeDocument/2006/extended-properties',
  dc: 'http://purl.org/dc/elements/1.1/',
  dcterms: 'http://purl.org/dc/terms/',
  dcmitype: 'http://purl.org/dc/dcmitype/',
  xsi: 'http://www.w3.org/2001/XMLSchema-instance',
} as const;

export const REL = {
  worksheet: `${NS.docRels}/worksheet`,
  styles: `${NS.docRels}/styles`,
  theme: `${NS.docRels}/theme`,
  officeDocument: `${NS.docRels}/officeDocument`,
  coreProps: `${NS.packageRels}/metadata/core-properties`,
  extProps: `${NS.docRels}/extended-properties`,
  table: `${NS.docRels}/table`,
  comments: `${NS.docRels}/comments`,
  vmlDrawing: `${NS.docRels}/vmlDrawing`,
  drawing: `${NS.docRels}/drawing`,
  printerSettings: `${NS.docRels}/printerSettings`,
  image: `${NS.docRels}/image`,
  hyperlink: `${NS.docRels}/hyperlink`,
  sharedStrings: `${NS.docRels}/sharedStrings`,
  pivotTable: `${NS.docRels}/pivotTable`,
  pivotCacheDefinition: `${NS.docRels}/pivotCacheDefinition`,
  pivotCacheRecords: `${NS.docRels}/pivotCacheRecords`,
} as const;

export function relationship(id: string, type: string, target: string): string {
  return `<Relationship Id="${id}" Type="${type}" Target="${target}"/>`;
}

export function preservedRelsXml(
  rels: readonly {id: string; type: string; target: string}[],
): string {
  const body = rels.map((rel) => relationship(rel.id, rel.type, escapeAttr(rel.target))).join('');
  return `${XML_DECLARATION}<Relationships xmlns="${NS.packageRels}">${body}</Relationships>`;
}

// A `.rels` part for a generated part chain (pivot table → cache definition → cache records). Targets
// are writer-controlled package paths, so no attribute escaping is needed.
export function relsPartXml(rels: readonly {id: string; type: string; target: string}[]): string {
  const body = rels.map((rel) => relationship(rel.id, rel.type, rel.target)).join('');
  return `${XML_DECLARATION}<Relationships xmlns="${NS.packageRels}">${body}</Relationships>`;
}
