// Small pure helpers for OPC part paths and numeric serialisation, shared across the writer's
// module cluster.

// The extension of a part path (`xl/media/image1.jpeg` → `jpeg`), or '' when it carries none.
export function extensionOf(partPath: string): string {
  const dot = partPath.lastIndexOf('.');
  const slash = partPath.lastIndexOf('/');
  return dot > slash ? partPath.slice(dot + 1) : '';
}

// The relationships part path for `dir/name.ext` → `dir/_rels/name.ext.rels`.
export function relsPathForPart(partPath: string): string {
  const slash = partPath.lastIndexOf('/');
  const dir = slash === -1 ? '' : partPath.slice(0, slash + 1);
  const base = slash === -1 ? partPath : partPath.slice(slash + 1);
  return `${dir}_rels/${base}.rels`;
}

// A relationship target expressed relative to the part that carries it: the `..` hops out of the
// referencing part's directory up to the common ancestor, then down to the target. Both paths are
// package-absolute (`xl/drawings/preservedP1.vml` → `xl/media/preservedP2.jpeg` → `../media/preservedP2.jpeg`).
export function relativePartPath(fromPath: string, toPath: string): string {
  const fromDir = fromPath.split('/').slice(0, -1);
  const toSegments = toPath.split('/');
  let common = 0;
  while (
    common < fromDir.length &&
    common < toSegments.length - 1 &&
    fromDir[common] === toSegments[common]
  ) {
    common++;
  }
  const up = fromDir.length - common;
  return [...Array<string>(up).fill('..'), ...toSegments.slice(common)].join('/');
}

// A finite number serialises as its shortest round-trippable decimal; a non-finite one
// has no OOXML numeric representation, so the writer refuses it rather than emit `NaN`.
export function numberText(value: number): string {
  if (!Number.isFinite(value)) {
    throw new Error(`cannot write a non-finite number (${value}) — it has no OOXML representation`);
  }
  return String(value);
}

export function range(n: number): number[] {
  return Array.from({length: n}, (_, i) => i);
}
