/**
 * The emitted `.xlsx` as what it actually is: a zip of XML parts.
 *
 * This is the panel no competing library's site offers, and the reason it can exist here is
 * that the reader already has the bytes in their own tab. Unzipping is done with `fflate`,
 * which is the library's one runtime dependency and therefore already in the bundle; adding
 * a second zip reader to look inside a file this page has just written would be absurd.
 *
 * Nothing here touches a DOM.
 */

import {unzipSync} from 'fflate';

export interface PackageEntry {
  /** The path inside the package: `xl/worksheets/sheet1.xml`. */
  readonly path: string;
  readonly byteLength: number;
  /** Text when the part is text, so the panel can show it. Absent for a binary part. */
  readonly text: string | undefined;
}

/** Extensions an OOXML package stores as text. Everything else is shown as a size only. */
const TEXT_EXTENSIONS = new Set(['xml', 'rels', 'txt', 'vml', 'json']);

// `ignoreBOM` is left at its default, which strips a leading byte-order mark rather than
// handing it back as a character. The parts this reads are UTF-8 with a BOM.
const decoder = new TextDecoder('utf-8');

/**
 * Every part of the package, in path order.
 *
 * Sorted rather than left in zip order because the panel is a table of contents, and a
 * reader looking for `xl/styles.xml` should find it where the alphabet puts it. Zip order is
 * a writer's choice and carries no meaning for them.
 */
export function listParts(bytes: Uint8Array): readonly PackageEntry[] {
  const unzipped = unzipSync(bytes);
  return Object.entries(unzipped)
    .map(([path, content]) => ({
      path,
      byteLength: content.length,
      text: isText(path) ? decoder.decode(content) : undefined,
    }))
    .sort((left, right) => (left.path < right.path ? -1 : left.path > right.path ? 1 : 0));
}

function isText(path: string): boolean {
  const dot = path.lastIndexOf('.');
  return dot !== -1 && TEXT_EXTENSIONS.has(path.slice(dot + 1).toLowerCase());
}

const TOKEN = /<[^>]*>|[^<]+/g;

/**
 * XML with one element per line, indented.
 *
 * For display only, and the panel says so beside it: the writer emits these parts with no
 * padding at all, so a reader who took this indentation for the emitted bytes would be
 * reading a file that does not exist.
 *
 * Text is never rewritten, only moved: an element whose whole content is text keeps that
 * text on its own line, verbatim, spaces included. That matters more here than it looks,
 * because `xml:space="preserve"` is how a cell holding `" "` says so, and a pretty-printer
 * that trimmed would show a different workbook from the one in the file.
 */
export function prettyXml(xml: string, indent = '  '): string {
  const tokens = xml.match(TOKEN) ?? [];
  const lines: string[] = [];
  let depth = 0;
  let index = 0;

  const pad = (): string => indent.repeat(depth);

  while (index < tokens.length) {
    const token = tokens[index] ?? '';
    index += 1;

    if (!token.startsWith('<')) {
      // Whitespace between two tags is the source's own layout, not content: the only text
      // that can be significant sits alone inside one element, and the open-text-close case
      // below takes that verbatim before it ever reaches here. Keeping it would put a blank
      // line under every declaration and reproduce a dropped file's indentation on top of
      // this one.
      if (token.trim() !== '') lines.push(pad() + token);
      continue;
    }
    if (token.startsWith('</')) {
      depth = Math.max(0, depth - 1);
      lines.push(pad() + token);
      continue;
    }
    if (token.startsWith('<?') || token.startsWith('<!') || token.endsWith('/>')) {
      lines.push(pad() + token);
      continue;
    }

    const text = tokens[index];
    const close = tokens[index + 1];
    if (
      text !== undefined &&
      !text.startsWith('<') &&
      close !== undefined &&
      close.startsWith('</')
    ) {
      lines.push(pad() + token + text + close);
      index += 2;
      continue;
    }
    lines.push(pad() + token);
    depth += 1;
  }
  return lines.join('\n');
}
