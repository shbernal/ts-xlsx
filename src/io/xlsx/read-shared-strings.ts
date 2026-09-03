// Parser for the shared-string table (`xl/sharedStrings.xml`), the pool that `t="s"` cells index into.
// Split out of read.ts beside its sibling parsers (read-styles.ts, rich-runs.ts) so read.ts stays
// orchestration; the run structure it decodes is owned by RunAccumulator, shared with inline strings.

import {parseXml} from '../../xml/xml-read.ts';
import {localName} from '../../xml/xml-scan.ts';
import type {SharedString} from './cell-value.ts';
import {RunAccumulator} from './rich-runs.ts';

const SHARED_STRING_EMPTY_CLOSES: ReadonlySet<string> = new Set(['si']);

// Shared strings resolve `t="s"` cells. Each `<si>` is one entry: a plain `<si><t>…</t>` decodes to a
// string, while a rich `<si><r><rPr>…</rPr><t>…</t></r>…` decodes to a {@link RichTextValue} whose runs
// carry their per-run fonts, so rich text Excel pooled reads back formatted, not flattened to text.
// The run structure inside an `<si>` is identical to an inline string's `<is>`, which is why the two
// readers share one element machine rather than a comment saying they agree: everything below the
// `<si>` is `RunAccumulator`'s, and what is left here is what committing an entry means.
export function parseSharedStrings(xml: string): SharedString[] {
  if (xml === '') return [];
  const strings: SharedString[] = [];
  const runs = new RunAccumulator({container: 'si', readRuns: true});
  parseXml(
    xml,
    {
      onOpen(name, attrs, selfClosing) {
        runs.open(localName(name), attrs, selfClosing);
      },
      onText(chunk) {
        runs.text(chunk);
      },
      onClose(name) {
        // An `<si>` is rich the moment it holds one `<r>`, at which point its runs, not its bare
        // `<t>` text, become the entry.
        if (runs.close(localName(name)) !== 'container') return;
        strings.push(runs.isRich ? {richText: runs.runs} : runs.plainText);
      },
    },
    // An empty pooled string is legally written `<si/>`, which fires an open and no close. Without
    // this it commits no entry, and every later index shifts by one, so every `t="s"` cell past it
    // resolves to the wrong string rather than to a missing one.
    {closeEmptyElements: SHARED_STRING_EMPTY_CLOSES},
  );
  return strings;
}
