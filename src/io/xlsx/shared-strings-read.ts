// Parser for the shared-string table (`xl/sharedStrings.xml`), the pool that `t="s"` cells index into.
// Split out of read.ts beside its sibling parsers (read-styles.ts, rich-runs.ts) so read.ts stays
// orchestration; the run structure it decodes is owned by RunAccumulator, shared with inline strings.

import type {SharedString} from './cell-value.ts';
import {RunAccumulator} from './rich-runs.ts';
import {localName, parseXml} from './xml-read.ts';

// Shared strings resolve `t="s"` cells. Each `<si>` is one entry: a plain `<si><t>…</t>` decodes to a
// string, while a rich `<si><r><rPr>…</rPr><t>…</t></r>…` decodes to a {@link RichTextValue} whose runs
// carry their per-run fonts — so rich text Excel pooled reads back formatted, not flattened to text.
// The run structure inside an `<si>` is identical to an inline string's `<is>`, so it is parsed the
// same way (see the inline-run accumulation in `parseWorksheet`).
export function parseSharedStrings(xml: string): SharedString[] {
  if (xml === '') return [];
  const strings: SharedString[] = [];
  // Per-`<si>` accumulation: `plain` gathers a bare `<t>`; `runs` gathers `<r>` runs. An `<si>` is
  // rich the moment it holds one `<r>`, at which point its runs — not `plain` — become the entry.
  let plain = '';
  const runs = new RunAccumulator();
  let isRich = false;
  let capture = false;
  let text = '';
  parseXml(xml, {
    onOpen(name, attrs) {
      const local = localName(name);
      switch (local) {
        case 'si':
          plain = '';
          runs.reset();
          isRich = false;
          break;
        case 'r':
          isRich = true;
          runs.beginRun();
          break;
        case 'rPr':
          runs.beginProperties();
          break;
        case 't':
          capture = true;
          text = '';
          break;
        default:
          runs.applyProperty(local, attrs);
          break;
      }
    },
    onText(chunk) {
      if (capture) text += chunk;
    },
    onClose(name) {
      const local = localName(name);
      switch (local) {
        case 't':
          // A `<t>` inside a run is that run's text; a bare `<t>` directly in the `<si>` is plain.
          if (!runs.appendText(text)) plain += text;
          capture = false;
          break;
        case 'r':
          runs.endRun();
          break;
        case 'si':
          strings.push(isRich ? {richText: runs.runs} : plain);
          break;
      }
    },
  });
  return strings;
}
