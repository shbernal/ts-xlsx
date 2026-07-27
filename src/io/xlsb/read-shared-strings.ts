// The BIFF12 shared-string table: `xl/sharedStrings.bin` in, the pool a `BrtCellIsst` cell indexes.
//
// The binary table is a flat run of `BrtSSTItem` records between a Begin/End pair, each holding one
// `RichStr`. Only the text is taken: a pooled string's per-run formatting is not modelled in this cut,
// so the runs are left unread rather than half-decoded (see `primitives.ts`). Reading stops at the
// closing record, so trailing future-record blocks are never framed.

import {RecordReader} from './primitives.ts';
import {readRecords} from './record-stream.ts';
import {BRT} from './record-types.ts';

/** Parse `xl/sharedStrings.bin` into the pool, in index order. An absent part is an empty pool. */
export function parseSharedStrings(part: Uint8Array | undefined): string[] {
  if (part === undefined) return [];
  const strings: string[] = [];
  for (const record of readRecords(part)) {
    if (record.type === BRT.SSTItem) strings.push(new RecordReader(record.data).richString());
    else if (record.type === BRT.EndSst) break;
  }
  return strings;
}
