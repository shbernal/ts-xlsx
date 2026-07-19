// Type-level tests over the public barrel (`src/index.ts`). Each imported symbol
// that no longer exists breaks compilation here (an export-removal guard), and each
// `Expect<Equal<...>>` locks a contract the runtime tests cannot see — synchronous
// I/O, the optionality of an address's col/row, the membership of CellValue.

import type {
  CellAddress,
  CellValue,
  decodeAddress,
  readXlsx,
  Workbook,
  Worksheet,
  writeXlsx,
} from '../index.ts';
import type {Equal, Expect, Extends} from './expect.ts';

// decodeAddress yields the canonical CellAddress, whose col/row stay optional so a
// column-only (`$A`) or row-only (`$1`) reference is representable.
export type AddressContracts = [
  Expect<Equal<ReturnType<typeof decodeAddress>, CellAddress>>,
  Expect<Equal<CellAddress['address'], string>>,
  Expect<Equal<CellAddress['col'], number | undefined>>,
  Expect<Equal<CellAddress['row'], number | undefined>>,
];

// CellValue admits the primitive/Date leaves and never `undefined`: an absent cell
// is `null`, not `undefined`, and the writer relies on that distinction.
export type ValueContracts = [
  Expect<Extends<null, CellValue>>,
  Expect<Extends<number, CellValue>>,
  Expect<Extends<string, CellValue>>,
  Expect<Extends<boolean, CellValue>>,
  Expect<Extends<Date, CellValue>>,
  Expect<Equal<Extends<undefined, CellValue>, false>>,
];

// The buffered I/O surface is synchronous: writeXlsx returns bytes and readXlsx a
// Workbook directly — never a Promise. getWorksheet is partial (a miss is undefined).
export type IoContracts = [
  Expect<Equal<ReturnType<typeof writeXlsx>, Uint8Array>>,
  Expect<Equal<ReturnType<typeof readXlsx>, Workbook>>,
  Expect<Equal<ReturnType<Workbook['getWorksheet']>, Worksheet | undefined>>,
];
