// Public entry point for the ts-xlsx rewrite (Phase 3).
//
// This barrel is intentionally thin and provisional: the definitive public API is
// designed module by module as the rewrite lands, corpus-first. Only surfaces that
// are implemented, strict-typed, and corpus-backed are re-exported here.

export {
  type CellAddress,
  type RangeAddress,
  MAX_COLUMN,
  columnToNumber,
  numberToColumn,
  decodeAddress,
  decodeRange,
  encodeAddress,
} from './core/address.ts';
