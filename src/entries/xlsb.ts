// Entry barrel for `@shbernal/ts-xlsx/xlsb` — the BIFF12 reader.
//
// `readXlsx` auto-detects and reads a binary `.xlsb` too, so most callers never name this; it is
// public for the caller that holds bytes it already knows to be `.xlsb` and wants to say so. That
// dispatch also means this entry is not the way to *avoid* the BIFF12 codec — importing
// `/xlsx` loads it either way (see the per-entry budgets in `scripts/size-budget.ts`).

export {readXlsb} from '../io/xlsb/read.ts';
