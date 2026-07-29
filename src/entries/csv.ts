// Entry barrel for `@shbernal/ts-xlsx/csv` — delimited text in and out.
//
// The two directions are asymmetric in what they cost: writing needs only a worksheet and the
// value vocabulary, while reading builds a `Workbook` and therefore pulls the model in whole.

export {type CsvReadOptions, readCsv} from '../io/csv/read.ts';
export {type CsvWriteOptions, writeCsv, writeCsvText} from '../io/csv/write.ts';
