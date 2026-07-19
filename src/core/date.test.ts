import assert from 'node:assert/strict';
import {test} from 'node:test';

import {DEFAULT_DATE_NUMFMT, dateToSerial, isDateFormat, serialToDate} from './date.ts';

const iso = (serial: number): string => serialToDate(serial).toISOString();

test('serial 1 reads as 1900-01-01, not 1899-12-31', () => {
  assert.equal(iso(1), '1900-01-01T00:00:00.000Z');
});

test('consecutive serials map to consecutive days', () => {
  assert.equal(iso(2), '1900-01-02T00:00:00.000Z');
  assert.equal(iso(3), '1900-01-03T00:00:00.000Z');
});

test('serial 59 is the real 1900-02-28 just below the phantom leap day', () => {
  assert.equal(iso(59), '1900-02-28T00:00:00.000Z');
});

test('serial 61 is 1900-03-01, the day after the phantom 1900-02-29', () => {
  assert.equal(iso(61), '1900-03-01T00:00:00.000Z');
});

test('a modern date round-trips through serial and back exactly', () => {
  const date = new Date('2020-03-04T00:00:00.000Z');
  assert.equal(serialToDate(dateToSerial(date)).toISOString(), date.toISOString());
});

test('a fractional serial carries the time of day', () => {
  const noon = new Date('2020-03-04T12:00:00.000Z');
  const serial = dateToSerial(noon);
  assert.equal(serial % 1, 0.5, 'noon is half a day past midnight');
  assert.equal(serialToDate(serial).toISOString(), noon.toISOString());
});

test('dateToSerial reproduces the leap-year boundary in reverse', () => {
  assert.equal(dateToSerial(new Date('1900-01-01T00:00:00.000Z')), 1);
  assert.equal(dateToSerial(new Date('1900-02-28T00:00:00.000Z')), 59);
  assert.equal(dateToSerial(new Date('1900-03-01T00:00:00.000Z')), 61);
});

test('isDateFormat recognises date and time codes', () => {
  for (const code of [
    'yyyy-mm-dd',
    'DD/MM/YYYY',
    'mm-dd-yy',
    'h:mm:ss',
    '[$-409]mmmm d, yyyy',
    DEFAULT_DATE_NUMFMT,
  ]) {
    assert.equal(isDateFormat(code), true, `${code} should be a date format`);
  }
});

test('isDateFormat rejects number, currency, percent, and text codes', () => {
  for (const code of [
    'General',
    '0.00',
    '#,##0',
    '0.00%',
    '_("$"* #,##0.00_)',
    '@',
    '#,##0" days"',
  ]) {
    assert.equal(isDateFormat(code), false, `${code} should not be a date format`);
  }
});
