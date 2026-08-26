import assert from 'node:assert/strict';
import {test} from 'node:test';

import {AuthoringError} from '../../errors.ts';
import {
  applyThemeOverrides,
  DEFAULT_THEME_XML,
  parseThemeColorScheme,
  parseThemeFontScheme,
} from './theme-xml.ts';

test('parseThemeColorScheme reads srgbClr and sysClr slots', () => {
  const xml =
    '<a:theme><a:themeElements><a:clrScheme name="X">' +
    '<a:dk1><a:sysClr val="windowText" lastClr="1A1A1A"/></a:dk1>' +
    '<a:lt1><a:sysClr val="window" lastClr="FAFAFA"/></a:lt1>' +
    '<a:accent1><a:srgbClr val="BB2649"/></a:accent1>' +
    '</a:clrScheme>' +
    '<a:fmtScheme><a:gs><a:srgbClr val="DEADBE"/></a:gs></a:fmtScheme>' +
    '</a:themeElements></a:theme>';
  const scheme = parseThemeColorScheme(xml);
  // A sysClr's `val` names an operating-system colour; only its `lastClr` is a usable value.
  assert.equal(scheme.dk1, '1A1A1A');
  assert.equal(scheme.lt1, 'FAFAFA');
  assert.equal(scheme.accent1, 'BB2649');
  // Slots the scheme does not declare stay absent, and nothing outside <clrScheme> is picked up.
  assert.equal(scheme.accent2, undefined);
  assert.equal(Object.values(scheme).includes('DEADBE'), false);
});

test('parseThemeColorScheme drops a slot it cannot decode rather than guessing', () => {
  const xml =
    '<a:clrScheme><a:accent1><a:hslClr hue="0" sat="0" lum="0"/></a:accent1>' +
    '<a:accent2><a:srgbClr val="zzzzzz"/></a:accent2></a:clrScheme>';
  assert.deepEqual(parseThemeColorScheme(xml), {});
});

test('a theme part with no colour scheme yields nothing', () => {
  assert.deepEqual(parseThemeColorScheme('<a:theme/>'), {});
});

test('parseThemeFontScheme reads the major and minor latin faces', () => {
  const scheme = parseThemeFontScheme(DEFAULT_THEME_XML);
  assert.equal(scheme.major, 'Calibri Light');
  assert.equal(scheme.minor, 'Calibri');
  assert.deepEqual(parseThemeFontScheme('<a:theme/>'), {}, 'and nothing out of a part without one');
});

test('applyThemeOverrides replaces only the slots the caller named', () => {
  const xml = applyThemeOverrides(DEFAULT_THEME_XML, {colors: {accent1: '#BB2649'}});
  assert.ok(xml.includes('<a:accent1><a:srgbClr val="BB2649"/></a:accent1>'));
  assert.ok(xml.includes('<a:accent2><a:srgbClr val="ED7D31"/></a:accent2>'), 'the rest stand');
  // dk1/lt1 are sysClr in the source; re-serialising them as srgbClr would pin them to one machine.
  assert.ok(xml.includes('<a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>'));
  assert.ok(xml.includes('<a:fmtScheme'), 'and the format scheme is untouched');
});

test('applyThemeOverrides refuses a colour that is not a hex triplet', () => {
  assert.throws(
    () => applyThemeOverrides(DEFAULT_THEME_XML, {colors: {accent1: 'rebeccapurple'}}),
    AuthoringError,
  );
});

test('an authored typeface is escaped into the attribute, every character of it', () => {
  // The escape this replaced handled `& < > "` and stopped there, so a face carrying an apostrophe
  // or a newline went into the attribute raw.
  const face = 'A&B <C> "D" \'E\'\nF';
  const xml = applyThemeOverrides(DEFAULT_THEME_XML, {fonts: {minor: face}});
  const typeface = /<a:minorFont>[\s\S]*?<a:latin typeface="([^"]*)"/.exec(xml)?.[1];
  assert.equal(typeface, 'A&amp;B &lt;C&gt; &quot;D&quot; &apos;E&apos;&#10;F');
});

test('an authored typeface XML cannot represent is refused rather than written raw', () => {
  assert.throws(() => applyThemeOverrides(DEFAULT_THEME_XML, {fonts: {major: 'Bad\u0001Face'}}), {
    name: 'AuthoringError',
  });
});
