import assert from 'node:assert/strict';
import test from 'node:test';

import {
  applyTint,
  DEFAULT_INDEXED_COLORS,
  resolveColor,
  SYSTEM_INDEXED_COLORS,
} from './color-resolution.ts';
import {DEFAULT_THEME_COLOR_SCHEME, parseThemeColorScheme, THEME_COLOR_SLOTS} from './theme.ts';

test('an explicit argb resolves to itself, keeping its declared alpha', () => {
  assert.equal(resolveColor({argb: 'FF336699'}), 'FF336699');
  assert.equal(resolveColor({argb: '80336699'}), '80336699');
  // The two developer conveniences the writer accepts are accepted here too.
  assert.equal(resolveColor({argb: '336699'}), 'FF336699');
  assert.equal(resolveColor({argb: '#336699'}), 'FF336699');
});

test('a malformed argb resolves to nothing rather than a half-parsed value', () => {
  assert.equal(resolveColor({argb: 'nope'}), undefined);
  assert.equal(resolveColor({argb: 'FF33669'}), undefined);
});

test('theme index 0 is lt1 and index 1 is dk1, not the clrScheme child order', () => {
  // The whole point of THEME_COLOR_SLOTS: the on-disk index swaps each dark/light pair relative to
  // the order the slots appear in the theme part. Getting this backwards inverts text on background.
  assert.equal(THEME_COLOR_SLOTS[0], 'lt1');
  assert.equal(THEME_COLOR_SLOTS[1], 'dk1');
  assert.equal(THEME_COLOR_SLOTS[2], 'lt2');
  assert.equal(THEME_COLOR_SLOTS[3], 'dk2');
  assert.equal(resolveColor({theme: 0}), 'FFFFFFFF');
  assert.equal(resolveColor({theme: 1}), 'FF000000');
});

test('the stylesheet default font colour resolves to black', () => {
  // Every workbook the writer emits carries `<color theme="1"/>` on font 0. If the index mapping were
  // the clrScheme child order, that default body text would resolve to white.
  assert.equal(resolveColor({theme: 1}), 'FF000000');
});

test('accent and hyperlink slots resolve through the workbook theme', () => {
  const theme = {...DEFAULT_THEME_COLOR_SCHEME, accent1: 'BB2649', folHlink: '7C4A8B'};
  assert.equal(resolveColor({theme: 4}, {theme}), 'FFBB2649');
  assert.equal(resolveColor({theme: 11}, {theme}), 'FF7C4A8B');
});

test('a theme index outside the scheme resolves to nothing', () => {
  assert.equal(resolveColor({theme: 12}), undefined);
  assert.equal(resolveColor({theme: -1}), undefined);
  assert.equal(resolveColor({theme: 4}, {theme: {}}), undefined);
});

test('an indexed colour resolves through the built-in palette, fully opaque', () => {
  // The spec tabulates entry 10 as 00FF0000; the leading 00 is a legacy artefact, not transparency.
  assert.equal(DEFAULT_INDEXED_COLORS[10], '00FF0000');
  assert.equal(resolveColor({indexed: 10}), 'FFFF0000');
  assert.equal(resolveColor({indexed: 22}), 'FFC0C0C0');
});

test('indices 0-7 duplicate 8-15, as the spec preserves them', () => {
  for (let i = 0; i < 8; i++) {
    assert.equal(resolveColor({indexed: i}), resolveColor({indexed: i + 8}));
  }
});

test('the system foreground and background indices stay unresolved', () => {
  assert.deepEqual([...SYSTEM_INDEXED_COLORS].sort(), [64, 65]);
  // indexed="64" is the placeholder on every solid fill Excel writes; inventing black for it would
  // paint every such fill.
  assert.equal(resolveColor({indexed: 64}), undefined);
  assert.equal(resolveColor({indexed: 65}), undefined);
  assert.equal(resolveColor({indexed: 200}), undefined);
});

test('a workbook custom palette replaces the built-in one wholesale', () => {
  const indexed = ['00112233', '00445566', '00778899'];
  assert.equal(resolveColor({indexed: 2}, {indexed}), 'FF778899');
  // Past the end of a custom palette there is nothing — it replaces the built-in table, it does not
  // extend it, so falling through would resolve to a colour the workbook explicitly overrode away.
  assert.equal(resolveColor({indexed: 10}, {indexed}), undefined);
});

test('tint -1 darkens to black and +1 lightens to white', () => {
  assert.equal(applyTint('FF156082', -1), 'FF000000');
  assert.equal(applyTint('FF156082', 1), 'FFFFFFFF');
});

test('tint 0 and a non-finite tint leave the colour alone', () => {
  assert.equal(resolveColor({argb: 'FF156082', tint: 0}), 'FF156082');
  assert.equal(resolveColor({argb: 'FF156082', tint: Number.NaN}), 'FF156082');
});

test('tint tracks what Excel renders, within the measured tolerance', () => {
  // Sampled from test/corpus/fixtures/excel-oracle/theme-color-tint-luminance.json — accent1 156082
  // and accent2 E97132 as Excel itself painted them. The bound is 2/255 per channel; see applyTint.
  const samples: [string, number, string][] = [
    ['FF156082', -0.5, '0B3040'],
    ['FF156082', -0.25, '104861'],
    ['FF156082', 0.25, '229ACE'],
    ['FF156082', 0.5, '64BEE6'],
    ['FFE97132', -0.5, '7E350E'],
    ['FFE97132', 0.25, 'EE9564'],
    ['FFE8E8E8', 0.5, 'F3F3F3'],
  ];
  for (const [base, tint, excel] of samples) {
    const got = applyTint(base, tint);
    for (const at of [2, 4, 6]) {
      const delta = Math.abs(
        Number.parseInt(got.slice(at, at + 2), 16) - Number.parseInt(excel.slice(at - 2, at), 16),
      );
      assert.ok(delta <= 2, `${base} tint ${tint}: got ${got}, Excel ${excel} (Δ${delta})`);
    }
  }
});

test('tint applies on top of a resolved theme colour', () => {
  const plain = resolveColor({theme: 4});
  const tinted = resolveColor({theme: 4, tint: 0.6});
  assert.notEqual(tinted, plain);
  assert.equal(tinted, applyTint(plain as string, 0.6));
});

test('a colour that states nothing resolves to nothing', () => {
  assert.equal(resolveColor({}), undefined);
});

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
  // Slots the scheme does not declare stay absent — and nothing outside <clrScheme> is picked up.
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
