import {strict as assert} from 'node:assert';
import {test} from 'node:test';

import {strToU8, zipSync} from 'fflate';

import {readXlsx} from '../io/xlsx/read.ts';
import {CustomUiParseError} from './errors.ts';
import {isCustomUiRelType, parseCustomUi} from './ribbon.ts';

// The two dialects' minimal ribbon shape, matching the real fixtures the round-trip test exercises.
const CUSTOM_UI_2007 =
  '<customUI xmlns="http://schemas.microsoft.com/office/2006/01/customui">' +
  '<ribbon><tabs><tab id="t07" label="Legacy"><group id="g07" label="G">' +
  '<button id="b07" label="Run" onAction="LegacyMacro"/></group></tab></tabs></ribbon></customUI>';

const CUSTOM_UI_2009 =
  '<customUI xmlns="http://schemas.microsoft.com/office/2009/07/customui">' +
  '<ribbon><tabs><tab id="t14" label="Macros"><group id="g14" label="G">' +
  '<button id="b14" label="Run" onAction="MyMacro"/></group></tab></tabs></ribbon></customUI>';

test('parses the 2007 ribbon tree, lifting id/label/onAction and the dialect', () => {
  const doc = parseCustomUi(CUSTOM_UI_2007);

  assert.equal(doc.dialect, '2007');
  assert.ok(doc.ribbon);
  assert.equal(doc.ribbon.startFromScratch, false);
  assert.equal(doc.ribbon.tabs.length, 1);

  const tab = doc.ribbon.tabs[0];
  assert.equal(tab?.id, 't07');
  assert.equal(tab?.label, 'Legacy');
  assert.equal(tab?.groups.length, 1);

  const group = tab?.groups[0];
  assert.equal(group?.id, 'g07');

  const button = group?.controls[0];
  assert.equal(button?.kind, 'button');
  assert.equal(button?.id, 'b07');
  assert.equal(button?.label, 'Run');
  assert.equal(button?.onAction, 'LegacyMacro');
  assert.equal(button?.children, undefined);
});

test('derives the 2010 dialect from the customUI14 namespace, not the relationship type', () => {
  const doc = parseCustomUi(CUSTOM_UI_2009);
  assert.equal(doc.dialect, '2010');
  assert.equal(doc.ribbon?.tabs[0]?.groups[0]?.controls[0]?.onAction, 'MyMacro');
});

test('accepts raw UTF-8 bytes as well as a decoded string', () => {
  const doc = parseCustomUi(strToU8(CUSTOM_UI_2007));
  assert.equal(doc.dialect, '2007');
  assert.equal(doc.ribbon?.tabs[0]?.id, 't07');
});

test('resolves the dialect through a prefixed customUI element', () => {
  const doc = parseCustomUi(
    '<mso:customUI xmlns:mso="http://schemas.microsoft.com/office/2009/07/customui">' +
      '<mso:ribbon><mso:tabs><mso:tab id="t"><mso:group id="g"/></mso:tab></mso:tabs>' +
      '</mso:ribbon></mso:customUI>',
  );
  assert.equal(doc.dialect, '2010');
  assert.equal(doc.ribbon?.tabs[0]?.groups[0]?.id, 'g');
});

test('models a container control as children, and preserves every attribute verbatim', () => {
  const xml =
    '<customUI xmlns="http://schemas.microsoft.com/office/2006/01/customui">' +
    '<ribbon startFromScratch="true"><tabs><tab idMso="TabHome"><group id="g">' +
    '<menu id="m" label="M" getEnabled="IsReady">' +
    '<button id="b1" onAction="A"/><menuSeparator id="s"/><button id="b2" onAction="B"/>' +
    '</menu>' +
    '<checkBox id="c" idQ="ns:c" label="On"/>' +
    '<separator id="sep"/>' +
    '<wackyControl id="w" foo="bar"/>' +
    '</group></tab></tabs></ribbon></customUI>';
  const doc = parseCustomUi(xml);

  assert.equal(doc.ribbon?.startFromScratch, true);
  const tab = doc.ribbon?.tabs[0];
  assert.equal(tab?.idMso, 'TabHome');

  const controls = tab?.groups[0]?.controls ?? [];
  const menu = controls[0];
  assert.equal(menu?.kind, 'menu');
  assert.equal(menu?.attributes.getEnabled, 'IsReady'); // a dynamic callback not lifted, kept in the raw map
  assert.equal(menu?.children?.length, 3);
  assert.equal(menu?.children?.[0]?.kind, 'button');
  assert.equal(menu?.children?.[0]?.onAction, 'A');
  assert.equal(menu?.children?.[1]?.kind, 'menuSeparator');

  const checkBox = controls[1];
  assert.equal(checkBox?.kind, 'checkBox');
  assert.equal(checkBox?.idQ, 'ns:c');

  assert.equal(controls[2]?.kind, 'separator');

  const unknown = controls[3];
  assert.equal(unknown?.kind, 'unknown'); // an element outside the known set is surfaced, never dropped
  assert.equal(unknown?.id, 'w');
  assert.equal(unknown?.attributes.foo, 'bar');
});

test('a document that customises only backstage yields no ribbon', () => {
  const doc = parseCustomUi(
    '<customUI xmlns="http://schemas.microsoft.com/office/2009/07/customui">' +
      '<backstage><tab id="bt"/></backstage></customUI>',
  );
  assert.equal(doc.dialect, '2010');
  assert.equal(doc.ribbon, undefined);
});

test('throws CustomUiParseError on malformed, unbalanced, unrooted, or unknown-namespace XML', () => {
  assert.throws(
    () =>
      parseCustomUi(
        '<customUI xmlns="http://schemas.microsoft.com/office/2006/01/customui"><ribbon',
      ),
    CustomUiParseError,
  );
  assert.throws(
    () =>
      parseCustomUi(
        '<customUI xmlns="http://schemas.microsoft.com/office/2006/01/customui"></customUI></customUI>',
      ),
    CustomUiParseError,
  );
  assert.throws(() => parseCustomUi('<notCustomUI/>'), CustomUiParseError);
  assert.throws(
    () => parseCustomUi('<customUI xmlns="http://example.com/other"><ribbon/></customUI>'),
    CustomUiParseError,
  );
});

test('caps nesting depth so a hostile part cannot overflow the walk', () => {
  const deep =
    '<customUI xmlns="http://schemas.microsoft.com/office/2006/01/customui">' +
    '<box>'.repeat(300) +
    '</box>'.repeat(300) +
    '</customUI>';
  assert.throws(() => parseCustomUi(deep), CustomUiParseError);
});

test('isCustomUiRelType matches both real ribbon relationship types, not the namespace', () => {
  assert.ok(
    isCustomUiRelType('http://schemas.microsoft.com/office/2006/relationships/ui/extensibility'),
  );
  assert.ok(
    isCustomUiRelType('http://schemas.microsoft.com/office/2007/relationships/ui/extensibility'),
  );
  assert.ok(!isCustomUiRelType('http://schemas.microsoft.com/office/2009/07/customui'));
  assert.ok(
    !isCustomUiRelType(
      'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument',
    ),
  );
});

// --- Workbook.customUI integration: the accessor sources bytes from the package-root preserved refs. ---

function packageWith(rootRelExtra: string, extraParts: Record<string, string>): Uint8Array {
  const parts: Record<string, string> = {
    '_rels/.rels':
      '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>' +
      rootRelExtra +
      '</Relationships>',
    '[Content_Types].xml':
      '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
      '<Default Extension="xml" ContentType="application/xml"/>' +
      '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
      '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' +
      '</Types>',
    'xl/workbook.xml':
      '<?xml version="1.0"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" ' +
      'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
      '<sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets></workbook>',
    'xl/_rels/workbook.xml.rels':
      '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>' +
      '</Relationships>',
    'xl/worksheets/sheet1.xml':
      '<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>',
    ...extraParts,
  };
  const files: Record<string, Uint8Array> = {};
  for (const [name, xml] of Object.entries(parts)) files[name] = strToU8(xml);
  return zipSync(files);
}

const RIBBON_REL = (id: string, type: string, target: string): string =>
  `<Relationship Id="${id}" Type="${type}" Target="${target}"/>`;

test('Workbook.customUI parses both ribbon parts, tagged by dialect, and memoises', () => {
  const wb = readXlsx(
    packageWith(
      RIBBON_REL(
        'rId4',
        'http://schemas.microsoft.com/office/2006/relationships/ui/extensibility',
        'customUI/customUI.xml',
      ) +
        RIBBON_REL(
          'rId5',
          'http://schemas.microsoft.com/office/2007/relationships/ui/extensibility',
          'customUI/customUI14.xml',
        ),
      {'customUI/customUI.xml': CUSTOM_UI_2007, 'customUI/customUI14.xml': CUSTOM_UI_2009},
    ),
  );

  const docs = wb.customUI;
  assert.equal(docs.length, 2);
  assert.deepEqual(
    docs.map((d) => d.dialect),
    ['2007', '2010'],
  );
  assert.equal(docs[0]?.ribbon?.tabs[0]?.groups[0]?.controls[0]?.onAction, 'LegacyMacro');
  assert.equal(docs[1]?.ribbon?.tabs[0]?.groups[0]?.controls[0]?.onAction, 'MyMacro');

  assert.equal(wb.customUI, docs, 'memoised: the same array is returned on re-access');
});

test('Workbook.customUI is empty for a workbook with no ribbon parts', () => {
  const wb = readXlsx(packageWith('', {}));
  assert.deepEqual(wb.customUI, []);
});

test('Workbook.customUI throws CustomUiParseError when a ribbon part is malformed', () => {
  const wb = readXlsx(
    packageWith(
      RIBBON_REL(
        'rId4',
        'http://schemas.microsoft.com/office/2006/relationships/ui/extensibility',
        'customUI/customUI.xml',
      ),
      {'customUI/customUI.xml': '<customUI xmlns="http://example.com/other"/>'},
    ),
  );
  assert.throws(() => wb.customUI, CustomUiParseError);
});
