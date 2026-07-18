import {strict as assert} from 'node:assert';
import {test} from 'node:test';

import {strFromU8, strToU8, unzipSync, zipSync} from 'fflate';

import {readXlsx} from './read.ts';
import {writeXlsx} from './write.ts';

// A minimal OPC package skeleton shared by the scenarios below: content types, the root and workbook
// relationships, and a single-sheet workbook. Each test overlays the worksheet, its rels, and the
// unmodeled parts (a shape drawing, a header/footer VML + image) that exercise the passthrough.
function packageParts(overlay: Record<string, string | Uint8Array>): Record<string, Uint8Array> {
  const base: Record<string, string | Uint8Array> = {
    '_rels/.rels':
      '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>' +
      '</Relationships>',
    'xl/workbook.xml':
      '<?xml version="1.0"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" ' +
      'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
      '<sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets></workbook>',
    'xl/_rels/workbook.xml.rels':
      '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>' +
      '</Relationships>',
  };
  const files: Record<string, Uint8Array> = {};
  for (const [name, data] of Object.entries({...base, ...overlay})) {
    files[name] = typeof data === 'string' ? strToU8(data) : data;
  }
  return files;
}

const contentTypes = (extra: string): string =>
  '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
  '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
  '<Default Extension="xml" ContentType="application/xml"/>' +
  extra +
  '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
  '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' +
  '</Types>';

const worksheet = (tail: string): string =>
  '<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" ' +
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
  '<sheetData/>' +
  tail +
  '</worksheet>';

const SHAPE_DRAWING =
  '<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" ' +
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><xdr:twoCellAnchor>' +
  '<xdr:from><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>' +
  '<xdr:to><xdr:col>3</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>4</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>' +
  '<xdr:sp macro="" textlink=""><xdr:nvSpPr><xdr:cNvPr id="2" name="Rectangle 1"/><xdr:cNvSpPr/></xdr:nvSpPr>' +
  '<xdr:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr>' +
  '<xdr:txBody><a:bodyPr/><a:p/></xdr:txBody></xdr:sp><xdr:clientData/></xdr:twoCellAnchor></xdr:wsDr>';

// A single drawing part holding BOTH a modeled picture (an <xdr:pic> with a blip embed) AND a chart
// (an <xdr:graphicFrame> naming a chart part by r:id) — the shape Excel produces for a sheet that
// carries an image and a chart together. Modeling only the picture and re-serialising the drawing
// from it would drop the chart; the whole part must ride through preservation instead.
const MIXED_DRAWING =
  '<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" ' +
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
  '<xdr:oneCellAnchor>' +
  '<xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>' +
  '<xdr:ext cx="100" cy="100"/><xdr:pic><xdr:nvPicPr><xdr:cNvPr id="1" name="p"/><xdr:cNvPicPr/></xdr:nvPicPr>' +
  '<xdr:blipFill><a:blip r:embed="rId1"/></xdr:blipFill><xdr:spPr/></xdr:pic><xdr:clientData/>' +
  '</xdr:oneCellAnchor>' +
  '<xdr:twoCellAnchor>' +
  '<xdr:from><xdr:col>4</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>4</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>' +
  '<xdr:to><xdr:col>10</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>20</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>' +
  '<xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="2" name="Chart 1"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>' +
  '<xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm>' +
  '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">' +
  '<c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" r:id="rId2"/>' +
  '</a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/>' +
  '</xdr:twoCellAnchor></xdr:wsDr>';

const CHART =
  '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" ' +
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><c:chart><c:plotArea><c:barChart/></c:plotArea></c:chart></c:chartSpace>';

const HF_VML =
  '<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office">' +
  '<v:shape id="RH" type="#_x0000_t75"><v:imagedata o:relid="rId1" o:title="pic"/></v:shape></xml>';

const rels = (entries: string): string =>
  '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
  entries +
  '</Relationships>';

const relationship = (id: string, type: string, target: string): string =>
  `<Relationship Id="${id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/${type}" Target="${target}"/>`;

// A slicer / slicer-cache relationship carries a Microsoft-namespaced Type; only its suffix
// (`/slicer`, `/slicerCache`) is what marks it for preservation.
const msRelationship = (id: string, type: string, target: string): string =>
  `<Relationship Id="${id}" Type="http://schemas.microsoft.com/office/2007/relationships/${type}" Target="${target}"/>`;

function partNames(bytes: Uint8Array): string[] {
  return Object.keys(unzipSync(bytes));
}

function partText(bytes: Uint8Array, rx: RegExp): string {
  const files = unzipSync(bytes);
  const name = Object.keys(files).find(n => rx.test(n));
  return name ? strFromU8(files[name] as Uint8Array) : '';
}

test('a worksheet drawing holding only a vector shape survives read→write', () => {
  const src = zipSync(
    packageParts({
      '[Content_Types].xml': contentTypes(
        '<Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>'
      ),
      'xl/worksheets/sheet1.xml': worksheet('<drawing r:id="rId1"/>'),
      'xl/worksheets/_rels/sheet1.xml.rels': rels(relationship('rId1', 'drawing', '../drawings/drawing1.xml')),
      'xl/drawings/drawing1.xml': SHAPE_DRAWING,
    })
  );

  const out = writeXlsx(readXlsx(src));

  assert.ok(partNames(out).some(n => /xl\/drawings\/drawing\d+\.xml$/.test(n)), 'the drawing part survives');
  assert.match(partText(out, /worksheets\/sheet1\.xml$/), /<drawing r:id="[^"]+"\/>/, 'the worksheet still references it');
  assert.match(partText(out, /drawings\/drawing1\.xml$/), /<xdr:sp\b/, 'the vector shape survives inside');
  // The rewritten package must re-read without error, and re-writing it must keep preserving the shape.
  assert.match(partText(writeXlsx(readXlsx(out)), /drawings\/drawing1\.xml$/), /<xdr:sp\b/, 'idempotent across a second round-trip');
});

test('a drawing holding both a picture and a chart preserves the chart across read→write', () => {
  const png = Uint8Array.from([0x89, 0x50, 0x4e, 0x47, 1, 2, 3]);
  const src = zipSync(
    packageParts({
      '[Content_Types].xml': contentTypes(
        '<Default Extension="png" ContentType="image/png"/>' +
          '<Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>' +
          '<Override PartName="/xl/charts/chart1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>'
      ),
      'xl/worksheets/sheet1.xml': worksheet('<drawing r:id="rId1"/>'),
      'xl/worksheets/_rels/sheet1.xml.rels': rels(relationship('rId1', 'drawing', '../drawings/drawing1.xml')),
      'xl/drawings/drawing1.xml': MIXED_DRAWING,
      'xl/drawings/_rels/drawing1.xml.rels': rels(
        relationship('rId1', 'image', '../media/image1.png') + relationship('rId2', 'chart', '../charts/chart1.xml')
      ),
      'xl/media/image1.png': png,
      'xl/charts/chart1.xml': CHART,
    })
  );

  const out = writeXlsx(readXlsx(src));
  const names = partNames(out);

  assert.match(partText(out, /worksheets\/sheet1\.xml$/), /<drawing r:id="[^"]+"\/>/, 'the worksheet still references the drawing');
  const drawing = partText(out, /drawings\/drawing\d+\.xml$/);
  assert.match(drawing, /<xdr:graphicFrame\b/, 'the chart anchor survives inside the drawing');
  assert.match(drawing, /<xdr:pic\b/, 'the picture anchor survives alongside it, not dropped');
  assert.ok(names.some(n => /xl\/charts\/chart\d+\.xml$/.test(n)), 'the chart part survives');
  assert.ok(names.some(n => /xl\/media\/.+\.png$/.test(n)), 'the picture media survives');
  // The whole mixed drawing rides through preservation, so a second round-trip keeps the chart too.
  assert.ok(
    partNames(writeXlsx(readXlsx(out))).some(n => /xl\/charts\/chart\d+\.xml$/.test(n)),
    'idempotent across a second round-trip'
  );
});

test('a header/footer image (legacyDrawingHF VML + media) and its &G token survive read→write', () => {
  const image = Uint8Array.from([0xff, 0xd8, 0xff, 0xe0, 1, 2, 3, 4]);
  const src = zipSync(
    packageParts({
      '[Content_Types].xml': contentTypes(
        '<Default Extension="vml" ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing"/>' +
          '<Default Extension="jpeg" ContentType="image/jpeg"/>'
      ),
      'xl/worksheets/sheet1.xml': worksheet(
        '<headerFooter><oddHeader>&amp;R&amp;G</oddHeader></headerFooter><legacyDrawingHF r:id="rId1"/>'
      ),
      'xl/worksheets/_rels/sheet1.xml.rels': rels(relationship('rId1', 'vmlDrawing', '../drawings/vmlDrawing1.vml')),
      'xl/drawings/vmlDrawing1.vml': HF_VML,
      'xl/drawings/_rels/vmlDrawing1.vml.rels': rels(relationship('rId1', 'image', '../media/image1.jpeg')),
      'xl/media/image1.jpeg': image,
    })
  );

  const out = writeXlsx(readXlsx(src));
  const ws = partText(out, /worksheets\/sheet1\.xml$/);

  assert.match(ws, /<legacyDrawingHF r:id="[^"]+"\/>/, 'the legacyDrawingHF reference survives');
  assert.match(ws, /&amp;R&amp;G/, 'the &G header/footer picture token survives');
  assert.ok(partNames(out).some(n => /vmlDrawing\d+\.vml$/.test(n)), 'the VML drawing survives');
  assert.ok(partNames(out).some(n => /xl\/media\//.test(n)), 'the image media survives');
  // The VML's image relationship must re-resolve to the media part's new path.
  assert.match(partText(out, /drawings\/_rels\/vmlDrawing\d+\.vml\.rels$/), /Target="\.\.\/media\/image\d+\.jpeg"/, 'the VML relinks its image');
});

test('a preserved header/footer VML is numbered clear of a modeled anchored image', () => {
  // A sheet that anchors a real image (modeled → drawing1.xml, image1.jpeg, media #1) AND carries a
  // preserved header/footer VML whose own image would also be image1.jpeg in the source. The writer
  // must renumber the preserved media past the modeled one rather than clobbering it.
  const modeledPng = Uint8Array.from([0x89, 0x50, 0x4e, 0x47, 9, 9]);
  const hfJpeg = Uint8Array.from([0xff, 0xd8, 0xff, 5, 6]);
  const src = zipSync(
    packageParts({
      '[Content_Types].xml': contentTypes(
        '<Default Extension="vml" ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing"/>' +
          '<Default Extension="png" ContentType="image/png"/><Default Extension="jpeg" ContentType="image/jpeg"/>' +
          '<Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>'
      ),
      'xl/worksheets/sheet1.xml': worksheet('<drawing r:id="rId1"/><legacyDrawingHF r:id="rId2"/>'),
      'xl/worksheets/_rels/sheet1.xml.rels': rels(
        relationship('rId1', 'drawing', '../drawings/drawing1.xml') +
          relationship('rId2', 'vmlDrawing', '../drawings/vmlDrawing1.vml')
      ),
      'xl/drawings/drawing1.xml':
        '<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" ' +
        'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><xdr:oneCellAnchor>' +
        '<xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>' +
        '<xdr:ext cx="100" cy="100"/><xdr:pic><xdr:nvPicPr><xdr:cNvPr id="1" name="p"/><xdr:cNvPicPr/></xdr:nvPicPr>' +
        '<xdr:blipFill><a:blip xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="rId1"/></xdr:blipFill>' +
        '<xdr:spPr/></xdr:pic><xdr:clientData/></xdr:oneCellAnchor></xdr:wsDr>',
      'xl/drawings/_rels/drawing1.xml.rels': rels(relationship('rId1', 'image', '../media/image1.png')),
      'xl/media/image1.png': modeledPng,
      'xl/drawings/vmlDrawing1.vml': HF_VML,
      'xl/drawings/_rels/vmlDrawing1.vml.rels': rels(relationship('rId1', 'image', '../media/image1.jpeg')),
      'xl/media/image1.jpeg': hfJpeg,
    })
  );

  const out = writeXlsx(readXlsx(src));
  const names = partNames(out);
  const media = names.filter(n => /xl\/media\//.test(n));
  assert.equal(new Set(media).size, media.length, 'no two media parts share a path');
  assert.ok(media.some(n => /\.png$/.test(n)) && media.some(n => /\.jpeg$/.test(n)), 'both the modeled png and the preserved jpeg survive');
  assert.ok(names.some(n => /vmlDrawing\d+\.vml$/.test(n)), 'the header/footer VML survives');
  assert.match(partText(out, /worksheets\/sheet1\.xml$/), /<legacyDrawingHF r:id="[^"]+"\/>/, 'the legacyDrawingHF reference survives');
});

// A workbook overlay that adds a `<pivotCaches>` registration and a workbook relationship reaching a
// pivot cache — the wiring a real pivot-bearing workbook carries, which the base package omits.
const workbookWithPivotCache =
  '<?xml version="1.0"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" ' +
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
  '<sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets>' +
  '<pivotCaches><pivotCache cacheId="42" r:id="rId2"/></pivotCaches></workbook>';

test('a pivot table and its pivot cache survive read→write, cacheId wiring intact', () => {
  const src = zipSync(
    packageParts({
      '[Content_Types].xml': contentTypes(
        '<Override PartName="/xl/pivotTables/pivotTable1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml"/>' +
          '<Override PartName="/xl/pivotCache/pivotCacheDefinition1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml"/>' +
          '<Override PartName="/xl/pivotCache/pivotCacheRecords1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheRecords+xml"/>'
      ),
      'xl/workbook.xml': workbookWithPivotCache,
      'xl/_rels/workbook.xml.rels':
        '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        relationship('rId1', 'worksheet', 'worksheets/sheet1.xml') +
        relationship('rId2', 'pivotCacheDefinition', 'pivotCache/pivotCacheDefinition1.xml') +
        '</Relationships>',
      // The pivot table is discovered through a sheet relationship — there is no worksheet child.
      'xl/worksheets/sheet1.xml': worksheet(''),
      'xl/worksheets/_rels/sheet1.xml.rels': rels(relationship('rId1', 'pivotTable', '../pivotTables/pivotTable1.xml')),
      'xl/pivotTables/pivotTable1.xml': '<pivotTableDefinition cacheId="42"/>',
      'xl/pivotTables/_rels/pivotTable1.xml.rels': rels(
        relationship('rId1', 'pivotCacheDefinition', '../pivotCache/pivotCacheDefinition1.xml')
      ),
      'xl/pivotCache/pivotCacheDefinition1.xml': '<pivotCacheDefinition/>',
      'xl/pivotCache/_rels/pivotCacheDefinition1.xml.rels': rels(
        relationship('rId1', 'pivotCacheRecords', 'pivotCacheRecords1.xml')
      ),
      'xl/pivotCache/pivotCacheRecords1.xml': '<pivotCacheRecords/>',
    })
  );

  const out = writeXlsx(readXlsx(src));
  const names = partNames(out);

  assert.ok(names.includes('xl/pivotTables/pivotTable1.xml'), 'the pivot table part survives');
  assert.equal(
    names.filter(n => /xl\/pivotCache\/.+\.xml$/.test(n)).length,
    2,
    'both the pivot cache definition and its records survive'
  );
  assert.match(partText(out, /worksheets\/_rels\/sheet1\.xml\.rels$/), /pivotTable/, 'the sheet still references the pivot table');

  // The <pivotCaches> registration is re-emitted with its cacheId, wired to the workbook relationship
  // that reaches the (surviving) cache definition — so a pivot table can resolve its cache on reopen.
  const wb = partText(out, /xl\/workbook\.xml$/);
  const cache = /<pivotCache cacheId="42" r:id="(rId\d+)"\/>/.exec(wb);
  assert.ok(cache, `workbook registers the pivot cache with its cacheId; got ${wb}`);
  const wbRels = partText(out, /_rels\/workbook\.xml\.rels$/);
  assert.match(
    wbRels,
    new RegExp(`Id="${cache?.[1]}"[^>]*Target="pivotCache/pivotCacheDefinition1\\.xml"`),
    'the pivotCaches relationship id resolves to the cache definition'
  );

  // Re-reading and re-writing the output keeps everything, so the passthrough is idempotent.
  const again = partNames(writeXlsx(readXlsx(out)));
  assert.ok(again.includes('xl/pivotTables/pivotTable1.xml'), 'idempotent across a second round-trip');
  assert.equal(again.filter(n => /xl\/pivotCache\/.+\.xml$/.test(n)).length, 2, 'the cache survives a second round-trip');
});

test('slicer and slicer-cache parts survive read→write', () => {
  const src = zipSync(
    packageParts({
      '[Content_Types].xml': contentTypes(
        '<Override PartName="/xl/slicers/slicer1.xml" ContentType="application/vnd.ms-excel.slicer+xml"/>' +
          '<Override PartName="/xl/slicerCaches/slicerCache1.xml" ContentType="application/vnd.ms-excel.slicerCache+xml"/>'
      ),
      'xl/_rels/workbook.xml.rels':
        '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        relationship('rId1', 'worksheet', 'worksheets/sheet1.xml') +
        msRelationship('rId2', 'slicerCache', 'slicerCaches/slicerCache1.xml') +
        '</Relationships>',
      'xl/worksheets/sheet1.xml': worksheet(''),
      'xl/worksheets/_rels/sheet1.xml.rels': rels(msRelationship('rId1', 'slicer', '../slicers/slicer1.xml')),
      'xl/slicers/slicer1.xml': '<slicers/>',
      'xl/slicerCaches/slicerCache1.xml': '<slicerCacheDefinition/>',
    })
  );

  const out = writeXlsx(readXlsx(src));
  const names = partNames(out);

  assert.ok(names.includes('xl/slicers/slicer1.xml'), 'the slicer part survives');
  assert.ok(names.includes('xl/slicerCaches/slicerCache1.xml'), 'the slicer cache part survives');
  assert.match(partText(out, /worksheets\/_rels\/sheet1\.xml\.rels$/), /\/slicer"/, 'the sheet still references the slicer');
  assert.match(partText(out, /_rels\/workbook\.xml\.rels$/), /slicerCaches\/slicerCache1\.xml/, 'the workbook still references the slicer cache');

  // Parts surviving is not enough — Excel only rediscovers a slicer through its x14 wiring, which
  // references the relationship ids the writer reassigns, so the ext blocks must name the *new* ids.
  const slicerRelId = partText(out, /worksheets\/_rels\/sheet1\.xml\.rels$/).match(/Id="(rId\d+)"[^>]*\/slicer"/)?.[1];
  assert.ok(slicerRelId, 'the re-emitted slicer rel has an id');
  assert.match(
    partText(out, /worksheets\/sheet1\.xml$/),
    new RegExp(`<x14:slicerList><x14:slicer r:id="${slicerRelId}"/></x14:slicerList>`),
    'the sheet body reactivates the slicer through a slicerList extension wired to its rel'
  );
  const cacheRelId = partText(out, /_rels\/workbook\.xml\.rels$/).match(/Id="(rId\d+)"[^>]*\/slicerCache"/)?.[1];
  assert.ok(cacheRelId, 'the re-emitted slicer-cache rel has an id');
  assert.match(
    partText(out, /xl\/workbook\.xml$/),
    new RegExp(`<x14:slicerCaches><x14:slicerCache r:id="${cacheRelId}"/></x14:slicerCaches>`),
    'the workbook registers the slicer cache in its x14 slicerCaches extension'
  );
});
