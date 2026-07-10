import {createRequire} from 'node:module';
const require = createRequire(import.meta.url);
const ExcelJS = require('../../lib/exceljs.nodejs.js');
const JSZip = require('jszip');
const line = s => console.log(s);

// 1454 insert row with style inheritance freezes cells
async function probeInsertFreeze() {
  line('\n=== 1454 insert-row style-inheritance freeze ===');
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('S');
  ws.getCell('A1').value = 'header'; ws.getCell('A1').font = {bold:true};
  ws.getCell('A2').value = 'data';
  // insertRow with style inheritance ('i+' = inherit from above? ExcelJS: insertRow(pos, value, style))
  let err = null;
  try {
    const r = ws.insertRow(2, ['inserted'], 'i'); // 'i' = inherit style
    ws.getCell('A2').numFmt = '$#,##0.00;[Red]-$#,##0.00';
  } catch(e){ err = String(e.message||e); }
  line('  insert+style with inheritance: ' + (err ? 'THREW '+err : 'ok'));
  // try 'o' (no style)
  const wb2 = new ExcelJS.Workbook();
  const ws2 = wb2.addWorksheet('S');
  ws2.getCell('A1').value='h'; ws2.getCell('A2').value='d';
  let err2=null;
  try { ws2.insertRow(2, ['x'], 'o'); ws2.getCell('A2').numFmt = '0.00'; } catch(e){err2=String(e.message||e);}
  line('  insert+style no-inheritance (o): ' + (err2?'THREW '+err2:'ok'));
}

// 1445 write to slave cell of merge
async function probeSlaveCell() {
  line('\n=== 1445 slave cell value write ===');
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('S');
  ws.mergeCells('A1:B2'); // 2x2 merge, A1 master
  ws.getCell('B2').value = 'slave-write'; // write to bottom-right slave
  const buf = await wb.xlsx.writeBuffer();
  const zip = await JSZip.loadAsync(buf);
  const sx = await zip.file('xl/worksheets/sheet1.xml').async('string');
  const cells = [...sx.matchAll(/<c r="([A-Z]\d+)"[^>]*>([\s\S]*?)<\/c>/g)].map(m=>m[1]+'='+(m[2].match(/<v>([\s\S]*?)<\/v>/)||[])[1]);
  line('  populated cells: ' + JSON.stringify(cells));
  line('  merges: ' + JSON.stringify([...sx.matchAll(/<mergeCell ref="([^"]*)"/g)].map(m=>m[1])));
  const wb2 = new ExcelJS.Workbook(); await wb2.xlsx.load(buf);
  const s = wb2.getWorksheet('S');
  line('  reload A1=' + JSON.stringify(s.getCell('A1').value) + ' B2=' + JSON.stringify(s.getCell('B2').value));
}

// 1189 column-letter case in cross-sheet reference
async function probeRefCase() {
  line('\n=== 1189 column-letter case ===');
  const wb = new ExcelJS.Workbook();
  const main = wb.addWorksheet('Main');
  const data = wb.addWorksheet('Data');
  for (let r=2;r<=8;r++) data.getCell('A'+r).value = 'opt'+r;
  main.getCell('B2').dataValidation = {type:'list', allowBlank:true, formulae:['Data!$A$2:$A$8']};
  main.getCell('C2').value = {formula:'Data!$A$2', result:'opt2'};
  const buf = await wb.xlsx.writeBuffer();
  const zip = await JSZip.loadAsync(buf);
  const sx = await zip.file('xl/worksheets/sheet1.xml').async('string');
  const dv = (sx.match(/<formula1>([\s\S]*?)<\/formula1>/)||[])[1];
  const f = (sx.match(/<c r="C2"[^>]*><f>([\s\S]*?)<\/f>/)||[])[1];
  line('  DV formula1: ' + dv + ' | cell formula: ' + f);
  line('  DV lowercased: ' + (dv && /\$a\$/.test(dv)) + ' | formula lowercased: ' + (f && /\$a\$/.test(f)));
}

// 1162 numeric-looking string preserved
async function probeStringType() {
  line('\n=== 1162 numeric string type ===');
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('S');
  ws.getCell('A1').value = '1000.80'; // string
  const buf = await wb.xlsx.writeBuffer();
  const wb2 = new ExcelJS.Workbook(); await wb2.xlsx.load(buf);
  const c = wb2.getWorksheet('S').getCell('A1');
  line('  A1 type=' + c.type + ' value=' + JSON.stringify(c.value));
}

// 0762 non-finite numeric (NaN/Infinity)
async function probeNaN() {
  line('\n=== 0762 non-finite numeric ===');
  for (const [label, val] of [['NaN', NaN], ['Infinity', Infinity], ['-Infinity', -Infinity]]) {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('S');
    ws.getCell('A1').value = val;
    let werr=null, buf;
    try{ buf = await wb.xlsx.writeBuffer(); }catch(e){werr=String(e.message||e);}
    if(werr){ line('  '+label+': WRITE THREW '+werr); continue; }
    const zip = await JSZip.loadAsync(buf);
    const sx = await zip.file('xl/worksheets/sheet1.xml').async('string');
    const cA1 = (sx.match(/<c r="A1"[^>]*>[\s\S]*?<\/c>|<c r="A1"[^>]*\/>/)||[''])[0];
    let rerr=null;
    try{ const wb2=new ExcelJS.Workbook(); await wb2.xlsx.load(buf);}catch(e){rerr=String(e.message||e);}
    line('  '+label+': cell='+cA1+' hasToken='+new RegExp(label.replace('-','-?')).test(cA1)+' reload='+(rerr?'THREW':'ok'));
  }
}

await probeInsertFreeze();
await probeSlaveCell();
await probeRefCase();
await probeStringType();
await probeNaN();
