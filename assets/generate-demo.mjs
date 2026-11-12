// Generates demo-light.svg and demo-dark.svg: the README quick-start, typed line by line
// on the left while the grid on the right fills in. CSS-only animation; every element's
// resting style is the final frame, so non-animating renderers show the finished scene.
// Run after `pnpm run build`: node assets/generate-demo.mjs
import {writeFileSync} from 'node:fs';
import {dirname, join} from 'node:path';
import {fileURLToPath} from 'node:url';

import {Workbook, writeXlsx} from '../dist/index.js';

const here = dirname(fileURLToPath(import.meta.url));

// Measure the real output so the caption never lies.
const wb = new Workbook();
const sheet = wb.addWorksheet('People');
sheet.getCell('A1').value = 'Name';
sheet.getCell('B1').value = 'Joined';
sheet.addRow(['Ada', new Date('2026-01-01')]);
sheet.addRow(['Grace', new Date('2026-02-14')]);
sheet.getCell('C1').value = {formula: 'COUNTA(A:A)', result: 3};
const byteCount = writeXlsx(wb).length;

const W = 900;
const H = 360;
const MONO = "ui-monospace, SFMono-Regular, Menlo, Consolas, 'Liberation Mono', monospace";
const SANS = "-apple-system, BlinkMacSystemFont, 'Segoe UI', Helvetica, Arial, sans-serif";

const themes = {
  light: {
    bg: '#ffffff',
    panel: '#f6f8fa',
    border: '#d0d7de',
    text: '#1f2328',
    muted: '#656d76',
    brand: '#146c43',
    brand2: '#178050',
    kw: '#8250df',
    str: '#0a3069',
    fn: '#953800',
    num: '#0550ae',
    punct: '#57606a',
    hl: 'rgba(20,108,67,0.10)',
    header: '#eef1f4',
    cellFlash: 'rgba(23,128,80,0.18)',
    dot: ['#ff5f57', '#febc2e', '#28c840'],
  },
  dark: {
    bg: '#0d1117',
    panel: '#161b22',
    border: '#30363d',
    text: '#e6edf3',
    muted: '#8b949e',
    brand: '#5cc79a',
    brand2: '#8adcba',
    kw: '#d2a8ff',
    str: '#a5d6ff',
    fn: '#ffa657',
    num: '#79c0ff',
    punct: '#8b949e',
    hl: 'rgba(92,199,154,0.12)',
    header: '#1c2128',
    cellFlash: 'rgba(92,199,154,0.22)',
    dot: ['#ff5f57', '#febc2e', '#28c840'],
  },
};

// Each code line is a list of [tokenClass, text]. `cells` are the grid effects that line triggers.
const code = [
  {
    t: [
      ['kw', 'import'],
      ['p', ' {'],
      ['id', 'Workbook'],
      ['p', ', '],
      ['id', 'writeXlsx'],
      ['p', '} '],
      ['kw', 'from'],
      ['s', " '@shbernal/ts-xlsx'"],
      ['p', ';'],
    ],
  },
  {t: []},
  {
    t: [
      ['kw', 'const'],
      ['id', ' wb'],
      ['p', ' = '],
      ['kw', 'new'],
      ['fn', ' Workbook'],
      ['p', '();'],
    ],
  },
  {
    t: [
      ['kw', 'const'],
      ['id', ' sheet'],
      ['p', ' = '],
      ['id', 'wb'],
      ['p', '.'],
      ['fn', 'addWorksheet'],
      ['p', '('],
      ['s', "'People'"],
      ['p', ');'],
    ],
    grid: 'sheet',
  },
  {
    t: [
      ['id', 'sheet'],
      ['p', '.'],
      ['fn', 'getCell'],
      ['p', '('],
      ['s', "'A1'"],
      ['p', ').'],
      ['id', 'value'],
      ['p', ' = '],
      ['s', "'Name'"],
      ['p', ';'],
    ],
    cells: ['A1'],
  },
  {
    t: [
      ['id', 'sheet'],
      ['p', '.'],
      ['fn', 'getCell'],
      ['p', '('],
      ['s', "'B1'"],
      ['p', ').'],
      ['id', 'value'],
      ['p', ' = '],
      ['s', "'Joined'"],
      ['p', ';'],
    ],
    cells: ['B1'],
  },
  {
    t: [
      ['id', 'sheet'],
      ['p', '.'],
      ['fn', 'addRow'],
      ['p', '(['],
      ['s', "'Ada'"],
      ['p', ', '],
      ['kw', 'new'],
      ['fn', ' Date'],
      ['p', '('],
      ['s', "'2026-01-01'"],
      ['p', ')]);'],
    ],
    cells: ['A2', 'B2'],
  },
  {
    t: [
      ['id', 'sheet'],
      ['p', '.'],
      ['fn', 'addRow'],
      ['p', '(['],
      ['s', "'Grace'"],
      ['p', ', '],
      ['kw', 'new'],
      ['fn', ' Date'],
      ['p', '('],
      ['s', "'2026-02-14'"],
      ['p', ')]);'],
    ],
    cells: ['A3', 'B3'],
  },
  {
    t: [
      ['id', 'sheet'],
      ['p', '.'],
      ['fn', 'getCell'],
      ['p', '('],
      ['s', "'C1'"],
      ['p', ').'],
      ['id', 'value'],
      ['p', ' = {'],
      ['id', 'formula'],
      ['p', ': '],
      ['s', "'COUNTA(A:A)'"],
      ['p', ', '],
      ['id', 'result'],
      ['p', ': '],
      ['n', '3'],
      ['p', '};'],
    ],
    cells: ['C1'],
  },
  {t: []},
  {
    t: [
      ['kw', 'const'],
      ['id', ' bytes'],
      ['p', ' = '],
      ['fn', 'writeXlsx'],
      ['p', '('],
      ['id', 'wb'],
      ['p', ');'],
    ],
    grid: 'caption',
  },
];

const cellText = {
  A1: {v: 'Name', bold: true},
  B1: {v: 'Joined', bold: true},
  C1: {v: '3', num: true, formula: true},
  A2: {v: 'Ada'},
  B2: {v: '2026-01-01', num: true},
  A3: {v: 'Grace'},
  B3: {v: '2026-02-14', num: true},
};

// Timeline (seconds). One line every STEP; empty lines take a short beat.
const START = 0.6;
const STEP = 0.75;
const BLANK = 0.25;
const HOLD = 10;
const FADE = 0.6;
const appear = [];
let clock = START;
for (const line of code) {
  appear.push(clock);
  clock += line.t.length ? STEP : BLANK;
}
const END = clock + 0.4; // caption lands here
const CYCLE = END + HOLD + FADE;
const pct = (s) => `${((s / CYCLE) * 100).toFixed(2)}%`;
const fadeOutAt = pct(CYCLE - FADE);

let kfId = 0;
const keyframes = [];
// Element hidden until `at`, visible through the hold, faded out for the loop restart.
function reveal(at, dur = 0.25, extra = '') {
  const name = `r${kfId++}`;
  keyframes.push(
    `@keyframes ${name}{0%,${pct(at)}{opacity:0${extra ? `;transform:${extra}` : ''}}` +
      `${pct(at + dur)},${fadeOutAt}{opacity:1${extra ? ';transform:none' : ''}}100%{opacity:0}}`,
  );
  return `animation:${name} ${CYCLE.toFixed(2)}s linear infinite`;
}
// A brief highlight pulse; resting opacity 0, so the static frame stays clean.
function pulse(at, dur = 0.9) {
  const name = `p${kfId++}`;
  keyframes.push(
    `@keyframes ${name}{0%,${pct(at)}{opacity:0}${pct(at + 0.12)}{opacity:1}${pct(at + dur)},100%{opacity:0}}`,
  );
  return `opacity:0;animation:${name} ${CYCLE.toFixed(2)}s linear infinite`;
}

const esc = (s) => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');

function build(th) {
  const out = [];
  const css = [];
  keyframes.length = 0;
  kfId = 0;

  // ---- Editor pane ----
  const ex = 12,
    ey = 12,
    ew = 556,
    eh = H - 24;
  out.push(
    `<rect x="${ex}" y="${ey}" width="${ew}" height="${eh}" rx="10" fill="${th.panel}" stroke="${th.border}"/>`,
  );
  out.push(
    `<line x1="${ex}" y1="${ey + 34}" x2="${ex + ew}" y2="${ey + 34}" stroke="${th.border}"/>`,
  );
  th.dot.forEach((c, i) =>
    out.push(`<circle cx="${ex + 18 + i * 16}" cy="${ey + 17}" r="5" fill="${c}" opacity="0.85"/>`),
  );
  out.push(
    `<text x="${ex + 72}" y="${ey + 21.5}" font-family="${MONO}" font-size="12" fill="${th.muted}">quick-start.ts</text>`,
  );

  const codeX = ex + 44,
    codeY0 = ey + 62,
    LH = 22;
  // "Running line" marker: steps down the code, resting on the last line.
  const lastIdx = code.length - 1;
  const hlName = 'hl';
  const stops = code
    .map((l, i) =>
      l.t.length ? `${pct(appear[i])}{transform:translateY(${(i - lastIdx) * LH}px)}` : '',
    )
    .join('');
  css.push(
    `@keyframes ${hlName}{0%{transform:translateY(${-lastIdx * LH}px);opacity:0}${pct(appear[0])}{opacity:1}${stops}${fadeOutAt}{opacity:1}100%{opacity:0}}`,
  );
  const hlY = codeY0 + lastIdx * LH - 15;
  out.push(
    `<rect x="${ex + 1}" y="${hlY}" width="${ew - 2}" height="${LH}" fill="${th.hl}" style="animation:${hlName} ${CYCLE.toFixed(2)}s steps(1,end) infinite"/>`,
  );
  out.push(
    `<rect x="${ex + 1}" y="${hlY}" width="3" height="${LH}" fill="${th.brand}" style="animation:${hlName} ${CYCLE.toFixed(2)}s steps(1,end) infinite"/>`,
  );

  const color = {kw: th.kw, s: th.str, fn: th.fn, n: th.num, p: th.punct, id: th.text};
  code.forEach((line, i) => {
    const y = codeY0 + i * LH;
    out.push(
      `<text x="${codeX - 14}" y="${y}" font-family="${MONO}" font-size="12" fill="${th.muted}" text-anchor="end" opacity="0.7">${i + 1}</text>`,
    );
    if (!line.t.length) return;
    const spans = line.t.map(([k, s]) => `<tspan fill="${color[k]}">${esc(s)}</tspan>`).join('');
    out.push(
      `<text x="${codeX}" y="${y}" font-family="${MONO}" font-size="12.5" xml:space="preserve" style="${reveal(appear[i], 0.2, 'translateX(-6px)')}">${spans}</text>`,
    );
  });

  // ---- Grid pane ----
  const gx = 588,
    gy = 12;
  const cols = [
    {k: '', w: 30},
    {k: 'A', w: 80},
    {k: 'B', w: 114},
    {k: 'C', w: 76},
  ];
  const RH = 30,
    gridRows = 5;
  const gw = cols.reduce((a, c) => a + c.w, 0);
  const gridTop = gy + 34; // leave room for formula bar
  const sheetIdx = code.findIndex((l) => l.grid === 'sheet');
  const sheetStyle = reveal(appear[sheetIdx], 0.35);

  // Formula bar showing the C1 formula once it lands.
  const c1Idx = code.findIndex((l) => l.cells?.includes('C1'));
  out.push(`<g style="${sheetStyle}">`);
  out.push(
    `<rect x="${gx}" y="${gy}" width="${gw}" height="24" rx="4" fill="${th.panel}" stroke="${th.border}"/>`,
  );
  out.push(
    `<text x="${gx + 10}" y="${gy + 16.5}" font-family="${SANS}" font-size="12" font-style="italic" fill="${th.muted}">fx</text>`,
  );
  out.push(
    `<line x1="${gx + 28}" y1="${gy + 5}" x2="${gx + 28}" y2="${gy + 19}" stroke="${th.border}"/>`,
  );
  out.push(`</g>`);
  out.push(
    `<text x="${gx + 38}" y="${gy + 16.5}" font-family="${MONO}" font-size="12" fill="${th.text}" style="${reveal(appear[c1Idx] + 0.15)}"><tspan fill="${th.muted}">C1</tspan>  =COUNTA(A:A)</text>`,
  );

  out.push(`<g style="${sheetStyle}">`);
  out.push(
    `<rect x="${gx}" y="${gridTop}" width="${gw}" height="${RH * (gridRows + 1)}" fill="${th.bg}" stroke="${th.border}"/>`,
  );
  out.push(`<rect x="${gx}" y="${gridTop}" width="${gw}" height="${RH}" fill="${th.header}"/>`);
  out.push(
    `<rect x="${gx}" y="${gridTop}" width="${cols[0].w}" height="${RH * (gridRows + 1)}" fill="${th.header}"/>`,
  );
  let cx = gx;
  const colX = {};
  for (const c of cols) {
    colX[c.k] = cx;
    if (c.k)
      out.push(
        `<text x="${cx + c.w / 2}" y="${gridTop + 19.5}" font-family="${SANS}" font-size="11.5" fill="${th.muted}" text-anchor="middle">${c.k}</text>`,
      );
    cx += c.w;
    out.push(
      `<line x1="${cx}" y1="${gridTop}" x2="${cx}" y2="${gridTop + RH * (gridRows + 1)}" stroke="${th.border}"/>`,
    );
  }
  for (let r = 1; r <= gridRows; r++) {
    const y = gridTop + r * RH;
    out.push(`<line x1="${gx}" y1="${y}" x2="${gx + gw}" y2="${y}" stroke="${th.border}"/>`);
    out.push(
      `<text x="${gx + cols[0].w / 2}" y="${y + 19.5}" font-family="${SANS}" font-size="11.5" fill="${th.muted}" text-anchor="middle">${r}</text>`,
    );
  }
  // Sheet tab
  const tabY = gridTop + RH * (gridRows + 1);
  out.push(
    `<path d="M${gx + 34} ${tabY} v20 a3 3 0 0 0 3 3 h70 a3 3 0 0 0 3 -3 v-20" fill="${th.bg}" stroke="${th.border}"/>`,
  );
  out.push(`<rect x="${gx + 35}" y="${tabY - 1}" width="74" height="3" fill="${th.bg}"/>`);
  out.push(
    `<rect x="${gx + 38}" y="${tabY + 19}" width="66" height="2" rx="1" fill="${th.brand}"/>`,
  );
  out.push(
    `<text x="${gx + 72}" y="${tabY + 15}" font-family="${SANS}" font-size="12" font-weight="600" fill="${th.brand}" text-anchor="middle">People</text>`,
  );
  out.push(`</g>`);

  const colW = Object.fromEntries(cols.map((c) => [c.k, c.w]));
  code.forEach((line, i) => {
    for (const ref of line.cells ?? []) {
      const col = ref[0],
        row = Number(ref.slice(1));
      const x = colX[col],
        y = gridTop + row * RH,
        w = colW[col];
      const cell = cellText[ref];
      const at = appear[i] + 0.15;
      out.push(
        `<rect x="${x + 1.5}" y="${y + 1.5}" width="${w - 3}" height="${RH - 3}" fill="${th.cellFlash}" stroke="${th.brand2}" stroke-width="1.5" style="${pulse(at)}"/>`,
      );
      const tx = cell.num ? x + w - 8 : x + 8;
      const anchor = cell.num ? 'end' : 'start';
      const fill = cell.formula ? th.brand : th.text;
      const font = cell.num ? MONO : SANS;
      const size = cell.num ? 12 : 13;
      out.push(
        `<text x="${tx}" y="${y + 20}" font-family="${font}" font-size="${size}"${cell.bold ? ' font-weight="600"' : ''} fill="${fill}" text-anchor="${anchor}" style="${reveal(at, 0.2)}">${esc(cell.v)}</text>`,
      );
      if (cell.formula) {
        out.push(
          `<path d="M${x + w - 1} ${y + 1} h-7 l7 7 z" fill="${th.brand}" style="${reveal(at, 0.2)}"/>`,
        );
      }
    }
  });

  // ---- Caption ----
  const capIdx = code.findIndex((l) => l.grid === 'caption');
  // Result box, baseline-aligned with the writeXlsx line in the editor.
  const capBase = codeY0 + capIdx * LH;
  const capStyle = reveal(appear[capIdx] + 0.3, 0.3, 'translateY(4px)');
  out.push(`<g style="${capStyle}">`);
  out.push(
    `<rect x="${gx}" y="${capBase - 21}" width="${gw}" height="32" rx="6" fill="${th.hl}" stroke="${th.brand}" stroke-opacity="0.55"/>`,
  );
  out.push(
    `<text x="${gx + gw / 2}" y="${capBase}" font-family="${MONO}" font-size="11.5" text-anchor="middle" xml:space="preserve">` +
      `<tspan fill="${th.fn}">writeXlsx</tspan><tspan fill="${th.punct}">(wb)</tspan><tspan fill="${th.brand}" font-weight="600"> → </tspan>` +
      `<tspan fill="${th.text}">Uint8Array</tspan><tspan fill="${th.muted}"> (</tspan><tspan fill="${th.brand}" font-weight="600">${byteCount.toLocaleString('en-US')} bytes</tspan><tspan fill="${th.muted}">)</tspan></text>`,
  );
  out.push(`</g>`);

  return `<svg xmlns="http://www.w3.org/2000/svg" width="${W}" height="${H}" viewBox="0 0 ${W} ${H}" role="img" aria-labelledby="t d">
<title id="t">ts-xlsx quick start</title>
<desc id="d">TypeScript code builds a workbook with a People sheet (Name, Joined, Ada 2026-01-01, Grace 2026-02-14, C1 =COUNTA(A:A) = 3); writeXlsx returns a ${byteCount}-byte Uint8Array.</desc>
<style>
text{white-space:pre}
${css.join('\n')}
${keyframes.join('\n')}
@media (prefers-reduced-motion: reduce){*{animation:none!important}}
</style>
<rect width="${W}" height="${H}" rx="12" fill="${th.bg}"/>
${out.join('\n')}
</svg>
`;
}

for (const [name, th] of Object.entries(themes)) {
  writeFileSync(join(here, `demo-${name}.svg`), build(th));
}
process.stdout.write(
  `wrote demo-light.svg, demo-dark.svg (caption: ${byteCount} bytes, cycle ${CYCLE.toFixed(2)}s)\n`,
);
