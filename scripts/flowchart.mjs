#!/usr/bin/env node
// DOT -> SVG -> PNG（本机无 Graphviz，用 @viz-js/viz 的 Graphviz WASM 构建）
//   node flowchart.mjs -Dot fig.dot [-Png fig.png] [-Svg fig.svg] [-Scale 2] [-AllowCurves]
// 审计规则来自 SKILL.md 的"流程图 强制"：边只能是直线/直角折线；朴素样式下节点不得用椭圆或圆角。
'use strict';
import { createRequire } from 'node:module';
import { readFileSync, writeFileSync, existsSync, mkdirSync } from 'node:fs';
import { execSync } from 'node:child_process';
import path from 'node:path';
import os from 'node:os';

const arg = (n, d) => { const i = process.argv.indexOf('-' + n); return i > -1 ? process.argv[i + 1] : d; };
const dotPath = path.resolve(arg('Dot', ''));
if (!dotPath || !existsSync(dotPath)) throw new Error('用法: node flowchart.mjs -Dot <file.dot>');
const base = dotPath.replace(/\.dot$/, '');
const svgOut = path.resolve(arg('Svg', base + '.svg'));
const pngOut = path.resolve(arg('Png', base + '.png'));
const scale = Number(arg('Scale', 2));
const allowCurves = process.argv.includes('-AllowCurves');

// 依次在几个已知位置找 @viz-js/viz
const requireFrom = p => createRequire(p);
const candidates = [
  process.env.VIZ_JS_PATH,                                              // 显式指定：VIZ_JS_PATH=/path/to/@viz-js/viz
  path.resolve(import.meta.dirname, '..', 'node_modules', '@viz-js', 'viz'),
  path.resolve(process.cwd(), 'node_modules', '@viz-js', 'viz'),
  path.join(os.homedir(), 'AppData/Roaming/npm/node_modules/@viz-js/viz'),
].filter(Boolean);
let viz = null, usedPath = '';
for (const c of candidates) {
  try {
    const mod = requireFrom(import.meta.url)(c);
    viz = await mod.instance(); usedPath = c; break;
  } catch { /* try next */ }
}
if (!viz) {
  throw new Error('找不到 @viz-js/viz。安装方式：在任一目录执行 npm i @viz-js/viz，' +
    '或 --Viz <路径>。已尝试: ' + candidates.join(' , '));
}

const src = readFileSync(dotPath, 'utf8');
const res = viz.render(src, { format: 'svg', engine: 'dot' });
const errs = (res.errors || []).map(e => `[${e.level || 'info'}] ${String(e.message).trim()}`);
console.log('graphviz', res.graphvizVersion || viz.graphvizVersion, '| status:', res.status, '| from', usedPath);
if (errs.length) console.log(errs.join('\n'));
if (res.status !== 'success') throw new Error('DOT 渲染失败');

const svg = res.output;
writeFileSync(svgOut, svg, 'utf8');

// ---- 审计：按 <g class> 分组，量几何而不是数指令 ----
// Graphviz 的 ortho 把每段折线写成"控制点与端点重合"的退化 C，肉眼是直线。
// 所以判据是：每段必须共线（偏差≈0）且轴对齐（水平或垂直）。
const groups = [...svg.matchAll(/<g id="[^"]*"\s+class="(edge|node)">([\s\S]*?)<\/g>/g)];
const EPS = 0.02;
let segs = 0, maxBend = 0, maxOffAxis = 0, curvedNodes = [];
for (const [, kind, body] of groups) {
  if (kind === 'node') {
    if (/<ellipse/.test(body)) curvedNodes.push('ellipse');
    const rx = body.match(/<rect[^>]*\s(?:r|rx)="([\d.]+)"/);
    if (rx && Number(rx[1]) > 0.5) curvedNodes.push('rounded-rect');
    continue;
  }
  for (const m of body.matchAll(/<path[^>]*\sd="([^"]*)"/g)) {
    // 一个 C/L 命令可以带多组坐标（polybezier），必须按"到下一个命令字母为止"切分
    const parts = m[1].match(/[MLCZ][^MLCZ]*/g) || [];
    let cur = null;
    for (const part of parts) {
      const cmd = part[0];
      const nums = (part.slice(1).match(/-?[\d.]+/g) || []).map(Number);
      if (cmd === 'M') { cur = [nums[0], nums[1]]; continue; }
      const step = cmd === 'C' ? 6 : 2;
      for (let j = 0; j + step <= nums.length; j += step) {
        const g = nums.slice(j, j + step);
        const end = [g[step - 2], g[step - 1]];
        const ctrl = cmd === 'C' ? [[g[0], g[1]], [g[2], g[3]]] : [];
        const [a, b] = [cur, end];
        const len = Math.hypot(b[0] - a[0], b[1] - a[1]) || 1;
        for (const p of ctrl) {
          const dev = Math.abs((b[0] - a[0]) * (p[1] - a[1]) - (p[0] - a[0]) * (b[1] - a[1])) / len;
          maxBend = Math.max(maxBend, dev);
        }
        maxOffAxis = Math.max(maxOffAxis, Math.min(Math.abs(b[0] - a[0]), Math.abs(b[1] - a[1])));
        segs++;
        cur = end;
      }
    }
  }
}
console.log(`边分段: ${segs} | 最大弯曲偏差: ${maxBend.toFixed(4)}pt | 最大斜出量: ${maxOffAxis.toFixed(3)}pt`);
// 中文字宽兜底检查：这个构建按空格估 CJK 宽度，框会画窄，浏览器里真字体会溢出
const overflow = [];
for (const [, kind, body] of groups) {
  if (kind !== 'node') continue;
  const pts = [...body.matchAll(/points="([^"]*)"/g)].flatMap(m => m[1].trim().split(/\s+/))
    .map(s => s.split(',').map(Number)).filter(p => p.length === 2);
  if (!pts.length) continue;
  const boxW = Math.max(...pts.map(p => p[0])) - Math.min(...pts.map(p => p[0]));
  let need = 0;
  for (const t of body.matchAll(/font-size="([\d.]+)"[^>]*>([^<]*)<\/text>/g)) {
    const fsz = Number(t[1]);
    const w = [...t[2]].reduce((a, ch) =>
      a + (ch.codePointAt(0) > 0x2e80 ? fsz : ch === ' ' ? fsz * 0.26 : fsz * 0.52), 0);
    need = Math.max(need, w);
  }
  if (need > boxW - 2) overflow.push({ label: (body.match(/>([^<]*)<\/text>/) || [, '?'])[1], boxW: boxW.toFixed(1), need: need.toFixed(1) });
}
if (overflow.length) {
  console.log('文字可能溢出节点（给这些节点加大 width= 或缩短标签）:');
  overflow.forEach(o => console.log(`  「${o.label}」框宽 ${o.boxW}pt < 需要 ${o.need}pt`));
}
if (!allowCurves) {
  if (maxBend > EPS) throw new Error(`边有真实弯曲（最大偏差 ${maxBend.toFixed(3)}pt）：splines=ortho 未生效`);
  if (maxOffAxis > EPS) throw new Error(`边有斜线（最大斜出 ${maxOffAxis.toFixed(2)}pt）：报告样式要求横平竖直`);
  if (curvedNodes.length) {
    throw new Error('节点用了非朴素形状: ' + [...new Set(curvedNodes)].join(', ') +
      '。报告样式要求矩形/菱形/平行四边形，起止也用直边矩形（不要用 ellipse、圆角）。');
  }
}

const ptW = Number(svg.match(/width="([\d.]+)pt"/)[1]);
const ptH = Number(svg.match(/height="([\d.]+)pt"/)[1]);
const pxW = Math.ceil(ptW * 4 / 3), pxH = Math.ceil(ptH * 4 / 3);
console.log(`尺寸: ${ptW}pt x ${ptH}pt = ${pxW}px x ${pxH}px | svg ${svg.length} bytes`);

const edge = 'C:/Program Files (x86)/Microsoft/Edge/Application/msedge.exe';
if (!existsSync(edge)) throw new Error('Edge 不存在: ' + edge);
mkdirSync(path.dirname(pngOut), { recursive: true });
execSync(`"${edge}" --headless=new --disable-gpu --hide-scrollbars --force-device-scale-factor=${scale} ` +
  `--window-size=${pxW},${pxH} --default-background-color=FFFFFFFF --screenshot="${pngOut}" ` +
  `"file:///${svgOut.replace(/\\/g, '/')}"`, { stdio: 'ignore' });
const png = readFileSync(pngOut);
console.log(`PNG: ${png.readUInt32BE(16)}x${png.readUInt32BE(20)} ${(png.length / 1024).toFixed(0)}KB -> ${pngOut}`);
if (Math.abs(png.readUInt32BE(16) - pxW * scale) > 4) console.log('提示: 截图宽度与预期不符，检查 window-size 是否被裁');
