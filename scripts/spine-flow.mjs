#!/usr/bin/env node
// 单主线流程图确定性路由器：spine spec JSON -> SVG -> PNG（160dpi）
//   node spine-flow.mjs -Spec fig.json [-Svg fig.svg] [-Png fig.png]
// 约定（SKILL.md「流程图 强制」）：
//   全部节点沿竖直中轴线自上而下；分支/回边从侧顶点引出、外侧绕行、箭头汇回主线汇合点或目标侧顶点；
//   汇合点实心圆点且只有一条向下出边；只接顶点不穿框；横平竖直；标签紧贴分支出口；
//   同侧侧线跨度嵌套者短在内道；跨度交错即交叉，直接报错（不允许跨线桥），须改 spec：换侧/换汇合点/调主线顺序。
// 布局是算出来的，不是 Graphviz 排的：ortho 实测保证不了侧入/不穿框，故路由权在本脚本。
'use strict';
import { readFileSync, writeFileSync, existsSync } from 'node:fs';
import { execSync } from 'node:child_process';
import path from 'node:path';
const arg = (n, d) => { const i = process.argv.indexOf('-' + n); return i > -1 ? process.argv[i + 1] : d; };
const specPath = path.resolve(arg('Spec', ''));
if (!specPath || !existsSync(specPath)) throw new Error('用法: node spine-flow.mjs -Spec <fig.json>');
const spec = JSON.parse(readFileSync(specPath, 'utf8'));
const svgOut = path.resolve(arg('Svg', specPath.replace(/\.json$/, '') + '.svg'));
const pngOut = arg('Png', '');
const FS = spec.fontsize || 12, GAP = spec.gap || 36, SKEW = 18, DOTR = 3;
const LANE_STEP = 30;

// ---- 主干布局：x 全 0，y 累加 ----
const N = {};
const lines = l => String(l).split('\n');
const boxH = n => n.shape === 'diamond' ? (n.h || 0.9) * 72 : lines(n.label).length * 15 + 18;
let y = 0;
for (const n of spec.spine) {
  if (N[n.id]) throw new Error('节点 id 重复: ' + n.id);
  n.h = boxH(n); n.wpt = (n.w || 1.5) * 72; n.x = 0; n.y = y + n.h / 2; y += n.h + GAP; N[n.id] = n;
}
const M = {};
for (const m of spec.merges || []) {
  const i = spec.spine.findIndex(n => n.id === m.between[0]);
  if (i < 0 || !spec.spine[i + 1] || spec.spine[i + 1].id !== m.between[1])
    throw new Error(`汇合点 ${m.id} 的 between 必须是主干上相邻两节点: ${m.between}`);
  const a = N[m.between[0]], b = N[m.between[1]];
  m.x = 0; m.y = (a.y + a.h / 2 + b.y - b.h / 2) / 2; M[m.id] = m;
}
const port = (id, p) => {
  const n = N[id], hw = n.wpt / 2, hh = n.h / 2;
  if (n.shape === 'parallelogram') {
    if (p === 'w') return [n.x - hw + SKEW / 2, n.y];
    if (p === 'e') return [n.x + hw - SKEW / 2, n.y];
  }
  return { n: [n.x, n.y - hh], s: [n.x, n.y + hh], w: [n.x - hw, n.y], e: [n.x + hw, n.y] }[p];
};
const DIR = { n: [0, 1], s: [0, -1], w: [1, 0], e: [-1, 0] };

// ---- 端口占用检查：一个顶点只许一条线接出/接入 ----
const used = {};
const claim = (id, p, role) => {
  const k = id + ':' + p;
  if (used[k]) throw new Error(`端口 ${k} 被 ${used[k]} 与 ${role} 共用：侧线/回边请换顶点`);
  used[k] = role;
};
for (let i = 0; i + 1 < spec.spine.length; i++) { claim(spec.spine[i].id, 's', 'spine'); claim(spec.spine[i + 1].id, 'n', 'spine'); }
for (const s of spec.sides || []) {
  if (!M[s.to]) throw new Error(`侧线 ${s.from}->${s.to} 的终点必须是汇合点`);
  if (s.tail !== 'w' && s.tail !== 'e') throw new Error(`侧线 ${s.from} 出口只能是 w/e 侧顶点`);
  claim(s.from, s.tail, 'side');
}
for (const l of spec.loops || []) {
  if (!N[l.to]) throw new Error(`回边 ${l.from}->${l.to} 的终点必须是节点`);
  if (l.tail !== 'w' && l.tail !== 'e') throw new Error(`回边 ${l.from} 出口只能是 w/e 侧顶点`);
  if (l.head !== 'w' && l.head !== 'e') throw new Error(`回边 ${l.to} 入口只能是 w/e 侧顶点（水平接入）`);
  claim(l.from, l.tail, 'loop'); claim(l.to, l.head, 'loop');
}

// ---- 车道自动分配：同侧按跨度长度升序（嵌套短者内道）；交错即交叉，后面直接报错 ----
const maxHalf = Math.max(...Object.values(N).map(n => n.wpt / 2));
const laneBase = Math.ceil(maxHalf + LANE_STEP);
const sideEdges = [];
for (const s of spec.sides || []) sideEdges.push({ kind: 'side', ...s });
for (const l of spec.loops || []) sideEdges.push({ kind: 'loop', ...l });
for (const e of sideEdges) {
  const a = port(e.from, e.tail)[1];
  const b = e.kind === 'side' ? M[e.to].y : port(e.to, e.head)[1];
  e.y0 = Math.min(a, b); e.y1 = Math.max(a, b);
}
for (const e of sideEdges) {
  e.side = e.lane ? Math.sign(e.lane) : e.side === 'L' ? -1 : e.side === 'R' ? 1 : 0;
  if (!e.side) throw new Error(`侧线/回边 ${e.from}->${e.to} 需给 side: "L"/"R" 或带符号 lane`);
}
const lanes = {};
for (const sign of [1, -1]) {
  const group = sideEdges.filter(e => e.side === sign);
  // 显式 lane 优先（兼容旧 spec）；自动者按跨度升序排道：嵌套短者内道
  const auto = group.filter(e => !e.lane).sort((a, b) => (a.y1 - a.y0) - (b.y1 - b.y0));
  let k = group.filter(e => e.lane).length;
  for (const e of auto) { e.lane = sign * (laneBase + LANE_STEP * k); k++; }
  for (const e of group) lanes[e.from + '->' + e.to] = e.lane;
}

// ---- 边表 ----
const edges = [];
const push = (name, pts, arrowDir, label, tail) => edges.push({ name, pts, arrowDir, label, tail });
for (let i = 0; i + 1 < spec.spine.length; i++) {
  const a = spec.spine[i], b = spec.spine[i + 1];
  const m = (spec.merges || []).find(x => x.between[0] === a.id && x.between[1] === b.id);
  const lb = (spec.spineLabels || {})[`${a.id}->${b.id}`];
  if (m) {
    // noUpper：上一节点唯一出口是侧线时，框底与汇合点之间留空，不画假流向
    if (!m.noUpper) push(`${a.id}->${m.id}`, [port(a.id, 's'), [0, m.y]], null, lb, 's');
    push(`${m.id}->${b.id}`, [[0, m.y], port(b.id, 'n')], DIR.n);
  } else push(`${a.id}->${b.id}`, [port(a.id, 's'), port(b.id, 'n')], DIR.n, lb, 's');
}
for (const e of sideEdges) {
  const P = port(e.from, e.tail);
  if (e.kind === 'side') {
    const m = M[e.to], dir = e.lane > 0 ? -1 : 1;
    const tipX = m.x + (e.lane > 0 ? DOTR : -DOTR);
    push(`${e.from}->${e.to}`, [P, [e.lane, P[1]], [e.lane, m.y], [tipX, m.y]], [dir, 0], e.label, e.tail);
  } else {
    const Q = port(e.to, e.head);
    push(`${e.from}->${e.to}`, [P, [e.lane, P[1]], [e.lane, Q[1]], Q], DIR[e.head], e.label, e.tail);
  }
}

// ---- 交叉检测：H×V 真交叉（水平段穿过另一条边的竖直段内部） ----
const segList = () => edges.flatMap(e => e.pts.slice(0, -1).map((p, i) => ({ e: e.name, p, q: e.pts[i + 1], i })));
const crossings = () => {
  const segs = segList(), out = [];
  for (let i = 0; i < segs.length; i++) for (let j = i + 1; j < segs.length; j++) {
    const s1 = segs[i], s2 = segs[j];
    if (s1.e === s2.e) continue;
    const h = s1.p[1] === s1.q[1] ? s1 : (s2.p[1] === s2.q[1] ? s2 : null);
    const v = h === s1 ? s2 : s1;
    if (!h || v.p[0] !== v.q[0]) continue;
    const x = v.p[0], yy = h.p[1];
    const hx0 = Math.min(h.p[0], h.q[0]), hx1 = Math.max(h.p[0], h.q[0]);
    const vy0 = Math.min(v.p[1], v.q[1]), vy1 = Math.max(v.p[1], v.q[1]);
    if (x > hx0 + 0.5 && x < hx1 - 0.5 && yy > vy0 + 0.5 && yy < vy1 - 0.5) out.push({ x, y: yy, v: v.e, vi: v.i, h: h.e });
  }
  return out;
};
// ---- 交叉检测：不允许任何交叉，也不允许跨线桥（桥拱是曲线，违反横平竖直） ----
// 交错即报错，作者必须改 spec：换侧、换汇合点位置、或调整主线顺序使跨度嵌套/分离。
const cs = crossings();
if (cs.length)
  throw new Error('侧线交叉（禁止跨线桥）: ' + cs.map(c => `${c.h} x ${c.v} @(${c.x.toFixed(0)},${c.y.toFixed(0)})`).join('; ') +
    ' —— 把其中一条换到另一侧、或移动汇合点/主线顺序使跨度嵌套或分离');

// ---- 标签：紧贴 tail 端口、离线 ----
const labels = [];
for (const e of edges) {
  if (!e.label) continue;
  const P = e.pts[0];
  const pos = e.tail === 's' ? [P[0] - 7, P[1] + 15, 'end']
    : e.tail === 'e' ? [P[0] + 10, P[1] - 7, 'start']
      : [P[0] - 10, P[1] - 7, 'end'];
  labels.push({ text: e.label, pos, edge: e.name });
}

// ---- 出 SVG ----
const xs = [], ys = [];
for (const n of Object.values(N)) { xs.push(n.x - n.wpt / 2, n.x + n.wpt / 2); ys.push(n.y - n.h / 2, n.y + n.h / 2); }
for (const m of Object.values(M)) { xs.push(m.x); ys.push(m.y); }
for (const e of edges) for (const p of e.pts) { xs.push(p[0]); ys.push(p[1]); }
for (const l of labels) { xs.push(l.pos[0]); ys.push(l.pos[1]); }
const MG = 30;
const x0 = Math.min(...xs) - MG, x1 = Math.max(...xs) + MG, y0 = Math.min(...ys) - MG, y1 = Math.max(...ys) + MG;
const W = x1 - x0, H = y1 - y0;
const esc = s => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
let out = `<svg xmlns="http://www.w3.org/2000/svg" width="${(W * 160 / 72).toFixed(1)}px" height="${(H * 160 / 72).toFixed(1)}px" viewBox="${x0} ${y0} ${W} ${H}">\n`;
out += `<rect x="${x0}" y="${y0}" width="${W}" height="${H}" fill="white"/>\n`;
for (const e of edges) {
  const pts = e.pts.map(p => p.slice());
  if (e.arrowDir) {
    const d = e.arrowDir, last = pts[pts.length - 1];
    pts[pts.length - 1] = [last[0] - d[0] * 8, last[1] - d[1] * 8];
  }
  let d = `M ${pts[0].join(' ')}`;
  for (let i = 1; i < pts.length; i++) d += ' L ' + pts[i].join(' ');
  out += `<g class="edge"><title>${esc(e.name)}</title><path d="${d}" fill="none" stroke="black" stroke-width="1"/>\n`;
  if (e.arrowDir) {
    const Q = e.pts[e.pts.length - 1], d = e.arrowDir, px = -d[1], py = d[0];
    const b = [Q[0] - d[0] * 8, Q[1] - d[1] * 8];
    out += `<polygon points="${Q.join(',')} ${b[0] + px * 3.5},${b[1] + py * 3.5} ${b[0] - px * 3.5},${b[1] - py * 3.5}" fill="black"/>\n`;
  }
  out += `</g>\n`;
}
for (const n of Object.values(N)) {
  const hw = n.wpt / 2, hh = n.h / 2, X = n.x, Y = n.y;
  out += `<g class="node"><title>${esc(n.id)}</title>`;
  if (n.shape === 'diamond')
    out += `<polygon points="${X},${Y - hh} ${X + hw},${Y} ${X},${Y + hh} ${X - hw},${Y}" fill="white" stroke="black" stroke-width="1"/>`;
  else if (n.shape === 'parallelogram')
    out += `<polygon points="${X - hw + SKEW},${Y - hh} ${X + hw},${Y - hh} ${X + hw - SKEW},${Y + hh} ${X - hw},${Y + hh}" fill="white" stroke="black" stroke-width="1"/>`;
  else
    out += `<rect x="${X - hw}" y="${Y - hh}" width="${n.wpt}" height="${n.h}" rx="${Math.min(12, n.h * 0.35)}" fill="white" stroke="black" stroke-width="1"/>`;
  const ls = lines(n.label);
  ls.forEach((t, i) => {
    out += `<text x="${X}" y="${Y + (i - (ls.length - 1) / 2) * 15 + 4}" font-family="SimSun" font-size="${FS}" text-anchor="middle" fill="black">${esc(t)}</text>`;
  });
  out += `</g>\n`;
}
for (const m of Object.values(M))
  out += `<circle class="merge" id="${esc(m.id)}" cx="${m.x}" cy="${m.y}" r="${DOTR}" fill="black"/>\n`;
for (const l of labels)
  out += `<text class="blabel" data-edge="${esc(l.edge)}" x="${l.pos[0]}" y="${l.pos[1]}" font-family="SimSun" font-size="11" text-anchor="${l.pos[2]}" fill="black">${esc(l.text)}</text>\n`;
out += `</svg>\n`;
writeFileSync(svgOut, out, 'utf8');
console.log(`SVG ${W.toFixed(0)}pt x ${H.toFixed(0)}pt | 车道 ${Object.entries(lanes).map(([k, v]) => k + ':' + v).join(' ')} | 交叉 0 处 -> ${svgOut}`);
if (pngOut) {
  const pxW = Math.ceil(W * 160 / 72), pxH = Math.ceil(H * 160 / 72);
  const edge = 'C:/Program Files (x86)/Microsoft/Edge/Application/msedge.exe';
  if (!existsSync(edge)) throw new Error('Edge 不存在: ' + edge);
  execSync(`"${edge}" --headless=new --disable-gpu --hide-scrollbars --force-device-scale-factor=1 ` +
    `--window-size=${pxW},${pxH} --default-background-color=FFFFFFFF --screenshot="${path.resolve(pngOut)}" ` +
    `"file:///${svgOut.replace(/\\/g, '/')}"`, { stdio: 'ignore' });
  const png = readFileSync(path.resolve(pngOut));
  console.log(`PNG ${png.readUInt32BE(16)}x${png.readUInt32BE(20)} -> ${pngOut}`);
}