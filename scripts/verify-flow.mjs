#!/usr/bin/env node
// 单主线流程图独立自检：只读 SVG + spine spec，不复用渲染器任何中间数据
//   node verify-flow.mjs <fig.svg> <fig.json>
// 查：中轴对齐 / 边集合与 spec 一致 / 起点端口 / 箭头不悬空(节点端口或汇合圆点边缘) /
//     不穿框(菱形/平行四边形逐 y 跨度) / 横平竖直 / 侧线零交叉(禁止跨线桥与任何曲线) /
//     标签齐全+归属+贴出口+不压线 / 判断框向下出边<=1 / 汇合点出入度
'use strict';
import { readFileSync } from 'node:fs';
const svg = readFileSync(process.argv[2], 'utf8');
const spec = JSON.parse(readFileSync(process.argv[3], 'utf8'));
const dec = s => s.replace(/&#45;/g, '-').replace(/&gt;/g, '>').replace(/&lt;/g, '<').replace(/&amp;/g, '&');
const nodes = {}, dots = {}, edges = [], labels = [];
for (const m of svg.matchAll(/<g class="node"><title>([^<]*)<\/title>([\s\S]*?)<\/g>/g)) {
  const id = dec(m[1]), body = m[2];
  let pts = [];
  const poly = body.match(/<polygon points="([^"]*)"/);
  const rect = body.match(/<rect x="([\d.-]+)" y="([\d.-]+)" width="([\d.]+)" height="([\d.]+)"/);
  if (poly) pts = poly[1].trim().split(/\s+/).map(s => s.split(',').map(Number));
  else if (rect) pts = [[+rect[1], +rect[2]], [+rect[1] + +rect[3], +rect[2] + +rect[4]]];
  if (!pts.length) continue;
  const xs = pts.map(p => p[0]), ys = pts.map(p => p[1]);
  nodes[id] = { x0: Math.min(...xs), x1: Math.max(...xs), y0: Math.min(...ys), y1: Math.max(...ys),
    shape: poly ? (pts.length === 4 && pts[0][0] === pts[2][0] ? 'diamond' : 'para') : 'rect' };
}
for (const m of svg.matchAll(/<circle class="merge" id="([^"]*)" cx="([\d.-]+)" cy="([\d.-]+)" r="([\d.]+)"/g))
  dots[dec(m[1])] = { p: [+m[2], +m[3]], r: +m[4] };
for (const m of svg.matchAll(/<g class="edge"><title>([^<]*)<\/title>([\s\S]*?)<\/g>/g)) {
  const name = dec(m[1]), body = m[2];
  const d = (body.match(/<path d="([^"]*)"/) || [])[1] || '';
  const pts = (d.match(/-?[\d.]+ -?[\d.]+/g) || []).map(s => s.split(' ').map(Number));
  const tip = (body.match(/<polygon points="([^"]*)"/) || [])[1];
  edges.push({ name, pts, tip: tip ? tip.trim().split(/\s+/)[0].split(',').map(Number) : null });
}
for (const m of svg.matchAll(/<text class="blabel" data-edge="([^"]*)" x="([\d.-]+)" y="([\d.-]+)"[^>]*>([^<]*)</g))
  labels.push({ edge: dec(m[1]), p: [+m[2], +m[3]], text: m[4] });

let bad = 0;
const fail = msg => { bad++; console.log('BAD ' + msg); };
const near = (a, b, t = 1.5) => Math.hypot(a[0] - b[0], a[1] - b[1]) <= t;
const ports = n => {
  const cx = (n.x0 + n.x1) / 2, cy = (n.y0 + n.y1) / 2;
  if (n.shape === 'para') return { n: [cx, n.y0], s: [cx, n.y1], w: [n.x0 + 9, cy], e: [n.x1 - 9, cy] };
  return { n: [cx, n.y0], s: [cx, n.y1], w: [n.x0, cy], e: [n.x1, cy] };
};
const spanAt = (n, y) => {
  if (y <= n.y0 || y >= n.y1) return null;
  const t = (y - n.y0) / (n.y1 - n.y0);
  if (n.shape === 'diamond') {
    const cx = (n.x0 + n.x1) / 2, hw = (n.x1 - n.x0) / 2 * (1 - Math.abs(2 * t - 1));
    return [cx - hw, cx + hw];
  }
  if (n.shape === 'para') return [n.x0 + 18 * (1 - t), n.x1 - 18 * t];
  return [n.x0, n.x1];
};
// 1 中轴对齐
const cxs = Object.values(nodes).map(n => (n.x0 + n.x1) / 2);
if (Math.max(...cxs) - Math.min(...cxs) > 0.5) fail('节点未沿同一竖直中轴线');
// 2 边集合 / 标签 / 圆点 与 spec 推导的期望一致
const want = [], wantLabels = {};
for (let i = 0; i + 1 < spec.spine.length; i++) {
  const a = spec.spine[i].id, b = spec.spine[i + 1].id;
  const m = (spec.merges || []).find(x => x.between[0] === a && x.between[1] === b);
  const lb = (spec.spineLabels || {})[`${a}->${b}`];
  if (m) {
    if (!m.noUpper) { want.push(`${a}->${m.id}`); if (lb) wantLabels[`${a}->${m.id}`] = lb; }
    want.push(`${m.id}->${b}`);
  } else { want.push(`${a}->${b}`); if (lb) wantLabels[`${a}->${b}`] = lb; }
}
for (const s of spec.sides || []) { want.push(`${s.from}->${s.to}`); if (s.label) wantLabels[`${s.from}->${s.to}`] = s.label; }
for (const l of spec.loops || []) { want.push(`${l.from}->${l.to}`); if (l.label) wantLabels[`${l.from}->${l.to}`] = l.label; }
const got = edges.map(e => e.name).sort();
if (JSON.stringify(got) !== JSON.stringify(want.slice().sort()))
  fail(`边集合不符: 多 ${got.filter(x => !want.includes(x))} 少 ${want.filter(x => !got.includes(x))}`);
const wantDots = (spec.merges || []).map(m => m.id).sort();
const gotDots = Object.keys(dots).sort();
if (JSON.stringify(gotDots) !== JSON.stringify(wantDots)) fail(`汇合圆点不符: ${gotDots} 期望 ${wantDots}`);
// 3 逐边：端口/悬空/穿框/轴对齐
for (const e of edges) {
  const [a, b] = e.name.split('->');
  const A = nodes[a] || dots[a], B = nodes[b] || dots[b];
  if (!A || !B) { fail(`${e.name} 端点缺失`); continue; }
  if (nodes[a]) {
    const pa = ports(nodes[a]);
    if (!Object.keys(pa).some(k => near(e.pts[0], pa[k], 2))) fail(`${e.name} 起点不在 ${a} 端口`);
  }
  if (e.tip) {
    if (dots[b]) {
      if (Math.abs(Math.hypot(e.tip[0] - dots[b].p[0], e.tip[1] - dots[b].p[1]) - dots[b].r) > 1)
        fail(`${e.name} 箭头未停在汇合圆点边缘: ${e.tip}`);
      if (Math.abs(e.tip[1] - dots[b].p[1]) > 0.5) fail(`${e.name} 汇入方向非水平`);
    } else if (nodes[b]) {
      const pb = ports(nodes[b]);
      const head = Object.keys(pb).find(k => near(e.tip, pb[k], 1));
      if (!head) fail(`${e.name} 箭头悬空: ${e.tip}`);
      else if (nodes[b].shape === 'diamond' && head === 's') fail(`${e.name} 从下顶点进判断框`);
    }
  } else if (nodes[b] || dots[b]) {
    // 无箭头边只允许 spine 上段 a->m（终点即圆点中心）
    if (!dots[b] || !near(e.pts[e.pts.length - 1], dots[b].p, 0.5)) fail(`${e.name} 无箭头且终点不是汇合点`);
  }
  for (let i = 0; i + 1 < e.pts.length; i++) {
    const [p, q] = [e.pts[i], e.pts[i + 1]];
    if (Math.abs(p[0] - q[0]) > 0.01 && Math.abs(p[1] - q[1]) > 0.01) fail(`${e.name} 第${i + 1}段斜线`);
    for (const [id, n] of Object.entries(nodes)) {
      if (id === a || id === b) continue;
      if (p[1] === q[1]) {
        const sp = spanAt(n, p[1]); if (!sp) continue;
        if (Math.max(p[0], q[0]) > sp[0] + 0.5 && Math.min(p[0], q[0]) < sp[1] - 0.5) fail(`${e.name} 水平段穿 ${id}`);
      } else if (Math.abs(p[0]) > 0.5) {
        for (let yy = Math.min(p[1], q[1]) + 1; yy < Math.max(p[1], q[1]); yy += 4) {
          const sp = spanAt(n, yy); if (!sp) continue;
          if (p[0] > sp[0] + 0.5 && p[0] < sp[1] - 0.5) { fail(`${e.name} 竖直段穿 ${id}`); break; }
        }
      }
    }
  }
}
// 4 侧线交叉：不允许任何交叉，也不允许跨线桥（桥拱是曲线）。出现交叉即不合格，须改 spec。
const segs = edges.flatMap(e => e.pts.slice(0, -1).map((p, i) => ({ e: e.name, p, q: e.pts[i + 1] })));
const cross = [];
for (let i = 0; i < segs.length; i++) for (let j = i + 1; j < segs.length; j++) {
  const s1 = segs[i], s2 = segs[j];
  if (s1.e === s2.e) continue;
  const h = s1.p[1] === s1.q[1] ? s1 : (s2.p[1] === s2.q[1] ? s2 : null);
  const v = h === s1 ? s2 : s1;
  if (!h || v.p[0] !== v.q[0]) continue;
  const x = v.p[0], y = h.p[1];
  const hx0 = Math.min(h.p[0], h.q[0]), hx1 = Math.max(h.p[0], h.q[0]);
  const vy0 = Math.min(v.p[1], v.q[1]), vy1 = Math.max(v.p[1], v.q[1]);
  if (x > hx0 + 0.5 && x < hx1 - 0.5 && y > vy0 + 0.5 && y < vy1 - 0.5) cross.push([x, y, s1.e, s2.e]);
}
for (const [x, y, e1, e2] of cross) fail(`交叉 ${e1} x ${e2} @(${x},${y})：禁止跨线桥，改 spec 换侧/换汇合点/调主线顺序`);
if (svg.includes('class="hop"') || svg.includes('class="bridge"')) fail('SVG 含跨线桥/桥拱（曲线），不允许');
// 5 标签齐全 / 归属 / 贴出口 / 不压线
for (const [edge, text] of Object.entries(wantLabels))
  if (!labels.some(l => l.edge === edge && l.text === text)) fail(`缺标签 ${text} @ ${edge}`);
if (labels.length !== Object.keys(wantLabels).length) fail(`标签数 ${labels.length} != ${Object.keys(wantLabels).length}`);
for (const l of labels) {
  const e = edges.find(x => x.name === l.edge);
  if (!e) { fail(`标签 ${l.text} 无边 ${l.edge}`); continue; }
  if (Math.hypot(l.p[0] - e.pts[0][0], l.p[1] - e.pts[0][1]) > 25) fail(`标签 ${l.text} 离出口过远`);
  for (const s of segs) {
    const onH = s.p[1] === s.q[1] && Math.abs(l.p[1] - s.p[1]) < 4 && l.p[0] > Math.min(s.p[0], s.q[0]) - 2 && l.p[0] < Math.max(s.p[0], s.q[0]) + 2;
    const onV = s.p[0] === s.q[0] && Math.abs(l.p[0] - s.p[0]) < 4 && l.p[1] > Math.min(s.p[1], s.q[1]) - 2 && l.p[1] < Math.max(s.p[1], s.q[1]) + 2;
    if (onH || onV) fail(`标签 ${l.text} 压线 ${s.e}`);
  }
}
// 6 判断框向下出边 <=1
for (const n of spec.spine.filter(n => n.shape === 'diamond')) {
  const downs = edges.filter(e => e.name.startsWith(n.id + '->') && nodes[n.id] && near(e.pts[0], ports(nodes[n.id]).s, 2));
  if (downs.length > 1) fail(`${n.id} 向下出边 ${downs.length} 条`);
}
// 7 汇合点出入度：入 = 上段(非 noUpper) + 指向它的侧线；出 = 1 条向下
for (const m of spec.merges || []) {
  const inUp = m.noUpper ? 0 : 1;
  const inSide = (spec.sides || []).filter(s => s.to === m.id).length;
  const gotIn = edges.filter(e => e.name.endsWith('->' + m.id)).length;
  const gotOut = edges.filter(e => e.name.startsWith(m.id + '->')).length;
  if (gotIn !== inUp + inSide) fail(`${m.id} 入边 ${gotIn} != ${inUp + inSide}`);
  if (gotOut !== 1) fail(`${m.id} 出边 ${gotOut} != 1`);
}
console.log(`汇合圆点: ${gotDots.join(',')} | 交叉: ${cross.length} 处 | 边 ${edges.length} 条 | 标签 ${labels.length} 个`);
console.log(bad ? `${bad} 项不合格` : '全部自检通过：中轴/边集合/端口/不悬空/不穿框/轴对齐/零交叉/标签/向下出边/汇合度');
process.exit(bad ? 1 : 0);