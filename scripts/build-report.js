#!/usr/bin/env node
// spec 驱动的中文实验/上机报告生成器。
//   node build-report.js -Spec report.json [-Out x.docx]
// 编号（一、/1./（1）/①）由脚本按层级自动生成并在父项下重置，不要手写在 spec 里。
// 依赖：Node、PowerShell 7（pwsh，用于解包/打包）、Word 16 COM（由 pack-docx.ps1 调用）、Edge（流程图截图）。
'use strict';
const fs = require('fs');
const os = require('os');
const path = require('path');
const cp = require('child_process');

const SKILL = path.resolve(__dirname, '..');
const RT = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const arg = (name, def) => {
  const i = process.argv.indexOf('-' + name);
  return i > -1 ? process.argv[i + 1] : def;
};
const SPEC_PATH = path.resolve(arg('Spec', ''));
if (!SPEC_PATH || !fs.existsSync(SPEC_PATH)) throw new Error('用法: node build-report.js -Spec <spec.json>');
const HERE = path.dirname(SPEC_PATH);
// 素材路径写相对路径时，以 spec 文件所在目录为基准，示例因此可以在任意工作目录下跑
const fromSpec = p => (path.isAbsolute(p) ? p : path.resolve(HERE, p));
const spec = JSON.parse(fs.readFileSync(SPEC_PATH, 'utf8'));

const TEMPLATES = {
  'upper-report': { file: 'assets/template-upper-report.docx', needFooter: true, gridFix: [1271, 2750, 993, 950, 1383, 949] },
  'ai-lab': { file: 'assets/template-ai-lab.docx', needFooter: false, gridFix: null },
};
const tpl = TEMPLATES[spec.template];
if (!tpl) throw new Error('template 必须是 ' + Object.keys(TEMPLATES).join(' / '));
const OUT = path.resolve(spec.out || arg('Out', 'report.docx'));
fs.mkdirSync(path.dirname(OUT), { recursive: true });

// ---------- OOXML 片段 ----------
const esc = s => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
const TNR = 'w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"';
const KAI = `<w:rFonts w:hint="eastAsia" ${TNR} w:eastAsia="楷体_GB2312"/><w:sz w:val="28"/><w:szCs w:val="28"/>`;
const BODY = `<w:rFonts w:hint="eastAsia" ${TNR} w:eastAsia="宋体"/><w:sz w:val="21"/><w:szCs w:val="21"/>`;
// 代码：三线表 + Courier New（等宽），无注释（脚本会拦），小五
const MONO = '<w:rFonts w:ascii="Courier New" w:hAnsi="Courier New" w:eastAsia="宋体" w:cs="Courier New"/><w:sz w:val="18"/><w:szCs w:val="18"/>';
// 表题/图题：五号（10.5 pt = sz 21）
const CAP = `<w:rFonts w:hint="eastAsia" ${TNR} w:eastAsia="宋体"/><w:sz w:val="21"/><w:szCs w:val="21"/>`;
const KEEP = '<w:keepNext/><w:keepLines/>';
const SP = '<w:snapToGrid w:val="0"/><w:spacing w:line="360" w:lineRule="auto"/>';
const SINGLE = '<w:snapToGrid w:val="0"/><w:spacing w:line="240" w:lineRule="auto"/>';

const P = (rpr, ppr, text) => `<w:p><w:pPr>${ppr}<w:rPr>${rpr}</w:rPr></w:pPr><w:r><w:rPr>${rpr}</w:rPr><w:t xml:space="preserve">${esc(text)}</w:t></w:r></w:p>`;
// keepNext 只用在真正需要绑定的地方：标题→首段、代码小标题→代码行、图→图题。
// 加在每个条目上会把它们串成长链，整链被推到下一页，反而留出大片空白。
const H1 = t => P(KAI, KEEP + SP + '<w:jc w:val="left"/>', t);
const IT = (lv, t) => P(BODY, SP + `<w:ind w:left="${420 * lv}" w:hanging="420" w:firstLine="0" w:firstLineChars="0"/><w:jc w:val="both"/>`, t);
const ITK = (lv, t) => P(BODY, KEEP + SP + `<w:ind w:left="${420 * lv}" w:hanging="420" w:firstLine="0" w:firstLineChars="0"/><w:jc w:val="both"/>`, t);
// 表题在表上方、图题在图下方，都是五号、居中；表题用 keepNext 绑住后面的表
const TCAPT = t => P(CAP, KEEP + '<w:snapToGrid w:val="0"/><w:spacing w:before="120" w:after="60" w:line="240" w:lineRule="auto"/><w:jc w:val="center"/>', t);
const FIGCAP = t => P(CAP, '<w:snapToGrid w:val="0"/><w:spacing w:before="60" w:after="120" w:line="240" w:lineRule="auto"/><w:jc w:val="center"/>', t);
const BLANK = `<w:p><w:pPr>${SP}</w:pPr></w:p>`;

// ---------- 三线表 ----------
// 只有三条横线：顶线 1.5 pt（sz 12）、栏目线 0.75 pt（sz 6）、底线 1.5 pt；无竖线、无内格线。
const TBLPR = w => `<w:tblPr><w:tblW w:w="${w}" w:type="dxa"/><w:jc w:val="center"/>`
  + '<w:tblBorders><w:top w:val="single" w:sz="12" w:space="0" w:color="000000"/><w:left w:val="nil"/>'
  + '<w:bottom w:val="single" w:sz="12" w:space="0" w:color="000000"/><w:right w:val="nil"/>'
  + '<w:insideH w:val="nil"/><w:insideV w:val="nil"/></w:tblBorders>'
  + '<w:tblLayout w:type="fixed"/><w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/>'
  + '<w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr>';
const HEAD_RULE = '<w:tcBorders><w:bottom w:val="single" w:sz="6" w:space="0" w:color="000000"/></w:tcBorders>';
const TC = (w, xml, extra) => `<w:tc><w:tcPr><w:tcW w:w="${w}" w:type="dxa"/>${extra || ''}<w:vAlign w:val="center"/></w:tcPr>${xml}</w:tc>`;
const TCELL = (rpr, jc, text) => P(rpr, SINGLE + `<w:jc w:val="${jc}"/>`, text === '' || text == null ? ' ' : String(text));

function TBL({ widths, header, rows, font, aligns }) {
  const total = widths.reduce((a, b) => a + b, 0);
  const alignsOf = i => (aligns && aligns[i]) || 'left';
  const row = (cells, head) => '<w:tr><w:trPr><w:cantSplit/>' + (head ? '<w:tblHeader/>' : '') + '</w:trPr>'
    + cells.map((c, i) => TC(widths[i], TCELL(font, alignsOf(i), c), head ? HEAD_RULE : '')).join('') + '</w:tr>';
  const head = header ? row(header, true) : '';
  const body = rows.map(r => row(r, false)).join('');
  return '<w:tbl>' + TBLPR(total) + '<w:tblGrid>' + widths.map(w => `<w:gridCol w:w="${w}"/>`).join('') + '</w:tblGrid>' + head + body + '</w:tbl>';
}

// 代码一律进三线表：行号｜代码，Courier New 小五，行间不断、跨页重复表头。
function CODE_TABLE(lines) {
  const numW = 560, codeW = TEXTW - numW;
  return TBL({
    widths: [numW, codeW],
    header: ['行号', '代码'],
    rows: lines.map((l, i) => [String(i + 1), l]),
    font: MONO,
    aligns: ['right', 'left'],
  });
}
// 代码里出现注释就报错（"不加注释"是硬要求）。只拦开头注释和空白后的 //，避免把 "http://" 误判。
const hasComment = l => /(^|\s)\/\//.test(l) || /\/\*/.test(l);

function PIC(relId, id, name, pxW, pxH, ptW) {
  const cx = Math.round(ptW * 12700), cy = Math.round(ptW * pxH / pxW * 12700);
  return `<w:p><w:pPr>${KEEP}${SP}<w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:noProof/></w:rPr><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0">`
    + `<wp:extent cx="${cx}" cy="${cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:docPr id="${id}" name="${name}"/>`
    + `<wp:cNvGraphicFramePr><a:graphicFrameLocks xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" noChangeAspect="1"/></wp:cNvGraphicFramePr>`
    + `<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">`
    + `<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:nvPicPr><pic:cNvPr id="${id}" name="${name}"/><pic:cNvPicPr/></pic:nvPicPr>`
    + `<pic:blipFill><a:blip r:embed="${relId}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>`
    + `<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>`
    + `</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>`;
}

const FOOTER_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:p><w:pPr><w:jc w:val="center"/><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:eastAsia="宋体"/><w:sz w:val="18"/><w:szCs w:val="18"/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:eastAsia="宋体"/><w:sz w:val="18"/><w:szCs w:val="18"/></w:rPr><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="18"/></w:rPr><w:instrText xml:space="preserve"> PAGE  \\* MERGEFORMAT </w:instrText></w:r><w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:eastAsia="宋体"/><w:sz w:val="18"/><w:szCs w:val="18"/></w:rPr><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="18"/></w:rPr><w:t>1</w:t></w:r><w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:eastAsia="宋体"/><w:sz w:val="18"/><w:szCs w:val="18"/></w:rPr><w:fldChar w:fldCharType="end"/></w:r></w:p></w:ftr>`;

// ---------- 编号自动生成 ----------
const CN = ['一', '二', '三', '四', '五', '六', '七', '八', '九', '十',
  '十一', '十二', '十三', '十四', '十五', '十六', '十七', '十八', '十九', '二十'];
const CIRC = n => (n >= 1 && n <= 20) ? String.fromCharCode(0x2460 + n - 1) : null;
function prefix(lv, n) {
  if (lv === 1) return CN[n - 1] + '、';
  if (lv === 2) return n + '. ';
  if (lv === 3) return '（' + n + '）';
  if (lv === 4) { const c = CIRC(n); if (!c) throw new Error('四级编号超过 ⑳，请拆父项'); return c + ' '; }
  throw new Error('非法层级 ' + lv + '（只允许 1-4）');
}

// ---------- 流程图：交给 flowchart.mjs（DOT -> ortho SVG -> PNG，并做无曲线审计） ----------
const FLOW_SCRIPT = path.join(__dirname, 'flowchart.mjs');
function makeFlowchart(fc, dir, n) {
  let dotPath = fc.dot ? fromSpec(fc.dot) : null;
  if (!dotPath && fc.dotText) {
    dotPath = path.join(dir, `fig${n}.dot`);
    fs.writeFileSync(dotPath, fc.dotText, 'utf8');
  }
  if (!dotPath || !fs.existsSync(dotPath)) throw new Error('flowchart 需要 dot 或 dotText，且文件要存在: ' + dotPath);
  const png = path.join(dir, `fig${n}.png`);
  const svg = path.join(dir, `fig${n}.svg`);
  const out = cp.execSync(`node "${FLOW_SCRIPT}" -Dot "${dotPath}" -Png "${png}" -Svg "${svg}"` +
    (fc.allowCurves ? ' -AllowCurves' : ''), { encoding: 'utf8' });
  console.log(out.trim());
  if (fc.keepDot) fs.copyFileSync(dotPath, path.join(path.dirname(OUT), path.basename(dotPath)));
  return png;
}

function pngSize(p) {
  const b = fs.readFileSync(p);
  if (b.readUInt32BE(0) !== 0x89504e47) throw new Error('不是 PNG: ' + p);
  return { w: b.readUInt32BE(16), h: b.readUInt32BE(20), bytes: b.length };
}

// ---------- 解包模板（中间产物放独立目录，绝不能留在包目录里） ----------
const PS = 'pwsh -NoProfile -Command';
const stage = fs.mkdtempSync(path.join(os.tmpdir(), 'labrep-'));
const figdir = fs.mkdtempSync(path.join(os.tmpdir(), 'labrep-fig-'));
cp.execSync(`${PS} "Copy-Item -LiteralPath '${path.join(SKILL, tpl.file)}' '${path.join(stage, 't.zip')}'; ` +
  `Expand-Archive -Path '${path.join(stage, 't.zip')}' -DestinationPath '${stage}' -Force; ` +
  `Remove-Item -LiteralPath '${path.join(stage, 't.zip')}' -Force"`);

let doc = fs.readFileSync(path.join(stage, 'word/document.xml'), 'utf8');

// 版心宽：从 sectPr 的 pgSz 与 pgMar 反算，三线表按它定宽
const TEXTW = (() => {
  const pg = /<w:pgSz[^>]*w:w="(\d+)"/.exec(doc);
  const mar = /<w:pgMar[^>]*\/>/.exec(doc);
  const pick = (tag, def) => { const m = mar && new RegExp('w:' + tag + '="(\\d+)"').exec(mar[0]); return m ? +m[1] : def; };
  return pg ? (+pg[1] - pick('left', 1800) - pick('right', 1800)) : 8306;
})();

// ---------- 封面：按标签找同行空单元格填入 ----------
function fillCover(xml, label, value) {
  const rows = xml.split('<w:tr ');
  for (let r = 1; r < rows.length; r++) {
    const cells = rows[r].split('<w:tc>');
    let hit = -1;
    for (let c = 1; c < cells.length; c++) {
      const txt = [...cells[c].matchAll(/<w:t[^>]*>([^<]*)<\/w:t>/g)].map(m => m[1]).join('').trim();
      if (txt === label || txt === label + '：' || txt === label + ':') { hit = c; break; }
    }
    if (hit < 0) continue;
    for (let c = hit + 1; c < cells.length; c++) {
      if (!/<w:t[ >]/.test(cells[c].split('</w:tc>')[0])) {
        const run = `<w:r><w:rPr><w:rFonts w:hint="eastAsia" w:asciiTheme="minorEastAsia" w:hAnsiTheme="minorEastAsia" w:eastAsiaTheme="minorEastAsia"/></w:rPr><w:t xml:space="preserve">${esc(value)}</w:t></w:r>`;
        const seg = cells[c];
        const j = seg.indexOf('</w:p>');
        if (j < 0) throw new Error('目标单元格没有段落: ' + label);
        cells[c] = seg.slice(0, j) + run + seg.slice(j);
        rows[r] = cells.join('<w:tc>');
        return rows.join('<w:tr ');
      }
    }
    throw new Error('「' + label + '」同一行没有空值单元格');
  }
  throw new Error('表格里找不到标签「' + label + '」');
}
for (const [k, v] of Object.entries(spec.cover || {})) doc = fillCover(doc, k, v);

// ---------- 列宽修正（模板已知缺陷） ----------
if (tpl.gridFix) {
  const g = tpl.gridFix;
  if (g.reduce((a, b) => a + b, 0) !== 8296) throw new Error('gridFix 合计必须等于 8296');
  doc = doc.replace(/<w:tblGrid>[\s\S]*?<\/w:tblGrid>/, '<w:tblGrid>' + g.map(w => `<w:gridCol w:w="${w}"/>`).join('') + '</w:tblGrid>');
  doc = doc.split('<w:tr ').map((row, n) => {
    if (n === 0) return row;
    const cnt = (row.match(/<w:tcW w:w="\d+" w:type="dxa"\/>/g) || []).length;
    const target = cnt === 6 ? g : [1271, 7025];
    let k = 0;
    return row.replace(/<w:tcW w:w="\d+" w:type="dxa"\/>/g, () => {
      const w = target[k] !== undefined ? target[k] : target[target.length - 1]; k++;
      return `<w:tcW w:w="${w}" w:type="dxa"/>`;
    });
  }).join('<w:tr ');
}

// ---------- 正文 ----------
const media = [];
const B = [BLANK];
const cnt = [0, 0, 0, 0];
let figNo = 1, tblNo = 1, docPrId = 900;
for (const el of spec.body) {
  if (el.blank) { B.push(BLANK); continue; }
  if (el.h1 !== undefined) { cnt[1]++; cnt[2] = cnt[3] = cnt[4] = 0; B.push(H1(prefix(1, cnt[1]) + el.h1)); continue; }
  if (el.para !== undefined) { B.push(P(BODY, SP + '<w:ind w:firstLine="420" w:firstLineChars="200"/><w:jc w:val="both"/>', el.para)); continue; }
  if (el.item) {
    const [lv, text] = el.item;
    cnt[lv]++; for (let d = lv + 1; d <= 4; d++) cnt[d] = 0;
    // 该条目若还挂着更深层子项，就必须和首子项同页，否则会出现孤立标题
    const next = spec.body[spec.body.indexOf(el) + 1];
    const hasChild = next && next.item && next.item[0] > lv;
    B.push((hasChild ? ITK : IT)(lv, prefix(lv, cnt[lv]) + text));
    continue;
  }
  if (el.code) {
    if (el.label) { cnt[3]++; cnt[4] = 0; B.push(ITK(3, prefix(3, cnt[3]) + el.label)); }
    el.code.forEach(l => { if (hasComment(l)) throw new Error('代码不要写注释，请先在 spec 里删掉这一行: ' + l.trim()); });
    B.push(TCAPT('表' + tblNo + ' ' + (el.caption || '核心代码')));
    B.push(CODE_TABLE(el.code));
    tblNo++;
    continue;
  }
  if (el.table) {
    const t = el.table;
    const n = (t.header || t.rows[0]).length;
    const widths = t.widths || Array(n).fill(Math.floor(TEXTW / n));
    if (widths.length !== n) throw new Error('table 的 widths 个数必须等于列数 ' + n);
    if (t.caption) B.push(TCAPT('表' + tblNo + ' ' + t.caption));
    B.push(TBL({ widths, header: t.header, rows: t.rows, font: BODY, aligns: t.align }));
    tblNo++;
    continue;
  }
  if (el.image || el.flowchart) {
    const cfg = el.flowchart || el.image;
    const src = el.flowchart ? makeFlowchart(cfg, figdir, figNo) : fromSpec(cfg.src);
    const widthPt = cfg.widthPt || 300;
    if (!fs.existsSync(src)) throw new Error('图片不存在: ' + src);
    const { w, h } = pngSize(src);
    const relId = 'rId' + (101 + media.length);
    const name = `image${media.length + 1}.png`;
    media.push({ relId, name, src });
    B.push(PIC(relId, ++docPrId, name, w, h, widthPt));
    B.push(FIGCAP(cfg.caption.replace(/^图\d+/, '图' + figNo)));
    figNo++;
    continue;
  }
  throw new Error('spec.body 里有不认识的元素: ' + JSON.stringify(el).slice(0, 80));
}

const tblEnd = doc.indexOf('</w:tbl>') + 8;
const sectAt = doc.indexOf('<w:sectPr');
if (sectAt < tblEnd || !doc.includes('</w:body>')) throw new Error('模板结构异常（找不到表格结尾或 sectPr）');
doc = doc.slice(0, tblEnd) + B.join('') + doc.slice(sectAt);

// ---------- 页脚（模板没有时补上） ----------
const rels = [];
if (tpl.needFooter && !doc.includes('<w:footerReference')) {
  fs.writeFileSync(path.join(stage, 'word/footer1.xml'), FOOTER_XML, 'utf8');
  rels.push(['rId100', 'footer', 'footer1.xml']);
  doc = doc.replace('<w:sectPr>', '<w:sectPr><w:footerReference w:type="default" r:id="rId100"/>');
  let ct = fs.readFileSync(path.join(stage, '[Content_Types].xml'), 'utf8');
  if (!ct.includes('footer1.xml')) {
    ct = ct.replace('</Types>', '<Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/></Types>');
    fs.writeFileSync(path.join(stage, '[Content_Types].xml'), ct, 'utf8');
  }
}

// ---------- 图片部件 ----------
if (media.length) {
  fs.mkdirSync(path.join(stage, 'word/media'), { recursive: true });
  for (const m of media) fs.copyFileSync(m.src, path.join(stage, 'word/media', m.name));
  let ct = fs.readFileSync(path.join(stage, '[Content_Types].xml'), 'utf8');
  if (!ct.includes('Extension="png"')) {
    ct = ct.replace('</Types>', '<Default Extension="png" ContentType="image/png"/></Types>');
    fs.writeFileSync(path.join(stage, '[Content_Types].xml'), ct, 'utf8');
  }
  for (const m of media) rels.push([m.relId, 'image', 'media/' + m.name]);
}
if (rels.length) {
  const rp = path.join(stage, 'word/_rels/document.xml.rels');
  let rel = fs.readFileSync(rp, 'utf8');
  for (const [id, kind, target] of rels) {
    if (!rel.includes(`Id="${id}"`)) {
      rel = rel.replace('</Relationships>', `<Relationship Id="${id}" Type="${RT}/${kind}" Target="${target}"/></Relationships>`);
    }
  }
  fs.writeFileSync(rp, rel, 'utf8');
}

// ---------- 命名空间 ----------
const open = doc.indexOf('<w:document'), gt = doc.indexOf('>', open);
let tag = doc.slice(open, gt);
if (!tag.includes('xmlns:a=')) tag += ' xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"';
if (!tag.includes('xmlns:pic=')) tag += ' xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"';
doc = doc.slice(0, open) + tag + doc.slice(gt);
fs.writeFileSync(path.join(stage, 'word/document.xml'), doc, 'utf8');

// ---------- 打包 + 校验 ----------
const pack = path.join(SKILL, 'scripts/pack-docx.ps1');
console.log(cp.execSync(`pwsh -NoProfile -File "${pack}" -SrcDir "${stage}" -OutFile "${OUT}"`,
  { encoding: 'utf8', stdio: ['ignore', 'pipe', 'pipe'] }).trim());
console.log('spec 元素:', spec.body.length, '| 图片:', media.length, '| 一级标题:', cnt[1], '| 输出:', OUT);
fs.rmSync(stage, { recursive: true, force: true });
fs.rmSync(figdir, { recursive: true, force: true });
