/* 论文格式助手 —— Node 环境回归测试
 * 用真实的毕业设计 docx 验证：结构识别 → 格式改写 → 重新打包 → XML 合法性
 * 运行：npm test
 */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { DOMParser, XMLSerializer } from '@xmldom/xmldom';
import JSZip from 'jszip';

globalThis.DOMParser = DOMParser;
globalThis.XMLSerializer = XMLSerializer;
globalThis.JSZip = JSZip; // 浏览器中 JSZip 是全局变量，Node 里需注入

const FormatTool = (await import('../js/formatter.js')).default;
const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const ROOT = path.resolve(__dirname, '../..');
/* 夹具查找：先扫根目录里版本号最大的 CYX_<n>.<n>.docx，扫不到才回退历史名单。
   顺序不能反 —— 作者每改一版论文文件名就跟着变（CYX_20.0 → CYX_21.0），死名单
   排在前面的话新版本永远选不中。2026-09-26 实测：名单首位的 CYX_21.0.docx 是
   9-08 的旧件，当天产出的 CYX_21.1.docx 一次都没被跑过 —— 原先那段"防止改名后
   静悄悄不跑"的注释（2026-09-16 写下）因为顺序反了，从写下那天起就没生效。 */
function findFixture(names) {
  const ver = (n) => { const m = /^CYX_(\d+)\.(\d+)\.docx$/.exec(n); return m ? +m[1] * 1000 + +m[2] : -1; };
  const best = fs.readdirSync(ROOT).filter((n) => ver(n) >= 0).sort((a, b) => ver(b) - ver(a))[0];
  if (best) return path.join(ROOT, best);
  const hit = names.map((n) => path.join(ROOT, n)).find((f) => fs.existsSync(f));
  return hit || null;
}
const src = findFixture(['CYX_21.0.docx', 'CYX_20.0.docx', 'CYX_14.0.docx', 'CYX_12.0.docx', 'CYX_11.0.docx', 'CYX_10.0.docx', 'CYX_8.0.docx', 'CYX_6.0.docx', 'CYX_5.0.docx', 'CYX_4.0.docx', 'CYX_3.0.docx', 'CYX_毕业设计.docx', 'CYX2.0.docx']);
const outFile = path.join(__dirname, 'output_格式化.docx');

if (!src) {
  console.error('未找到测试文档，请将论文放在项目根目录下（如 CYX_3.0.docx / CYX_毕业设计.docx）');
  process.exit(1);
}

const buf = fs.readFileSync(src);
console.log(`输入: ${path.basename(src)}  (${(buf.length / 1024 / 1024).toFixed(2)} MB)`);

const t0 = Date.now();
const result = await FormatTool.formatDocx(buf, FormatTool.DEFAULTS, { format: 'nodebuffer' });
console.log(`耗时: ${Date.now() - t0} ms\n`);

console.log('===== 结构识别结果 =====');
const c = result.counts;
console.log(`封面/前部段落: ${c.front}（保持原样不动）`);
console.log(`论文题目: ${c.title} | 英文题目: ${c.enTitle}`);
console.log(`摘要: 中文标题 ${c.absHead} / 正文 ${c.absBody} / 关键词 ${c.kw}`);
console.log(`英文摘要: 标题 ${c.absEnHead} / 正文 ${c.absEnBody} / Keywords ${c.kwEn}`);
console.log(`目录: 标题 ${c.tocHead} / 条目 ${c.tocItem}`);
console.log(`一级标题: ${c.h1} | 二级标题: ${c.h2} | 三级标题: ${c.h3}`);
console.log(`正文段落: ${c.body}`);
console.log(`图表题注: ${c.caption} | 公式段落: ${c.formula} | 纯图片段: ${c.figure}`);
console.log(`参考文献: 标题 ${c.refHead} / 条目 ${c.refItem}`);
console.log(`表格: ${c.tables} | 图片: ${c.images} | 公式(内联+段): ${c.formulas}`);
console.log(`页脚页码: ${c.footerAdded ? '已添加' : '未添加（原文档已有或已禁用）'}`);
console.log(`识别标志: ${JSON.stringify(result.info)}`);

// 真实文档目录识别：目录标题必须被识别；条目可能为 0（输入目录已是 sdt 占位域，
// 条目由格式化时预填充生成，见"自动目录域"校验）
if (!result.info.tocDetected || c.tocHead < 1) {
  throw new Error(`真实文档目录未被识别（tocDetected=${result.info.tocDetected} tocHead=${c.tocHead} tocItem=${c.tocItem}）`);
}
console.log(`✓ 真实文档目录识别: 标题 ${c.tocHead} | 条目 ${c.tocItem}（条目由自动目录域预填充）`);

/* ---------- 重新解包校验 ---------- */
console.log('\n===== 输出校验 =====');
const zip = await JSZip.loadAsync(result.data);
fs.writeFileSync(outFile, result.data);
console.log(`已写出: ${outFile} (${(result.data.length / 1024 / 1024).toFixed(2)} MB)`);

const docXml = await zip.file('word/document.xml').async('string');
const doc = new DOMParser().parseFromString(docXml, 'application/xml'); // 解析失败会抛异常
if (doc.documentElement.localName !== 'document') throw new Error('document.xml 根元素异常');
console.log('✓ document.xml 解析通过');

// 校验 sectPr 子元素顺序
const SECTPR_ORDER = ['headerReference','footerReference','footnotePr','endnotePr','type','pgSz','pgMar',
  'paperSrc','pgBorders','lnNumType','pgNumType','cols','formProt','vAlign','noEndnote','titlePg',
  'textDirection','bidi','rtlGutter','docGrid','printerSettings','sectPrChange'];
const sectPrs = doc.getElementsByTagNameNS(W_NS, 'sectPr');
let orderOk = true;
for (const sp of Array.from(sectPrs)) {
  let last = -1;
  for (const ch of Array.from(sp.childNodes).filter(n => n.nodeType === 1)) {
    // headerReference/footerReference 是 schema 选择组，可交替出现，视作同一索引
    let idx = SECTPR_ORDER.indexOf(ch.localName);
    if (idx === 0 || idx === 1) idx = 0;
    if (idx < last) { orderOk = false; console.error(`✗ sectPr 子元素顺序错误: ${ch.localName}`); }
    if (idx >= 0) last = idx;
  }
}
console.log(orderOk ? `✓ ${sectPrs.length} 个分节 pgSz/pgMar/docGrid 顺序正确` : '✗ sectPr 顺序错误');
/* sectPr 子元素顺序错了 Word 会直接判「文档已损坏」——原先这里只打印 ✗ 不抛错，
   跑完照样退出码 0 打印「全部校验通过」，等于没查 */
if (!orderOk) throw new Error('sectPr 子元素顺序错误（Word 会判文档损坏）');

// 抽样检查正文段落样式
const paras = doc.getElementsByTagNameNS(W_NS, 'p');
let withSpacing = 0, withIndent = 0, withJc = 0;
const sample = [];
for (const p of Array.from(paras)) {
  const pPr = Array.from(p.childNodes).find(n => n.nodeType === 1 && n.localName === 'pPr');
  if (!pPr) continue;
  if (pPr.getElementsByTagNameNS(W_NS, 'spacing').length) withSpacing++;
  if (pPr.getElementsByTagNameNS(W_NS, 'ind').length) withIndent++;
  if (pPr.getElementsByTagNameNS(W_NS, 'jc').length) withJc++;
  if (sample.length < 3) sample.push(pPr);
}
console.log(`✓ 段落含 spacing: ${withSpacing} | ind: ${withIndent} | jc: ${withJc}（共 ${paras.length} 段）`);

// 校验 rPr 子元素顺序
const RPR_ORDER = ['rStyle','rFonts','b','bCs','i','iCs','caps','smallCaps','strike','dstrike','outline',
  'shadow','emboss','imprint','noProof','snapToGrid','vanish','webHidden','color','spacing','w','kern',
  'position','sz','szCs','highlight','u','effect','bdr','shd','fitText','vertAlign','rtl','cs','em',
  'lang','textFill','eastAsianLayout','specVanish','oMath'];
let rprOk = true, rprChecked = 0;
const rPrs = doc.getElementsByTagNameNS(W_NS, 'rPr');
for (const rPr of Array.from(rPrs).slice(0, 500)) {
  rprChecked++;
  let last = -1;
  for (const ch of Array.from(rPr.childNodes).filter(n => n.nodeType === 1)) {
    const idx = RPR_ORDER.indexOf(ch.localName);
    if (idx < last) { rprOk = false; console.error(`✗ rPr 子元素顺序错误: ${ch.localName}`); }
    if (idx >= 0) last = idx;
  }
}
console.log(rprOk ? `✓ 抽查 ${rprChecked} 个 rPr 元素顺序正确` : '✗ rPr 顺序错误');
if (!rprOk) throw new Error('rPr 子元素顺序错误（Word 会判文档损坏）');

// 页脚校验
if (c.footerAdded) {
  const ct = await zip.file('[Content_Types].xml').async('string');
  const rels = await zip.file('word/_rels/document.xml.rels').async('string');
  const hasCT = ct.includes('footer+xml');
  const hasRel = rels.includes('relationships/footer');
  const footers = Object.keys(zip.files).filter(f => /^word\/footer\d+\.xml$/.test(f));
  /* 原先这里无论有没有都打印 ✓（✓ 是写死的），缺部件也看不出来 */
  console.log(`${hasCT && hasRel && footers.length ? '✓' : '✗'} 页脚: ContentTypes ${hasCT ? 'OK' : 'MISSING'}` +
    ` | rels ${hasRel ? 'OK' : 'MISSING'} | 部件 ${footers.length ? footers.join(', ') : '无'}`);
  if (!hasCT || !hasRel || !footers.length) {
    throw new Error('声明了添加页脚却没有页脚部件，或缺 ContentTypes 声明 / rels 关联（正文会整篇没有页码）');
  }
  for (const f of footers) {
    new DOMParser().parseFromString(await zip.file(f).async('string'), 'application/xml');
  }
  console.log('✓ 页脚部件解析通过');
}

// 页边距校验（附件8：上下左右 2.5cm = 1418 twips，页眉 1.6cm ≈ 907，页脚 2.1cm ≈ 1191）
const mar = Array.from(sectPrs[0]?.childNodes || []).find(n => n.nodeType === 1 && n.localName === 'pgMar');
if (mar) {
  const top = mar.getAttributeNS(W_NS, 'top');
  const hdr = mar.getAttributeNS(W_NS, 'header');
  const ftr = mar.getAttributeNS(W_NS, 'footer');
  /* 原先只把数值念一遍、连比较都没有，括号里的期望值纯装饰 */
  const marOk = top === '1418' && hdr === '907' && ftr === '1191';
  console.log(`${marOk ? '✓' : '✗'} 页边距 top=${top} (应 1418) | 页眉距=${hdr} (应 907) | 页脚距=${ftr} (应 1191)`);
  if (!marOk) throw new Error(`页边距不符合附件8（top=${top} header=${hdr} footer=${ftr}）`);
} else {
  console.error('✗ 分节里找不到 w:pgMar，页边距未设置');
  throw new Error('分节缺少 w:pgMar');
}

/* ---------- 附件8 新规则校验 ---------- */
// 1. 固定20磅行距（line=400, lineRule=exact）
const spacings = doc.getElementsByTagNameNS(W_NS, 'spacing');
let exact400 = 0;
for (const sp of Array.from(spacings)) {
  if (sp.getAttributeNS(W_NS, 'line') === '400' && sp.getAttributeNS(W_NS, 'lineRule') === 'exact') exact400++;
}
console.log(`✓ 固定20磅行距(line=400 exact)段落: ${exact400}`);

// 2. 一级标题换页（pageBreakBefore）
const pbs = doc.getElementsByTagNameNS(W_NS, 'pageBreakBefore');
console.log(`✓ 一级标题换页标记: ${pbs.length} 处`);

/* 前置部分工具：摘要（第一个非 front 段落）之前的 body 直系元素不参与格式修改。
   放在这里是为了让下面的三线表校验也能用它 —— 分母必须是**正文表**（排除封面表），
   c.tables 是全量（含封面表），拿它当分母口径就错了。 */
const bodyKids0 = Array.from(doc.getElementsByTagNameNS(W_NS, 'body')[0].childNodes).filter(n => n.nodeType === 1);
const absParaIdx = bodyKids0.findIndex(e => e.localName === 'p' && /^摘\s*要/.test(
  Array.from(e.getElementsByTagNameNS(W_NS, 't')).map(x => x.textContent).join('').trim()));
const frontCount = absParaIdx > 0 ? absParaIdx : 0;
const isFront = (el) => {
  if (frontCount <= 0) return false;
  while (el) {
    const p = el.parentNode;
    if (p && p.nodeType === 1 && p.localName === 'body') return bodyKids0.indexOf(el) < frontCount;
    el = p;
  }
  return false;
};

// 3. 三线表（tblBorders top sz=12, bottom sz=12）。封面表受前置保护、本就保留原边框，
//    所以分母只能算正文表（与下面第 8 条 tblCount 同口径）
const tblBorders = doc.getElementsByTagNameNS(W_NS, 'tblBorders');
let threeLine = 0;
for (const tb of Array.from(tblBorders)) {
  const top = tb.getElementsByTagNameNS(W_NS, 'top')[0];
  const bottom = tb.getElementsByTagNameNS(W_NS, 'bottom')[0];
  const insideV = tb.getElementsByTagNameNS(W_NS, 'insideV')[0];
  if (top && bottom && insideV &&
      top.getAttributeNS(W_NS, 'sz') === '12' && bottom.getAttributeNS(W_NS, 'sz') === '12' &&
      insideV.getAttributeNS(W_NS, 'val') === 'none') threeLine++;
}
const bodyTblCount = Array.from(doc.getElementsByTagNameNS(W_NS, 'tbl')).filter(t => !isFront(t)).length;
console.log(`${threeLine === bodyTblCount ? '✓' : '✗'} 表格顶底线 1.5磅(sz=12): ${threeLine}/${bodyTblCount} 正文表` +
  `（全文 ${c.tables} 表，含封面表 ${c.tables - bodyTblCount}）`);
if (threeLine !== bodyTblCount) {
  throw new Error(`有正文表格未套三线表（${threeLine}/${bodyTblCount}）`);
}

// 4. 分节页码（罗马/阿拉伯）
const pgNumTypes = doc.getElementsByTagNameNS(W_NS, 'pgNumType');
let roman = 0, decimal = 0;
for (const pn of Array.from(pgNumTypes)) {
  const fmt = pn.getAttributeNS(W_NS, 'fmt');
  if (fmt === 'upperRoman') roman++;
  if (fmt === 'decimal') decimal++;
}
console.log(`✓ 页码分节: 罗马数字 ${roman} 节 | 阿拉伯数字 ${decimal} 节`);

// 4b. 页码连续：仅绪论所在分节从 1 编页（start=1 恰好 1 处），章节之间连续不重启
let decimalStart = 0;
for (const pn of Array.from(pgNumTypes)) {
  if (pn.getAttributeNS(W_NS, 'fmt') === 'decimal' && pn.getAttributeNS(W_NS, 'start') === '1') decimalStart++;
}
console.log(`✓ 页码连续: decimal start=1 共 ${decimalStart} 处（应为 1，章节间不重新编页）`);
if (decimalStart !== 1) throw new Error('页码未连续编页（start=1 应为 1 处）');

// 5. 标题字号（二级小三=30半磅、三级四号=28半磅）
let sz30 = 0, sz28 = 0;
for (const rPr of Array.from(rPrs)) {
  const sz = rPr.getElementsByTagNameNS(W_NS, 'sz')[0];
  if (!sz) continue;
  const v = sz.getAttributeNS(W_NS, 'val');
  if (v === '30') sz30++;
  if (v === '28') sz28++;
}
console.log(`✓ 小三(30) run: ${sz30} | 四号(28) run: ${sz28}`);

// 6. 二级标题固定20磅行距（附件8：全文行距固定20磅）
const RE_H2_OUT = /^\d{1,2}\.\d{1,2}(?![\d.])\s*[^\d.]/;
let h2n = 0, h2ok = 0;
for (const p of Array.from(doc.getElementsByTagNameNS(W_NS, 'body')[0].childNodes).filter(n => n.nodeType === 1 && n.localName === 'p')) {
  const t = Array.from(p.getElementsByTagNameNS(W_NS, 't')).map(x => x.textContent).join('').trim();
  if (!RE_H2_OUT.test(t)) continue;
  if (/\d\s*$/.test(t)) continue;   // 排除目录条目（标题+页码）
  if (t.length > 20) continue;
  h2n++;
  const pPr = Array.from(p.childNodes).find(n => n.nodeType === 1 && n.localName === 'pPr');
  const sp = pPr && Array.from(pPr.childNodes).find(n => n.nodeType === 1 && n.localName === 'spacing');
  if (sp && sp.getAttributeNS(W_NS, 'line') === '400' && sp.getAttributeNS(W_NS, 'lineRule') === 'exact') h2ok++;
}
console.log(`✓ 二级标题固定20磅行距: ${h2ok}/${h2n} 段（line=400 exact）`);
if (h2ok !== h2n) throw new Error('二级标题未改为固定 20 磅行距');

/* ---------- v1.2 新规则校验 ---------- */

// 7. 图片段落单倍行距：正文含 w:drawing / w:pict 的段落 spacing 必须为 line=240 lineRule=auto
//    （前置封面图片按"前两页不修改"保留原样）
let imgParas = 0, imgParasOk = 0;
for (const p of Array.from(doc.getElementsByTagNameNS(W_NS, 'p'))) {
  if (!p.getElementsByTagNameNS(W_NS, 'drawing').length && !p.getElementsByTagNameNS(W_NS, 'pict').length) continue;
  if (isFront(p)) continue;
  imgParas++;
  const pPr = Array.from(p.childNodes).find(n => n.nodeType === 1 && n.localName === 'pPr');
  if (!pPr) continue;
  const sp = Array.from(pPr.childNodes).find(n => n.nodeType === 1 && n.localName === 'spacing');
  if (sp && sp.getAttributeNS(W_NS, 'line') === '240' && sp.getAttributeNS(W_NS, 'lineRule') === 'auto') imgParasOk++;
}
console.log(`✓ 图片段落单倍行距: ${imgParasOk}/${imgParas} 段（line=240 auto，前置封面图片除外）`);
if (imgParasOk !== imgParas) throw new Error(`图片段落未全部改为单倍行距（${imgParasOk}/${imgParas}）`);

// 8. 表格三线表（参考样张）：表头行单元格顶线 single/12(1.5磅)+栏目线 single/6(0.75磅)，
//    末行单元格底线 single/12(1.5磅)，其余单元格无 tcBorders；前置封面表格保留原样
const directChildren = (elm, name) =>
  Array.from(elm.childNodes).filter(n => n.nodeType === 1 && n.localName === name);
const cellBorder = (tc, side) => {
  const tcPr = directChildren(tc, 'tcPr')[0];
  const tcb = tcPr && directChildren(tcPr, 'tcBorders')[0];
  const b = tcb && directChildren(tcb, side)[0];
  return b ? b.getAttributeNS(W_NS, 'val') + '/' + b.getAttributeNS(W_NS, 'sz') : null;
};
let tblOk = 0, tblCount = 0, hdrCellCount = 0, hdrOk = 0, lastOk = 0, lastCellCount = 0, midBorders = 0;
for (const tbl of Array.from(doc.getElementsByTagNameNS(W_NS, 'tbl'))) {
  if (isFront(tbl)) continue;
  tblCount++;
  const trs = directChildren(tbl, 'tr');
  if (!trs.length) continue;
  const hdrCells = directChildren(trs[0], 'tc');
  const lastCells = trs.length > 1 ? directChildren(trs[trs.length - 1], 'tc') : [];
  let hdrGood = true;
  for (const tc of hdrCells) {
    hdrCellCount++;
    const topB = cellBorder(tc, 'top'), botB = cellBorder(tc, 'bottom');
    if (topB === 'single/12' && botB === 'single/6') hdrOk++;
    else hdrGood = false;
  }
  if (lastCells.length) {
    for (const tc of lastCells) {
      lastCellCount++;
      if (cellBorder(tc, 'bottom') === 'single/12') lastOk++;
    }
  }
  for (const tr of trs.slice(1, trs.length > 1 ? trs.length - 1 : 1)) {
    for (const tc of directChildren(tr, 'tc')) {
      const tcPr = directChildren(tc, 'tcPr')[0];
      if (tcPr && directChildren(tcPr, 'tcBorders').length) midBorders++;
    }
  }
  if (hdrGood && lastOk === lastCellCount) tblOk++;
}
console.log(`✓ 表格三线表: ${tblOk}/${tblCount} 表 | 表头顶线1.5+栏目线0.75 ${hdrOk}/${hdrCellCount} 格 | 末行底线1.5 ${lastOk}/${lastCellCount} 格 | 中间行残留边框 ${midBorders} 处`);
if (tblOk !== tblCount || hdrOk !== hdrCellCount || lastOk !== lastCellCount || midBorders > 0) {
  throw new Error('表格不符合三线表要求（顶/底 1.5磅、栏目线 0.75磅、其余无线）');
}

// 8b. 表格行不得跨页断裂：每行 trPr 含 cantSplit；图片段带 keepNext+keepLines（与题注同页）且居中
let csTotal = 0, csOk = 0, keepImgOk = 0, keepImgTotal = 0;
for (const tbl of Array.from(doc.getElementsByTagNameNS(W_NS, 'tbl'))) {
  if (isFront(tbl)) continue;
  for (const tr of Array.from(tbl.childNodes).filter(n => n.nodeType === 1 && n.localName === 'tr')) {
    csTotal++;
    const trPr = directChildren(tr, 'trPr')[0];
    if (trPr && directChildren(trPr, 'cantSplit').length) csOk++;
  }
}
for (const p of Array.from(doc.getElementsByTagNameNS(W_NS, 'p'))) {
  if (!p.getElementsByTagNameNS(W_NS, 'drawing').length && !p.getElementsByTagNameNS(W_NS, 'pict').length) continue;
  if (isFront(p)) continue;
  keepImgTotal++;
  const pPr = Array.from(p.childNodes).find(n => n.nodeType === 1 && n.localName === 'pPr');
  const jc = pPr && Array.from(pPr.childNodes).find(n => n.nodeType === 1 && n.localName === 'jc');
  if (pPr && pPr.getElementsByTagNameNS(W_NS, 'keepNext').length && pPr.getElementsByTagNameNS(W_NS, 'keepLines').length &&
      jc && jc.getAttributeNS(W_NS, 'val') === 'center') keepImgOk++;
}
console.log(`✓ 表格行 cantSplit: ${csOk}/${csTotal} 行 | 图片段居中+keepNext+keepLines: ${keepImgOk}/${keepImgTotal} 段`);
if (csOk !== csTotal || keepImgOk !== keepImgTotal) throw new Error('表格行未防跨页断裂或图片段未居中/未与题注同页');

// 9. 空白页回归：一级标题前已有强制分页（nextPage 分节符 / 段尾分页 run / 空段分页 /
//    pageBreakBefore）时，不得再叠加 pageBreakBefore（双重分页会产生空白页）；
//    反之必须叠加（保证章节从新页开始）。
//    h1 集合复用引擎的 classifyParas 识别结果；段落只取 body 直系子元素
//    （与 collectParas 一致），排除 sdt/目录等嵌套内容
const M_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/math';
const body9 = doc.getElementsByTagNameNS(W_NS, 'body')[0];
const allParas = Array.from(body9.childNodes).filter(n => n.nodeType === 1 && n.localName === 'p');
const paraMetaOf = (p) => {
  const meta = { style: '', hasLink: false, outline: null, hasTab: false };
  const pPr = Array.from(p.childNodes).find(n => n.nodeType === 1 && n.localName === 'pPr');
  if (pPr) {
    const ps = Array.from(pPr.childNodes).find(n => n.nodeType === 1 && n.localName === 'pStyle');
    if (ps) meta.style = (ps.getAttributeNS(W_NS, 'val') || '').toLowerCase();
    const ol = Array.from(pPr.childNodes).find(n => n.nodeType === 1 && n.localName === 'outlineLvl');
    if (ol) meta.outline = ol.getAttributeNS(W_NS, 'val');
  }
  if (Array.from(p.childNodes).some(n => n.nodeType === 1 && n.localName === 'hyperlink')) meta.hasLink = true;
  if (p.getElementsByTagNameNS(W_NS, 'tab').length > 0) meta.hasTab = true;
  return meta;
};
const items = [];
const idxMap = [];
allParas.forEach((p, i) => {
  const text = Array.from(p.getElementsByTagNameNS(W_NS, 't')).map(x => x.textContent).join('').trim();
  const hasMath = p.getElementsByTagNameNS(M_NS, 'oMath').length > 0;
  const hasDrawing = p.getElementsByTagNameNS(W_NS, 'drawing').length > 0 ||
                     p.getElementsByTagNameNS(W_NS, 'pict').length > 0;
  if (!text && !hasDrawing) return; // 与 collectParas 一致：纯空段跳过
  idxMap.push(i);
  items.push({ p, text, hasMath, hasDrawing, hasPageBreak: false, meta: paraMetaOf(p) });
});
const cls9 = FormatTool.classifyParas(items);
const h1Idxs = [];
cls9.roles.forEach((r, k) => { if (r === 'h1') h1Idxs.push(idxMap[k]); });

const paraHasContent = (p) => p.getElementsByTagNameNS(W_NS, 't').length > 0 ||
  p.getElementsByTagNameNS(W_NS, 'drawing').length > 0 || p.getElementsByTagNameNS(W_NS, 'pict').length > 0;
const paraHasBreakRun = (p) =>
  Array.from(p.getElementsByTagNameNS(W_NS, 'br')).some(b => b.getAttributeNS(W_NS, 'type') === 'page');
const breakAtParaEnd = (p) => {
  for (let i = p.childNodes.length - 1; i >= 0; i--) {
    const c = p.childNodes[i];
    if (c.nodeType !== 1) continue;
    if (c.localName === 'r') {
      const t = Array.from(c.getElementsByTagNameNS(W_NS, 't')).map(x => x.textContent).join('');
      const br = Array.from(c.getElementsByTagNameNS(W_NS, 'br')).some(b => b.getAttributeNS(W_NS, 'type') === 'page');
      if (t) return false;
      if (br) return true;
      continue;
    }
    if (['hyperlink','fldSimple','oMath','drawing','pict','sdt'].includes(c.localName)) return false;
  }
  return false;
};
const needsPageBreak = (idx) => {
  let prev = idx - 1;
  while (prev >= 0) { // 与引擎一致：无限回溯（跨空段，遇最后内容段/分节符停止）
    const c = allParas[prev];
    const pPr = Array.from(c.childNodes).find(n => n.nodeType === 1 && n.localName === 'pPr');
    const sectPr = pPr && Array.from(pPr.childNodes).find(n => n.nodeType === 1 && n.localName === 'sectPr');
    if (sectPr) {
      const tp = Array.from(sectPr.childNodes).find(n => n.nodeType === 1 && n.localName === 'type');
      if ((tp ? tp.getAttributeNS(W_NS, 'val') || 'nextPage' : 'nextPage') !== 'continuous') return false;
    }
    const hasPbb = !!(pPr && pPr.getElementsByTagNameNS(W_NS, 'pageBreakBefore').length);
    if (!paraHasContent(c)) {
      if (paraHasBreakRun(c) || hasPbb) return false;
      prev--;
      continue;
    }
    if (breakAtParaEnd(c)) return false;
    break; // 最后内容段无强制分页 → 需要 pageBreakBefore
  }
  return true;
};
let blankOk = true;
for (const idx of h1Idxs) {
  const p = allParas[idx];
  const pPr = Array.from(p.childNodes).find(n => n.nodeType === 1 && n.localName === 'pPr');
  const has = !!(pPr && pPr.getElementsByTagNameNS(W_NS, 'pageBreakBefore').length);
  const should = needsPageBreak(idx);
  if (has !== should) {
    blankOk = false;
    console.error(`✗ 空白页回归: 「${Array.from(p.getElementsByTagNameNS(W_NS, 't')).map(x => x.textContent).join('').trim().slice(0, 15)}」 pageBreakBefore=${has} 预期=${should}`);
  }
}
console.log(`✓ 空白页回归: ${h1Idxs.length} 个一级标题分页标记与已有强制分页一致（无双重分页/无漏页）`);
if (!blankOk) throw new Error('一级标题分页标记与已有强制分页冲突（可能产生空白页）');

// 10. 一级标题/摘要/ABSTRACT/目录标题上下空行统一为五号字大小（10.5pt=210），
//     上方与下方各 1 个行高 210 的空段（标题自身 after=0，空行由空段提供）
const TITLE_RE = /^(摘\s*要|ABSTRACT|目\s*录|1\s+绪|2\s+系统|3\s+系统|4\s+系统|5\s+系统|6\s+总|致\s*谢|附录)/;
const bodyKids10 = Array.from(doc.getElementsByTagNameNS(W_NS, 'body')[0].childNodes).filter(n => n.nodeType === 1);
const paraLine10 = (e) => {
  if (!e || e.localName !== 'p') return null;
  const pPr = Array.from(e.childNodes).find(n => n.nodeType === 1 && n.localName === 'pPr');
  const sp = pPr && Array.from(pPr.childNodes).find(n => n.nodeType === 1 && n.localName === 'spacing');
  return sp && sp.getAttributeNS(W_NS, 'line');
};
const skipInert10 = (idx, dir) => {
  let k = idx + dir;
  while (k >= 0 && k < bodyKids10.length && ['bookmarkStart','bookmarkEnd','proofErr'].includes(bodyKids10[k].localName)) k += dir;
  return k;
};
let titleTotal = 0, titleOk = 0;
for (let i = 0; i < bodyKids10.length; i++) {
  const e = bodyKids10[i];
  if (e.localName !== 'p') continue;
  const t = Array.from(e.getElementsByTagNameNS(W_NS, 't')).map(x => x.textContent).join('').trim();
  if (!TITLE_RE.test(t)) continue;
  if (t.length > 20 || /\d\s*$/.test(t)) continue;   // 排除目录条目
  if (isFront(e)) continue;
  titleTotal++;
  const pPr = Array.from(e.childNodes).find(n => n.nodeType === 1 && n.localName === 'pPr');
  const sp = pPr && Array.from(pPr.childNodes).find(n => n.nodeType === 1 && n.localName === 'spacing');
  const after = sp && sp.getAttributeNS(W_NS, 'after');
  const pIdx = skipInert10(i, -1);
  const nIdx = skipInert10(i, +1);
  const prevE = pIdx >= 0 ? bodyKids10[pIdx] : null;
  const nextE = nIdx < bodyKids10.length ? bodyKids10[nIdx] : null;
  const prevT = prevE && prevE.localName === 'p' ? Array.from(prevE.getElementsByTagNameNS(W_NS, 't')).map(x => x.textContent).join('').trim() : '';
  const nextT = nextE && nextE.localName === 'p' ? Array.from(nextE.getElementsByTagNameNS(W_NS, 't')).map(x => x.textContent).join('').trim() : '';
  if ((!after || after === '0') && prevE && prevE.localName === 'p' && !prevT && paraLine10(prevE) === '210' &&
      nextE && nextE.localName === 'p' && !nextT && paraLine10(nextE) === '210') titleOk++;
}
console.log(`✓ 标题上下空行(五号210): ${titleOk}/${titleTotal} 处（摘要/ABSTRACT/目录/一级标题）`);
if (titleOk !== titleTotal) throw new Error('标题上下空行未统一为五号字大小（210）');

// 11. 三级标题与正文均首行缩进2字符（firstLineChars=200）
let h3n = 0, h3ok = 0, bodyN = 0, bodyOk = 0;
cls9.roles.forEach((r, k) => {
  if (r !== 'body' && r !== 'h3') return;
  const pPr = Array.from(items[k].p.childNodes).find(n => n.nodeType === 1 && n.localName === 'pPr');
  const ind = pPr && Array.from(pPr.childNodes).find(n => n.nodeType === 1 && n.localName === 'ind');
  if (r === 'h3') { h3n++; if (ind && ind.getAttributeNS(W_NS, 'firstLineChars') === '200') h3ok++; }
  else { bodyN++; if (ind && ind.getAttributeNS(W_NS, 'firstLineChars') === '200') bodyOk++; }
});
console.log(`✓ 首行缩进2字符: 三级标题 ${h3ok}/${h3n} | 正文 ${bodyOk}/${bodyN}`);
if (h3ok !== h3n || bodyOk !== bodyN) throw new Error('三级标题或正文未首行缩进 2 字符');

// 12. 题注格式与位置：五号宋体居中（10.5pt=21半磅），图题紧贴图片正下方、表题紧贴表格正上方
const bodyKids11 = Array.from(doc.getElementsByTagNameNS(W_NS, 'body')[0].childNodes).filter(n => n.nodeType === 1);
const isEmptyCapP = (e) => e.localName === 'p' && e.getElementsByTagNameNS(W_NS, 't').length === 0 &&
  e.getElementsByTagNameNS(W_NS, 'drawing').length === 0;
const isImgCapP = (e) => e.localName === 'p' && e.getElementsByTagNameNS(W_NS, 'drawing').length > 0;
let capTotal = 0, capPosOk = 0, capFmtOk = 0;
for (let i = 0; i < bodyKids11.length; i++) {
  const e = bodyKids11[i];
  if (e.localName !== 'p') continue;
  const t = Array.from(e.getElementsByTagNameNS(W_NS, 't')).map(x => x.textContent).join('').trim();
  if (!/^(图|表)\s*[A-Z]?\d/.test(t)) continue;   // 正文「图4.1」+ 附录「图A1」
  if (isFront(e)) continue;
  capTotal++;
  const r0 = directChildren(e, 'r')[0];
  const rPr = r0 && directChildren(r0, 'rPr')[0];
  const sz = rPr && directChildren(rPr, 'sz')[0];
  const rf = rPr && directChildren(rPr, 'rFonts')[0];
  if (sz && sz.getAttributeNS(W_NS, 'val') === '21' && rf && rf.getAttributeNS(W_NS, 'eastAsia') === '宋体') capFmtOk++;
  let prevIdx = i - 1, nextIdx = i + 1;
  while (prevIdx >= 0 && (isEmptyCapP(bodyKids11[prevIdx]) || ['bookmarkStart','bookmarkEnd','proofErr'].includes(bodyKids11[prevIdx].localName))) prevIdx--;
  while (nextIdx < bodyKids11.length && (isEmptyCapP(bodyKids11[nextIdx]) || ['bookmarkStart','bookmarkEnd','proofErr'].includes(bodyKids11[nextIdx].localName))) nextIdx++;
  const prevE = prevIdx >= 0 ? bodyKids11[prevIdx] : null;
  const nextE = nextIdx < bodyKids11.length ? bodyKids11[nextIdx] : null;
  if (isImgCapP(e)) capPosOk++;          // 图题与图片同段（图片+题注一个段落）也算位置正确
  else if (/^图/.test(t)) { if (prevE && isImgCapP(prevE)) capPosOk++; }
  else { if (nextE && nextE.localName === 'tbl') capPosOk++; }
}
console.log(`✓ 题注: 五号宋体格式 ${capFmtOk}/${capTotal} | 位置(图题下/表题上) ${capPosOk}/${capTotal}`);
if (capFmtOk !== capTotal || capPosOk !== capTotal) throw new Error('题注格式或位置不符合要求（五号宋体、图题在图片下、表题在表格上）');

// 12b. 题注编号：同一章内的图号（表号）必须从 1 开始连续、不重号。
//      这是「编号漏计」这一类 bug 的唯一有效探针 —— 2026-09-27 查出的三种形态
//      （题注段自带图片 / 图片段里被敲进游离字符如夹具里那个 "z" / 其他）
//      症状完全一样：某个图不占号，其后整章前移一位，出现两个「图3.16」，
//      正文里「如下图3.3、图3.4所示」跟着全指错。
//      按编号自身分组校验（图3.5 → 组「图3」第 5 号），不依赖章标题识别，
//      所以章标题判错时这条不会跟着一起失效。图号与表号是两个独立序列。
const capSeq = {};
for (let i = 0; i < bodyKids11.length; i++) {
  const e = bodyKids11[i];
  if (e.localName !== 'p' || isFront(e)) continue;
  const t = Array.from(e.getElementsByTagNameNS(W_NS, 't')).map(x => x.textContent).join('').trim();
  const m = /^(图|表)\s*([A-Z]?\d+)(?:[.．](\d+))?/.exec(t);
  if (!m) continue;
  const key = m[1] + m[2];
  (capSeq[key] = capSeq[key] || []).push(m[3] ? +m[3] : 1);   // 附录「图A1」无小数点，按第 1 号算
}
const badSeq = [];
const grpKeys = Object.keys(capSeq);
for (const k of grpKeys) {
  const ns = capSeq[k].slice().sort((a, b) => a - b);
  for (let i = 0; i < ns.length; i++) {
    if (ns[i] !== i + 1) { badSeq.push(`${k}(${ns.join(',')})`); break; }
  }
}
console.log(`✓ 题注编号: ${grpKeys.length} 组从 1 起连续无重号 ${grpKeys.length - badSeq.length}/${grpKeys.length}` +
  (badSeq.length ? ` | 异常: ${badSeq.join(' ')}` : ''));
if (badSeq.length) throw new Error(`题注编号不连续或有重号: ${badSeq.join(' ')}`);

// 13. 题注移位单测：图题原本在图片上方、表题原本在表格下方时，自动移到图片正下方/表格正上方
function buildCapTestDocx() {
  const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  const esc = (s) => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  const P = (t) => `<w:p><w:r><w:t>${esc(t)}</w:t></w:r></w:p>`;
  const imgP = '<w:p><w:r><w:drawing><w:inline/></w:drawing></w:r></w:p>';
  const tbl = '<w:tbl><w:tblPr><w:tblW w:w="5000" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="5000"/></w:tblGrid>' +
    '<w:tr><w:tc><w:p><w:r><w:t>表头</w:t></w:r></w:p></w:tc></w:tr></w:tbl>';
  const body = P('1 绪  论') + P('正文内容。') +
    P('图1.1 系统框图') + imgP +            // 图题在图片上方（反了）
    P('正文继续。') +
    tbl + P('表1.1 参数表') +               // 表题在表格下方（反了）
    P('正文结束。');
  const xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<w:document xmlns:w="' + W + '"><w:body>' + body +
    '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:bottom="1440" w:left="1800" w:right="1800"/></w:sectPr>' +
    '</w:body></w:document>';
  return JSZip().file('word/document.xml', xml)
    .file('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>')
    .generateAsync({ type: 'nodebuffer' });
}
const capR = await FormatTool.formatDocx(await buildCapTestDocx(), FormatTool.DEFAULTS, { format: 'nodebuffer' });
const capZip = await JSZip.loadAsync(capR.data);
const capDoc = new DOMParser().parseFromString(await capZip.file('word/document.xml').async('string'), 'application/xml');
const capBody = capDoc.getElementsByTagNameNS(W_NS, 'body')[0];
const capEls = Array.from(capBody.childNodes).filter(n => n.nodeType === 1);
const capIdx = (t) => capEls.findIndex(e => e.localName === 'p' && Array.from(e.getElementsByTagNameNS(W_NS, 't')).map(x => x.textContent).join('').trim() === t);
const capTxt = (i) => i >= 0 ? Array.from(capEls[i].getElementsByTagNameNS(W_NS, 't')).map(x => x.textContent).join('').trim() : '';
const figI = capIdx('图1.1 系统框图'), tabI = capIdx('表1.1 参数表');
const figBelowImg = figI > 0 && isImgCapP(capEls[figI - 1]);
const tabAboveTbl = tabI >= 0 && capEls[tabI + 1] && capEls[tabI + 1].localName === 'tbl';
console.log(`✓ 题注移位单测: 图题移到图片下方=${figBelowImg} | 表题移到表格上方=${tabAboveTbl}`);
if (!figBelowImg || !tabAboveTbl) throw new Error('题注未自动移到图片下方/表格上方');

// 13b. 公式段居中单测：w:jc 必须落在 w:pPr 里。
//      setJc 收的是 **pPr** 不是段落；传段落进去时 childByNs(p,'jc') 找不到东西，
//      insertInOrder 又把 pPr/r 都当成「不在 PPR_ORDER 里」（ci = -1），一路落到
//      appendChild —— w:jc 被写成 w:p 的最后一个子元素，排在所有 run 后面。
//      OOXML 里 w:jc 只能待在 w:pPr 内，Word 会判文档损坏或直接丢弃，且居中不生效。
//      夹具里公式段落数为 0，真实文档走不到这条分支，所以只能靠构造文档来守。
function buildFormulaTestDocx() {
  const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  const M = 'http://schemas.openxmlformats.org/officeDocument/2006/math';
  const esc = (s) => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  const P = (t) => `<w:p><w:r><w:t>${esc(t)}</w:t></w:r></w:p>`;
  // 两种公式段都要守：
  //   mathP   —— 带编号文字（w:t 非空），一直能进 paras，原先会写出非法 w:jc；
  //   pureMathP —— 只有 oMath、w:t 为空，原先被 collectParas 当空段整段丢掉，
  //                role='formula' 永远不成立，公式居中整个功能空转。
  const mathP = `<w:p><m:oMath xmlns:m="${M}"><m:r><m:t>y=2</m:t></m:r></m:oMath><w:r><w:t>（3-2）</w:t></w:r></w:p>`;
  const pureMathP = `<w:p><m:oMath xmlns:m="${M}"><m:r><m:t>x=1</m:t></m:r></m:oMath></w:p>`;
  const body = P('1 绪  论') + P('正文内容。') + mathP + pureMathP + P('正文继续。');
  const xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    `<w:document xmlns:w="${W}" xmlns:m="${M}"><w:body>${body}` +
    '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:bottom="1440" w:left="1800" w:right="1800"/></w:sectPr>' +
    '</w:body></w:document>';
  return JSZip().file('word/document.xml', xml)
    .file('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>')
    .generateAsync({ type: 'nodebuffer' });
}
const fmR = await FormatTool.formatDocx(await buildFormulaTestDocx(), FormatTool.DEFAULTS, { format: 'nodebuffer' });
const fmDoc = new DOMParser().parseFromString(
  await (await JSZip.loadAsync(fmR.data)).file('word/document.xml').async('string'), 'application/xml');
/* ① 通用结构守卫：全文任何 w:jc 的父元素都必须是 w:pPr */
const strayJc = Array.from(fmDoc.getElementsByTagNameNS(W_NS, 'jc'))
  .filter(j => !j.parentNode || j.parentNode.localName !== 'pPr')
  .map(j => `<w:${j.parentNode ? j.parentNode.localName : '?'}>`);
/* ② 公式段自身：pPr 必须在最前，且 jc=center */
const M_NS_TEST = 'http://schemas.openxmlformats.org/officeDocument/2006/math';
/* 一个段落里可能有多个 oMath，按段落去重后逐段检查 */
const fmParas = [];
for (const mo of Array.from(fmDoc.getElementsByTagNameNS(M_NS_TEST, 'oMath'))) {
  const par = mo.parentNode;
  if (par && par.localName === 'p' && fmParas.indexOf(par) < 0) fmParas.push(par);
}
let fmCentered = 0;
const fmOrders = [];
for (const par of fmParas) {
  const kids = Array.from(par.childNodes).filter(n => n.nodeType === 1);
  fmOrders.push(kids.map(n => n.localName).join(' → '));
  const pr = kids[0] && kids[0].localName === 'pPr' ? kids[0] : null;
  const jc = pr && Array.from(pr.childNodes).find(n => n.localName === 'jc');
  if (pr && jc && jc.getAttributeNS(W_NS, 'val') === 'center') fmCentered++;
}
const fmOk = fmParas.length === 2 && fmCentered === 2;
console.log(`✓ 公式段单测: ${fmParas.length} 段，pPr 置顶且居中 ${fmCentered}/${fmParas.length}` +
  ` | 结构: ${fmOrders.join(' ／ ')}` + (strayJc.length ? ` | 游离 w:jc 父元素: ${strayJc.join(', ')}` : ''));
if (strayJc.length || !fmOk) throw new Error('公式段未居中，或 w:jc 未写入 w:pPr（w:jc 不是 w:p 的合法子元素，Word 会判文档损坏）');

// 14. 章节换页规整化单测：
//     无分节符/分页符 → 在标题前的空段上加 pageBreakBefore（空段换页，标题紧随新页顶部）；
//     连续分节符 → 改为 nextPage；标题前恰好 1 个空段；标题段自身无 pageBreakBefore
function buildBreakTestDocx() {
  const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  const esc = (s) => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  const P = (t) => `<w:p><w:r><w:t>${esc(t)}</w:t></w:r></w:p>`;
  const Pempty = () => '<w:p/>';
  const Psect = (type) => `<w:p><w:pPr><w:sectPr><w:type w:val="${type}"/>` +
    '<w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:bottom="1440" w:left="1800" w:right="1800"/>' +
    '</w:sectPr></w:pPr></w:p>';
  const body = P('正文第一段。') + Pempty() + Pempty() + Pempty() +  // 无分节 → 空段承担分页
    P('1 绪  论') + P('本章正文内容。') + Pempty() +
    Psect('continuous') + Pempty() + Pempty() +                       // 连续分节 + 多余空段
    P('2 系统方案设计') + P('后续正文内容。');
  const xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<w:document xmlns:w="' + W + '"><w:body>' + body +
    '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:bottom="1440" w:left="1800" w:right="1800"/></w:sectPr>' +
    '</w:body></w:document>';
  return JSZip().file('word/document.xml', xml)
    .file('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>')
    .generateAsync({ type: 'nodebuffer' });
}
const brR = await FormatTool.formatDocx(await buildBreakTestDocx(), FormatTool.DEFAULTS, { format: 'nodebuffer' });
const brZip = await JSZip.loadAsync(brR.data);
const brDoc = new DOMParser().parseFromString(await brZip.file('word/document.xml').async('string'), 'application/xml');
const brBody = brDoc.getElementsByTagNameNS(W_NS, 'body')[0];
const brEls = Array.from(brBody.childNodes).filter(n => n.nodeType === 1);
const brTxt = (e) => Array.from(e.getElementsByTagNameNS(W_NS, 't')).map(x => x.textContent).join('').trim();
const brChild = (e, n) => { if (!e || !e.childNodes) return null; return Array.from(e.childNodes).find(x => x.nodeType === 1 && x.localName === n); };
const brPBB = (e) => { const pPr = e.localName === 'p' && brChild(e, 'pPr'); return !!(pPr && pPr.getElementsByTagNameNS(W_NS, 'pageBreakBefore').length); };
const h1a = brEls.findIndex(e => e.localName === 'p' && brTxt(e) === '1 绪  论');
const h1b = brEls.findIndex(e => e.localName === 'p' && brTxt(e) === '2 系统方案设计');
const aPrev = h1a > 0 ? brEls[h1a - 1] : null;                       // 第一章前的空段（分节符已接手换页 → 无 PBB）
const aNext = h1a >= 0 && h1a < brEls.length - 1 ? brEls[h1a + 1] : null; // 标题下方空段（210 五号）
const aNextSp = aNext && aNext.localName === 'p' ? brChild(brChild(aNext, 'pPr') || {}, 'spacing') : null;
/* 第一章前的换页由「分节符（下一页）」承担（chapterSection 新行为），
   所以空段上不应再有 pageBreakBefore——两者叠加会多出一张空白页。 */
const aSectEl = h1a > 1 ? brEls[h1a - 2] : null;
const aSect = aSectEl && aSectEl.localName === 'p' ? brChild(brChild(aSectEl, 'pPr') || {}, 'sectPr') : null;
const aType = aSect && brChild(aSect, 'type');
const aPrevOk = !!aSect && (!aType || aType.getAttributeNS(W_NS, 'val') === 'nextPage') &&
  !!aPrev && aPrev.localName === 'p' && !brTxt(aPrev) && !brPBB(aPrev) &&
  !!aNext && aNext.localName === 'p' && !brTxt(aNext) &&
  !!aNextSp && aNextSp.getAttributeNS(W_NS, 'line') === '210';   // 标题下方五号空段（与上方对称）
const bPrev1 = h1b > 0 ? brEls[h1b - 1] : null;                      // 第二章前的空段（无 PBB）
const bPrev2 = h1b > 1 ? brEls[h1b - 2] : null;                      // 分节段（应为 nextPage）
const bSect = bPrev2 && bPrev2.localName === 'p' ? brChild(brChild(bPrev2, 'pPr') || {}, 'sectPr') : null;
const bType = bSect && brChild(bSect, 'type');
const bSectSp = bPrev2 && bPrev2.localName === 'p' ? brChild(brChild(bPrev2, 'pPr') || {}, 'spacing') : null;
const bPrev1Sp = bPrev1 && bPrev1.localName === 'p' ? brChild(brChild(bPrev1, 'pPr') || {}, 'spacing') : null;
const bPrevOk = !!bPrev1 && bPrev1.localName === 'p' && !brTxt(bPrev1) && !brPBB(bPrev1) &&
  !!bSect && (!bType || bType.getAttributeNS(W_NS, 'val') === 'nextPage') &&
  !!bSectSp && bSectSp.getAttributeNS(W_NS, 'line') === '20' &&   // 空分节段 1pt 行高（章节页只留一个回车）
  !!bPrev1Sp && bPrev1Sp.getAttributeNS(W_NS, 'line') === '210';  // 标题上方空段五号字行高
const titleNoPBB = !brPBB(brEls[h1a]) && !brPBB(brEls[h1b]);
console.log(`✓ 章节换页单测: 第一章前补 nextPage 分节符=${aPrevOk} | 连续分节改nextPage+空段规整=${bPrevOk} | 标题无重复分页=${titleNoPBB}`);
if (!aPrevOk || !bPrevOk || !titleNoPBB) throw new Error('章节换页规整化不符合要求');

// 15. 表格与图片紧邻自动空一行（单测：构造 tbl+img 与 img+tbl 相邻的文档，验证被分隔）
function buildTblImgTestDocx() {
  const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  const esc = (s) => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  const P = (t) => `<w:p><w:r><w:t>${esc(t)}</w:t></w:r></w:p>`;
  const imgP = '<w:p><w:r><w:drawing><w:inline/></w:drawing></w:r></w:p>';
  const tbl = '<w:tbl><w:tblPr><w:tblW w:w="5000" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="5000"/></w:tblGrid>' +
    '<w:tr><w:tc><w:p><w:r><w:t>表头</w:t></w:r></w:p></w:tc></w:tr>' +
    '<w:tr><w:tc><w:p><w:r><w:t>数据</w:t></w:r></w:p></w:tc></w:tr></w:tbl>';
  const body = P('1 绪  论') + P('论文正文内容。') + tbl + imgP + P('正文继续。') + imgP + tbl + P('正文结束。');
  const xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<w:document xmlns:w="' + W + '"><w:body>' + body +
    '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:bottom="1440" w:left="1800" w:right="1800"/></w:sectPr>' +
    '</w:body></w:document>';
  return JSZip().file('word/document.xml', xml)
    .file('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>')
    .generateAsync({ type: 'nodebuffer' });
}
const tiR = await FormatTool.formatDocx(await buildTblImgTestDocx(), FormatTool.DEFAULTS, { format: 'nodebuffer' });
const tiZip = await JSZip.loadAsync(tiR.data);
const tiDoc = new DOMParser().parseFromString(await tiZip.file('word/document.xml').async('string'), 'application/xml');
const tiBody = tiDoc.getElementsByTagNameNS(W_NS, 'body')[0];
const tiEls = Array.from(tiBody.childNodes).filter(n => n.nodeType === 1);
const isEmptyP = (e) => e.localName === 'p' && e.getElementsByTagNameNS(W_NS, 't').length === 0 &&
  e.getElementsByTagNameNS(W_NS, 'drawing').length === 0 && e.getElementsByTagNameNS(W_NS, 'pict').length === 0;
const isImgEl = (e) => e.localName === 'p' && e.getElementsByTagNameNS(W_NS, 'drawing').length > 0;
let tiSep = 0, tiAdj = 0;
for (let i = 0; i < tiEls.length - 1; i++) {
  const a = tiEls[i], b = tiEls[i + 1], c = tiEls[i + 2];
  if ((a.localName === 'tbl' && isImgEl(b)) || (isImgEl(a) && b.localName === 'tbl')) tiAdj++;
  if ((a.localName === 'tbl' && isEmptyP(b) && isImgEl(c)) || (isImgEl(a) && isEmptyP(b) && c.localName === 'tbl')) tiSep++;
}
/* 改用真题注后，表格/图片之间多半已经被题注段隔开，空行计数不再是固定的 2。
   真正要守的性质是：①表格与图片不直接相邻 ②题注紧贴各自的对象
   （表题在表格正上方、图题在图片正下方），顺带保证编号连续无跳号。 */
const tiTxt = (e) => Array.from(e.getElementsByTagNameNS(W_NS, 't')).map(x => x.textContent).join('');
let tiTbl = 0, tiImg = 0, tiCapOk = 0;
for (let i = 0; i < tiEls.length; i++) {
  const e = tiEls[i];
  if (e.localName === 'tbl') {
    tiTbl++;
    const p = tiEls[i - 1];
    if (p && p.localName === 'p' && /^表/.test(tiTxt(p))) tiCapOk++;
  }
  if (isImgEl(e)) {
    tiImg++;
    const n = tiEls[i + 1];
    if (n && n.localName === 'p' && /^图/.test(tiTxt(n))) tiCapOk++;
  }
}
console.log(`✓ 表格/图片分隔单测: 残留紧邻 ${tiAdj} 处 | 题注紧贴对象 ${tiCapOk}/${tiTbl + tiImg}` +
  ` | 表号连续=${tiEls.filter(e => e.localName === 'p' && /^表/.test(tiTxt(e))).map(e => tiTxt(e)).join(',')}`);
if (tiAdj !== 0 || tiCapOk !== tiTbl + tiImg) throw new Error('表格与图片未正确分隔/题注未紧贴');

// 16. 前置部分（前两页）保护：摘要之前的封面/声明不进行任何修改操作
//     （封面表格边框保留、封面图片行距保留、封面分节边距保留；正文正常格式化）
function buildFrontTestDocx() {
  const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  const esc = (s) => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  const P = (t) => `<w:p><w:r><w:t>${esc(t)}</w:t></w:r></w:p>`;
  const imgP = (line) => `<w:p><w:pPr><w:spacing w:line="${line}" w:lineRule="exact"/></w:pPr>` +
    '<w:r><w:drawing><w:inline/></w:drawing></w:r></w:p>';
  const coverTbl = '<w:tbl><w:tblPr><w:tblBorders><w:top w:val="single" w:sz="8"/><w:bottom w:val="single" w:sz="8"/>' +
    '<w:left w:val="single" w:sz="8"/><w:right w:val="single" w:sz="8"/>' +
    '<w:insideH w:val="single" w:sz="4"/><w:insideV w:val="single" w:sz="4"/></w:tblBorders></w:tblPr>' +
    '<w:tblGrid><w:gridCol w:w="5000"/></w:tblGrid><w:tr><w:tc><w:p><w:r><w:t>封面信息</w:t></w:r></w:p></w:tc></w:tr></w:tbl>';
  const bodyTbl = '<w:tbl><w:tblPr></w:tblPr><w:tblGrid><w:gridCol w:w="5000"/></w:tblGrid>' +
    '<w:tr><w:tc><w:p><w:r><w:t>正文表格</w:t></w:r></w:p></w:tc></w:tr></w:tbl>';
  const sectP = (mar) => `<w:p><w:pPr><w:sectPr><w:pgSz w:w="11906" w:h="16838"/>` +
    `<w:pgMar w:top="${mar}" w:bottom="${mar}" w:left="${mar}" w:right="${mar}"/></w:sectPr></w:pPr></w:p>`;
  const body = P('毕业设计（论文）') + P('学号：20230001') + coverTbl + imgP('500') + sectP('1800') +
    P('摘  要') + P('本文针对转台控制精度问题开展研究。') + P('关键词：测试') +
    P('1 绪  论') + P('论文正文内容。') + bodyTbl + imgP('500') + P('2 总  结') + P('总结内容。');
  const xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<w:document xmlns:w="' + W + '"><w:body>' + body +
    '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1800" w:bottom="1800" w:left="1800" w:right="1800"/></w:sectPr>' +
    '</w:body></w:document>';
  return JSZip().file('word/document.xml', xml)
    .file('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>')
    .generateAsync({ type: 'nodebuffer' });
}
const frR = await FormatTool.formatDocx(await buildFrontTestDocx(), FormatTool.DEFAULTS, { format: 'nodebuffer' });
const frZip = await JSZip.loadAsync(frR.data);
const frDoc = new DOMParser().parseFromString(await frZip.file('word/document.xml').async('string'), 'application/xml');
const frBody = frDoc.getElementsByTagNameNS(W_NS, 'body')[0];
const frEls = Array.from(frBody.childNodes).filter(n => n.nodeType === 1);
const frTbls = frEls.filter(e => e.localName === 'tbl');
const marOf = (sp) => { const m = directChildren(sp, 'pgMar')[0]; return m && m.getAttributeNS(W_NS, 'top'); };
const lineOf = (p) => { const pPr = directChildren(p, 'pPr')[0]; const sp = pPr && directChildren(pPr, 'spacing')[0]; return sp && sp.getAttributeNS(W_NS, 'line'); };
const insHOf = (tbl) => {
  const tblPr = directChildren(tbl, 'tblPr')[0];
  const tb = tblPr && directChildren(tblPr, 'tblBorders')[0];
  const h = tb && directChildren(tb, 'insideH')[0];
  return h && h.getAttributeNS(W_NS, 'val');
};
const topOf = (tbl) => {
  const tblPr = directChildren(tbl, 'tblPr')[0];
  const tb = tblPr && directChildren(tblPr, 'tblBorders')[0];
  const t = tb && directChildren(tb, 'top')[0];
  return t && t.getAttributeNS(W_NS, 'val') + '/' + t.getAttributeNS(W_NS, 'sz');
};
const frImgPs = frEls.filter(e => e.localName === 'p' && e.getElementsByTagNameNS(W_NS, 'drawing').length > 0);
const frSects = Array.from(frDoc.getElementsByTagNameNS(W_NS, 'sectPr'));
const coverTblKept = insHOf(frTbls[0]) === 'single';          // 封面表格边框未被清除
const bodyTblLine = topOf(frTbls[frTbls.length - 1]) === 'single/12'; // 正文表格已是顶底线1.5磅
const coverImgKept = lineOf(frImgPs[0]) === '500';            // 封面图片行距未改
const bodyImgFixed = lineOf(frImgPs[frImgPs.length - 1]) === '240'; // 正文图片单倍行距
const coverMarKept = marOf(frSects[0]) === '1800';            // 封面分节边距未改
const bodyMarFixed = marOf(frSects[frSects.length - 1]) === '1418'; // 正文分节 2.5cm
console.log(`✓ 前置保护单测: 封面表格边框保留=${coverTblKept} 封面图片行距保留=${coverImgKept} 封面边距保留=${coverMarKept} | 正文表格=${bodyTblLine} 正文图片=${bodyImgFixed} 正文边距=${bodyMarFixed}`);
if (!coverTblKept || !coverImgKept || !coverMarKept || !bodyTblLine || !bodyImgFixed || !bodyMarFixed) {
  throw new Error('前置部分保护或正文格式化不符合要求');
}

// 17. 学校格式预设：js/presets.js 可加载，且每个预设覆盖全部可调字段
await import('../js/presets.js');
const PRESETS = globalThis.SCHOOL_PRESETS || [];
const presetFields = ['marginTop','marginBottom','marginLeft','marginRight','pageNumber','bodyFont','latinFont',
  'bodySize','firstLineChars','titleSize','titleBold','h1Size','h1Jc','h2Size','h3Size','headingBold',
  'absHeadSize','absBodySize','refSize','refHangingChars','captionFont','captionSize','pageNumberSplit',
  'threeLineTable','tableText','tableFont','tableSize',
  'chapterPageBreak','autoToc','lineSpacingMode','lineSpacing'];
let presetOk = true;
for (const p of PRESETS) {
  const missing = presetFields.filter(f => p.settings[f] === undefined);
  if (missing.length) { presetOk = false; console.error(`✗ 预设「${p.name}」缺少字段: ${missing.join(', ')}`); }
}
console.log(`✓ 学校格式预设: ${PRESETS.length} 套（${PRESETS.map(p => p.name).join(' / ')}）`);
if (!presetOk || !PRESETS.length) throw new Error('学校预设缺失或不完整');

/* ---------- v1.3 目录识别单测 ---------- */
// 18. 构造含目录的文档：TOC 样式条目 + tab页码条目 + 超链接条目，验证目录被识别并重排
function buildTocTestDocx() {
  const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  const esc = (s) => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  const P = (t) => `<w:p><w:r><w:t>${esc(t)}</w:t></w:r></w:p>`;
  const Pstyled = (t, style) => `<w:p><w:pPr><w:pStyle w:val="${style}"/></w:pPr><w:r><w:t>${esc(t)}</w:t></w:r></w:p>`;
  // Word 目录条目标准结构：标题 + tab + 页码（无点线、无空格）
  const Ptab = (t, page) => `<w:p><w:r><w:t>${esc(t)}</w:t></w:r><w:r><w:tab/><w:t>${page}</w:t></w:r></w:p>`;
  const Phlink = (t) => `<w:p><w:hyperlink w:anchor="_Toc1"><w:r><w:t>${esc(t)}</w:t></w:r></w:hyperlink></w:p>`;
  const body =
    P('毕业设计（论文）') +
    P('摘  要') + P('本文设计并实现了一套智能养老监护系统。') +
    P('目 录') +
    Ptab('摘  要', 'I') +                    // 罗马数字页码条目
    Pstyled('1 绪论', 'TOC1') + Pstyled('1.1 研究背景', 'TOC2') +
    Ptab('2 系统方案设计', '12') + Ptab('2.1 系统设计思路', '12') +
    Phlink('3 系统硬件电路设计') +
    P('1 绪  论') + P('1.1 研究背景') + P('随着我国老龄化程度不断加深，传统养老模式难以满足现代社会的需求。') +
    P('2 系统方案设计') +
    P('参考文献') + P('[1] 孙家广.计算机图形学[M].北京:清华大学出版社,1995:15-18.');
  const xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<w:document xmlns:w="' + W + '"><w:body>' + body +
    '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:bottom="1440" w:left="1800" w:right="1800"/></w:sectPr>' +
    '</w:body></w:document>';
  return JSZip().file('word/document.xml', xml)
    .file('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>')
    .generateAsync({ type: 'nodebuffer' });
}
const tocR = await FormatTool.formatDocx(await buildTocTestDocx(), FormatTool.DEFAULTS, { format: 'nodebuffer' });
console.log(`✓ 目录识别单测: 标题 ${tocR.counts.tocHead} | 条目 ${tocR.counts.tocItem}（罗马页码1 + TOC样式2 + tab页码2 + 超链接1）| 一级标题 ${tocR.counts.h1} | 识别标志 ${JSON.stringify(tocR.info)}`);
if (tocR.counts.tocHead !== 1 || tocR.counts.tocItem !== 6 || tocR.counts.h1 < 2) {
  throw new Error(`目录识别失败: tocHead=${tocR.counts.tocHead} tocItem=${tocR.counts.tocItem}`);
}
// 校验输出中目录已替换为自动目录域（sdt 包裹 + TOC 指令 + dirty 自动更新 + 预填充条目）
const tocZip = await JSZip.loadAsync(tocR.data);
const tocDoc = new DOMParser().parseFromString(await tocZip.file('word/document.xml').async('string'), 'application/xml');
const tocSdt = tocDoc.getElementsByTagNameNS(W_NS, 'sdt')[0];
const tocInstr = Array.from(tocDoc.getElementsByTagNameNS(W_NS, 'instrText')).map(x => x.textContent).join('');
const tocDirty = Array.from(tocDoc.getElementsByTagNameNS(W_NS, 'fldChar'))
  .some(f => f.getAttributeNS(W_NS, 'fldCharType') === 'begin' && f.getAttributeNS(W_NS, 'dirty') === 'true');
const tocEntryParas = tocSdt ? Array.from(tocSdt.getElementsByTagNameNS(W_NS, 'p')).filter(p => {
  const pPr = Array.from(p.childNodes).find(n => n.nodeType === 1 && n.localName === 'pPr');
  const ps = pPr && Array.from(pPr.childNodes).find(n => n.nodeType === 1 && n.localName === 'pStyle');
  return ps && /^TOC[123]$/.test(ps.getAttributeNS(W_NS, 'val') || '');
}) : [];
const tocEntryTxt = tocEntryParas.map(p => Array.from(p.getElementsByTagNameNS(W_NS, 't')).map(x => x.textContent).join(''));
const tocHasEntries = tocEntryParas.length >= 2 && tocEntryTxt.some(t => t.includes('绪')) &&
  tocEntryParas.every(p => p.getElementsByTagNameNS(W_NS, 'tab').length > 0);
const tocFieldOk = !!tocSdt && /TOC/.test(tocInstr) && tocDirty && tocHasEntries;
console.log(`✓ 自动目录域: sdt=${!!tocSdt} | 指令=${tocInstr.trim().slice(0, 22)} | 自动更新=${tocDirty} | 预填充条目=${tocEntryParas.length} 条`);
if (!tocFieldOk) throw new Error('目录未替换为自动目录域（sdt + TOC + dirty + 预填充条目）');

// 20. 章标题体例改写：「第1章 绪论」/「第二章 …」→「1 绪论」/「2 …」（数字与文字间隔一字符），
//     正文里指代章节的「第2章」原样保留；附录只一个时不编序号（「附录 原理图」），
//     多于一个依序编 A、B；附录里的图题改成「图A1」（字母与序号之间没有「.」），正文仍是「图1.1」
function buildH1TestDocx(multiApp) {
  const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  const esc = (s) => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  const P = (t) => `<w:p><w:r><w:t>${esc(t)}</w:t></w:r></w:p>`;
  const imgP = '<w:p><w:r><w:drawing><w:inline/></w:drawing></w:r></w:p>';
  const body = P('摘  要') + P('本文研究转台控制精度问题。') + P('关键词：转台；控制') +
    P('第1章 绪论') + P('本文将在第2章给出总体设计方案。') + imgP + P('图1.1 系统框图') +
    P('第二章 系统总体设计') + P('系统分为三个模块实现。') +
    P('参考文献') + P('[1] 孙家广.计算机图形学[M].北京:清华大学出版社,1995:15-18.') +
    P('附录1 原理图') + imgP + P('图1.2 电路原理图') +
    (multiApp ? P('附录2 部署步骤') + imgP + P('图1.3 部署流程图') : '');
  const xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<w:document xmlns:w="' + W + '"><w:body>' + body +
    '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:bottom="1440" w:left="1800" w:right="1800"/></w:sectPr>' +
    '</w:body></w:document>';
  return JSZip().file('word/document.xml', xml)
    .file('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>')
    .generateAsync({ type: 'nodebuffer' });
}
async function h1TextsOf(buf) {
  const z = await JSZip.loadAsync(buf);
  const d = new DOMParser().parseFromString(await z.file('word/document.xml').async('string'), 'application/xml');
  const b = d.getElementsByTagNameNS(W_NS, 'body')[0];
  return Array.from(b.childNodes).filter(n => n.nodeType === 1 && n.localName === 'p')
    .map(p => Array.from(p.getElementsByTagNameNS(W_NS, 't')).map(x => x.textContent).join('').trim());
}
const h1Single = await h1TextsOf((await FormatTool.formatDocx(await buildH1TestDocx(false), FormatTool.DEFAULTS, { format: 'nodebuffer' })).data);
const h1Multi = await h1TextsOf((await FormatTool.formatDocx(await buildH1TestDocx(true), FormatTool.DEFAULTS, { format: 'nodebuffer' })).data);
const has = (arr, s) => arr.includes(s);
const oneAppOk = has(h1Single, '1 绪论') && has(h1Single, '2 系统总体设计') &&
  has(h1Single, '附录 原理图') && has(h1Single, '图A1 电路原理图') &&
  has(h1Single, '图1.1 系统框图') &&                       // 正文图仍带「.」
  has(h1Single, '本文将在第2章给出总体设计方案。') &&        // 正文里的指代不动
  !h1Single.some(t => /^第[一二三四五六七八九十0-9]+章/.test(t));
const multiAppOk = has(h1Multi, '附录A 原理图') && has(h1Multi, '附录B 部署步骤') &&
  has(h1Multi, '图A1 电路原理图') && has(h1Multi, '图B1 部署流程图') &&
  !has(h1Multi, '附录 原理图');
console.log(`✓ 章标题体例单测: 单个附录=${oneAppOk} | 多个附录编 A/B=${multiAppOk}`);
console.log(`  单附录标题: ${h1Single.filter(t => /^(第|[0-9]+ |附)/.test(t)).join(' | ')}`);
console.log(`  多附录标题: ${h1Multi.filter(t => /^(第|[0-9]+ |附)/.test(t)).join(' | ')}`);
if (!oneAppOk || !multiAppOk) throw new Error('章标题体例改写不符合附件8（第N章 → N、附录序号、图A1）');

// 21. 导出文件名：原名 + 工具版本号（不再缀「_格式化」）；已带版本号或浏览器重名副本
//     后缀的先去掉再拼，反复跑不会越拼越长；作者自己的版本号（CYX_3.0）不能被吃掉
const V = '_v' + FormatTool.VERSION + '.docx';
const nameCases = [
  ['CYX_毕业设计.docx', 'CYX_毕业设计' + V],
  ['1.docx', '1' + V],
  ['CYX_毕业设计.DOCX', 'CYX_毕业设计' + V],
  ['CYX_毕业设计_v1.6.2.docx', 'CYX_毕业设计' + V],        // 旧版本后缀 → 换成本版
  ['CYX_毕业设计_v1.6.2 (1).docx', 'CYX_毕业设计' + V],    // 浏览器重名副本的「 (1)」
  ['CYX_毕业设计-v1.6.2.docx', 'CYX_毕业设计' + V],        // 连字符写法也认
  ['CYX_3.0.docx', 'CYX_3.0' + V],                          // 作者自己的版本号不动
  ['论文-终稿.docx', '论文-终稿' + V],
];
const nameBad = nameCases.filter(([i, o]) => FormatTool.outputName(i) !== o);
console.log(`✓ 导出文件名: ${nameCases.length - nameBad.length}/${nameCases.length} 例正确（例：CYX_毕业设计.docx → ${FormatTool.outputName('CYX_毕业设计.docx')}）`);
nameBad.forEach(([i, o]) => console.error(`  ✗ ${i} → ${FormatTool.outputName(i)}（预期 ${o}）`));
if (nameBad.length) throw new Error('导出文件名规则不正确');

// 22. 域字符（w:fldChar）必须配平 —— 这是「输出文档正文内容全部消失」的根因守卫
//     域字符不配平（典型：只剩一个没有 begin 的 end）时，Word 打开文档更新域会把
//     其后的内容当成域的一部分处理，正文就整片不见了。v1.6.3 及以前把已格式化过的
//     文档再格式化一遍时，旧目录只删掉前半截，正文里就留下这样的孤儿 end
function fieldStatsOf(xml) {
  const d = new DOMParser().parseFromString(xml, 'application/xml');
  const fc = d.getElementsByTagNameNS(W_NS, 'fldChar');
  let begin = 0, end = 0, depth = 0, min = 0;
  for (let i = 0; i < fc.length; i++) {
    const t = fc[i].getAttributeNS(W_NS, 'fldCharType');
    if (t === 'begin') { begin++; depth++; }
    else if (t === 'end') { end++; depth--; }
    if (depth < min) min = depth;
  }
  return { begin, end, depth, min };
}
async function docStats(buf) {
  const z = await JSZip.loadAsync(buf);
  const xml = await z.file('word/document.xml').async('string');
  const d = new DOMParser().parseFromString(xml, 'application/xml');
  const ps = d.getElementsByTagNameNS(W_NS, 'p');
  let chars = 0, paras = 0;
  for (let i = 0; i < ps.length; i++) {
    const t = (ps[i].textContent || '').replace(/\s+/g, '');
    if (t) { paras++; chars += t.length; }
  }
  /* 包级统计。原先这里只数字数和段数，于是「每跑一遍多一个 header 部件、每个
     _Toc 书签多一份副本」这类退化完全看不出来 —— 2026-09-27 实测连跑三遍
     header 16→17→18、_Toc 书签 459→501→543，断言照样打印 ✓。 */
  const names = Object.keys(z.files);
  const bmNames = Array.from(d.getElementsByTagNameNS(W_NS, 'bookmarkStart'))
    .map((b) => b.getAttributeNS(W_NS, 'name') || '').filter(Boolean);
  const bmSeen = {};
  let bmDup = 0, tocDup = 0;
  for (const n of bmNames) {
    if (bmSeen[n]) { bmDup++; if (/^_Toc/.test(n)) tocDup++; }
    bmSeen[n] = true;
  }
  return Object.assign(fieldStatsOf(xml), {
    chars, paras,
    parts: names.length,
    headers: names.filter((f) => /^word\/header\d*\.xml$/.test(f)).length,
    footers: names.filter((f) => /^word\/footer\d*\.xml$/.test(f)).length,
    bmTotal: bmNames.length,
    bmUnique: Object.keys(bmSeen).length,
    bmDup, tocDup,
  });
}
const balanced = (s) => s.begin === s.end && s.depth === 0 && s.min === 0;

const pass1 = await docStats(result.data);
const again = await FormatTool.formatDocx(result.data, FormatTool.DEFAULTS, { format: 'nodebuffer' });
const pass2 = await docStats(again.data);
/* 稳定性 = 字数/段数不退化，**且**包部件数、页眉页脚部件数、书签总数与唯一名数
   一模一样。_Toc 重名单独判死：它会让目录域的 HYPERLINK \l 跳到第一个同名书签，
   点目录条目跳错页（作者原有的其它书签重名不归工具管，只报告不判死）。 */
const idemOk = pass1.paras === pass2.paras && Math.abs(pass1.chars - pass2.chars) <= pass1.chars * 0.02 &&
  pass1.parts === pass2.parts && pass1.headers === pass2.headers && pass1.footers === pass2.footers &&
  pass1.bmTotal === pass2.bmTotal && pass1.bmUnique === pass2.bmUnique &&
  pass1.tocDup === 0 && pass2.tocDup === 0;
console.log(`✓ 域字符配平: 一遍 begin=${pass1.begin}/end=${pass1.end} 收尾深度=${pass1.depth} | ` +
  `二遍 begin=${pass2.begin}/end=${pass2.end} 收尾深度=${pass2.depth}`);
console.log(`✓ 反复格式化稳定: 一遍 ${pass1.chars}字/${pass1.paras}段 → 二遍 ${pass2.chars}字/${pass2.paras}段 | ` +
  `包部件 ${pass1.parts}→${pass2.parts}（页眉 ${pass1.headers}→${pass2.headers} / 页脚 ${pass1.footers}→${pass2.footers}）| ` +
  `书签 ${pass1.bmTotal} 个/${pass1.bmUnique} 唯一名（_Toc 重名 ${pass1.tocDup}）→ ` +
  `${pass2.bmTotal} 个/${pass2.bmUnique} 唯一名（_Toc 重名 ${pass2.tocDup}）`);
if (!balanced(pass1) || !balanced(pass2) || !idemOk) {
  throw new Error('域字符不配平，或反复格式化后结果退化：' +
    `一遍 ${pass1.chars}字/${pass1.paras}段 部件${pass1.parts}(页眉${pass1.headers}/页脚${pass1.footers}) ` +
    `书签${pass1.bmTotal}/${pass1.bmUnique}唯一 _Toc重名${pass1.tocDup}；` +
    `二遍 ${pass2.chars}字/${pass2.paras}段 部件${pass2.parts}(页眉${pass2.headers}/页脚${pass2.footers}) ` +
    `书签${pass2.bmTotal}/${pass2.bmUnique}唯一 _Toc重名${pass2.tocDup}`);
}

/* 22b. 目录条目引用的书签必须真实存在，且每一个都唯一。
       工具生成的目录条目是 PAGEREF _TocN 域（\h 超链接），书签不存在时 Word 显示
       「错误！未定义书签。」；重名时只认第一个同名书签，点条目跳到错误位置。
       这条同时能抓住两种退化：把作者的 _TocN 改名（引用悬空）、以及再叠加一份
       同名副本（重名）—— 前者 idempotency 断言看不出来（计数照样稳定）。 */
function tocRefStats(xml) {
  const d = new DOMParser().parseFromString(xml, 'application/xml');
  const defined = {};
  let definedDup = 0;
  const bms = Array.from(d.getElementsByTagNameNS(W_NS, 'bookmarkStart'));
  for (const b of bms) {
    const n = b.getAttributeNS(W_NS, 'name') || '';
    if (!n) continue;
    if (defined[n]) definedDup++;
    defined[n] = true;
  }
  const refs = {};
  const instrs = Array.from(d.getElementsByTagNameNS(W_NS, 'instrText'));
  for (const it of instrs) {
    const s = it.textContent || '';
    const m = /(?:PAGEREF|HYPERLINK\s+"?\\l"?)\s+"?(_Toc\d+)/i.exec(s);
    if (m) refs[m[1]] = true;
  }
  const missing = Object.keys(refs).filter((n) => !defined[n]);
  return { definedCount: Object.keys(defined).length, definedDup, refCount: Object.keys(refs).length, missing };
}
const rel1 = tocRefStats(await (await JSZip.loadAsync(result.data)).file('word/document.xml').async('string'));
const rel2 = tocRefStats(await (await JSZip.loadAsync(again.data)).file('word/document.xml').async('string'));
const linkOk = rel1.refCount > 0 && rel1.missing.length === 0 && rel1.definedDup === 0 &&
  rel2.refCount > 0 && rel2.missing.length === 0 && rel2.definedDup === 0;
console.log(`✓ 目录书签可解析: 一遍 书签 ${rel1.definedCount} 个（重名 ${rel1.definedDup}）/ 条目引用 ${rel1.refCount} 个，悬空 ${rel1.missing.length} | ` +
  `二遍 书签 ${rel2.definedCount} 个（重名 ${rel2.definedDup}）/ 条目引用 ${rel2.refCount} 个，悬空 ${rel2.missing.length}`);
if (!linkOk) {
  throw new Error('目录条目引用的书签悬空或重名（Word 会显示「错误！未定义书签。」或跳到错误位置）：' +
    `一遍 悬空[${rel1.missing.join(',')}] 重名${rel1.definedDup}；二遍 悬空[${rel2.missing.join(',')}] 重名${rel2.definedDup}`);
}

// 输入文档自带孤儿域字符（旧目录只剩下一个 end）时，输出必须被修好
const orphanXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<w:document xmlns:w="' + W_NS + '"><w:body>' +
  '<w:p><w:r><w:t>摘  要</w:t></w:r></w:p><w:p><w:r><w:t>本文研究转台控制精度问题。</w:t></w:r></w:p>' +
  '<w:p><w:r><w:t>关键词：转台；控制</w:t></w:r></w:p>' +
  '<w:p><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>' +          // 孤儿 end
  '<w:p><w:r><w:t>第1章 绪论</w:t></w:r></w:p><w:p><w:r><w:t>绪论正文内容。</w:t></w:r></w:p>' +
  '<w:p><w:r><w:t>参考文献</w:t></w:r></w:p>' +
  '<w:p><w:r><w:t>[1] 孙家广.计算机图形学[M].北京:清华大学出版社,1995:15-18.</w:t></w:r></w:p>' +
  '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:bottom="1440" w:left="1800" w:right="1800"/></w:sectPr>' +
  '</w:body></w:document>';
const orphanDocx = await JSZip().file('word/document.xml', orphanXml)
  .file('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>')
  .generateAsync({ type: 'nodebuffer' });
const orphanOut = await FormatTool.formatDocx(orphanDocx, FormatTool.DEFAULTS, { format: 'nodebuffer' });
const orphanStats = fieldStatsOf(await (await JSZip.loadAsync(orphanOut.data)).file('word/document.xml').async('string'));
console.log(`✓ 孤儿域字符修复: 输入含无 begin 的 end → 输出 begin=${orphanStats.begin}/end=${orphanStats.end} ` +
  `收尾深度=${orphanStats.depth}（fieldsFixed=${orphanOut.counts.fieldsFixed}）`);
if (!balanced(orphanStats) || orphanOut.counts.fieldsFixed < 1) {
  throw new Error('输入文档里的孤儿域字符没有被修掉');
}

// 23. 三级标题「编号与文字之间没有空格」必须照样认成标题、并把空格补上。
//     2026-09-16 曾鹏的 2.docx 第 3 章 9 个三级标题全写成「3.1.1系统目标」，
//     而 RE.h3 当时要求编号后必须有空白，h2 又因 (?![\d.]) 撞上第二个小数点，
//     于是整段掉进 body：宋体小四、两端对齐、缩进 480 —— 标题的字体字号对齐缩进
//     四处全丢。同时补空格也不能误伤：本来就有空格的不许补成两个，正文里指代
//     章节的「1.2.3节」不许动，带小数点的长句不许被当成标题。
function buildH3TestDocx() {
  const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  const esc = (s) => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  const P = (t) => `<w:p><w:r><w:t>${esc(t)}</w:t></w:r></w:p>`;
  const body = P('摘  要') + P('本文研究转台控制精度问题。') + P('关键词：转台；控制') +
    P('第1章 绪论') + P('1.1 研究背景') +
    P('1.1.1 研究背景与意义') +          // 本来就有空格：不许补成两个
    P('1.1.2国内外研究现状') +           // 缺空格：要补成「1.1.2 国内外研究现状」
    P('这是正文段落，1.2.3节给出了详细说明。') +   // 正文里的指代：一个字都不许动
    // 长句：不许当标题。这句必须超过 40 字 —— 二/三级标题判定有一条 tx.length <= 40
    // 的长度守卫，夹具写短了就测不到守卫本身（第一版夹具 38 字，被判成二级标题，
    // 是夹具的问题不是代码的问题）。带小数点的短句仍会被当成 h2，那是 h2 规则早就
    // 有的行为（RE.h2 一直是 \s*），不在本次改动范围内。
    P('1.5倍车速下的制动距离与理论计算值存在明显偏差，需要进一步分析其中的原因所在，以便对制动模型进行修正。') +
    P('第3章 系统设计') +
    P('3.1.1系统目标') +                 // 缺空格：要补
    P('3.1.2用户角色分析') +              // 缺空格：要补
    P('3.1.3 已带空格的三级标题') +        // 本来就有空格：不许补成两个
    P('参考文献') + P('[1] 孙家广.计算机图形学[M].北京:清华大学出版社,1995:15-18.');
  const xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<w:document xmlns:w="' + W + '"><w:body>' + body +
    '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:bottom="1440" w:left="1800" w:right="1800"/></w:sectPr>' +
    '</w:body></w:document>';
  return JSZip().file('word/document.xml', xml)
    .file('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>')
    .generateAsync({ type: 'nodebuffer' });
}
/* 逐段取出：文字 + 段落里所有 run 的字号/中文字体。
   三级标题是黑体四号（sz=28），正文是宋体小四（sz=24）——
   两者分得开，就说明这段确实按标题排了，没有掉进 body */
async function paraStylesOf(buf) {
  const z = await JSZip.loadAsync(buf);
  const d = new DOMParser().parseFromString(await z.file('word/document.xml').async('string'), 'application/xml');
  const b = d.getElementsByTagNameNS(W_NS, 'body')[0];
  return Array.from(b.childNodes).filter(n => n.nodeType === 1 && n.localName === 'p').map(p => {
    const szs = new Set(), fonts = new Set();
    Array.from(p.getElementsByTagNameNS(W_NS, 'sz')).forEach(s => szs.add(s.getAttributeNS(W_NS, 'val')));
    Array.from(p.getElementsByTagNameNS(W_NS, 'rFonts')).forEach(f => {
      const e = f.getAttributeNS(W_NS, 'eastAsia');
      if (e) fonts.add(e);
    });
    return {
      text: Array.from(p.getElementsByTagNameNS(W_NS, 't')).map(x => x.textContent).join('').trim(),
      szs: szs, fonts: fonts,
    };
  }).filter(x => x.text);
}
const h3Out = await paraStylesOf((await FormatTool.formatDocx(await buildH3TestDocx(), FormatTool.DEFAULTS, { format: 'nodebuffer' })).data);
const txt = h3Out.map(x => x.text);
const byTxt = (t) => h3Out.find(x => x.text === t);
/* 是标题还是正文，用字体字号判：黑体 sz28 = 三级标题，宋体 sz24 = 正文 */
const isH3 = (t) => { const x = byTxt(t); return !!x && x.szs.has('28') && !x.szs.has('24') && x.fonts.has('黑体'); };
const isBody = (t) => { const x = byTxt(t); return !!x && x.szs.has('24') && !x.szs.has('28') && x.fonts.has('宋体'); };

const h3SpaceOk =
  txt.includes('1.1.2 国内外研究现状') && txt.includes('3.1.1 系统目标') && txt.includes('3.1.2 用户角色分析') &&
  !txt.includes('1.1.2国内外研究现状') && !txt.includes('3.1.1系统目标') && !txt.includes('3.1.2用户角色分析');
const h3DoubleOk =                                // 已有空格的不许补成两个
  txt.includes('1.1.1 研究背景与意义') && txt.includes('3.1.3 已带空格的三级标题') &&
  !txt.some(t => /^\d+\.\d+\.\d+\s\s/.test(t));
const h3RoleOk =                                  // 补过空格的和本来就有空格的，都要按三级标题排版
  ['1.1.1 研究背景与意义', '1.1.2 国内外研究现状', '3.1.1 系统目标', '3.1.2 用户角色分析', '3.1.3 已带空格的三级标题'].every(isH3);
const bodyCiteOk = isBody('这是正文段落，1.2.3节给出了详细说明。') &&      // 正文里的指代：不动、排版也不动
  txt.includes('这是正文段落，1.2.3节给出了详细说明。');
const bodyLongOk = isBody('1.5倍车速下的制动距离与理论计算值存在明显偏差，需要进一步分析其中的原因所在，以便对制动模型进行修正。');
const h3BodyOk = bodyCiteOk && bodyLongOk;        // 正文不被牵连

console.log(`✓ 三级标题缺空格: 补空格=${h3SpaceOk} | 不补成两个=${h3DoubleOk} | 按标题排版=${h3RoleOk} | ` +
  `正文指代不动=${bodyCiteOk} | 正文长句不当标题=${bodyLongOk}`);
console.log(`  标题判定: ${txt.filter(t => /^\d+\.\d+\.\d+/.test(t)).map(t => t + (isH3(t) ? '[标题]' : '[正文✗]')).join(' | ')}`);
if (!h3SpaceOk || !h3DoubleOk || !h3RoleOk || !h3BodyOk) {
  throw new Error('三级标题（编号后无空格）未被正确识别 / 补空格有误');
}

// 24. 表格内文字统一五号宋体 + Times New Roman；表格与正文说明框的底纹一律清掉。
//     2026-09-16 作者反馈「最新输出版本中表格内的文字部分格式未修改，且表格整体
//     底纹需要设定为无」——根因是 collectParas 只收 body 的直接子段落，表格段落
//     嵌在 tbl 底下，正文那套 applyFormatting 一个字也改不到表内。
//     附件8：「表序、表名和表格内字体均为五号宋体」，没给表头开例外。
//     封面/声明页的表格属于前置部分，一个字都不许动（含它自己的底纹）。
function buildTableTestDocx() {
  const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  const esc = (s) => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  const P = (t) => `<w:p><w:r><w:t>${esc(t)}</w:t></w:r></w:p>`;
  // 带指定字体/字号/加粗的单元格段落
  const Pc = (t, { font = '黑体', latin = 'Consolas', sz = 20, bold = false } = {}) =>
    `<w:p><w:r><w:rPr><w:rFonts w:ascii="${latin}" w:hAnsi="${latin}" w:eastAsia="${font}" w:cs="${latin}"/>` +
    (bold ? '<w:b/>' : '') +
    `<w:sz w:val="${sz}"/><w:szCs w:val="${sz}"/></w:rPr><w:t>${esc(t)}</w:t></w:r></w:p>`;
  const tc = (inner, shd) =>
    `<w:tc><w:tcPr>${shd ? `<w:shd w:val="clear" w:color="auto" w:fill="${shd}"/>` : ''}</w:tcPr>${inner}</w:tc>`;
  const tr = (cells) => `<w:tr>${cells}</w:tr>`;

  // 封面表格：黑体加粗三号 + 底纹，全部必须原样保留
  const coverTbl = '<w:tbl>' + tr(
    tc(Pc('课题名称', { font: '黑体', latin: '黑体', sz: 32, bold: true }), 'D9E2F3') +
    tc(Pc('二手车交易数据分析系统', { font: '黑体', latin: '黑体', sz: 32, bold: true }), 'D9E2F3')
  ) + '</w:tbl>';

  // 正文表格：表头黑体加粗 10pt、数据行 Consolas 10pt、单元格带浅蓝底纹
  const bodyTbl = '<w:tbl>' + tr(
    tc(Pc('字段名', { font: '黑体', sz: 20, bold: true }), 'DCE6F1') +
    tc(Pc('类型', { font: '黑体', sz: 20, bold: true }), 'DCE6F1')
  ) + tr(
    tc(Pc('price', { latin: 'Consolas', sz: 20 })) +
    tc(Pc('decimal(10,2)', { latin: 'Consolas', sz: 20 }))
  ) + '</w:tbl>';

  // 正文里的说明框：段落级底纹（F2F4F7），要清掉
  const callout = `<w:p><w:pPr><w:shd w:val="clear" w:color="auto" w:fill="F2F4F7"/></w:pPr>` +
    `<w:r><w:t>本系统数据为仿真数据。</w:t></w:r></w:p>`;

  const body = P('重庆工程学院本科毕业设计（论文）') + coverTbl +
    P('摘  要') + P('本文研究二手车交易数据分析。') + P('关键词：二手车；数据分析') +
    P('第1章 绪论') + P('1.1 研究背景') + P('二手车市场持续增长。') + bodyTbl + callout +
    P('参考文献') + P('[1] 孙家广.计算机图形学[M].北京:清华大学出版社,1995:15-18.');
  const xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<w:document xmlns:w="' + W + '"><w:body>' + body +
    '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:bottom="1440" w:left="1800" w:right="1800"/></w:sectPr>' +
    '</w:body></w:document>';
  return JSZip().file('word/document.xml', xml)
    .file('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>')
    .generateAsync({ type: 'nodebuffer' });
}

const tblOut = (await FormatTool.formatDocx(await buildTableTestDocx(), FormatTool.DEFAULTS, { format: 'nodebuffer' })).data;
const tz = await JSZip.loadAsync(tblOut);
const tdoc = new DOMParser().parseFromString(await tz.file('word/document.xml').async('string'), 'application/xml');
const tBody = tdoc.getElementsByTagNameNS(W_NS, 'body')[0];
const tKids = Array.from(tBody.childNodes).filter((n) => n.nodeType === 1);
const tAbsIdx = tKids.findIndex((k) => /^摘\s*要/.test((k.textContent || '').trim()));
const coverTblEl = tKids.find((k, i) => k.localName === 'tbl' && i < tAbsIdx);
const bodyTblEl = tKids.find((k, i) => k.localName === 'tbl' && i > tAbsIdx);

/* 取表格里每个带文字的 run 的字体/字号/粗细 */
function runsOf(tblEl) {
  const out = [];
  Array.from(tblEl.getElementsByTagNameNS(W_NS, 'r')).forEach((r) => {
    if (!r.getElementsByTagNameNS(W_NS, 't').length) return;
    const rPr = Array.from(r.childNodes).find((n) => n.nodeType === 1 && n.localName === 'rPr');
    const get = (name, attr) => {
      if (!rPr) return null;
      const el = Array.from(rPr.childNodes).find((n) => n.nodeType === 1 && n.localName === name);
      return el ? el.getAttributeNS(W_NS, attr) : null;
    };
    out.push({
      text: Array.from(r.getElementsByTagNameNS(W_NS, 't')).map((x) => x.textContent).join(''),
      east: get('rFonts', 'eastAsia'), ascii: get('rFonts', 'ascii'),
      sz: get('sz', 'val'),
      bold: !!(rPr && Array.from(rPr.childNodes).some((n) => n.nodeType === 1 && n.localName === 'b')),
    });
  });
  return out;
}
const shdCount = (el) => el.getElementsByTagNameNS(W_NS, 'shd').length;

const bodyRuns = runsOf(bodyTblEl);
const bodyTblOk = bodyRuns.length > 0 && bodyRuns.every((r) =>
  r.east === '宋体' && r.ascii === 'Times New Roman' && r.sz === '21' && !r.bold);
const tblShdOk = shdCount(bodyTblEl) === 0;
const coverRuns = runsOf(coverTblEl);
const coverOk = coverRuns.length > 0 && coverRuns.every((r) => r.east === '黑体' && r.sz === '32' && r.bold) &&
  shdCount(coverTblEl) > 0;                       // 封面表格的底纹必须留着
const calloutEl = tKids.find((k) => k.localName === 'p' && /本系统数据为仿真数据/.test(k.textContent || ''));
const calloutOk = !!calloutEl && shdCount(calloutEl) === 0;
const tblTextOk = bodyTblOk && tblShdOk && coverOk && calloutOk;

console.log(`✓ 表格文字格式: 正文表内 run ${bodyRuns.length} 个全为五号宋体+TNR+不加粗=${bodyTblOk} | ` +
  `表内底纹清空=${tblShdOk} | 正文说明框底纹清空=${calloutOk} | 封面表格原样保留=${coverOk}`);
console.log(`  正文表内 run: ${bodyRuns.map((r) => `${r.text}[${r.east}/${
  r.ascii.split(' ')[0]}/${r.sz}${r.bold ? '/粗' : ''}]`).join(' ')}`);
if (!tblTextOk) {
  throw new Error('表格内文字未按附件8 统一为五号宋体 + Times New Roman，或底纹未清/封面被误改');
}

// 25. 页码页脚不得复用包里已有的页脚部件。
//     addFooter 原先是「在 [Content_Types].xml 里找第一条 footer 类型的 Override，拿它的
//     PartName 当 target；部件已经存在就不写内容」，于是复用到的必定是**别人的**页脚 ——
//     封面/摘要那一节自带的空页脚，或者上一版工具留在包里的残留。各分节被指过去之后，
//     正文一页页码都没有，而 counts.footerAdded 仍是 true，app.js 照报「已添加页码」。
//     2026-09-27 探针实测：输入里登记了 word/footer1.xml（空段落、无 PAGE 域）、正文没有
//     任何 footerReference，跑完 2 个分节引用全指向 footer1，含 PAGE 域的引用 0 个。
//     夹具 CYX_*.docx 自带 7 个页脚部件、正文本来就有页脚引用，hasFooterRef 为 true，
//     addFooter 根本不被调用 —— 所以只能靠构造文档来守。
function buildStaleFooterDocx() {
  const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const esc = (s) => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  const P = (t) => `<w:p><w:r><w:t>${esc(t)}</w:t></w:r></w:p>`;
  const stale = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    `<w:ftr xmlns:w="${W}"><w:p><w:pPr><w:jc w:val="center"/></w:pPr></w:p></w:ftr>`;
  const body = P('1 绪  论') + P('正文内容，测试页码用。') + P('1.1 研究背景') + P('正文第二段。');
  const xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    `<w:document xmlns:w="${W}" xmlns:r="${R}"><w:body>${body}` +
    '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:bottom="1440" w:left="1800" w:right="1800"/></w:sectPr>' +
    '</w:body></w:document>';
  const rels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    `<Relationship Id="rId1" Type="${R}/footer" Target="footer1.xml"/></Relationships>`;
  return {
    stale,
    zip: JSZip().file('word/document.xml', xml)
      .file('word/footer1.xml', stale)
      .file('word/_rels/document.xml.rels', rels)
      .file('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
        '<Default Extension="xml" ContentType="application/xml"/>' +
        '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
        '<Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/></Types>')
      .generateAsync({ type: 'nodebuffer' }),
  };
}
const sfIn = await buildStaleFooterDocx();
const sfRes = await FormatTool.formatDocx(await sfIn.zip, FormatTool.DEFAULTS, { format: 'nodebuffer' });
const sfz = await JSZip.loadAsync(sfRes.data);
const sfDocXml = await sfz.file('word/document.xml').async('string');
const sfRels = {};
for (const m of (await sfz.file('word/_rels/document.xml.rels').async('string')).matchAll(/<Relationship\b[^>]*>/g)) {
  const id = (/Id="([^"]+)"/.exec(m[0]) || [])[1];
  const t = (/Target="([^"]+)"/.exec(m[0]) || [])[1];
  if (id) sfRels[id] = String(t || '').replace(/^\.?\//, '').replace(/^word\//, '');
}
const sfRefs = [...sfDocXml.matchAll(/<w:footerReference\b[^>]*>/g)];
let sfWithPage = 0;
const sfTargets = [];
for (const fref of sfRefs) {
  const rid = (/r:id="([^"]+)"/.exec(fref[0]) || [])[1];
  const tgt = sfRels[rid];
  const part = tgt ? 'word/' + tgt : null;
  if (!part || !sfz.file(part)) { sfTargets.push(`${rid}→${part || '?'}(部件缺失)`); continue; }
  const xml = await sfz.file(part).async('string');
  const ok = /PAGE/.test(xml);
  if (ok) sfWithPage++;
  sfTargets.push(`${rid}→${part}${ok ? '(有PAGE)' : '(无PAGE)'}`);
}
const sfStaleKept = (await sfz.file('word/footer1.xml').async('string')) === sfIn.stale;
const sfOk = sfRefs.length > 0 && sfWithPage === sfRefs.length && sfStaleKept;
console.log(`✓ 页码页脚不复用旧部件: ${sfRefs.length} 个分节引用全部指向含 PAGE 域的页脚=` +
  `${sfWithPage === sfRefs.length}/${sfRefs.length} | 旧空页脚未被改写=${sfStaleKept} | ${sfTargets.join(' ')}`);
if (!sfOk) {
  throw new Error('页码页脚复用了包里已有的页脚部件（正文整篇没有页码，却仍报「已添加页码」），或旧部件被改写');
}

console.log('\n全部校验通过 ✔');
