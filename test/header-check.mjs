/* 论文格式助手 —— 页眉精准检测脚本
 * 用法: node test/header-check.mjs [docx路径]
 * 默认取项目根目录里版本号最大的 CYX_<n>.<n>.docx
 *   （原先写死 d:/毕业论文/CYX_20.0.docx，该文件已不在根目录，直接跑会报"文档不存在"。
 *     这里改成与 run-test.mjs / browser-e2e.mjs 同一套扫描口径，2026-09-26。）
 *
 * 正确检测要点(修正历史审查方法的缺陷):
 * 1. 遍历所有 headerReference,解析 rels 得到每个被引用 header 部件
 * 2. 检测段落内【所有】 run 的格式(不只第一个 run)
 * 3. 字号/字体未显式设置时,解析 styles.xml 的样式继承
 *    (run → 段落样式 → header 样式 → Normal → docDefaults),给出实际渲染值
 * 4. 完整输出:文字(含章名)、字号(含继承)、字体、下划线、对齐
 *
 * 【已知差异】header 内置样式的 rPr 若定义 sz=18(9pt),纯 XML 继承解析会得出
 * 9pt,但 Word 实际渲染页眉时忽略 header 样式的字号,跟随 Normal 样式
 * (CYX_20.0 实测:XML 解析 9pt,Word 显示五号 10.5pt)。
 * 结论:页眉字号以 Word 实际渲染为准,解析值仅供参照。
 */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { DOMParser } from '@xmldom/xmldom';
import JSZip from 'jszip';

const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const ROOT = path.resolve(__dirname, '../..');

/* 与 run-test.mjs / browser-e2e.mjs 同口径：先扫版本号最大的 CYX_<n>.<n>.docx */
function defaultFixture() {
  const ver = (n) => { const m = /^CYX_(\d+)\.(\d+)\.docx$/.exec(n); return m ? +m[1] * 1000 + +m[2] : -1; };
  let names;
  try { names = fs.readdirSync(ROOT); } catch (e) { return null; }
  const best = names.filter((n) => ver(n) >= 0).sort((a, b) => ver(b) - ver(a))[0];
  return best ? path.join(ROOT, best) : null;
}

const target = process.argv[2] || defaultFixture();

if (!target) {
  console.error('未指定文档，且 ' + ROOT + ' 下找不到 CYX_<n>.<n>.docx');
  process.exit(1);
}
if (!fs.existsSync(target)) {
  console.error('文档不存在:', target);
  process.exit(1);
}
console.log('夹具:', path.basename(target));

const zip = await JSZip.loadAsync(fs.readFileSync(target));
const doc = new DOMParser().parseFromString(await zip.file('word/document.xml').async('string'), 'application/xml');
const relsDoc = new DOMParser().parseFromString(await zip.file('word/_rels/document.xml.rels').async('string'), 'application/xml');

/* ---------- 样式继承解析 ---------- */
const child = (e, n) => {
  if (!e || !e.childNodes) return null;
  return Array.from(e.childNodes).find(x => x.nodeType === 1 && x.localName === n);
};
/* 继承链: run 自身 rPr → header 样式 → Normal 样式 → docDefaults */
const styleInherit = { header: null, normal: null, defaults: null };
if (zip.file('word/styles.xml')) {
  const stDoc = new DOMParser().parseFromString(await zip.file('word/styles.xml').async('string'), 'application/xml');
  // docDefaults
  const dd = stDoc.getElementsByTagNameNS(W_NS, 'docDefaults')[0];
  const rPrDef = dd && dd.getElementsByTagNameNS(W_NS, 'rPrDefault')[0];
  const defRPr = rPrDef && rPrDef.getElementsByTagNameNS(W_NS, 'rPr')[0];
  if (defRPr) {
    const sz = child(defRPr, 'sz');
    const rf = child(defRPr, 'rFonts');
    styleInherit.defaults = {
      sz: sz ? Number(sz.getAttributeNS(W_NS, 'val')) : null,
      eastAsia: rf ? rf.getAttributeNS(W_NS, 'eastAsia') : null,
      ascii: rf ? rf.getAttributeNS(W_NS, 'ascii') : null,
    };
  }
  // 按样式 name 匹配(styleId 可能是自定义值,不能只按 id)
  for (const st of Array.from(stDoc.getElementsByTagNameNS(W_NS, 'style'))) {
    const nm = st.getElementsByTagNameNS(W_NS, 'name')[0];
    const name = (nm && nm.getAttributeNS(W_NS, 'val') || '').toLowerCase();
    if (!['normal', 'header'].includes(name)) continue;
    const rPr = child(st, 'rPr');
    if (!rPr) continue;
    const sz = child(rPr, 'sz');
    const rf = child(rPr, 'rFonts');
    const info = {
      sz: sz ? Number(sz.getAttributeNS(W_NS, 'val')) : null,
      eastAsia: rf ? rf.getAttributeNS(W_NS, 'eastAsia') : null,
      ascii: rf ? rf.getAttributeNS(W_NS, 'ascii') : null,
    };
    if (name === 'header' && !styleInherit.header) styleInherit.header = info;
    if (name === 'normal' && !styleInherit.normal) styleInherit.normal = info;
  }
}

/* ---------- 解析一个 run 的实际格式(含继承) ---------- */
function resolveRun(rPr) {
  const out = { sz: null, eastAsia: null, ascii: null, u: false, b: false, inherited: false };
  if (rPr) {
    const sz = child(rPr, 'sz');
    const rf = child(rPr, 'rFonts');
    if (sz) out.sz = Number(sz.getAttributeNS(W_NS, 'val'));
    if (rf) {
      out.eastAsia = rf.getAttributeNS(W_NS, 'eastAsia');
      out.ascii = rf.getAttributeNS(W_NS, 'ascii');
    }
    out.u = !!child(rPr, 'u');
    out.b = !!child(rPr, 'b');
  }
  // 继承链补全(header 样式 → Normal → docDefaults)
  for (const layer of [styleInherit.header, styleInherit.normal, styleInherit.defaults]) {
    if (!layer) continue;
    if (out.sz === null && layer.sz !== null) { out.sz = layer.sz; out.inherited = true; }
    if (!out.eastAsia && layer.eastAsia) { out.eastAsia = layer.eastAsia; out.inherited = true; }
    if (!out.ascii && layer.ascii) { out.ascii = layer.ascii; out.inherited = true; }
  }
  return out;
}

/* ---------- 主检测 ---------- */
const rid2target = {};
Array.from(relsDoc.getElementsByTagName('Relationship')).forEach(x => {
  rid2target[x.getAttribute('Id')] = x.getAttribute('Target');
});

console.log('===== 页眉检测: ' + path.basename(target) + ' =====');
const hdrs = doc.getElementsByTagNameNS(W_NS, 'headerReference');
console.log('headerReference 数量:', hdrs.length);

let headerCount = 0, emptyHeaders = 0;
const seen = new Set();
for (const h of Array.from(hdrs)) {
  const type = h.getAttributeNS(W_NS, 'type') || 'default';
  const rid = h.getAttributeNS(R_NS, 'id');
  const tgt = rid2target[rid];
  if (!tgt) { console.log('  ✗ 引用 ' + rid + ' 无目标'); continue; }
  const f2 = (tgt.indexOf('word/') === 0 ? tgt : 'word/' + tgt).replace(/^\//, '');
  if (seen.has(f2) || !zip.file(f2)) continue;
  seen.add(f2);
  headerCount++;

  const hDoc = new DOMParser().parseFromString(await zip.file(f2).async('string'), 'application/xml');
  const ps = Array.from(hDoc.getElementsByTagNameNS(W_NS, 'p'));
  const nonEmpty = ps.find(p => Array.from(p.getElementsByTagNameNS(W_NS, 't')).map(x => x.textContent).join('').trim());
  if (!nonEmpty) { emptyHeaders++; console.log('  ' + f2 + ' (type=' + type + '): (空页眉)'); continue; }

  const pPr = child(nonEmpty, 'pPr');
  const jc = pPr && child(pPr, 'jc');
  const pBdr = pPr && child(pPr, 'pBdr');
  const fullText = Array.from(nonEmpty.getElementsByTagNameNS(W_NS, 't')).map(x => x.textContent).join('').trim();

  console.log('  ' + f2 + ' (type=' + type + '): 「' + fullText.slice(0, 40) + '」');
  console.log('      对齐=' + (jc ? jc.getAttributeNS(W_NS, 'val') : '(未设/继承)') + ' 段落边框(页眉线)=' + (pBdr ? '有' : '无') + ' 文字下划线=见各 run');

  // 【关键】遍历所有 run,不只第一个
  for (const r of Array.from(nonEmpty.childNodes).filter(x => x.nodeType === 1 && (x.localName === 'r' || x.localName === 'hyperlink'))) {
    const rt = Array.from(r.getElementsByTagNameNS(W_NS, 't')).map(x => x.textContent).join('');
    const rPr = child(r, 'rPr') || (r.localName === 'hyperlink' ? (() => { const rr = child(r, 'r'); return rr && child(rr, 'rPr'); })() : null);
    const fmt = resolveRun(rPr);
    if (!rt.trim() && !rPr) continue;
    console.log('      run「' + (rt.slice(0, 14) || '(空白)') + '」' +
      ' 字号=' + (fmt.sz ? fmt.sz / 2 : '?') + 'pt' + (fmt.sz === null ? '(继承)' : '') +
      (fmt.sz === 18 ? '(注: header样式9pt,Word实际按Normal五号渲染)' : '') +
      ' 中文=' + (fmt.eastAsia || '?') + ' 西文=' + (fmt.ascii || '?') +
      (fmt.u ? ' 下划线' : '') + (fmt.b ? ' 加粗' : ''));
  }
}

console.log('===== 汇总 =====');
console.log('被引用页眉:', headerCount, '| 空页眉(封面区):', emptyHeaders);
console.log('检测依据(附件8): 页眉五号宋体(10.5pt)、页眉之下有下划线、左对齐校名/右对齐章名');
