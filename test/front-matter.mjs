/* 「套用模板前置页」回归测试 —— 合成夹具，不依赖作者的真实论文
 *
 * 造两个内存里的最小 docx：
 *   模板：封面（含校徽图片 + 一个封面表格）+ 原创性声明页，末尾带 sectPr；
 *         并引用 3 个页眉页脚 —— 2 个空的、1 个有字（用来验证「有内容会被告知」）
 *   目标：一份自带旧封面的论文，摘要起是正文
 * 然后把模板前置页套到目标上，断言：
 *   ① 前置页逐元素顺序与模板完全一致（不是倒序、不是乱序）
 *   ② 校徽图片真搬过来了（文件 + 关系 + rId 不悬空）
 *   ③ 前置节的页面设置保持模板原样，且不编页码
 *   ④ 目标自己的旧封面被替换掉
 *   ⑤ 正文从摘要起，页码自罗马数字 1 开始
 * 运行：npm test
 */
import { DOMParser, XMLSerializer } from '@xmldom/xmldom';
import JSZip from 'jszip';

globalThis.DOMParser = DOMParser;
globalThis.XMLSerializer = XMLSerializer;
globalThis.JSZip = JSZip;

const FormatTool = (await import('../js/formatter.js')).default;

const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const PKG_REL_NS = 'http://schemas.openxmlformats.org/package/2006/relationships';
const CT_NS = 'http://schemas.openxmlformats.org/package/2006/content-types';
const CT_DOC = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml';
const CT_HEADER = 'application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml';
const CT_FOOTER = 'application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml';
const CT_STYLES = 'application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml';

/* 1×1 的合法 PNG，够用来验证「二进制被原样搬运」 */
const PNG_B64 =
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg==';

const W_DECL = 'xmlns:w="' + W_NS + '" xmlns:r="' + R_NS + '" ' +
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
  'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" ' +
  'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"';

const p = (text) => '<w:p><w:r><w:t xml:space="preserve">' + text + '</w:t></w:r></w:p>';
const head = (text) => '<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr>' +
  '<w:r><w:t>' + text + '</w:t></w:r></w:p>';
const sectPr = (left, extra) => '<w:sectPr>' + (extra || '') +
  '<w:pgSz w:w="11906" w:h="16838"/>' +
  '<w:pgMar w:top="1440" w:right="1418" w:bottom="1440" w:left="' + left + '" ' +
  'w:header="851" w:footer="992" w:gutter="0"/></w:sectPr>';
/* 带 sectPr 的段落 —— 表示「本段落之前构成一节」 */
const pSect = (text, left, extra) =>
  '<w:p><w:pPr>' + sectPr(left, extra) + '</w:pPr>' +
  (text ? '<w:r><w:t xml:space="preserve">' + text + '</w:t></w:r>' : '') + '</w:p>';

/* 封面上的校徽：现代 DrawingML 图片，走 a:blip 的 r:embed */
const logoP =
  '<w:p><w:r><w:drawing><wp:inline><wp:extent cx="914400" cy="914400"/>' +
  '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
  '<pic:pic><pic:nvPicPr><pic:cNvPr id="1" name="logo.png"/><pic:cNvPicPr/></pic:nvPicPr>' +
  '<pic:blipFill><a:blip r:embed="rIdImg1"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>' +
  '<pic:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic>' +
  '</a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>';

const COVER_TBL =
  '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr>' +
  '<w:tr><w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr>' +
  '<w:p><w:r><w:t>学    院</w:t></w:r></w:p></w:tc>' +
  '<w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr>' +
  '<w:p><w:r><w:t>大数据与人工智能学院</w:t></w:r></w:p></w:tc></w:tr></w:tbl>';

/* ---------- 模板 ---------- */
/* 前置页 7 个元素，顺序刻意有辨识度：空的 → 空 → 图 → 校名 → 表格 → 声明页 → 带 sectPr 的空段 */
const TPL_FRONT = [
  p(''),                                    // 1 顶端空段
  p('重庆工程学院'),                         // 2 校名
  logoP,                                    // 3 校徽
  p('毕业设计（论文）'),                      // 4
  COVER_TBL,                                // 5 封面信息表
  p('学生毕业设计（论文）原创性声明'),          // 6 声明页
  pSect('', 1418, '<w:headerReference w:type="default" r:id="rIdH1"/>' +
    '<w:headerReference w:type="even" r:id="rIdH2"/>' +
    '<w:footerReference w:type="default" r:id="rIdF1"/>')  // 7 前置节到此结束
];
const TPL_AFTER = [
  head('摘  要'),
  p('模板的摘要正文，套用后应被目标自己的摘要取代。'),
  p('关键词：甲；乙'),
  pSect('', 1418)
];

const templateXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
  '<w:document ' + W_DECL + '><w:body>' + TPL_FRONT.join('') + TPL_AFTER.join('') + '</w:body></w:document>';

const tplRels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
  '<Relationships xmlns="' + PKG_REL_NS + '">' +
  '<Relationship Id="rIdImg1" Type="' + R_NS + '/image" Target="media/logo.png"/>' +
  '<Relationship Id="rIdH1" Type="' + R_NS + '/header" Target="header1.xml"/>' +
  '<Relationship Id="rIdH2" Type="' + R_NS + '/header" Target="header2.xml"/>' +
  '<Relationship Id="rIdF1" Type="' + R_NS + '/footer" Target="footer1.xml"/>' +
  '</Relationships>';

const emptyHf = (tag) => '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
  '<w:' + tag + ' ' + W_DECL + '><w:p/></w:' + tag + '>';
/* header2 有字 —— 提取时要被告知「没搬」，但不能静默丢 */
const header2Xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
  '<w:hdr ' + W_DECL + '>' + p('重庆工程学院') + '</w:hdr>';

/* ---------- 目标 ---------- */
const TGT_OLD_COVER = [
  p('旧封面标题不该留痕'),
  p('旧封面副标题不该留痕')
];
const TGT_BODY = [
  head('摘  要'),
  p('本文研究二手车交易数据的分析与可视化。'),
  p('关键词：二手车；数据分析；可视化'),
  pSect('', 1418),
  head('1 绪论'),
  p('绪论正文。'),
  pSect('', 1418)
];
const targetXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
  '<w:document ' + W_DECL + '><w:body>' + TGT_OLD_COVER.join('') + TGT_BODY.join('') +
  '</w:body></w:document>';

const targetRels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
  '<Relationships xmlns="' + PKG_REL_NS + '">' +
  '<Relationship Id="rId1" Type="' + R_NS + '/styles" Target="styles.xml"/>' +
  '</Relationships>';

const stylesXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
  '<w:styles ' + W_DECL + '>' +
  '<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/>' +
  '<w:pPr><w:outlineLvl w:val="0"/></w:pPr></w:style>' +
  '</w:styles>';

function ctXml(parts) {
  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
    '<Types xmlns="' + CT_NS + '">' +
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
    '<Default Extension="xml" ContentType="application/xml"/>' +
    '<Default Extension="png" ContentType="image/png"/>' +
    '<Override PartName="/word/document.xml" ContentType="' + CT_DOC + '"/>' +
    '<Override PartName="/word/styles.xml" ContentType="' + CT_STYLES + '"/>' +
    parts.map(([n, ct]) => '<Override PartName="/word/' + n + '" ContentType="' + ct + '"/>').join('') +
    '</Types>';
}

const rootRels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
  '<Relationships xmlns="' + PKG_REL_NS + '">' +
  '<Relationship Id="rId1" Type="' + R_NS + '/officeDocument" Target="word/document.xml"/>' +
  '</Relationships>';

async function buildDocx(entries) {
  const z = new JSZip();
  for (const [name, val] of Object.entries(entries)) {
    z.file(name, val, val instanceof Uint8Array ? {} : undefined);
  }
  return z.generateAsync({ type: 'nodebuffer' });
}

const tplBuf = await buildDocx({
  '[Content_Types].xml': ctXml([['header1.xml', CT_HEADER], ['header2.xml', CT_HEADER],
    ['footer1.xml', CT_FOOTER]]),
  '_rels/.rels': rootRels,
  'word/document.xml': templateXml,
  'word/_rels/document.xml.rels': tplRels,
  'word/styles.xml': stylesXml,
  'word/header1.xml': emptyHf('hdr'),
  'word/header2.xml': header2Xml,
  'word/footer1.xml': emptyHf('ftr'),
  'word/media/logo.png': Buffer.from(PNG_B64, 'base64')
});

const tgtBuf = await buildDocx({
  '[Content_Types].xml': ctXml([]),
  '_rels/.rels': rootRels,
  'word/document.xml': targetXml,
  'word/_rels/document.xml.rels': targetRels,
  'word/styles.xml': stylesXml
});

/* ---------- 断言小工具 ---------- */
let passed = 0;
const checks = [];
function ok(label, cond, detail) {
  checks.push({ label, cond, detail });
  if (cond) passed++;
  console.log((cond ? '  ✓ ' : '  ✗ ') + label + (detail ? '  ' + detail : ''));
}

/* ---------- 1. 提取 ---------- */
console.log('===== 1. 提取模板前置页 =====');
const front = await FormatTool.extractFrontMatter(tplBuf);
ok('前置页元素数 = 7', front.els === 7, '实际 ' + front.els);
ok('界桩识别为「摘要」', front.boundary === '摘要', '实际「' + front.boundary + '」');
ok('校徽图片被提取', front.media.length === 1 && front.media[0].ext === '.png',
  front.media.map((m) => m.rid + m.ext).join(','));
ok('有字的页眉被告知未搬运', front.ignoredParts.indexOf('页眉') >= 0,
  '未搬运: ' + JSON.stringify(front.ignoredParts));
ok('空页脚不算「有内容」', front.ignoredParts.indexOf('页脚') < 0);

/* ---------- 2. 套用 ---------- */
console.log('\n===== 2. 套用并格式化 =====');
const res = await FormatTool.formatDocx(tgtBuf, FormatTool.DEFAULTS,
  { format: 'nodebuffer', templateFront: front });
ok('插入元素数 = 7', res.counts.frontMatter === 7, '实际 ' + res.counts.frontMatter);

/* ---------- 3. 校验输出 ---------- */
console.log('\n===== 3. 校验输出 =====');
const zip = await JSZip.loadAsync(res.data);
const outXml = await zip.file('word/document.xml').async('string');
const outDoc = new DOMParser().parseFromString(outXml, 'application/xml');
const all = (n, name) => n.getElementsByTagNameNS(W_NS, name);
const kid = (n, name) => Array.from(n.childNodes).find((c) => c.nodeType === 1 && c.localName === name);
const body = all(outDoc, 'body')[0];
const tops = Array.from(body.childNodes).filter((n) => n.nodeType === 1);
const textOf = (el) => Array.from(all(el, 't')).map((t) => t.textContent).join('');

/* 3a. 前置页逐元素顺序 —— 这是本功能的核心承诺：「与模板前两页相同」 */
const tplDoc = new DOMParser().parseFromString(templateXml, 'application/xml');
const tplTops = Array.from(all(tplDoc, 'body')[0].childNodes)
  .filter((n) => n.nodeType === 1).slice(0, 7);
const orderOk = tplTops.every((t, i) => {
  const o = tops[i];
  if (!o || o.localName !== t.localName) return false;
  if (t.localName === 'tbl') return true;
  const hasImg = all(t, 'drawing').length > 0;
  return hasImg ? all(o, 'drawing').length > 0 : textOf(o) === textOf(t);
});
ok('前置页 7 个元素逐元素顺序一致', orderOk,
  tops.slice(0, 7).map((e) => e.localName === 'tbl' ? '[表]' : JSON.stringify(textOf(e).slice(0, 10))).join(' '));
ok('第 3 个元素是校徽图片（顺序未反转）', all(tops[2], 'drawing').length > 0);
ok('第 5 个元素是封面表格', tops[4].localName === 'tbl');

/* 3b. 图片真的搬过来了 */
const media = Object.keys(zip.files).filter((f) => /^word\/media\/front\d+\./.test(f));
ok('图片落盘 word/media/front1.png', media.length === 1 && /front1\.png$/.test(media[0]), media.join(','));
const outRelsXml = await zip.file('word/_rels/document.xml.rels').async('string');
const relDoc = new DOMParser().parseFromString(outRelsXml, 'application/xml');
const rels = relDoc.getElementsByTagName('Relationship');
const known = new Set();
const tgtOf = {};
for (let i = 0; i < rels.length; i++) {
  known.add(rels[i].getAttribute('Id'));
  tgtOf[rels[i].getAttribute('Id')] = rels[i].getAttribute('Target');
}
const used = new Set();
for (const m of outXml.matchAll(/r:(?:embed|id)="([^"]+)"/g)) used.add(m[1]);
const dangling = [...used].filter((u) => !known.has(u));
ok('无悬空 rId 引用', dangling.length === 0, dangling.join(','));
const imgRid = tops[2].getElementsByTagNameNS(A_NS, 'blip')[0].getAttributeNS(R_NS, 'embed');
ok('校徽指向新关系', !!imgRid && /front1\.png$/.test(tgtOf[imgRid] || ''),
  imgRid + ' → ' + (tgtOf[imgRid] || '(无)'));

/* 3c. 前置节的页面设置不动、不编页码 */
const sects = all(outDoc, 'sectPr');
const frontSect = sects[0];
const fMar = kid(frontSect, 'pgMar');
ok('前置节页边距保持模板的 1418', fMar.getAttributeNS(W_NS, 'left') === '1418',
  '实际 ' + fMar.getAttributeNS(W_NS, 'left'));
ok('前置节没有 pgNumType（封面声明页不编页码）', !kid(frontSect, 'pgNumType'));
ok('前置节没有页眉页脚引用（模板的引用已删）',
  all(frontSect, 'headerReference').length === 0 && all(frontSect, 'footerReference').length === 0);

/* 3d. 目标自己的旧封面被替换 */
ok('旧封面段落已删除', outXml.indexOf('旧封面') < 0);
ok('模板的声明页在（前置页照搬成功）', textOf(tops[5]).indexOf('原创性声明') >= 0);
ok('校名不再是模板里那个孤立段落', textOf(tops[1]) === '重庆工程学院');

/* 3e. 正文从摘要起，页码从罗马 1 开始 */
const absIdx = tops.findIndex((e) => e.localName === 'p' && /^摘\s*要$/.test(textOf(e).trim()));

/* 最强的一条：单独格式化模板得到的前置区，与套用后的前置区必须逐元素一致。
   这正是「前两页与模板相同」的字面含义，而且格式化流程自己插的分隔空段
   两边都会插，比对不受它影响。 */
const tplAlone = await FormatTool.formatDocx(tplBuf, FormatTool.DEFAULTS, { format: 'nodebuffer' });
const za = await JSZip.loadAsync(tplAlone.data);
const zd = new DOMParser().parseFromString(await za.file('word/document.xml').async('string'), 'application/xml');
const zb = zd.getElementsByTagNameNS(W_NS, 'body')[0];
const zTops = Array.from(zb.childNodes).filter((n) => n.nodeType === 1);
const zAbsIdx = zTops.findIndex((e) => e.localName === 'p' && /^摘\s*要$/.test(textOf(e).trim()));
const shape = (el) => all(el, 'drawing').length ? '<图>' :
  (el.localName === 'tbl' ? '<表>' : textOf(el));
const frontAlone = zTops.slice(0, zAbsIdx).map(shape).join('¶');
const frontOut = tops.slice(0, absIdx).map(shape).join('¶');
ok('前置区与「模板单独格式化」逐元素一致', frontAlone === frontOut && frontAlone.length > 0,
  '模板 ' + frontAlone.length + ' 字 vs 输出 ' + frontOut.length + ' 字');
if (frontAlone !== frontOut) {
  console.log('    模板: ' + frontAlone.slice(0, 200));
  console.log('    输出: ' + frontOut.slice(0, 200));
}
ok('摘要标题在前置页之后', absIdx > 0 && zAbsIdx === absIdx, '模板下标 ' + zAbsIdx + ' / 输出下标 ' + absIdx);
const absSect = sects[1];
const absNum = kid(absSect, 'pgNumType');
ok('摘要节页码为罗马数字、从 1 开始',
  !!absNum && absNum.getAttributeNS(W_NS, 'fmt') === 'upperRoman' &&
  absNum.getAttributeNS(W_NS, 'start') === '1',
  absNum ? absNum.getAttributeNS(W_NS, 'fmt') + ' start=' + absNum.getAttributeNS(W_NS, 'start') : '(无)');
ok('摘要节已有页脚', all(absSect, 'footerReference').length > 0);
ok('目标自己的摘要正文还在', outXml.indexOf('本文研究二手车交易数据') >= 0);

/* 3f. 输出能被 Word 打开的前提：命名空间都绑上了 */
ok('输出里没有空命名空间节点', outXml.indexOf('xmlns=""') < 0);

console.log('\n' + passed + '/' + checks.length + ' 项断言通过');
if (passed !== checks.length) {
  console.error('\n✗ 失败项：' + checks.filter((x) => !x.cond).map((x) => x.label).join('；'));
  process.exit(1);
}
console.log('前置页套用校验通过 ✔');
