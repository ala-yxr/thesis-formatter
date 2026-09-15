/* 论文格式助手 —— 真实浏览器端到端测试
 * 运行：npm run test:browser （npm test 会连带跑）
 *
 * 为什么需要它：Node 单测用的 @xmldom/xmldom 与浏览器的 DOMParser/XMLSerializer
 * 行为并不一致，已经因此漏过两次事故（都是 Word 打开报「文件已损坏」）：
 *   1. createElement('Override') 在浏览器里会序列化成 <Override xmlns=""/>，
 *      使 [Content_Types].xml 里的声明失效；
 *   2. 浏览器解析 XML 时不把声明建成节点，序列化时又会原样吐出，
 *      于是 serialize() 补的那条声明变成两条。
 * 两者在 xmldom 下都看不出来。产出的 docx 必须在真浏览器里过一遍才作数。
 *
 * 实现说明：页面跑完后用 fetch POST 把报告回传给本脚本，脚本等这个 POST。
 * 早先用的是 --virtual-time-budget + --dump-dom，但虚拟时间不会为挂起的
 * 网络请求停表，常常在 fetch 还没回来时就把 DOM 导出了（报告停在 RUNNING）。
 */
import { spawn } from 'node:child_process';
import http from 'node:http';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const ROOT = path.resolve(__dirname, '../..');
const PORT = 8931;
const TIMEOUT_MS = 180000;

const EDGE_CANDIDATES = [
  'C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe',
  'C:\\Program Files\\Microsoft\\Edge\\Application\\msedge.exe',
  'C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe',
  'C:\\Program Files (x86)\\Google\\Chrome\\Application\\chrome.exe',
];
const browser = EDGE_CANDIDATES.find((p) => fs.existsSync(p));
if (!browser) {
  console.error('未找到 Edge / Chrome，跳过浏览器端到端测试');
  process.exit(1);
}

const candidates = ['CYX_20.0.docx', 'CYX_14.0.docx', 'CYX_12.0.docx', 'CYX_11.0.docx',
  'CYX_10.0.docx', 'CYX_8.0.docx', 'CYX_6.0.docx', 'CYX_5.0.docx', 'CYX_4.0.docx',
  'CYX_3.0.docx', 'CYX_毕业设计.docx', 'CYX2.0.docx', '1.docx'];
const fixture = candidates.map((n) => path.resolve(ROOT, n)).find((f) => fs.existsSync(f));
if (!fixture) {
  console.error('未找到测试文档，请将论文放在 ' + ROOT + ' 下');
  process.exit(1);
}
const fixtureName = path.basename(fixture);

const MIME = {
  '.html': 'text/html; charset=utf-8', '.js': 'text/javascript; charset=utf-8',
  '.mjs': 'text/javascript; charset=utf-8', '.css': 'text/css; charset=utf-8',
  '.docx': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
};

/* 测试页：加载真实的 js/formatter.js，走与网页完全相同的 formatDocx 调用 */
const PAGE = `<!DOCTYPE html><html><head><meta charset="UTF-8"><title>e2e</title></head>
<body><pre id="report">RUNNING</pre>
<script>
var REPORT_LINES = [];
function reportLog(s) {
  REPORT_LINES.push(s);
  document.getElementById('report').textContent = REPORT_LINES.join('\\n');
}
function reportSend(mark) {
  if (mark) reportLog(mark);
  if (window.__sent) return;
  window.__sent = 1;
  try { fetch('/__report', { method: 'POST', body: REPORT_LINES.join('\\n') }); } catch (e) {}
}
window.onerror = function (m, s, l, c) { reportLog('异常: ' + m + ' @' + s + ':' + l + ':' + c); reportSend(); };
window.addEventListener('unhandledrejection', function (e) {
  reportLog('未处理拒绝: ' + (e.reason && e.reason.message)); reportSend();
});
<\/script>
<script src="/论文格式助手/libs/jszip.min.js"><\/script>
<script src="/论文格式助手/js/presets.js"><\/script>
<script src="/论文格式助手/js/formatter.js"><\/script>
<script>
(async function () {
  try {
    if (typeof JSZip === 'undefined') throw new Error('JSZip 未加载');
    if (typeof FormatTool === 'undefined') throw new Error('FormatTool 未加载');

    var buf = await (await fetch('/${encodeURIComponent(fixtureName)}')).arrayBuffer();
    var r = await FormatTool.formatDocx(buf, FormatTool.DEFAULTS, { format: 'blob' });
    var zip = await JSZip.loadAsync(r.data);

    /* 内容不能丢：比对源文件与输出文件的正文文字量。
       曾经出过「导出文档正文全没了」这类问题，只校验 XML 合法性看不出来 */
    function textStats(xmlStr) {
      var d = new DOMParser().parseFromString(xmlStr, 'application/xml');
      var ps = d.getElementsByTagNameNS('http://schemas.openxmlformats.org/wordprocessingml/2006/main', 'p');
      var chars = 0, paras = 0;
      for (var i = 0; i < ps.length; i++) {
        var t = (ps[i].textContent || '').replace(/\\s+/g, '');
        if (t) { paras++; chars += t.length; }
      }
      return { paras: paras, chars: chars };
    }
    var srcStats = textStats(await (await JSZip.loadAsync(buf)).file('word/document.xml').async('string'));
    var outStats = textStats(await zip.file('word/document.xml').async('string'));
    reportLog('正文文字量: 源 ' + srcStats.chars + ' 字/' + srcStats.paras + ' 段 → 输出 ' +
      outStats.chars + ' 字/' + outStats.paras + ' 段（保留 ' +
      (srcStats.chars ? Math.round(outStats.chars / srcStats.chars * 100) : 0) + '%）');

    /* 域字符必须配平：不配平（如只剩一个没有 begin 的 end）时，Word 更新域会
       把其后的内容当成域的一部分，正文整片消失 —— 这就是「正文内容全部消失」 */
    function fieldStats(xmlStr) {
      var d = new DOMParser().parseFromString(xmlStr, 'application/xml');
      var fc = d.getElementsByTagNameNS('http://schemas.openxmlformats.org/wordprocessingml/2006/main', 'fldChar');
      var begin = 0, end = 0, depth = 0, min = 0;
      for (var i = 0; i < fc.length; i++) {
        var t = fc[i].getAttributeNS('http://schemas.openxmlformats.org/wordprocessingml/2006/main', 'fldCharType');
        if (t === 'begin') { begin++; depth++; }
        else if (t === 'end') { end++; depth--; }
        if (depth < min) min = depth;
      }
      return { begin: begin, end: end, depth: depth, min: min };
    }
    var outFld = fieldStats(await zip.file('word/document.xml').async('string'));
    var fldOk = outFld.begin === outFld.end && outFld.depth === 0 && outFld.min === 0;
    reportLog('域字符配平: begin=' + outFld.begin + ' end=' + outFld.end +
      ' 收尾深度=' + outFld.depth + (fldOk ? ' ✓' : ' ✗ 不配平（Word 会吞正文）'));

    var names = Object.keys(zip.files).filter(function (f) {
      return /^word\\/(header|footer)\\d*\\.xml$/.test(f);
    }).sort();
    var parts = ['[Content_Types].xml', 'word/_rels/document.xml.rels', 'word/styles.xml',
                 'word/document.xml', 'word/settings.xml'].concat(names);

    var declBad = [], nsBad = [], missDeclare = [], missRel = [];
    for (var i = 0; i < parts.length; i++) {
      var f = zip.file(parts[i]);
      if (!f) { nsBad.push(parts[i] + '(部件缺失)'); continue; }
      var s = await f.async('string');
      var decls = (s.match(/<\\?xml/g) || []).length;
      if (decls !== 1) declBad.push(parts[i] + '(' + decls + ' 条声明)');
      if (s.indexOf('xmlns=""') >= 0) nsBad.push(parts[i] + '(xmlns="")');
      var d = new DOMParser().parseFromString(s, 'application/xml');
      var err = d.getElementsByTagName('parsererror');
      if (err.length) {
        nsBad.push(parts[i] + '(XML 解析失败: ' +
          (err[0].textContent || '').replace(/\\s+/g, ' ').slice(0, 100) + ')');
        continue;
      }
      var els = d.getElementsByTagName('*');
      for (var k = 0; k < els.length; k++) {
        var e = els[k];
        if (e.namespaceURI) continue;
        var p = e.parentNode;
        if (!p || p.nodeType !== 1 || !p.namespaceURI) continue;
        nsBad.push(parts[i] + '→' + e.nodeName);
      }
    }

    var cd = new DOMParser().parseFromString(
      await zip.file('[Content_Types].xml').async('string'), 'application/xml');
    var declared = {}, ov = cd.getElementsByTagName('*');
    for (var m = 0; m < ov.length; m++) {
      if (ov[m].localName === 'Override') declared[ov[m].getAttribute('PartName')] = 1;
    }
    var rd = new DOMParser().parseFromString(
      await zip.file('word/_rels/document.xml.rels').async('string'), 'application/xml');
    var targets = {}, rs = rd.getElementsByTagName('*');
    for (var q = 0; q < rs.length; q++) {
      if (rs[q].localName === 'Relationship') targets[rs[q].getAttribute('Target')] = 1;
    }
    for (var j = 0; j < names.length; j++) {
      if (!declared['/' + names[j]]) missDeclare.push(names[j]);
      if (!targets[names[j].replace(/^word\\//, '')]) missRel.push(names[j]);
    }

    reportLog('源文件: ${fixtureName} (' + buf.byteLength + ' 字节)');
    reportLog('工具版本: ' + FormatTool.VERSION + ' | 输出: ' + r.data.size + ' 字节');
    reportLog('空命名空间节点计数: ' + r.counts.emptyNsNodes);
    reportLog('XML 声明数异常: ' + (declBad.length ? declBad.join(', ') : '无'));
    reportLog('命名空间异常: ' + (nsBad.length ? nsBad.join(', ') : '无'));
    reportLog('Content_Types 未声明: ' + (missDeclare.length ? missDeclare.join(', ') : '无'));
    reportLog('rels 未关联: ' + (missRel.length ? missRel.join(', ') : '无'));
    reportLog('页眉/页脚部件: ' + names.length + ' 个 | bodyHeader=' + r.counts.bodyHeader);

    /* 回归：段落之间带换行的 docx。
       Word 存出来是紧凑的一行，所以夹具永远测不到这个形态 —— 2026-09-15 就是这么
       漏掉了「格式助手处理真实产物直接抛 TypeError」：按缩进排过版的工具（本项目
       的 md2docx 就是）会在标签之间留换行，解析后它们是**空白文本节点**，而
       replaceTocWithField 拿 nextSibling 当元素用，一调 getElementsByTagNameNS 就炸。
       这里把夹具每个标签之间插上换行（/>/ 后面紧跟 /</ 的地方），必须一样跑通。 */
    var zin = await JSZip.loadAsync(buf);
    var dx = await zin.file('word/document.xml').async('string');
    zin.file('word/document.xml', dx.replace(/></g, '>\\n<'));
    var spacedBuf = await zin.generateAsync({ type: 'arraybuffer', compression: 'STORE' });
    var rSp = await FormatTool.formatDocx(spacedBuf, FormatTool.DEFAULTS, { format: 'blob' });
    var spXml = await (await JSZip.loadAsync(rSp.data)).file('word/document.xml').async('string');
    var spStats = textStats(spXml), spFld = fieldStats(spXml);
    reportLog('带换行变体: 输出 ' + spStats.chars + ' 字/' + spStats.paras + ' 段（保留 ' +
      (srcStats.chars ? Math.round(spStats.chars / srcStats.chars * 100) : 0) + '%）' +
      ' | 域配平 begin=' + spFld.begin + ' end=' + spFld.end + ' 收尾深度=' + spFld.depth);
  } catch (e) {
    reportLog('异常: ' + (e && e.message));
  }
  reportLog('DONE');
  reportSend();
})();
setTimeout(function () { reportSend('看门狗: 超时未完成'); }, 150000);
<\/script></body></html>`;

let onReport;
const reported = new Promise((r) => { onReport = r; });

const server = http.createServer((req, res) => {
  const url = decodeURIComponent(new URL(req.url, 'http://localhost').pathname);
  if (req.method === 'POST' && url === '/__report') {
    const chunks = [];
    req.on('data', (c) => chunks.push(c));
    req.on('end', () => {
      res.writeHead(200, { 'Content-Type': 'text/plain' });
      res.end('ok');
      onReport(Buffer.concat(chunks).toString('utf8'));
    });
    return;
  }
  if (url === '/__e2e') {
    res.writeHead(200, { 'Content-Type': 'text/html; charset=utf-8' });
    res.end(PAGE);
    return;
  }
  const fp = path.normalize(path.join(ROOT, url));
  if (!fp.startsWith(ROOT)) { res.writeHead(403); res.end('forbidden'); return; }
  fs.readFile(fp, (err, data) => {
    if (err) { res.writeHead(404); res.end('not found'); return; }
    res.writeHead(200, { 'Content-Type': MIME[path.extname(fp).toLowerCase()] || 'application/octet-stream' });
    res.end(data);
  });
});

await new Promise((r) => server.listen(PORT, r));

/* 独立的 profile 目录，避免复用用户正在用的浏览器进程；放在临时目录里 */
const profile = fs.mkdtempSync(path.join(os.tmpdir(), 'thesis-e2e-'));
const ps = spawn(browser, ['--headless=new', '--disable-gpu', '--no-sandbox',
  '--no-first-run', '--user-data-dir=' + profile,
  `http://127.0.0.1:${PORT}/__e2e`], { stdio: 'ignore' });

const timer = setTimeout(() => onReport('异常: 浏览器 ' + (TIMEOUT_MS / 1000) + ' 秒内没有回传报告\nDONE'), TIMEOUT_MS);
const report = (await reported).replace(/\r/g, '');
clearTimeout(timer);
try { ps.kill(); } catch (e) {}
try { server.close(); } catch (e) {}
try { fs.rmSync(profile, { recursive: true, force: true }); } catch (e) {}

console.log('===== 浏览器端到端（真实 DOMParser/XMLSerializer）=====');
console.log(report);

if (!/^DONE$/m.test(report)) {
  console.error('\n✗ 浏览器端到端测试未跑完');
  process.exit(1);
}
/* 必须逐条出现「无」的检查项（注意别用单个「异常」二字做匹配：
   「XML 声明数异常: 无」里也含「异常」，会误报） */
const MUST_CLEAN = ['空命名空间节点计数: 0', 'XML 声明数异常: 无', '命名空间异常: 无',
  'Content_Types 未声明: 无', 'rels 未关联: 无'];
/* 域字符配平（浏览器端 DOMParser 下再验一次：域修复逻辑跑在浏览器里） */
const fldMatch = report.match(/域字符配平: begin=(\d+) end=(\d+) 收尾深度=(-?\d+) (✓|✗)/);
if (!fldMatch || fldMatch[4] !== '✓' || fldMatch[1] !== fldMatch[2] || fldMatch[3] !== '0') {
  console.error('\n✗ 输出文档的域字符不配平（Word 打开会吞掉正文）：' + (fldMatch ? fldMatch[0] : '检查项缺失'));
  process.exit(1);
}
const missing = MUST_CLEAN.filter((s) => report.indexOf(s) < 0);
/* 正文文字量：输出不得少于源文件的 90%（目录域替换、删中西文空格会少一点点） */
const charsLine = (report.match(/正文文字量: 源 (\d+) 字.*→ 输出 (\d+) 字/) || []);
const srcChars = Number(charsLine[1] || 0), outChars = Number(charsLine[2] || 0);
if (!srcChars || outChars < srcChars * 0.9) {
  console.error('\n✗ 正文内容疑似丢失：源 ' + srcChars + ' 字 → 输出 ' + outChars + ' 字');
  process.exit(1);
}
/* 带换行变体（标签之间留空的 docx）：夹具是紧凑的，测不到这个形态，单列一条。
   2026-09-15 真实产物就是栽在这里 —— 用户丢进去"无法修改"，而测试全绿。 */
const spLine = (report.match(/带换行变体: 输出 (\d+) 字/) || []);
const spChars = Number(spLine[1] || 0);
if (!spChars || spChars < srcChars * 0.9) {
  console.error('\n✗ 带换行的文档处理失败：' + (spLine[0] || '检查项缺失'));
  process.exit(1);
}
const badLines = report.split('\n').filter((l) =>
  /^异常: /.test(l) || /看门狗/.test(l) || /解析失败|xmlns=""|部件缺失/.test(l));
if (missing.length || badLines.length) {
  console.error('\n✗ 浏览器端到端测试失败：');
  missing.forEach((s) => console.error('  缺少检查项: ' + s));
  badLines.forEach((l) => console.error('  ' + l));
  process.exit(1);
}
console.log('\n✓ 浏览器端到端通过（各部件 XML 合法、命名空间正确、页眉页脚已声明并关联）');
