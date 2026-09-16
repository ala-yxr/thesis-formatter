/* ============================================================
 * 论文格式助手 —— 核心格式化引擎（纯前端）
 *
 * 原理：直接解析并改写 .docx 内部的 OOXML（word/document.xml 等），
 *       保留原文档中的图片、公式(OMML)、表格与所有内容，仅重排格式。
 * 依赖：JSZip（解压/压缩）、浏览器原生 DOMParser / XMLSerializer。
 *
 * 用法：
 *   const r = await FormatTool.formatDocx(arrayBuffer, settings, {format:'blob'});
 *   r.data   → 格式化后的 docx（Blob 或 nodebuffer）
 *   r.counts → 结构统计 {h1: 5, body: 120, ...}
 *   r.info   → 识别信息 {absDetected: true, ...}
 * ============================================================ */
(function (global) {
  'use strict';

  var W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  var R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  var M_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/math';
  var XML_NS = 'http://www.w3.org/XML/1998/namespace';
  var CT_NS = 'http://schemas.openxmlformats.org/package/2006/content-types';
  var PKG_REL_NS = 'http://schemas.openxmlformats.org/package/2006/relationships';
  var CT_FOOTER = 'application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml';
  var REL_FOOTER = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer';
  var CT_HEADER = 'application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml';
  var REL_HEADER = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/header';
  var MIME_DOCX = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';

  /* ---------- 默认设置（重庆工程学院本科毕业设计（论文）撰写格式要求，附件8） ---------- */
  var DEFAULTS = {
    marginTop: 2.5,     // 上边距 cm（附件8：上、下、左、右均为 2.5cm）
    marginBottom: 2.5,
    marginLeft: 2.5,
    marginRight: 2.5,

    bodyFont: '宋体',              // 正文字体
    latinFont: 'Times New Roman',  // 西文字体（附件8：所有数字和字母一律 Times New Roman）
    bodySize: 12,                  // 正文字号 pt（小四 = 12）
    lineSpacing: 20,               // 行距：fixed 时单位磅 / multiple 时为倍数
    lineSpacingMode: 'exact',      // 'exact' 固定值20磅 | 'multiple' 倍数
    firstLineChars: 2,             // 首行缩进（字符）

    titleFont: '黑体',
    titleSize: 22,                 // 论文题目（二号 = 22pt）
    titleBold: false,

    h1Font: '黑体',
    h1Size: 16,                    // 一级标题（三号 = 16pt，居中）
    h1Jc: 'center',
    h2Font: '黑体',
    h2Size: 15,                    // 二级标题（小三 = 15pt，居左）
    h3Font: '黑体',
    h3Size: 14,                    // 三级标题（四号 = 14pt，居左、右缩进两字）
    headingBold: false,            // 标题是否加粗（黑体本身较粗，默认不加粗）

    absHeadSize: 16,               // 摘要标题（三号，居中）
    absBodySize: 12,               // 摘要正文（小四）

    refSize: 10.5,                 // 参考文献字号（五号 = 10.5pt）
    refHangingChars: 0,            // 参考文献悬挂缩进（附件8：序号左顶格，默认 0）

    captionFont: '宋体',           // 图表题注（附件8：图题/表题五号宋体）
    captionSize: 10.5,

    pageNumber: 'center',          // 页码：center | right | none
    pageNumberSplit: true,         // 前置部分罗马数字、主体阿拉伯数字单独编页
    threeLineTable: true,          // 表格仅保留顶线/底线 1.5 磅，清除其余所有线条
    tableText: true,               // 表格内文字统一格式（附件8：表格内字体均为五号宋体，表头不开例外）
    tableFont: '宋体',             // 表格内中文字体（西文跟随 latinFont = Times New Roman）
    tableSize: 10.5,               // 表格内字号（五号 = 10.5pt）
    chapterPageBreak: true,        // 一级标题之间换页
    chapterNumber: true,           // 标题编号体例（「第1章 绪论」→「1 绪论」；「3.1.1系统目标」→「3.1.1 系统目标」）
    autoToc: true,                 // 目录替换为 Word 自动目录域（可整体选择、可更新）
    bodyHeader: true,              // 正文及各后置分节页眉 = 一级标题（章名，STYLEREF 域）
    updateFields: true,            // 打开文档时自动更新域（目录页码自动刷新）
    citeSuperscript: true,         // 正文引用标注 [n] 改为上标（附件8：右上角上标标注）
    clearHighlight: true,          // 清除全文底色：突出显示 + 字符/段落/表格单元格底纹（论文不应带灰底彩底）
    cjkSpace: true,                // 删除正文中中文与英文/数字之间的空格（编号后的间隔保留）
    chapterSection: true,          // 每章末尾补「分节符（下一页）」，便于逐章单独设置页眉
    wordCaption: true              // 图片题注改用 Word 真题注（题注样式 + SEQ 域自动编号）
  };

  /* ---------- 结构识别正则 ---------- */
  var RE = {
    abs:    /^摘\s*要/,
    absEn:  /^Abstract/i,
    kw:     /^关键词/,
    kwEn:   /^(Key\s*[Ww]ords|Keywords)/,
    toc:    /^(目\s*录|目\s*次)/,
    ref:    /^(参\s*考\s*文\s*献|References)/,
    h1:     /^第[一二三四五六七八九十百千零0-9]+章/,
    h1Word: /^(绪\s*论|引\s*言|前\s*言|结\s*论|总\s*结|结\s*论\s*与\s*展\s*望)/,
    /* 编号首位禁止 0：正文里的十进制小数（如「0.96英寸…」）会误命中编号规则 */
    h1Alt:  /^[1-9]\d?\s+[^\d.\s]/,
    /* 三级标题：编号与标题之间**允许没有空格**（「3.1.1系统目标」同样是标题）。
       这里原先是 \s（必须有空格），与下面 h2 的 \s* 不一致 —— 于是「3.1.1系统目标」
       这种写法 h3 不命中、h2 又因 (?![\d.]) 撞上第二个小数点而不命中，整段掉进
       body：宋体小四、两端对齐、缩进 480，三级标题的格式全丢。
       2026-09-16 曾鹏的 2.docx 第3章 9 个三级标题就是这么被排成正文的。
       (?![\d.]) 是别让「1.1.1.1」这类四级编号的前缀被认成三级标题。 */
    h3:     /^[1-9]\d?\.\d{1,2}\.\d{1,2}(?![\d.])\s*[^\d.]/,
    h2:     /^[1-9]\d?\.\d{1,2}(?![\d.])\s*[^\d.]/,
    h2cn:   /^[一二三四五六七八九十]{1,3}[、.]/,
    /* 附录内小节：A.1 / B.2.1（附件8：附录中的图、表、式另行编号，与正文分开） */
    h2app:  /^[A-Z]\.\d{1,2}(?![\d.])\s*[^\d.]/,
    cap:    /^(图|表|Figure|Table)\s*[A-Z]?\.?\s*\d/,
    ack:    /^致\s*谢/,
    app:    /^附\s*录/,
    tocItem: /\.{2,}|\s\d{1,3}\s*$/,  // 目录条目：带点线引导符或以页码结尾
    cover:  /(学号|姓名|学院|专业|班级|日期|指导|教师|题目|答辩|成绩|职称|系|作者|签字|声明|原创|诚信|承诺)/, // 封面特征
    date:   /^\d|年\s*\d|月\s*\d+\s*日/ // 日期/编号类文本
  };

  /* ---------- OOXML 子元素顺序（CT_PPr / CT_RPr / CT_SectPr） ---------- */
  var PPR_ORDER = ['pStyle','keepNext','keepLines','pageBreakBefore','framePr','widowControl','numPr',
    'suppressLineNumbers','pBdr','shd','tabs','suppressAutoHyphens','kinsoku','wordWrap','overflowPunct',
    'topLinePunct','autoSpaceDE','autoSpaceDN','bidi','adjustRightInd','snapToGrid','spacing','ind',
    'contextualSpacing','mirrorIndents','suppressOverlap','jc','textDirection','textAlignment',
    'textboxTightWrap','outlineLvl','divId','cnfStyle','rPr','sectPr','pPrChange'];

  var RPR_ORDER = ['rStyle','rFonts','b','bCs','i','iCs','caps','smallCaps','strike','dstrike','outline',
    'shadow','emboss','imprint','noProof','snapToGrid','vanish','webHidden','color','spacing','w','kern',
    'position','sz','szCs','highlight','u','effect','bdr','shd','fitText','vertAlign','rtl','cs','em',
    'lang','textFill','eastAsianLayout','specVanish','oMath'];

  var SECTPR_ORDER = ['headerReference','footerReference','footnotePr','endnotePr','type','pgSz','pgMar',
    'paperSrc','pgBorders','lnNumType','pgNumType','cols','formProt','vAlign','noEndnote','titlePg',
    'textDirection','bidi','rtlGutter','docGrid','printerSettings','sectPrChange'];

  /* ---------- 基础工具 ---------- */
  /* elm 未必是元素。有些工具生成的 docx 段落之间留着换行（按缩进排过版的更是如此），
     解析后这些换行是**空白文本节点**；遍历 body 子节点、拿 nextSibling 当元素用时
     就会摸到它们 —— 文本节点没有 getElementsByTagNameNS，直接抛 TypeError。
     这里兜住返回空集：文本节点里当然不会有 w:fldChar 这类元素。
     注意别用 nodeType === 1 去判 —— Document 也要走这条路（ensureTocStyles 传的是文档）。 */
  function allByNs(elm, name, ns) {
    if (!elm || typeof elm.getElementsByTagNameNS !== 'function') return [];
    return elm.getElementsByTagNameNS(ns, name);
  }
  function wAll(elm, name) { return allByNs(elm, name, W_NS); }
  function childByNs(elm, name, ns) {
    if (!elm) return null;
    for (var i = 0; i < elm.childNodes.length; i++) {
      var c = elm.childNodes[i];
      if (c.nodeType === 1 && c.localName === name && (!ns || c.namespaceURI === ns)) return c;
    }
    return null;
  }
  function createW(doc, name) { return doc.createElementNS(W_NS, 'w:' + name); }
  function createR(doc, name) { return doc.createElementNS(R_NS, 'r:' + name); }
  function insertInOrder(parent, el, order) {
    var idx = order.indexOf(el.localName);
    for (var i = 0; i < parent.childNodes.length; i++) {
      var c = parent.childNodes[i];
      if (c.nodeType !== 1 || c === el) continue;
      var ci = order.indexOf(c.localName);
      if (ci > idx) { parent.insertBefore(el, c); return; }
    }
    parent.appendChild(el);
  }
  function paraText(p) {
    var s = '';
    var ts = wAll(p, 't');
    for (var i = 0; i < ts.length; i++) s += ts[i].textContent || '';
    return s;
  }
  function runText(r) {
    var s = '';
    var ts = wAll(r, 't');
    for (var i = 0; i < ts.length; i++) s += ts[i].textContent || '';
    return s;
  }
  function hasPageBreak(p) {
    var brs = wAll(p, 'br');
    for (var i = 0; i < brs.length; i++) {
      if ((brs[i].getAttributeNS(W_NS, 'type') || '') === 'page') return true;
    }
    return wAll(p, 'lastRenderedPageBreak').length > 0;
  }
  /* 空命名空间节点计数（见 assertNamespaces）。模块级，每次 formatDocx 重置。 */
  var emptyNsNodes = [];

  function assertNamespaces(doc, part) {
    /* 浏览器把 createElement('Foo') 建出来的元素当成「无命名空间」节点，
       序列化时为了保住这个语义会补一个 xmlns="" —— 在 Word 眼里这个元素就
       等于不存在。曾经因此让 [Content_Types].xml 里的 Override 失效，报
       「文件已损坏」。node 端的 xmldom 不补这个声明，所以单测全绿也照样出
       问题，只能在这里自己拦。 */
    var els = doc.getElementsByTagName('*');
    for (var i = 0; i < els.length; i++) {
      var e = els[i];
      if (e.namespaceURI) continue;
      var parent = e.parentNode;
      if (!parent || parent.nodeType !== 1 || !parent.namespaceURI) continue;
      emptyNsNodes.push(part + ':' + e.nodeName);
    }
  }

  var XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';

  function serialize(doc, part) {
    /* XML 声明在两个环境下表现相反，必须两边都照顾：
       xmldom（Node 单测）：声明是文档的一个 PI 子节点，序列化器不再另出声明
                            —— 所以要先摘掉，否则成品里会带两条；
       浏览器：            DOMParser 根本不把声明建节点（doc.childNodes[0] 就是根
                            元素），而 XMLSerializer 会把声明原样吐出来
                            —— 所以还要防它自带的那条。
       两头都防的办法：先摘 PI 节点，序列化后再把开头的声明削掉，最后统一补一条。
       （曾经只按 xmldom 的行为写，浏览器产出的每个部件都带两条 XML 声明，
         Word 打开直接报「文件已损坏」。） */
    while (doc.childNodes.length && doc.childNodes[0].nodeType === 7) {
      doc.removeChild(doc.childNodes[0]);
    }
    if (part) assertNamespaces(doc, part);
    var out = new XMLSerializer().serializeToString(doc);
    out = out.replace(/^\s*<\?xml[\s\S]*?\?>\s*/, '');
    return XML_DECL + out;
  }

  /* ---------- 收集正文段落（跳过表格内部与空段落） ---------- */
  /* 段落元信息：pStyle（目录条目的 Word 样式名如 TOC1）、是否含超链接、
     outlineLvl（大纲级别，目录条目常带）、是否含制表符（Word 目录条目
     的标准结构：标题 + tab + 页码）——用于可靠识别目录 */
  function paraMeta(p) {
    var meta = { style: '', hasLink: false, outline: null, hasTab: false };
    var pPr = childByNs(p, 'pPr', W_NS);
    if (pPr) {
      var ps = childByNs(pPr, 'pStyle', W_NS);
      if (ps) meta.style = (ps.getAttributeNS(W_NS, 'val') || '').toLowerCase();
      var ol = childByNs(pPr, 'outlineLvl', W_NS);
      if (ol) meta.outline = ol.getAttributeNS(W_NS, 'val');
    }
    for (var i = 0; i < p.childNodes.length; i++) {
      var c = p.childNodes[i];
      if (c.nodeType === 1 && c.localName === 'hyperlink') meta.hasLink = true;
    }
    if (wAll(p, 'tab').length > 0) meta.hasTab = true; // tab 在 run 内部，需深层查找
    return meta;
  }

  function collectParas(doc) {
    var body = allByNs(doc, 'body', W_NS)[0];
    var items = [];
    if (!body) return items;
    for (var i = 0; i < body.childNodes.length; i++) {
      var c = body.childNodes[i];
      if (c.nodeType !== 1) continue;
      if (c.localName === 'p') {
        var text = paraText(c).trim();
        var hasMath = allByNs(c, 'oMath', M_NS).length > 0;
        var hasDrawing = wAll(c, 'drawing').length > 0 || wAll(c, 'pict').length > 0;
        if (!text && !hasDrawing) continue; // 纯空段落跳过
        items.push({ p: c, text: text, hasMath: hasMath, hasDrawing: hasDrawing,
                     hasPageBreak: hasPageBreak(c), meta: paraMeta(c) });
      }
    }
    return items;
  }

  /* ---------- 结构识别 ---------- */
  function findTitleUp(paras, idx) {
    for (var i = idx - 1; i >= 0; i--) {
      var p = paras[i];
      if (p.hasPageBreak) return -1;          // 跨页了 → 不是摘要页题目
      if (!p.text) continue;                  // 空段落继续向上
      var t = p.text;
      if (t.length < 2 || t.length > 45) return -1;
      if (RE.date.test(t)) return -1;         // 日期/编号行 → 不是题目
      if (RE.cover.test(t)) return -1;        // 含封面特征字段 → 不是题目
      return i;
    }
    return -1;
  }

  function isH(t, hasChapterStyle) {
    return RE.h1.test(t) || RE.h1Word.test(t) || (!hasChapterStyle && RE.h1Alt.test(t));
  }

  /* 目录条目特征（多特征命中其一即可）：
     1. 文本带点线引导符或以页码结尾（RE.tocItem）
     2. Word 目录样式（pStyle 含 toc）
     3. 超链接段落且文本像标题（Word 生成的目录条目默认是超链接）
     4. 含制表符 + 以页码结尾（Word 目录条目标准结构：标题 + tab + 页码） */
  function isTocEntry(p, text) {
    if (!text) return false;
    if (RE.tocItem.test(text)) return true;
    /* 本工具生成的目录条目：正文是标题文本（不带页码），页码由 PAGEREF _TocN 域给出。
       域结果在重新格式化时是空的，文本里既没有点线也不以页码结尾，靠下面几条
       特征全都认不出来 —— 于是第二次格式化时目录区边界会落在目录中间，
       旧目录删一半留在正文里，还留下一个没有 begin 的 end 域字符
       （Word 打开更新域时会把后面的正文一起吞掉，表现为「正文全没了」）。
       PAGEREF _TocN / HYPERLINK _TocN 是目录条目独有的，见到即可判定 */
    var el = p.p || p;                       // 调用方传的是条目对象（p 为段落元素）
    var instrs = allByNs(el, 'instrText', W_NS);
    for (var k = 0; k < instrs.length; k++) {
      var ins = instrs[k].textContent || '';
      if (/^\s*PAGEREF\s+"?_Toc\d+/i.test(ins) || /HYPERLINK\s+"?\\l"?\s+"?_Toc\d+/i.test(ins)) return true;
    }
    var m = p.meta;
    if (!m) return false;
    if (/toc|目录/.test(m.style)) return true;
    if (m.hasLink && text.length <= 40 && !/\s{2,}/.test(text) && !RE.cover.test(text)) return true;
    // tab + 页码结尾（页码可为阿拉伯数字或罗马数字，如目录中「摘 要I」「ABSTRACTII」）
    if (m.hasTab && text.length <= 40 && /[\dIVX]+\s*$/.test(text)) return true;
    return false;
  }

  function tocEnd(paras, startIdx, hasChapterStyle) {
    for (var i = startIdx + 1; i < paras.length; i++) {
      var t = paras[i].text;
      if (!t) continue;
      if (isTocEntry(paras[i], t)) continue;            // 目录条目 → 继续
      var h = isH(t, hasChapterStyle);
      if (h) {
        // 疑似真正的一级标题：向后看几步是否仍是标题样式的目录条目
        var looksToc = false;
        for (var j = i + 1; j < Math.min(i + 4, paras.length); j++) {
          if (!paras[j].text) continue;
          if (isTocEntry(paras[j], paras[j].text) || isH(paras[j].text, hasChapterStyle)) { looksToc = true; }
          break;
        }
        if (!looksToc) return i;
        continue;
      }
      return i;                                          // 普通正文 → 目录结束
    }
    return paras.length;
  }

  function classifyParas(paras) {
    var n = paras.length;
    var texts = paras.map(function (x) { return x.text; });
    var roles = new Array(n).fill('body');

    /* 定位关键段落 */
    var absIdx = -1, kwIdx = -1, absEnIdx = -1, kwEnIdx = -1, tocIdx = -1, refIdx = -1;
    for (var i = 0; i < n; i++) {
      var t = texts[i];
      if (absIdx < 0 && RE.abs.test(t)) absIdx = i;
      if (absIdx >= 0 && kwIdx < 0 && i > absIdx && RE.kw.test(t)) kwIdx = i;
      if (absEnIdx < 0 && absIdx >= 0 && i > Math.max(kwIdx, absIdx) && RE.absEn.test(t)) absEnIdx = i;
      if (absEnIdx >= 0 && kwEnIdx < 0 && i > absEnIdx && RE.kwEn.test(t)) kwEnIdx = i;
      if (tocIdx < 0 && RE.toc.test(t)) tocIdx = i;
      // 目录条目里的「参考文献36」这类带页码的条目不能当真正的参考文献标题
      if (refIdx < 0 && RE.ref.test(t) && !isTocEntry(paras[i], t)) refIdx = i;
    }
    // 无中文摘要时的兜底：全文查找英文摘要（跳过目录条目）
    if (absEnIdx < 0) {
      for (var k = 0; k < n; k++) { if (RE.absEn.test(texts[k]) && !isTocEntry(paras[k], texts[k])) { absEnIdx = k; break; } }
      if (absEnIdx >= 0) {
        kwEnIdx = -1;
        for (var m = absEnIdx + 1; m < n; m++) { if (RE.kwEn.test(texts[m])) { kwEnIdx = m; break; } }
      }
    }

    var hasChapterStyle = texts.some(function (x) { return RE.h1.test(x); });
    var titleIdx = absIdx >= 0 ? findTitleUp(paras, absIdx) : -1;
    var enTitleIdx = absEnIdx >= 0 ? findTitleUp(paras, absEnIdx) : -1;
    /* 无英文题目时 findTitleUp 会向上找到「关键词：…」行 —— 那不是题目，排除 */
    if (enTitleIdx >= 0 && (enTitleIdx === kwIdx || enTitleIdx === absIdx || enTitleIdx === titleIdx)) enTitleIdx = -1;

    // 第一个一级标题（须先于下方目录兜底识别计算，否则取不到值）
    var firstH = -1;
    for (var h = 0; h < n; h++) {
      var ht = texts[h];
      if (ht && (RE.h1.test(ht) || RE.h1Word.test(ht) || (!hasChapterStyle && RE.h1Alt.test(ht)))) { firstH = h; break; }
    }
    // 摘要正文区域边界：无"关键词"行时延伸到英文摘要标题或第一个章节标题
    var absBodyEnd = kwIdx >= 0 ? kwIdx : (absEnIdx >= 0 ? absEnIdx : (firstH >= 0 ? firstH : n));
    var absEnBodyEnd = kwEnIdx >= 0 ? kwEnIdx : (firstH >= 0 ? firstH : n);

    /* 无目录标题（「目 录」）时的兜底识别：摘要/关键词结束之后、第一个一级标题
       之前，若存在 ≥2 个目录条目（tab+页码等特征），则视为目录章节。
       注意：firstH/absBodyEnd/absEnBodyEnd 必须先赋值（原实现因 var 提升
       拿到 undefined，该兜底从未生效，此处已修正顺序） */
    if (tocIdx < 0 && firstH >= 0) {
      var scanFrom = Math.max(absBodyEnd, absEnBodyEnd, 0);
      var tocStart = -1, tocHits = 0;
      for (var f = scanFrom; f < n && f <= firstH; f++) {
        if (!texts[f]) continue;
        if (isTocEntry(paras[f], texts[f])) { if (tocStart < 0) tocStart = f; tocHits++; }
      }
      if (tocHits >= 2) tocIdx = tocStart;
    }
    var tocHasHead = tocIdx >= 0 && RE.toc.test(texts[tocIdx]);
    var tocEndIdx = tocIdx >= 0 ? tocEnd(paras, tocIdx, hasChapterStyle) : -1;

    /* 逐段归类 */
    var inBack = false;
    for (var i2 = 0; i2 < n; i2++) {
      if (i2 === absIdx) { roles[i2] = 'absHead'; continue; }
      if (i2 === kwIdx) { roles[i2] = 'kw'; continue; }
      if (i2 === absEnIdx) { roles[i2] = 'absEnHead'; continue; }
      if (i2 === kwEnIdx) { roles[i2] = 'kwEn'; continue; }
      if (i2 === tocIdx) { roles[i2] = tocHasHead ? 'tocHead' : 'tocItem'; continue; }
      if (i2 === refIdx) { roles[i2] = 'refHead'; inBack = false; continue; }
      if (i2 === titleIdx) { roles[i2] = 'title'; continue; }
      if (i2 === enTitleIdx) { roles[i2] = 'enTitle'; continue; }

      var tx = texts[i2];
      if (absIdx >= 0 && i2 < absIdx) { roles[i2] = 'front'; continue; }                    // 封面/前部：不动
      // 目录条目优先于摘要正文判定：无关键词的文档中 absBodyEnd 可能延伸到
      // 第一个标题样式段落，目录条目不能被误判为摘要正文
      if (tocIdx >= 0 && i2 > tocIdx && i2 < tocEndIdx) { roles[i2] = 'tocItem'; continue; }
      if (absIdx >= 0 && i2 > absIdx && i2 < absBodyEnd) { roles[i2] = 'absBody'; continue; }
      if (absEnIdx >= 0 && i2 > absEnIdx && i2 < absEnBodyEnd) { roles[i2] = 'absEnBody'; continue; }
      if (RE.cap.test(tx)) { roles[i2] = 'caption'; continue; }  // 题注优先于参考文献/后置部分判定（附录里的图1、表1等）
      // 后置部分（附录）里的纯图片段也要认成 figure，否则会掉进下面的 body 分支，
      // 图题就永远不会被真题注化（附件8 要求附录内图表另行编号）
      if (refIdx >= 0 && i2 > refIdx && !tx && paras[i2].hasDrawing) { roles[i2] = 'figure'; continue; }
      if (refIdx >= 0 && i2 > refIdx) {
        if (RE.ack.test(tx) || RE.app.test(tx)) { roles[i2] = 'h1'; inBack = true; }
        // 附录内小节（A.1、B.2 …）按二级标题排版；题注已在上面判过
        else if (inBack && RE.h2app.test(tx) && tx.length <= 40) roles[i2] = 'h2';
        else roles[i2] = inBack ? 'body' : 'refItem';
        continue;
      }
      if (paras[i2].hasMath) { roles[i2] = 'formula'; continue; }
      if (!tx && paras[i2].hasDrawing) { roles[i2] = 'figure'; continue; }                  // 纯图片段：不动
      if (RE.h1.test(tx) || (!hasChapterStyle && RE.h1Alt.test(tx)) || RE.h1Word.test(tx)) { roles[i2] = 'h1'; continue; }
      /* 二/三级标题判定加长度守卫：正文长句（如「1.5倍…」「0.96英寸…」）不当作标题 */
      if (RE.h3.test(tx) && tx.length <= 40) { roles[i2] = 'h3'; continue; }
      if ((RE.h2.test(tx) || RE.h2cn.test(tx) || RE.h2app.test(tx)) && tx.length <= 40) { roles[i2] = 'h2'; continue; }
      roles[i2] = 'body';
    }

    /* 统计 */
    var ROLE_KEYS = ['front','title','absHead','absBody','kw','enTitle','absEnHead','absEnBody','kwEn',
      'tocHead','tocItem','h1','h2','h3','body','caption','formula','figure','refHead','refItem'];
    var counts = {};
    ROLE_KEYS.forEach(function (r) { counts[r] = 0; });
    roles.forEach(function (r) { counts[r]++; });

    var info = {
      absDetected: absIdx >= 0,
      absEnDetected: absEnIdx >= 0,
      tocDetected: tocIdx >= 0,
      refDetected: refIdx >= 0,
      titleDetected: titleIdx >= 0,
      enTitleDetected: enTitleIdx >= 0,
      chapterStyle: hasChapterStyle
    };

    return { roles: roles, counts: counts, info: info };
  }

  /* ---------- 样式表 ---------- */
  function buildStyles(s) {
    // 行距：fixed 固定值(磅) → twips；multiple 倍数 → 1/240 行单位
    var line = s.lineSpacingMode === 'exact'
      ? Math.round((s.lineSpacing || 20) * 20)
      : Math.round((s.lineSpacing || 1.5) * 240);
    var lineRule = s.lineSpacingMode === 'exact' ? 'exact' : 'auto';
    var east = s.bodyFont, latin = s.latinFont;
    function mk(eastFont, latinFont, size, jc, extra) {
      extra = extra || {};
      var st = {
        eastFont: eastFont, latinFont: latinFont, size: size, jc: jc,
        line: extra.line !== undefined ? extra.line : line,
        lineRule: extra.lineRule !== undefined ? extra.lineRule : lineRule,
        before: extra.before || 0, after: extra.after || 0,
        indent: extra.indent, hanging: extra.hanging, bold: extra.bold,
        rightChars: extra.rightChars, leftChars: extra.leftChars
      };
      return st;
    }
    return {
      /* 标题段前段后均为 0、行距固定 20 磅，与参考样张（基于PLC的转台伺服控制系统）一致。
         一级标题/摘要/ABSTRACT/目录标题的上下空行统一为五号字大小（10.5pt=210），
         与标题自身字号（三号）分割开；上方空行由 normalizeChapterBreaks 保留的
         1 个空段提供（行高同为 210），下方由 after 提供 */
      title:     mk(s.titleFont, latin, s.titleSize, 'center', { bold: s.titleBold, before: 240, after: 120 }),
      absHead:   mk(s.h1Font, latin, s.absHeadSize, 'center', { bold: false }),
      absBody:   mk(east, latin, s.absBodySize, 'both', { indent: 2 }),
      kw:        mk(east, latin, s.absBodySize, 'left', { bold: undefined }),
      absEnHead: mk(s.h1Font, latin, s.absHeadSize, 'center', { bold: true }),
      absEnBody: mk(east, latin, s.absBodySize, 'both', { indent: 2 }),
      kwEn:      mk(east, latin, s.absBodySize, 'left'),
      enTitle:   mk(east, latin, s.titleSize, 'center', { bold: true, before: 240, after: 120 }),
      tocHead:   mk(s.h1Font, latin, s.absHeadSize, 'center', { bold: false }),
      tocItem:   mk(east, latin, s.bodySize, 'left'),   // 目录条目缩进按层级动态设置
      h1:        mk(s.h1Font, latin, s.h1Size, s.h1Jc, { bold: s.headingBold }),
      h2:        mk(s.h2Font, latin, s.h2Size, 'left', { bold: s.headingBold }), // 行距固定 20 磅（附件8）
      h3:        mk(s.h3Font, latin, s.h3Size, 'left', { bold: s.headingBold, indent: 2 }), // 首行缩进两字符
      body:      mk(east, latin, s.bodySize, 'both', { indent: s.firstLineChars }),
      caption:   mk(s.captionFont, latin, s.captionSize, 'center', { bold: false }),
      refHead:   mk(s.h1Font, latin, s.h1Size, 'center', { bold: false }),
      refItem:   mk(east, latin, s.refSize, 'left', { hanging: s.refHangingChars, before: 157, after: 157 }),
      formula:   { jc: 'center' }
    };
  }

  /* ---------- 应用 rPr（字体/字号/加粗）到单个 run ---------- */
  function applyRPr(rPr, doc, st, boldMode) {
    var rf = childByNs(rPr, 'rFonts', W_NS);
    if (!rf) { rf = createW(doc, 'rFonts'); insertInOrder(rPr, rf, RPR_ORDER); }
    rf.setAttributeNS(W_NS, 'w:ascii', st.latinFont);
    rf.setAttributeNS(W_NS, 'w:hAnsi', st.latinFont);
    rf.setAttributeNS(W_NS, 'w:eastAsia', st.eastFont);
    rf.setAttributeNS(W_NS, 'w:cs', st.latinFont);

    var sz = childByNs(rPr, 'sz', W_NS);
    if (!sz) { sz = createW(doc, 'sz'); insertInOrder(rPr, sz, RPR_ORDER); }
    sz.setAttributeNS(W_NS, 'w:val', String(Math.round(st.size * 2)));
    var szCs = childByNs(rPr, 'szCs', W_NS);
    if (!szCs) { szCs = createW(doc, 'szCs'); insertInOrder(rPr, szCs, RPR_ORDER); }
    szCs.setAttributeNS(W_NS, 'w:val', String(Math.round(st.size * 2)));

    if (boldMode === true) {
      var b = childByNs(rPr, 'b', W_NS);
      if (!b) { b = createW(doc, 'b'); insertInOrder(rPr, b, RPR_ORDER); }
    } else {
      // 非加粗角色（含未指定 bold 的正文/题注/参考文献等）一律移除加粗，保证文本常规格式
      var b0 = childByNs(rPr, 'b', W_NS); if (b0) rPr.removeChild(b0);
      var b1 = childByNs(rPr, 'bCs', W_NS); if (b1) rPr.removeChild(b1);
    }
  }

  function applyRun(r, doc, st) {
    var rPr = childByNs(r, 'rPr', W_NS);
    if (!rPr) { rPr = createW(doc, 'rPr'); r.insertBefore(rPr, r.firstChild); }
    applyRPr(rPr, doc, st, st.bold);
  }

  function formatRuns(p, doc, st) {
    var runs = wAll(p, 'r');
    for (var i = 0; i < runs.length; i++) applyRun(runs[i], doc, st);
  }

  /* 关键词行：标签（"关键词："）加粗黑体，内容正文样式 */
  function formatKeywords(p, doc, st, role) {
    var runs = wAll(p, 'r');
    var text = paraText(p);
    var m = text.match(role === 'kw' ? /^(关键词)\s*[:：]?/ : /^(Key\s*[Ww]ords|Keywords)\s*[:：]?/);
    var labelLen = m ? m[0].length : (role === 'kw' ? 3 : 9);
    // 中文关键词标签黑体不加粗（参考样张），英文 Keywords 加粗
    var labelSt = { eastFont: '黑体', latinFont: st.latinFont, size: st.size, bold: role === 'kwEn' };
    var bodySt = { eastFont: st.eastFont, latinFont: st.latinFont, size: st.size, bold: false };
    var cursor = 0;
    for (var i = 0; i < runs.length; i++) {
      var t = runText(runs[i]);
      if (!t) continue;
      applyRun(runs[i], doc, cursor < labelLen ? labelSt : bodySt);
      cursor += t.length;
    }
  }

  /* 引用文献标注上标：《附件8》要求"正文中引用的文献应在所引用原文内容最末句的
     右上角以上标方式进行标注，并按先后顺序连续编号置于方括号内"。
     只处理形如 [1] / [1,2] / [1-3] 的标注；参考文献表自身的序号（refItem）不动。 */
  var CITE_RE = /\[\d+(?:\s*[,\-–—]\s*\d+)*\]/g;

  function applyCiteSuperscript(p, doc) {
    var runs = wAll(p, 'r');
    for (var i = 0; i < runs.length; i++) {
      var r = runs[i];
      var t = runText(r);
      if (!t || t.indexOf('[') < 0) continue;
      CITE_RE.lastIndex = 0;
      if (!CITE_RE.test(t)) continue;
      /* 含制表符/换行/域代码的 run 直接跳过（拆分 runText 会丢失这些结构） */
      var struct = false;
      for (var c = 0; c < r.childNodes.length; c++) {
        var cn = r.childNodes[c];
        if (cn.nodeType === 1 && cn.localName !== 'rPr' && cn.localName !== 't') { struct = true; break; }
      }
      if (struct) continue;
      var rPr0 = childByNs(r, 'rPr', W_NS);
      var va0 = rPr0 ? childByNs(rPr0, 'vertAlign', W_NS) : null;
      if (va0 && va0.getAttributeNS(W_NS, 'val') === 'superscript') continue;  // 已是上标
      CITE_RE.lastIndex = 0;
      var segs = [], last = 0, m;
      while ((m = CITE_RE.exec(t))) {
        if (m.index > last) segs.push({ s: t.slice(last, m.index), sup: false });
        segs.push({ s: m[0], sup: true });
        last = m.index + m[0].length;
      }
      if (last < t.length) segs.push({ s: t.slice(last), sup: false });
      if (segs.length < 2) continue;
      var parent = r.parentNode;
      for (var k = 0; k < segs.length; k++) {
        var nr = createW(doc, 'r');
        var rPr = rPr0 ? rPr0.cloneNode(true) : null;
        if (rPr) nr.appendChild(rPr);
        if (segs[k].sup) {
          if (!rPr) { rPr = createW(doc, 'rPr'); nr.appendChild(rPr); }
          var va = childByNs(rPr, 'vertAlign', W_NS);
          if (!va) { va = createW(doc, 'vertAlign'); insertInOrder(rPr, va, RPR_ORDER); }
          va.setAttributeNS(W_NS, 'w:val', 'superscript');
        }
        var nt = createW(doc, 't');
        nt.setAttributeNS(XML_NS, 'xml:space', 'preserve');
        nt.appendChild(doc.createTextNode(segs[k].s));
        nr.appendChild(nt);
        parent.insertBefore(nr, r);
      }
      parent.removeChild(r);
    }
  }

  /* ---------- 清除底色 ----------
     论文不应带任何灰底/彩底标记（写作时标英文术语很常见），交付前一律清掉。
     两样都清：
     ① w:highlight —— 工具条上的「突出显示」；
     ② w:shd —— 字符/段落/表格单元格的底纹。
     ②原先是不动的（注释写着「常被模板用来做表头底色等正常排版」），但附件8 的
     表格是纯三线表：顶线 1.5 磅、栏目线 0.75 磅、底线 1.5 磅，通篇没有底色，
     源文档里那种浅蓝表头（DCE6F1）和正文说明框的浅灰蓝（F2F4F7）都要去掉。
     2026-09-16 按作者要求改为底纹一并清除。 */
  function removeAll(doc, name) {
    var list = allByNs(doc, name, W_NS);   // 可能是 live NodeList，先快照
    var els = [];
    for (var i = 0; i < list.length; i++) els.push(list[i]);
    var n = 0;
    for (var j = els.length - 1; j >= 0; j--) {
      if (els[j].parentNode) { els[j].parentNode.removeChild(els[j]); n++; }
    }
    return n;
  }

  function clearHighlights(doc) {
    return removeAll(doc, 'highlight');
  }

  /* 底纹分布得很散：可能在 rPr（字符）、pPr（段落）、tcPr（单元格）、
     trPr/tblPrEx（行）、tblPr（整表）里，所以按元素扫一遍。
     前置部分（封面/学生声明）里的底纹一律不动 —— 与三线表、换页等工序同一口径，
     封面表格是学校模板的一部分，改了就不像模板了。 */
  function clearShading(doc, frontBoundary) {
    var body = allByNs(doc, 'body', W_NS)[0];
    if (!frontBoundary || !body) return removeAll(doc, 'shd');
    var n = 0, inBody = false;
    for (var i = 0; i < body.childNodes.length; i++) {
      var c = body.childNodes[i];
      if (c.nodeType !== 1) continue;
      if (!inBody) {
        /* 边界之前 → 前置部分（封面/声明），跳到边界为止，一个都不动 */
        if (c === frontBoundary) inBody = true; else continue;
      }
      n += removeAll(c, 'shd');
    }
    return n;
  }

  /* ---------- 页眉/页脚里的孤儿标点清理 ----------
     来源文档（模板复制、WPS 导出最常见）常在页码域前留下一个只含顿号的 run，
     排出来就是页脚上的「、33」。判定条件收得很紧：该 run 的文字**只有**一个
     全角标点，且与 PAGE 域同段，且在文本框之外——页码语境里这一定是残留。
     「第3页，共10页」这类正常写法里标点不会单独成 run，不会被误删。 */
  function isInTextbox(el, stop) {
    for (var a = el.parentNode; a && a !== stop; a = a.parentNode) {
      if (a.localName === 'txbxContent') return true;
    }
    return false;
  }
  async function cleanFooterStrayPunct(zip) {
    var removed = 0;
    var names = Object.keys(zip.files).filter(function (f) {
      return /^word\/(?:header|footer)\d*\.xml$/.test(f);
    });
    for (var i = 0; i < names.length; i++) {
      var entry = zip.file(names[i]);
      if (!entry) continue;
      var xml = await entry.async('string');
      if (xml.indexOf('PAGE') < 0) continue;
      var hDoc = new DOMParser().parseFromString(xml, 'application/xml');
      var ps = allByNs(hDoc, 'p', W_NS);
      var touched = false;
      for (var k = 0; k < ps.length; k++) {
        var instr = allByNs(ps[k], 'instrText', W_NS), hasPage = false;
        for (var q = 0; q < instr.length; q++) {
          if (/PAGE/.test(instr[q].textContent || '')) { hasPage = true; break; }
        }
        if (!hasPage) continue;
        var runs = allByNs(ps[k], 'r', W_NS);
        for (var r = runs.length - 1; r >= 0; r--) {
          var run = runs[r];
          if (isInTextbox(run, ps[k])) continue;
          var ts = wAll(run, 't'), s = '';
          for (var t = 0; t < ts.length; t++) s += ts[t].textContent || '';
          if (/^[、，。；：,;]$/.test(s) && run.parentNode) {
            run.parentNode.removeChild(run); removed++; touched = true;
          }
        }
      }
      if (touched) zip.file(names[i], serialize(hDoc, names[i]));
    }
    return removed;
  }

  /* ---------- 删除中西文之间的空格 ----------
     附件8 要求标题/题注「数字与文字间隔一字符」（1 绪论、图4.1 系统图），
     所以**段首编号之后的那个空格必须保留**，其余中文与英文/数字之间的空格
     一律删除（「采用 Node.js」→「采用Node.js」）。
     中文侧含汉字与中文标点（全角括号、顿号等），但排除全角字母数字。 */
  var CJK_CH = '[\\u3000-\\u303f\\u3400-\\u4dbf\\u4e00-\\u9fff\\uf900-\\ufaff' +
               '\\uff01-\\uff0f\\uff1a-\\uff20\\uff3b-\\uff40\\uff5b-\\uff65]';
  var LAT_CH = '[0-9A-Za-z]';
  /* 段首编号：1 / 1.1 / 1.1.1 / A.1 / 图4.1 / 表7.2 —— 连同其后的空格一起保留 */
  var NUM_HEAD = /^(?:\d{1,2}(?:\.\d{1,2}){0,2}|[A-Z]\.\d{1,2}(?:\.\d{1,2}){0,2}|[图表]\s*\d{1,2}(?:\.\d{1,2}){0,2}|附\s*录\s*[A-Za-z0-9])\s+/;
  /* 需要保留段首编号间隔的角色：标题、目录条目、图表题注 */
  var KEEP_NUM_ROLE = { h1: 1, h2: 1, h3: 1, tocItem: 1, caption: 1, refHead: 1 };

  /* 段落正文文字节点（run 内的直接 w:t，以及超链接/域内的 run；跳过文本框与图形） */
  function paraTextNodes(p) {
    var nodes = [];
    function walk(el) {
      for (var i = 0; i < el.childNodes.length; i++) {
        var c = el.childNodes[i];
        if (c.nodeType !== 1) continue;
        if (c.localName === 'r') {
          for (var j = 0; j < c.childNodes.length; j++) {
            var t = c.childNodes[j];
            if (t.nodeType === 1 && t.localName === 't') nodes.push(t);
          }
        } else if (c.localName === 'hyperlink' || c.localName === 'fldSimple' ||
                   c.localName === 'smartTag' || c.localName === 'sdt') {
          walk(c);
        }
        /* drawing / pict / txbxContent 内的文字与段落正文无连续关系，跳过 */
      }
    }
    walk(p);
    return nodes;
  }

  function removeCjkLatinSpaces(doc, paras, roles) {
    var changed = 0;
    for (var i = 0; i < paras.length; i++) {
      var role = roles[i];
      if (role === 'front' || role === 'figure' || role === 'formula') continue;
      changed += joinCjkSpaceInPara(paras[i].p, role);
    }
    /* 表格单元格：正文清了表内不清会两套口径。表内只删空格字符，
       字号/字体/列宽/对齐一律不动，不会撑破版面。
       单元格段落同样套用「段首编号后保留一字符」的守卫。 */
    if (doc) {
      var tbls = wAll(doc, 'tbl');
      for (var t = 0; t < tbls.length; t++) {
        var cps = wAll(tbls[t], 'p');
        for (var c = 0; c < cps.length; c++) changed += joinCjkSpaceInPara(cps[c], 'body');
      }
    }
    return changed;
  }

  /* 删除单段里中文↔英文/数字之间的空格，返回删除的空格数 */
  function joinCjkSpaceInPara(p, role) {
    var changed = 0;
    {
      var nodes = paraTextNodes(p);
      if (!nodes.length) return 0;

      /* 拼出整段文字；owner[全局下标] = {n: 第几个节点, o: 节点内偏移} */
      var full = '', owner = [];
      for (var k = 0; k < nodes.length; k++) {
        var s = nodes[k].textContent || '';
        for (var q = 0; q < s.length; q++) owner.push({ n: k, o: q });
        full += s;
      }
      /* 标题/目录/题注：段首编号连同其后的空格一起保留（附件8：数字与文字间隔一字符） */
      var guard = KEEP_NUM_ROLE[role] ? (NUM_HEAD.exec(full) || [''])[0].length : 0;

      var del = {};
      var pats = [
        new RegExp('(' + CJK_CH + ')([ \\t]+)(' + LAT_CH + ')', 'g'),
        new RegExp('(' + LAT_CH + ')([ \\t]+)(' + CJK_CH + ')', 'g')
      ];
      for (var pi = 0; pi < pats.length; pi++) {
        var re = pats[pi], m;
        while ((m = re.exec(full))) {
          var start = m.index + m[1].length;
          for (var si = 0; si < m[2].length; si++) {
            if (start + si >= guard) del[start + si] = 1;
          }
          re.lastIndex = start + m[2].length;   // 从空格后的字符继续，允许连排命中
        }
      }
      var pos = Object.keys(del);
      if (!pos.length) return 0;

      /* 按节点回写：只删字符、不增不改，节点边界不受影响 */
      var hit = {};                             // 节点序号 → { 节点内偏移: 1 }
      for (var d = 0; d < pos.length; d++) {
        var ow = owner[+pos[d]];
        if (!ow) continue;
        (hit[ow.n] = hit[ow.n] || {})[ow.o] = 1;
      }
      for (var nk in hit) {
        if (!Object.prototype.hasOwnProperty.call(hit, nk)) continue;
        var node = nodes[+nk];
        var txt = node.textContent || '';
        var out = '';
        for (var ci = 0; ci < txt.length; ci++) { if (!hit[nk][ci]) out += txt.charAt(ci); }
        node.textContent = out;
        changed += txt.length - out.length;    // 计数按「删掉的空格数」，不是「改过的节点数」
      }
    }
    return changed;
  }

  /* ---------- 应用 pPr（行距/缩进/对齐） ---------- */
  function setSpacing(doc, pPr, line, before, after, lineRule) {
    var old = childByNs(pPr, 'spacing', W_NS);
    if (old) pPr.removeChild(old);
    if (!line) return;
    var sp = createW(doc, 'spacing');
    sp.setAttributeNS(W_NS, 'w:line', String(line));
    sp.setAttributeNS(W_NS, 'w:lineRule', lineRule || 'auto');
    if (before) sp.setAttributeNS(W_NS, 'w:before', String(before));
    if (after) sp.setAttributeNS(W_NS, 'w:after', String(after));
    insertInOrder(pPr, sp, PPR_ORDER);
  }

  function setIndent(doc, pPr, st) {
    var old = childByNs(pPr, 'ind', W_NS);
    if (old) pPr.removeChild(old);
    if (!(st.indent || st.hanging || st.rightChars || st.leftChars)) return;
    var ind = createW(doc, 'ind');
    if (st.indent) {
      ind.setAttributeNS(W_NS, 'w:firstLineChars', String(st.indent * 100));
      ind.setAttributeNS(W_NS, 'w:firstLine', String(Math.round(st.indent * st.size * 20)));
    }
    if (st.hanging) {
      ind.setAttributeNS(W_NS, 'w:hangingChars', String(st.hanging * 100));
      ind.setAttributeNS(W_NS, 'w:hanging', String(Math.round(st.hanging * st.size * 20)));
    }
    if (st.rightChars) {
      ind.setAttributeNS(W_NS, 'w:rightChars', String(st.rightChars * 100));
      ind.setAttributeNS(W_NS, 'w:right', String(Math.round(st.rightChars * st.size * 20)));
    }
    if (st.leftChars) {
      ind.setAttributeNS(W_NS, 'w:leftChars', String(st.leftChars * 100));
      ind.setAttributeNS(W_NS, 'w:left', String(Math.round(st.leftChars * st.size * 20)));
    }
    insertInOrder(pPr, ind, PPR_ORDER);
  }

  /* 设置段落"与下段同页/本段不跨页"属性（keepNext / keepLines） */
  function setKeep(pPr, doc, name) {
    if (!childByNs(pPr, name, W_NS)) {
      var el = createW(doc, name);
      insertInOrder(pPr, el, PPR_ORDER);
    }
  }

  function setJc(doc, pPr, jc) {
    var old = childByNs(pPr, 'jc', W_NS);
    if (old) pPr.removeChild(old);
    if (!jc) return;
    var j = createW(doc, 'jc');
    j.setAttributeNS(W_NS, 'w:val', jc);
    insertInOrder(pPr, j, PPR_ORDER);
  }

  function setParaFormat(p, doc, st) {
    var pPr = childByNs(p, 'pPr', W_NS);
    if (!pPr) { pPr = createW(doc, 'pPr'); p.insertBefore(pPr, p.firstChild); }
    setSpacing(doc, pPr, st.line, st.before, st.after, st.lineRule);
    setIndent(doc, pPr, st);
    setJc(doc, pPr, st.jc);

    // 段落标记 rPr：让空行/回车行字号一致
    var markRPr = childByNs(pPr, 'rPr', W_NS);
    if (!markRPr) { markRPr = createW(doc, 'rPr'); insertInOrder(pPr, markRPr, PPR_ORDER); }
    applyRPr(markRPr, doc, st, st.bold === true ? true : false);
  }

  /* 强制分页 run（w:br w:type="page"）；lastRenderedPageBreak 只是上次排版的
     自然断页位置，不算强制分页 */
  function hasPageBreakRun(p) {
    var brs = wAll(p, 'br');
    for (var i = 0; i < brs.length; i++) {
      if ((brs[i].getAttributeNS(W_NS, 'type') || '') === 'page') return true;
    }
    return false;
  }

  /* 段落是否以分页 run 结尾（段尾分页 → 后续内容必然从新页开始） */
  function breakAtParaEnd(p) {
    for (var i = p.childNodes.length - 1; i >= 0; i--) {
      var c = p.childNodes[i];
      if (c.nodeType !== 1) continue;
      if (c.localName === 'r') {
        var t = runText(c);
        var hasBr = hasPageBreakRun(c);
        if (t) return false;   // 最后的内容 run 带文字 → 分页不在段尾
        if (hasBr) return true;
        continue;              // 空 run（bookmark/格式等）→ 继续向前
      }
      if (c.localName === 'hyperlink' || c.localName === 'fldSimple' || c.localName === 'oMath' ||
          c.localName === 'drawing' || c.localName === 'pict' || c.localName === 'sdt') return false;
    }
    return false;
  }

  /* 分节段（pPr/sectPr 或 body 级 sectPr）是否强制换页：
     w:type 缺省即 nextPage；仅 continuous 不换页 */
  function sectPrBreaksPage(sectPr) {
    var tp = childByNs(sectPr, 'type', W_NS);
    var t = tp ? (tp.getAttributeNS(W_NS, 'val') || 'nextPage') : 'nextPage';
    return t !== 'continuous';
  }

  /* 惰性元素：书签/校对标记等不占版面的节点，回溯分页检测时跨过 */
  var INERT_ELEMS = { bookmarkStart: 1, bookmarkEnd: 1, proofErr: 1, permStart: 1, permEnd: 1,
    commentRangeStart: 1, commentRangeEnd: 1, commentReference: 1 };

  /* 一级标题换页：h1 前若已有强制分页（分节符、段尾分页 run、空段上的分页/
     pageBreakBefore），说明标题已必然位于新页顶部，不再叠加 pageBreakBefore，
     避免双重分页产生空白页。回溯时跨过空段落与书签等惰性元素——原文档常以
     「内容 + 空段 + 分节段 + 空段 + 标题」的形态分章，只检查紧邻段会漏判 */
  function addPageBreakIfNeeded(p, doc) {
    var pPr = childByNs(p, 'pPr', W_NS);
    if (pPr && childByNs(pPr, 'sectPr', W_NS)) return;   // 分节段必然换页
    if (hasPageBreakRun(p)) return;                      // 标题段自带分页 → 已换页

    var prev = p.previousSibling;
    while (prev && prev.nodeType !== 1) prev = prev.previousSibling;
    while (prev) {
      if (prev.localName === 'p') {
        var ppPr = childByNs(prev, 'pPr', W_NS);
        var ss = ppPr ? childByNs(ppPr, 'sectPr', W_NS) : null;
        if (ss) { if (sectPrBreaksPage(ss)) return; }    // 分节换页 → 已在新页顶部
        var hasContent = wAll(prev, 't').length > 0 || wAll(prev, 'drawing').length > 0 ||
                         wAll(prev, 'pict').length > 0 || allByNs(prev, 'oMath', M_NS).length > 0;
        if (!hasContent) {
          if (hasPageBreakRun(prev)) return;             // 空段上的分页 run
          var hasPbb = ppPr && childByNs(ppPr, 'pageBreakBefore', W_NS);
          if (hasPbb) return;                            // 空段 pageBreakBefore → 标题紧随新页顶部
          prev = prev.previousSibling;                   // 纯空段 → 继续向上
          while (prev && prev.nodeType !== 1) prev = prev.previousSibling;
          continue;
        }
        if (breakAtParaEnd(prev)) return;                // 内容段段尾分页 → 标题从新页开始
        break;                                           // 最后内容段无强制分页 → 需要换页
      }
      if (prev.localName === 'sectPr') { if (sectPrBreaksPage(prev)) return; break; }  // body 级分节符
      if (INERT_ELEMS[prev.localName]) {                 // 书签等惰性元素 → 跨过继续
        prev = prev.previousSibling;
        while (prev && prev.nodeType !== 1) prev = prev.previousSibling;
        continue;
      }
      break;                                             // 表格等实质内容 → 需要换页
    }

    if (!pPr) { pPr = createW(doc, 'pPr'); p.insertBefore(pPr, p.firstChild); }
    if (!childByNs(pPr, 'pageBreakBefore', W_NS)) {
      var pb = createW(doc, 'pageBreakBefore');
      insertInOrder(pPr, pb, PPR_ORDER);
    }
  }

  /* 目录条目按层级左缩进（下级目录依次右缩进两个字符） */
  function tocIndentChars(text) {
    if (/^\d{1,2}\.\d{1,2}\.\d{1,2}/.test(text)) return 4;   // 三级
    if (/^\d{1,2}\.\d{1,2}/.test(text)) return 2;            // 二级
    return 0;                                                 // 一级/其他
  }

  /* ---------- 应用全部格式 ---------- */
  function applyFormatting(doc, paras, roles, settings) {
    var S = buildStyles(settings);
    for (var i = 0; i < paras.length; i++) {
      var role = roles[i];
      if (role === 'front' || role === 'figure') continue;
      var st = S[role];
      if (!st) continue;
      var p = paras[i].p;
      if (role === 'formula') { setJc(doc, p, 'center'); continue; } // 公式：仅居中，字体不动
      // 章节换页（一级标题之间换页、目录另起一页）由 normalizeChapterBreaks 统一处理
      if (role === 'tocItem') {
        var lc = tocIndentChars(paras[i].text);
        if (lc) st = { eastFont: st.eastFont, latinFont: st.latinFont, size: st.size, jc: st.jc,
                       line: st.line, lineRule: st.lineRule, before: st.before, after: st.after,
                       leftChars: lc };
      }
      setParaFormat(p, doc, st);
      // 标题加大纲级别（供自动目录域 TOC 收集；不影响视觉排版）
      if (role === 'h1' || role === 'h2' || role === 'h3') {
        var olPPr = childByNs(p, 'pPr', W_NS);
        var ol = childByNs(olPPr, 'outlineLvl', W_NS);
        if (!ol) { ol = createW(doc, 'outlineLvl'); insertInOrder(olPPr, ol, PPR_ORDER); }
        ol.setAttributeNS(W_NS, 'w:val', String(role === 'h1' ? 0 : role === 'h2' ? 1 : 2));
      }
      if (role === 'caption') {
        // 题注与图片/表格排版在同一页：图题随上方图片（图片段带 keepNext），
        // 表题 keepNext 跟随下方表格；题注自身不跨页
        var cpPr = childByNs(p, 'pPr', W_NS);
        if (/^表/.test(paras[i].text)) setKeep(cpPr, doc, 'keepNext');
        setKeep(cpPr, doc, 'keepLines');
      }
      if (role === 'kw' || role === 'kwEn') formatKeywords(p, doc, st, role);
      else formatRuns(p, doc, st);
      /* 引用标注上标：参考文献表序号（refItem）与封面（front）不处理 */
      if (settings.citeSuperscript !== false &&
          role !== 'front' && role !== 'figure' && role !== 'refItem' && role !== 'formula') {
        applyCiteSuperscript(p, doc);
      }
    }
  }

  /* ---------- 表格（参考样张：顶线/底线 1.5 磅 + 栏目线 0.75 磅的三线表） ---------- */
  var TBLPR_ORDER = ['tblStyle','tblpPr','tblOverlap','bidiVisual','tblStyleRowBandSize','tblStyleColBandSize',
    'tblW','jc','tblCellSpacing','tblInd','tblBorders','shd','tblLayout','tblCellMar','tblLook','tblCaption','tblDescription'];
  var TCPR_ORDER = ['cnfStyle','tcW','gridSpan','hMerge','vMerge','tcBorders','shd','noWrap','tcMar',
    'textDirection','tcFitText','vAlign','hideMark','headers','cellIns','cellDel','cellMerge','tcPrChange'];
  var TRPR_ORDER = ['cnfStyle','divId','gridBefore','gridAfter','wBefore','wAfter','cantSplit','trHeight',
    'tblHeader','tblCellSpacing','jc','hidden','ins','del','trPrChange'];
  var BORDER_ORDER = ['top','left','bottom','right','insideH','insideV'];

  function setBorder(borders, doc, name, sz, val) {
    var el = childByNs(borders, name, W_NS);
    if (!el) { el = createW(doc, name); insertInOrder(borders, el, BORDER_ORDER); }
    el.setAttributeNS(W_NS, 'w:val', val || 'single');
    el.setAttributeNS(W_NS, 'w:sz', String(sz));
    el.setAttributeNS(W_NS, 'w:space', '0');
    el.setAttributeNS(W_NS, 'w:color', '000000');
  }

  /* 单元格上/下边线显式设置（top/bottom 为磅值或空），其余五边显式 none，
     防止表格样式回渗导致线条不显示 */
  function setCellLines(tc, doc, top, bottom) {
    var tcPr = childByNs(tc, 'tcPr', W_NS);
    if (!tcPr) { tcPr = createW(doc, 'tcPr'); tc.insertBefore(tcPr, tc.firstChild); }
    var tcb = childByNs(tcPr, 'tcBorders', W_NS);
    if (tcb) tcPr.removeChild(tcb);
    tcb = createW(doc, 'tcBorders');
    setBorder(tcb, doc, 'top', top || 4, top ? 'single' : 'none');
    setBorder(tcb, doc, 'left', 4, 'none');
    setBorder(tcb, doc, 'bottom', bottom || 4, bottom ? 'single' : 'none');
    setBorder(tcb, doc, 'right', 4, 'none');
    setBorder(tcb, doc, 'insideH', 4, 'none');
    setBorder(tcb, doc, 'insideV', 4, 'none');
    insertInOrder(tcPr, tcb, TCPR_ORDER);
  }

  /* ---------- 表格内文字格式 ----------
     附件8：「表序、表名和表格内字体均为五号宋体」——只说了「表格内」，没给表头
     开例外，所以表头与数据行一律五号宋体、不加粗；西文与数字用 Times New Roman
     （与正文口径一致）。样章里「五号宋体」也重复标注了三次，没有别的要求。
     表格内的段落**不在 collectParas 的收集范围**里（那里只收 body 的直接子段落，
     表格嵌套在 tbl 下面），所以正文那套 applyFormatting 一个字也改不到表内 ——
     必须单独走这一道。2026-09-16 作者反馈「表格内的文字部分格式未修改」就是这个
     原因：表里混着宋体/黑体/Consolas，字号 20（10pt）和 32（16pt）并存。

     这里只管字体字号；单元格底纹归 clearShading() 管（全文一次清干净），
     免得「清除底色」关掉之后表底纹只清一半，变成两套口径。 */
  function runsInPara(p) {
    /* 只要本段的 run：段里若嵌了嵌套表格，那张表里的 run 不归本段管 */
    var out = [];
    var all = allByNs(p, 'r', W_NS);
    for (var i = 0; i < all.length; i++) {
      var crossed = false;
      for (var a = all[i].parentNode; a && a !== p; a = a.parentNode) {
        if (a.nodeType === 1 && a.localName === 'tbl') { crossed = true; break; }
      }
      if (!crossed) out.push(all[i]);
    }
    return out;
  }

  function formatTableText(doc, tbl, st) {
    var n = 0;
    var tcs = wAll(tbl, 'tc');
    for (var i = 0; i < tcs.length; i++) {
      var tc = tcs[i];
      for (var c = 0; c < tc.childNodes.length; c++) {
        var p = tc.childNodes[c];
        if (p.nodeType !== 1 || p.localName !== 'p') continue;
        var runs = runsInPara(p);
        for (var k = 0; k < runs.length; k++) {
          if (!wAll(runs[k], 't').length) continue;   // 只带域/图片的 run 不动
          var rPr = childByNs(runs[k], 'rPr', W_NS);
          if (!rPr) { rPr = createW(doc, 'rPr'); runs[k].insertBefore(rPr, runs[k].firstChild); }
          applyRPr(rPr, doc, st, false);              // 表头也走 false：不加粗
          n++;
        }
      }
    }
    return n;
  }

  function applyTableTopBottom(tbl, doc) {
    var tblPr = childByNs(tbl, 'tblPr', W_NS);
    if (!tblPr) { tblPr = createW(doc, 'tblPr'); tbl.insertBefore(tblPr, tbl.firstChild); }
    var borders = childByNs(tblPr, 'tblBorders', W_NS);
    if (borders) tblPr.removeChild(borders);
    borders = createW(doc, 'tblBorders');
    setBorder(borders, doc, 'top', 12, 'single');       // 顶线 1.5磅 = 12/8
    setBorder(borders, doc, 'bottom', 12, 'single');    // 底线 1.5磅
    setBorder(borders, doc, 'left', 4, 'none');
    setBorder(borders, doc, 'right', 4, 'none');
    setBorder(borders, doc, 'insideH', 4, 'none');
    setBorder(borders, doc, 'insideV', 4, 'none');
    insertInOrder(tblPr, borders, TBLPR_ORDER);

    // 清除所有单元格既有边框（原栏目线、内部线、竖线全部清除）
    var allTcs = wAll(tbl, 'tc');
    for (var a = 0; a < allTcs.length; a++) {
      var tcPr0 = childByNs(allTcs[a], 'tcPr', W_NS);
      if (!tcPr0) continue;
      var oldTcb = childByNs(tcPr0, 'tcBorders', W_NS);
      if (oldTcb) tcPr0.removeChild(oldTcb);
    }

    /* 三线表线条全部在单元格级显式设置（参考样张），确保显示：
       表头行：顶线 1.5 磅 + 栏目线 0.75 磅；末行：底线 1.5 磅 */
    var trs = [];
    for (var hr = 0; hr < tbl.childNodes.length; hr++) {
      var hc = tbl.childNodes[hr];
      if (hc.nodeType === 1 && hc.localName === 'tr') trs.push(hc);
    }
    if (trs.length > 0) {
      var first = trs[0], last = trs[trs.length - 1];
      for (var hx = 0; hx < first.childNodes.length; hx++) {
        var htc = first.childNodes[hx];
        if (htc.nodeType === 1 && htc.localName === 'tc') setCellLines(htc, doc, 12, 6);
      }
      if (last !== first) {
        for (var lx = 0; lx < last.childNodes.length; lx++) {
          var ltc = last.childNodes[lx];
          if (ltc.nodeType === 1 && ltc.localName === 'tc') setCellLines(ltc, doc, null, 12);
        }
      }
    }

    // 表格行不得跨页断裂（防表格断层）：每行加 cantSplit
    for (var rr = 0; rr < tbl.childNodes.length; rr++) {
      var rc = tbl.childNodes[rr];
      if (rc.nodeType !== 1 || rc.localName !== 'tr') continue;
      /* 行级表格属性例外 w:tblPrEx 里的 tblBorders 优先于表级 tblBorders：
         源文档每行都带蓝色全网格（#9AA5B8），不清掉则三线表失效 */
      var exs = wAll(rc, 'tblPrEx');
      for (var e2 = exs.length - 1; e2 >= 0; e2--) {
        if (exs[e2].parentNode === rc) rc.removeChild(exs[e2]);
      }
      var trPr = childByNs(rc, 'trPr', W_NS);
      if (!trPr) { trPr = createW(doc, 'trPr'); rc.insertBefore(trPr, rc.firstChild); }
      if (!childByNs(trPr, 'cantSplit', W_NS)) {
        var cs = createW(doc, 'cantSplit');
        insertInOrder(trPr, cs, TRPR_ORDER);
      }
    }
  }

  /* 表格与图片紧邻（<w:tbl> 直接挨着含图片的段落，或反之）时，
     在两者之间插入一个空段落，避免排版重叠。前置部分（封面）不插入 */
  function separateTablesAndImages(doc, frontBoundary) {
    var body = allByNs(doc, 'body', W_NS)[0];
    if (!body) return 0;
    var kids = [];
    for (var i = 0; i < body.childNodes.length; i++) {
      var c = body.childNodes[i];
      if (c.nodeType === 1) kids.push(c);
    }
    var inserted = 0;
    for (var k = 0; k < kids.length - 1; k++) {
      var a = kids[k], b = kids[k + 1];
      if (isFrontElement(a, frontBoundary) || isFrontElement(b, frontBoundary)) continue;
      var aImg = a.localName === 'p' && (wAll(a, 'drawing').length > 0 || wAll(a, 'pict').length > 0);
      var bImg = b.localName === 'p' && (wAll(b, 'drawing').length > 0 || wAll(b, 'pict').length > 0);
      if ((a.localName === 'tbl' && bImg) || (aImg && b.localName === 'tbl')) {
        body.insertBefore(createW(doc, 'p'), b);
        inserted++;
      }
    }
    return inserted;
  }

  /* ---------- 前置部分（摘要之前，通常为封面+声明两页）保护 ----------
     格式修改一律从检测到的摘要开始；摘要之前的段落（front 角色）本就不动，
     这里再对表格、图片段、分隔插入与分节设置做同样的边界保护 */
  /* 前置部分判定：以「第一个非 front 段落」（通常为摘要标题）的元素引用为边界。
     用元素引用而非静态索引，因为 normalize 会删除/插入元素导致索引漂移 */
  function isFrontElement(el, frontBoundary) {
    if (!frontBoundary) return false;
    if (el === frontBoundary) return false;
    var k = el.previousSibling;
    while (k) {
      if (k === frontBoundary) return false;   // 在边界之后 → 非前置
      k = k.previousSibling;
    }
    return true;                               // 在边界之前 → 前置
  }

  function collectFrontSectPrs(doc, frontBoundary) {
    var out = [];
    if (!frontBoundary) return out;
    var body = allByNs(doc, 'body', W_NS)[0];
    if (!body) return out;
    for (var i = 0; i < body.childNodes.length; i++) {
      var c = body.childNodes[i];
      if (c.nodeType !== 1) continue;
      if (c === frontBoundary) break;
      if (c.localName === 'p') {
        var pPr = childByNs(c, 'pPr', W_NS);
        var ss = pPr && childByNs(pPr, 'sectPr', W_NS);
        if (ss) out.push(ss);
      } else if (c.localName === 'sectPr') {
        out.push(c);
      }
    }
    return out;
  }

  /* 将 body 级 w:sdt（Word 目录域、内容控件）解包为普通段落：
     目录条目常被包在 sdt 里，不展开就无法参与结构识别与格式修改 */
  function unwrapSdt(doc) {
    var body = allByNs(doc, 'body', W_NS)[0];
    if (!body) return 0;
    var n = 0;
    for (var i = 0; i < body.childNodes.length; i++) {
      var c = body.childNodes[i];
      if (c.nodeType !== 1 || c.localName !== 'sdt') continue;
      var content = childByNs(c, 'sdtContent', W_NS);
      if (content) {
        while (content.firstChild) body.insertBefore(content.firstChild, c);
      }
      body.removeChild(c);
      n++;
    }
    return n;
  }

  /* 域字符（w:fldChar）配平兜底 —— 保证导出的文档不存在不配平的域。
     出问题的场景：文档里本来就有 Word 目录域，重新格式化时若目录区只删掉一半，
     正文里就会留下一个没有 begin 的 end；Word 打开（更新域）后会把其后的内容
     当成域的一部分处理，用户看到的就是「正文内容全部消失」。
     这里删掉没有 begin 配对的 end/separate，并给没有 end 收尾的 begin 就地补 end。 */
  function repairFields(doc) {
    var body = allByNs(doc, 'body', W_NS)[0];
    if (!body) return 0;
    var stack = [], fixed = 0;

    function dropRunOf(fc) {
      var r = fc.parentNode;
      if (r && r.nodeType === 1 && r.localName === 'r') {
        r.removeChild(fc);
        if (!r.firstChild && r.parentNode) r.parentNode.removeChild(r);
      } else if (r) {
        r.removeChild(fc);
      }
    }

    (function walk(node, curPara) {
      for (var n = node.firstChild; n; n = n.nextSibling) {
        if (n.nodeType !== 1) continue;
        var ln = n.localName;
        var p = (ln === 'p') ? n : curPara;
        if (p !== curPara) {
          /* 进入新段落：尚未闭合的域若继续延伸，最后一段跟着前移 */
          for (var s = 0; s < stack.length; s++) stack[s].para = p;
        }
        if (ln === 'fldChar') {
          var t = n.getAttributeNS(W_NS, 'fldCharType') || '';
          if (t === 'begin') stack.push({ node: n, para: p });
          else if (t === 'end') {
            if (stack.length) stack.pop();
            else { dropRunOf(n); fixed++; }         // 孤儿 end
          } else if (t === 'separate') {
            if (!stack.length) { dropRunOf(n); fixed++; }   // 孤儿 separate
          }
          continue;
        }
        walk(n, p);
      }
    })(body, null);

    for (var i = 0; i < stack.length; i++) {        // 未闭合的 begin → 补 end
      var host = stack[i].para;
      if (!host) continue;
      var r = createW(doc, 'r');
      var fc = createW(doc, 'fldChar');
      fc.setAttributeNS(W_NS, 'w:fldCharType', 'end');
      r.appendChild(fc);
      host.appendChild(r);
      fixed++;
    }
    return fixed;
  }

  /* 图片所在段落统一改为单倍行距（line=240, lineRule=auto）：
     固定行距（如 20 磅）会把高于行距的图片截断显示不全，
     单倍行距行高随内容自适应，图片可完整显示。
     段前段后（before/after）保留不动。前置部分（封面）图片不修改 */
  function singleSpaceImageParas(doc, frontBoundary) {
    var ps = wAll(doc, 'p');
    for (var i = 0; i < ps.length; i++) {
      var p = ps[i];
      if (!(wAll(p, 'drawing').length > 0 || wAll(p, 'pict').length > 0)) continue;
      if (isFrontElement(p, frontBoundary)) continue;
      var pPr = childByNs(p, 'pPr', W_NS);
      if (!pPr) { pPr = createW(doc, 'pPr'); p.insertBefore(pPr, p.firstChild); }
      var sp = childByNs(pPr, 'spacing', W_NS);
      if (!sp) { sp = createW(doc, 'spacing'); insertInOrder(pPr, sp, PPR_ORDER); }
      sp.setAttributeNS(W_NS, 'w:line', '240');
      sp.setAttributeNS(W_NS, 'w:lineRule', 'auto');
      // 图片与题注排版在同一页：图片段 keepNext（跟随图题）+ keepLines（图片自身不跨页）
      setKeep(pPr, doc, 'keepNext');
      setKeep(pPr, doc, 'keepLines');
      // 图片居中对齐（附件8：插图居中排版）
      setJc(doc, pPr, 'center');
    }
  }

  /* ---------- 一级标题改「1 绪论」体例（附件8 理工类） ----------
     附件8：一级标题写「1 绪论」——章号用阿拉伯数字，数字与文字间隔一字符，
     不写「第1章 绪论」。只动标题段与目录里的对应条目；正文里指代章节的
     「第5章」保持不动——那是叙述不是标题（「数据生成规则将在第5章如实给出」），
     改了反而不通顺。
     三级标题同一句话的另一半：附件8 要求「数字与文字间隔一字符」，
     「3.1.1系统目标」要补成「3.1.1 系统目标」（只补空格，编号一个字符不改）。
     附录：附件8「附录如果为多个附件，依序用附录A、附录B、附录C……编序号，
     否则只用『附录』」，所以单个附录去掉序号（附录1 → 附录）、多个附录依序
     给字母（附录1 / 附录2 → 附录A / 附录B）；附录里的图表随之用字母编号
     （图A1，字母与序号之间没有「.」）。 */
  var CN_DIG = { 零: 0, 一: 1, 二: 2, 两: 2, 三: 3, 四: 4, 五: 5, 六: 6, 七: 7, 八: 8, 九: 9 };
  /* 「二十三」→ 23；认不出来（如「第一章」之外的花样）返回空串，调用处放弃改写 */
  function cnNumToArabic(s) {
    if (/^\d+$/.test(s)) return String(parseInt(s, 10));
    var total = 0, cur = 0;
    for (var i = 0; i < s.length; i++) {
      var ch = s.charAt(i);
      if (ch === '十') { cur = (cur || 1) * 10; total += cur; cur = 0; }
      else if (ch === '百') { cur = (cur || 1) * 100; total += cur; cur = 0; }
      else if (ch === '千') { cur = (cur || 1) * 1000; total += cur; cur = 0; }
      else if (CN_DIG[ch] !== undefined) cur = CN_DIG[ch];
      else return '';
    }
    total += cur;
    return total > 0 ? String(total) : '';
  }

  /* 改写文本节点内容。首尾带空格的（章号常被拆成「1 」+「绪论」两个 run）
     要补 xml:space="preserve" —— 否则 Word 会把那个空格吃掉，标题成「1绪论」 */
  function setRunText(node, s) {
    while (node.firstChild) node.removeChild(node.firstChild);
    node.appendChild(node.ownerDocument.createTextNode(s));
    if (/^\s|\s$/.test(s)) node.setAttributeNS(XML_NS, 'xml:space', 'preserve');
  }

  /* 把段落开头的 n 个字符换成 newText。章号常被拆在「第」「1」「章」几个 run
     里，所以整段拼串定位，再按下标写回各文本节点（只改文字，格式不动） */
  function replaceParaPrefix(p, n, newText) {
    var nodes = paraTextNodes(p);
    if (!nodes.length) return false;
    var texts = [], starts = [], acc = 0;
    for (var i = 0; i < nodes.length; i++) {
      texts.push(nodes[i].textContent || '');
      starts.push(acc);
      acc += texts[i].length;
    }
    if (acc < n) return false;
    /* 新编号写进「前缀最后一个字符所在」的那个节点：前缀正好在节点边界结束
       （「第1章 」自成一个 run、标题在下一个 run）时也算它，否则编号会被清空。
       它前面的节点清空，它后面的节点原样保留 */
    var last = 0;
    for (var k = 0; k < nodes.length; k++) {
      if (starts[k] < n) last = k; else break;
    }
    for (var c = 0; c < last; c++) setRunText(nodes[c], '');
    setRunText(nodes[last], newText + texts[last].slice(n - starts[last]));
    return true;
  }

  /* 返回改写的段落数；paras[i].text 同步更新（后续几个 pass 还在读它） */
  function rewriteChapterHeadings(paras, roles) {
    /* 附录先点名：附件8 只有一个附录时不编序号，多个才依序编 A、B、C。
       若各附录本来就有字母，保留原字母（正文里「详见附录B」这类指代不会错位） */
    var apps = [], allLetter = true;
    for (var a = 0; a < paras.length; a++) {
      if (roles[a] !== 'h1' || !RE.app.test(paras[a].text)) continue;
      apps.push(a);
      if (!/^附\s*录\s*[A-Za-z]/.test(paras[a].text)) allLetter = false;
    }
    var appLetter = {};
    for (var b = 0; b < apps.length; b++) {
      if (apps.length < 2) appLetter[apps[b]] = '';               // 单个附录：只用「附录」
      else if (allLetter) appLetter[apps[b]] = /^附\s*录\s*([A-Za-z])/.exec(paras[apps[b]].text)[1].toUpperCase();
      else appLetter[apps[b]] = b < 26 ? String.fromCharCode(65 + b) : String(b + 1);
    }

    var changed = 0, tocAppSeen = 0;
    for (var i = 0; i < paras.length; i++) {
      var role = roles[i];
      if (role !== 'h1' && role !== 'h3' && role !== 'tocItem') continue;
      var live = paraText(paras[i].p);
      var lead = live.length - live.replace(/^\s+/, '').length;   // 段首空格不参与匹配
      var head = live.slice(lead);
      var n = 0, newText = '';

      var m = /^第\s*([一二三四五六七八九十百千零两0-9]+)\s*章\s*/.exec(head);
      if (m) {
        /* 「第4章 系统总体设计」→「4 系统总体设计」；「第一章 绪论」一并转成阿拉伯数字 */
        var num = cnNumToArabic(m[1]);
        if (!num) continue;
        n = m[0].length;
        newText = head.slice(n) ? num + ' ' : num;                // 编号与标题间隔一字符（附件8）
      } else if (RE.app.test(head)) {
        var letter;
        if (role === 'h1') letter = appLetter[i] || '';
        else {                                                    // 目录条目按顺序对应各附录
          var src = apps[tocAppSeen]; tocAppSeen++;
          letter = src === undefined ? '' : (appLetter[src] || '');
        }
        var ma = /^附\s*录\s*[A-Za-z0-9一二三四五六七八九十]*[.\s]*/.exec(head);
        if (!ma) continue;
        n = ma[0].length;
        newText = (head.slice(n) ? '附录' + letter + ' ' : '附录' + letter);
      } else if (role === 'h3') {
        /* 「3.1.1系统目标」→「3.1.1 系统目标」：附件8 要求编号与文字间隔一字符。
           三级标题的识别已经允许编号后不带空格（见 RE.h3），这里只把缺的那个空格补上。
           本来就有空格的（「3.5.1 技术可行性」）直接跳过，免得补成两个空格。 */
        var m3 = /^[1-9]\d?\.\d{1,2}\.\d{1,2}(?![\d.])/.exec(head);
        if (!m3 || /^\s/.test(head.slice(m3[0].length))) continue;
        n = m3[0].length;
        newText = m3[0] + ' ';
      }
      if (!newText) continue;
      var result = newText + head.slice(n);
      if (result === head) continue;                              // 已是目标体例：不动
      if (!replaceParaPrefix(paras[i].p, lead + n, newText)) continue;
      paras[i].text = result.trim();
      changed++;
    }
    return changed;
  }

  /* ---------- 图片题注改用 Word 真题注（题注样式 + SEQ 域自动编号） ----------
     附件8：图按章编号、图题置于图片正下方。不用手打数字，而是套 Word 的
     「题注」样式 + SEQ 域：
         图<章号>.{ SEQ 图 \* ARABIC \r 1 }  名称
     每章第一条题注带 \r 1（把序号重置为 1），本章其余题注用普通 SEQ 递增，
     于是增删图片时章内序号自动重排，也能用「引用 → 插入表目录」生成图表目录。
     章号写成固定文字：本工具的章标题是手打编号，STYLEREF \s 取不到章号。
     附录里的图用附录字母编号（图A1，附件8：附录图表另行编号）。 */
  function captionChapterNo(text) {
    var t = String(text || '').trim();
    var m = /^第\s*([1-9]\d?)\s*章/.exec(t);         // 「第4章 系统总体设计」（旧体例）
    if (m) return m[1];
    m = /^([1-9]\d?)\s+\S/.exec(t);                  // 「4 系统总体设计」（附件8 体例）
    if (m) return m[1];
    m = /^附\s*录\s*([A-Za-z0-9])/.exec(t);          // 「附录A 部署与运行步骤」
    if (m) return m[1].toUpperCase();
    if (RE.app.test(t)) return 'A';                  // 只一个附录时不编序号，图仍用字母（图A1）
    return '';
  }
  /* 从现有题注里取出名称：「图4-1 系统总体架构图」→「系统总体架构图」 */
  function captionName(text) {
    return String(text || '')
      .replace(/^(图|表|Figure|Table)\s*[A-Z]?\.?\s*\d+(?:\s*[.\-–—]\s*\d+)*\s*/, '')
      .replace(/\s+/g, ' ').trim();
  }

  async function ensureCaptionStyle(zip) {
    var entry = zip.file('word/styles.xml');
    if (!entry) return null;
    var stDoc = new DOMParser().parseFromString(await entry.async('string'), 'application/xml');
    var styles = stDoc.getElementsByTagNameNS(W_NS, 'style');
    var used = {}, found = null;
    for (var i = 0; i < styles.length; i++) {
      var sid = styles[i].getAttributeNS(W_NS, 'styleId') || '';
      if (sid) used[sid] = true;
      var nm = childByNs(styles[i], 'name', W_NS);
      if (!nm) continue;
      var n = (nm.getAttributeNS(W_NS, 'val') || '').trim();
      /* 内置「题注」样式的规范名是 caption（中文版 Word 显示为「题注」） */
      if (!found && /^(caption|题注)$/i.test(n)) found = sid;
    }
    if (!found) {
      var newId = 'Caption', k = 1;
      while (used[newId]) { k++; newId = 'Caption' + k; }
      var st = createW(stDoc, 'style');
      st.setAttributeNS(W_NS, 'w:type', 'paragraph');
      st.setAttributeNS(W_NS, 'w:styleId', newId);
      var nEl = createW(stDoc, 'name');
      nEl.setAttributeNS(W_NS, 'w:val', 'caption');
      st.appendChild(nEl);
      var bo = createW(stDoc, 'basedOn');
      bo.setAttributeNS(W_NS, 'w:val', 'Normal');
      st.appendChild(bo);
      var nx = createW(stDoc, 'next');
      nx.setAttributeNS(W_NS, 'w:val', 'Normal');
      st.appendChild(nx);
      stDoc.documentElement.appendChild(st);
      zip.file('word/styles.xml', serialize(stDoc, 'word/styles.xml'));
      found = newId;
    }
    return found;
  }

  /* 重写题注段落内容（保留 pPr）：图<章>. + SEQ 域 + 空格 + 名称。
     seqName 为「图」或「表」——SEQ 域按名字各自计数，图和表互不干扰。 */
  function setCaptionContent(doc, p, chap, no, name, seqName, isApp) {
    for (var i = p.childNodes.length - 1; i >= 0; i--) {
      var c = p.childNodes[i];
      if (c.nodeType === 1 && c.localName === 'pPr') continue;
      p.removeChild(c);
    }
    function txtRun(s) {
      var r = createW(doc, 'r');
      var t = createW(doc, 't');
      t.setAttributeNS(XML_NS, 'xml:space', 'preserve');
      t.appendChild(doc.createTextNode(s));
      r.appendChild(t);
      return r;
    }
    function fldRun(kind, instr) {
      var r = createW(doc, 'r');
      if (kind === 'instr') {
        var it = createW(doc, 'instrText');
        it.setAttributeNS(XML_NS, 'xml:space', 'preserve');
        it.appendChild(doc.createTextNode(instr));
        r.appendChild(it);
      } else if (kind === 'text') {
        var t = createW(doc, 't');
        t.setAttributeNS(XML_NS, 'xml:space', 'preserve');
        t.appendChild(doc.createTextNode(instr));
        r.appendChild(t);
      } else {
        var fc = createW(doc, 'fldChar');
        fc.setAttributeNS(W_NS, 'w:fldCharType', kind);
        r.appendChild(fc);
      }
      return r;
    }
    var sn = seqName || '图';
    /* 附录里用字母编号，附件8 的写法是「图A1」（字母与序号之间没有「.」）；
       正文里是「图4.1」 */
    var sep = (isApp && /^[A-Z]$/.test(String(chap))) ? '' : '.';
    p.appendChild(txtRun(sn + chap + sep));
    p.appendChild(fldRun('begin'));
    p.appendChild(fldRun('instr', ' SEQ ' + sn + ' \\* ARABIC' + (no === 1 ? ' \\r 1' : '') + ' '));
    p.appendChild(fldRun('separate'));
    p.appendChild(fldRun('text', String(no)));
    p.appendChild(fldRun('end'));
    if (name) p.appendChild(txtRun(' ' + name));
    return p;
  }

  function setPStyle(doc, p, styleId) {
    if (!styleId) return;
    var pPr = childByNs(p, 'pPr', W_NS);
    if (!pPr) { pPr = createW(doc, 'pPr'); p.insertBefore(pPr, p.firstChild); }
    var ps = childByNs(pPr, 'pStyle', W_NS);
    if (!ps) { ps = createW(doc, 'pStyle'); pPr.insertBefore(ps, pPr.firstChild); }
    ps.setAttributeNS(W_NS, 'w:val', styleId);
  }

  /* 遍历正文，把图片下方的题注换成分域真题注；没有题注的补一条占位题注。
     新写入的 run 不带格式，这里按题注样式重新排一遍。
     返回 { done, placeholders }，占位题注的名称留空，由作者补写。 */
  function applyWordCaptions(doc, body, paras, roles, frontBoundary, capSt, capStyleId) {
    var curChap = '', curApp = false, figNo = 0, tabNo = 0, done = 0, placeholders = 0;

    /* 段落元素 → 角色。图和表共用同一个章号，必须顺着 body 一次走完，
       不能先遍历图、再遍历表，否则跨章的序号会错。 */
    var roleByEl = {};
    for (var r = 0; r < paras.length; r++) roleByEl[paras[r].p] = roles[r];

    /* 题注要跟图片/表格排在同一页，且自身不跨页 */
    function finishCaption(capEl, chap, no, name, seqName, isApp) {
      setCaptionContent(doc, capEl, chap, no, name, seqName, isApp);
      var cpPr = childByNs(capEl, 'pPr', W_NS);
      if (!cpPr) { cpPr = createW(doc, 'pPr'); capEl.insertBefore(cpPr, capEl.firstChild); }
      setKeep(cpPr, doc, 'keepLines');
      /* 新 run 没有字体字号：按题注样式补排一遍 */
      if (capSt) { setParaFormat(capEl, doc, capSt); formatRuns(capEl, doc, capSt); }
      setPStyle(doc, capEl, capStyleId);
      done++;
    }
    /* 跨过空段找相邻的实质元素（dir=1 向下、-1 向上） */
    function skipBlanks(from, dir) {
      var s = from;
      while (s) {
        if (s.nodeType !== 1) { s = dir > 0 ? s.nextSibling : s.previousSibling; continue; }
        if (s.localName === 'p' && !paraText(s).trim() &&
            !wAll(s, 'drawing').length && !wAll(s, 'pict').length) {
          s = dir > 0 ? s.nextSibling : s.previousSibling; continue;
        }
        return s;
      }
      return null;
    }

    /* 必须先快照 body 的子元素：给没有题注的表补占位题注时要插在表格**之前**，
       按 live 的 childNodes 下标遍历会让同一张表再被访问一次，表号跳成 1、3、5… */
    var kids = [];
    for (var c0 = 0; c0 < body.childNodes.length; c0++) {
      var c1 = body.childNodes[c0];
      if (c1.nodeType === 1) kids.push(c1);
    }
    for (var i = 0; i < kids.length; i++) {
      var el = kids[i];

      if (el.localName === 'p') {
        var role = roleByEl[el];
        if (role === 'h1' || role === 'refHead') {
          /* 认不出章号（参考文献/致谢等）→ 停用编号，这些区域不误编成上一章的号 */
          var hText = paraText(el).trim();
          curChap = captionChapterNo(hText);
          curApp = RE.app.test(hText);          // 附录：图题用「图A1」而不是「图A.1」
          figNo = 0; tabNo = 0;
          continue;
        }
        if (role !== 'figure') continue;
        if (isFrontElement(el, frontBoundary)) continue;   // 封面图片不动
        if (!curChap) continue;                           // 不在任何章节内 → 不动

        /* 图题在图片正下方 */
        var nxt = skipBlanks(el.nextSibling, 1);
        var capEl = null, name = '';
        if (nxt && nxt.localName === 'p') {
          var nt = paraText(nxt).trim();
          if (RE.cap.test(nt) && /^(图|Figure)/.test(nt)) { capEl = nxt; name = captionName(nt); }
        }
        figNo++;
        if (!capEl) {
          capEl = createW(doc, 'p');
          body.insertBefore(capEl, nxt);
          placeholders++;
        }
        finishCaption(capEl, curChap, figNo, name, '图', curApp);
        continue;
      }

      if (el.localName === 'tbl') {
        if (isFrontElement(el, frontBoundary)) continue;
        if (!curChap) continue;

        /* 表题在表格正上方（附件8） */
        var prv = skipBlanks(el.previousSibling, -1);
        var tCapEl = null, tName = '';
        if (prv && prv.localName === 'p') {
          var pt = paraText(prv).trim();
          if (RE.cap.test(pt) && /^(表|Table)/.test(pt)) { tCapEl = prv; tName = captionName(pt); }
        }
        tabNo++;
        if (!tCapEl) {
          tCapEl = createW(doc, 'p');
          body.insertBefore(tCapEl, el);
          placeholders++;
        }
        finishCaption(tCapEl, curChap, tabNo, tName, '表', curApp);
      }
    }
    return { done: done, placeholders: placeholders };
  }

  /* ---------- 正文交叉引用同步 ----------
     题注改成「图4.1」「表4.1」后，正文里的「如图4-1所示」「表4-2~表4-4」
     必须一起改，否则题注和引用对不上，比不改还糟。
     「图4-1」→「图4.1」是等长替换，所以可以直接按字符位改，
     不必重建文本节点（匹配可能跨 run，按整段拼串定位）。 */
  function normalizeFigureRefs(paras, roles) {
    var fixed = 0;
    for (var i = 0; i < paras.length; i++) {
      var role = roles[i];
      if (role === 'front' || role === 'tocItem' || role === 'tocHead' ||
          role === 'caption' || role === 'figure' || role === 'formula') continue;
      var nodes = paraTextNodes(paras[i].p);
      if (!nodes.length) continue;
      var texts = [], starts = [], acc = 0;
      for (var n = 0; n < nodes.length; n++) {
        var s = nodes[n].textContent || '';
        texts.push(s); starts.push(acc); acc += s.length;
      }
      var joined = texts.join('');
      if (joined.indexOf('-') < 0) continue;
      var re = /[图表]\s*(\d{1,2})\s*-\s*(\d{1,2})/g, m, hits = [];
      while ((m = re.exec(joined))) {
        /* 「图表-1」这种「图表」连写不是题注引用，跳过 */
        var pre = m.index > 0 ? joined.charAt(m.index - 1) : '';
        if (pre === '图' || pre === '表') continue;
        hits.push(m.index + m[0].indexOf('-'));
      }
      if (!hits.length) continue;
      var touched = false;
      for (var h = 0; h < hits.length; h++) {
        var g = hits[h], k = -1;
        for (var n2 = texts.length - 1; n2 >= 0; n2--) { if (g >= starts[n2]) { k = n2; break; } }
        if (k < 0) continue;
        var li = g - starts[k];
        if (texts[k].charAt(li) !== '-') continue;
        texts[k] = texts[k].slice(0, li) + '.' + texts[k].slice(li + 1);
        touched = true; fixed++;
      }
      if (!touched) continue;
      for (var n3 = 0; n3 < nodes.length; n3++) {
        if ((nodes[n3].textContent || '') !== texts[n3]) nodes[n3].textContent = texts[n3];
      }
    }
    return fixed;
  }

  /* ---------- 题注位置规范化（附件8：图题置于图片正下方、表题置于表格正上方） ---------- */
  function repositionCaptions(doc, frontBoundary) {
    var body = allByNs(doc, 'body', W_NS)[0];
    if (!body) return 0;
    var kids = [];
    for (var i = 0; i < body.childNodes.length; i++) {
      var c = body.childNodes[i];
      if (c.nodeType === 1) kids.push(c);
    }
    function isEmptyP(e) {
      return e.localName === 'p' && wAll(e, 't').length === 0 &&
        wAll(e, 'drawing').length === 0 && wAll(e, 'pict').length === 0;
    }
    function isImageP(e) {
      return e.localName === 'p' && (wAll(e, 'drawing').length > 0 || wAll(e, 'pict').length > 0);
    }
    function isInert(e) { return !!INERT_ELEMS[e.localName]; }
    var moved = 0;
    for (var k = 0; k < kids.length; k++) {
      var p = kids[k];
      if (p.localName !== 'p') continue;
      var t = paraText(p).trim();
      if (!RE.cap.test(t)) continue;
      if (isFrontElement(p, frontBoundary)) continue;   // 前置部分不动
      var isFig = /^图/.test(t);
      // 跨过空段/惰性元素，找前后最近的实质元素
      var prevIdx = k - 1, nextIdx = k + 1;
      while (prevIdx >= 0 && (isEmptyP(kids[prevIdx]) || isInert(kids[prevIdx]))) prevIdx--;
      while (nextIdx < kids.length && (isEmptyP(kids[nextIdx]) || isInert(kids[nextIdx]))) nextIdx++;
      var prev = prevIdx >= 0 ? kids[prevIdx] : null;
      var next = nextIdx < kids.length ? kids[nextIdx] : null;
      if (isFig) {
        // 图题：前方是图片 → 非紧邻则移到图片正下方；前方不是图片且后方是图片
        // （题注原本在图片上方）→ 移到图片下方
        if (prev && isImageP(prev)) {
          if (prevIdx < k - 1) { body.insertBefore(p, prev.nextSibling); moved++; }
        } else if (next && isImageP(next)) {
          body.insertBefore(p, next.nextSibling);
          moved++;
        }
      } else {
        // 表题：后方是表格 → 非紧邻则移到表格正上方；后方不是表格且前方是表格
        // （题注原本在表格下方）→ 移到表格上方
        if (next && next.localName === 'tbl') {
          if (nextIdx > k + 1) { body.insertBefore(p, next); moved++; }
        } else if (prev && prev.localName === 'tbl') {
          body.insertBefore(p, prev);
          moved++;
        }
      }
    }
    return moved;
  }

  /* ---------- 章节换页规整化 ----------
     每个章节标题（h1/目录标题）之前必须：恰好 1 个空段 + 恰好 1 个换页机制，
     分节符/分页符与上一章内容连续（中间无多余空段），避免：
     - 分节段前的空段溢出到下一页造成空白页（如绪论与目录之间）
     - 标题前多个空段造成页首大片空白（如第四章上方）
     - 连续分节符不换页 + 标题段 pageBreakBefore 的组合导致空行被吞/双重分页
        （如第五章上方）
     规则：
     - 前有分节段：continuous 改 nextPage；删除分节段前的空段；标题前保留恰好 1 个空段
     - 无分节段且前导内容段有段尾分页 run / pageBreakBefore：保留，仅压缩空段
     - 无任何换页机制：分页符加在标题前的空段上（空段换页，标题紧随新页顶部）
     - 标题段自身的 pageBreakBefore 一律移除（换页由分节段/空段承担，防双重分页） */
  function isEmptySeparatorPara(p) {
    if (wAll(p, 't').length > 0 || wAll(p, 'drawing').length > 0 || wAll(p, 'pict').length > 0) return false;
    if (allByNs(p, 'oMath', M_NS).length > 0) return false;
    if (allByNs(p, 'instrText', W_NS).length > 0) return false;   // 含域指令的段落（旧 TOC 域残留）不算空段
    /* 含任何域字符（begin/separate/end）的段落也不算空段：旧目录域末尾那个
       只有 end 的段落会被当成空段「保留一个」，域字符就孤零零留在正文里 */
    if (wAll(p, 'fldChar').length > 0) return false;
    if (hasPageBreakRun(p)) return false;
    var pPr = childByNs(p, 'pPr', W_NS);
    return !(pPr && childByNs(pPr, 'pageBreakBefore', W_NS));
  }

  /* 纯空段（无文字/图片/公式/域，忽略分页符——标题后空行压缩用） */
  function isBlankPara(p) {
    return wAll(p, 't').length === 0 && wAll(p, 'drawing').length === 0 &&
      wAll(p, 'pict').length === 0 && allByNs(p, 'oMath', M_NS).length === 0 &&
      allByNs(p, 'instrText', W_NS).length === 0;
  }

  /* 移除段落上的强制分页（分页 run 与 pageBreakBefore），lastRenderedPageBreak 保留 */
  function removePageBreaksFrom(p) {
    var removed = 0;
    var brs = wAll(p, 'br');
    for (var i = 0; i < brs.length; i++) {
      if ((brs[i].getAttributeNS(W_NS, 'type') || '') === 'page') {
        brs[i].parentNode.removeChild(brs[i]);
        removed++;
      }
    }
    var pPr = childByNs(p, 'pPr', W_NS);
    if (pPr) {
      var pbb = childByNs(pPr, 'pageBreakBefore', W_NS);
      if (pbb) { pPr.removeChild(pbb); removed++; }
    }
    return removed;
  }

  function normalizeOneChapterBreak(doc, body, head) {
    var changed = 0;
    var headPPr = childByNs(head, 'pPr', W_NS);
    var headPBB = headPPr && childByNs(headPPr, 'pageBreakBefore', W_NS);

    /* 扫描标题前的空段（从近到远，含带分页符的空段）与最近的实质元素 */
    var empties = [], breakEmpties = [], lead = null;
    var prev = head.previousSibling;
    while (prev && prev.nodeType !== 1) prev = prev.previousSibling;
    while (prev) {
      if (prev.localName === 'p') {
        var ppPr = childByNs(prev, 'pPr', W_NS);
        var sect = ppPr && childByNs(ppPr, 'sectPr', W_NS);
        var hasContent = wAll(prev, 't').length > 0 || wAll(prev, 'drawing').length > 0 ||
                         wAll(prev, 'pict').length > 0 || allByNs(prev, 'oMath', M_NS).length > 0;
        var brRun = hasPageBreakRun(prev);
        var pbb = ppPr && childByNs(ppPr, 'pageBreakBefore', W_NS);
        if (!hasContent && !sect) {
          empties.push(prev);
          if (brRun || pbb) breakEmpties.push(prev);   // 空段上的分页符：换页机制候选，继续向前找分节段
          prev = prev.previousSibling;
          while (prev && prev.nodeType !== 1) prev = prev.previousSibling;
          continue;
        }
        lead = { el: prev, sect: sect, pPr: ppPr };
        break;
      }
      if (prev.localName === 'sectPr') { lead = { el: prev, sect: prev, pPr: null }; break; }
      if (INERT_ELEMS[prev.localName]) {
        prev = prev.previousSibling;
        while (prev && prev.nodeType !== 1) prev = prev.previousSibling;
        continue;
      }
      lead = { el: prev, sect: null, pPr: null };
      break;
    }

    /* 换页机制判定 */
    var mechEmpty = null;        // 承担换页的空段
    var addBreakToEmpty = false; // 需要给保留的空段加 pageBreakBefore
    if (lead && lead.sect) {
      /* 分节段换页：空段上的分页符全部移除（分节段+分页 run 双重换页会产生空白页） */
      for (var b1 = 0; b1 < breakEmpties.length; b1++) changed += removePageBreaksFrom(breakEmpties[b1]);
      var tp = childByNs(lead.sect, 'type', W_NS);
      if (tp && tp.getAttributeNS(W_NS, 'val') === 'continuous') {   // 连续分节不换页 → 改下一页
        tp.setAttributeNS(W_NS, 'w:val', 'nextPage');
        changed++;
      }
      /* 分节符与上一章内容连续：删除分节段前的空段（避免空段溢出产生空白页） */
      var s = lead.el.previousSibling;
      while (s && s.nodeType === 1 && s.localName === 'p' && isEmptySeparatorPara(s)) {
        var ns = s.previousSibling;
        body.removeChild(s);
        changed++;
        s = ns;
      }
      /* 空的分节段行高缩到 1pt：分节符段落即使溢出到章节页顶部也不显示为空行，
         章节页顶部只保留标题前那 1 个空行（避免页眉下方出现两个回车） */
      if (isEmptySeparatorPara(lead.el)) {
        var sPPr = childByNs(lead.el, 'pPr', W_NS);
        if (!sPPr) { sPPr = createW(doc, 'pPr'); lead.el.insertBefore(sPPr, lead.el.firstChild); }
        setSpacing(doc, sPPr, 20, 0, 0, 'exact');
        changed++;
      }
    } else if (breakEmpties.length > 0) {
      /* 空段上的分页符承担换页：保留最近的，其余空段的分页符移除 */
      mechEmpty = breakEmpties[0];
      for (var b2 = 0; b2 < breakEmpties.length; b2++) {
        if (breakEmpties[b2] !== mechEmpty) changed += removePageBreaksFrom(breakEmpties[b2]);
      }
    } else if (!(lead && lead.el.localName === 'p' &&
                 (breakAtParaEnd(lead.el) || (lead.pPr && childByNs(lead.pPr, 'pageBreakBefore', W_NS))))) {
      addBreakToEmpty = true;    // 无任何换页机制：由保留的空段承担分页
    }

    /* 标题前空段：压缩到恰好 1 个（标题上方空一行），换页空段优先保留 */
    if (mechEmpty) {
      for (var k2 = 0; k2 < empties.length; k2++) {
        if (empties[k2] !== mechEmpty) { body.removeChild(empties[k2]); changed++; }
      }
    } else {
      while (empties.length > 1) {
        body.removeChild(empties.pop());
        changed++;
      }
      if (empties.length === 0) {
        var sep = createW(doc, 'p');
        body.insertBefore(sep, head);
        empties.push(sep);
        changed++;
      }
    }
    var keep = mechEmpty || empties[0];
    /* 标题上方空段行高 = 五号字大小（10.5pt=210），与标题字体大小分割开 */
    var kPPr = childByNs(keep, 'pPr', W_NS);
    if (!kPPr) { kPPr = createW(doc, 'pPr'); keep.insertBefore(kPPr, keep.firstChild); }
    var kSp = childByNs(kPPr, 'spacing', W_NS);
    if (!kSp) { kSp = createW(doc, 'spacing'); insertInOrder(kPPr, kSp, PPR_ORDER); }
    kSp.setAttributeNS(W_NS, 'w:line', '210');
    kSp.setAttributeNS(W_NS, 'w:lineRule', 'exact');

    /* 标题下方与上方对称：确保恰好 1 个五号字空段（行高 210=10.5pt）。
       带分页符/分页 run 的空段（Word 重存产物）也纳入压缩，并在保留的空段上移除分页 */
    var nx = head.nextSibling;
    while (nx && nx.nodeType !== 1) nx = nx.nextSibling;
    var afterEmpties = [];
    while (nx && nx.localName === 'p' && isBlankPara(nx) &&
           !(function (pp) { return pp ? childByNs(pp, 'sectPr', W_NS) : null; })(childByNs(nx, 'pPr', W_NS))) {
      afterEmpties.push(nx);
      var nx2 = nx.nextSibling;
      while (nx2 && nx2.nodeType !== 1) nx2 = nx2.nextSibling;
      nx = nx2;
    }
    while (afterEmpties.length > 1) {
      body.removeChild(afterEmpties.pop());
      changed++;
    }
    if (afterEmpties.length === 0) {
      var sepA = createW(doc, 'p');
      body.insertBefore(sepA, head.nextSibling);
      afterEmpties.push(sepA);
      changed++;
    }
    changed += removePageBreaksFrom(afterEmpties[0]);   // 移除保留空段上的分页符
    var aPPr = childByNs(afterEmpties[0], 'pPr', W_NS);
    if (!aPPr) { aPPr = createW(doc, 'pPr'); afterEmpties[0].insertBefore(aPPr, afterEmpties[0].firstChild); }
    var aSp = childByNs(aPPr, 'spacing', W_NS);
    if (!aSp) { aSp = createW(doc, 'spacing'); insertInOrder(aPPr, aSp, PPR_ORDER); }
    aSp.setAttributeNS(W_NS, 'w:line', '210');
    aSp.setAttributeNS(W_NS, 'w:lineRule', 'exact');
    changed++;
    if (addBreakToEmpty) {
      var ePPr = childByNs(keep, 'pPr', W_NS);
      if (!childByNs(ePPr, 'pageBreakBefore', W_NS)) {
        var pb = createW(doc, 'pageBreakBefore');
        insertInOrder(ePPr, pb, PPR_ORDER);
        changed++;
      }
    }
    /* 标题段自身分页符移除（分节段/空段已承担换页，双重分页会产生空白页） */
    if (headPBB) {
      headPPr.removeChild(headPBB);
      changed++;
    }
    return changed;
  }

  /* ---------- 每章末尾补「分节符（下一页）」 ----------
     目的：让每一章独立成一节，在 Word 里可以逐章单独设置页眉（页眉区会
     出现「第 N 节」）。Word 的「分节符（下一页）」本身就带分页效果，
     所以不再叠加分页符——两者叠加会在章节之间多出一张空白页。
     已有分节符的边界跳过；缺的用一个空段承载分节符，再交给
     normalizeChapterBreaks 统一压缩上下空行、去掉重复的分页机制。 */
  function ensureChapterSectionBreaks(doc, body, paras, roles, frontBoundary) {
    var added = 0, kept = 0, dropped = 0;
    for (var i = 0; i < paras.length; i++) {
      var role = roles[i] === 'refHead' ? 'h1' : roles[i];
      if (role !== 'h1') continue;
      var head = paras[i].p;
      if (isFrontElement(head, frontBoundary)) continue;

      var prev = head.previousSibling;
      while (prev && prev.nodeType !== 1) prev = prev.previousSibling;
      if (prev && prev.localName === 'sectPr') { kept++; continue; }   // body 级分节符

      /* 往回跨过空段找已有分节符。原文档常见写法：上一节末尾那个空段自带分节符，
         后面再跟一个空段才是章标题——此时标题其实已经处在新的一节里了，
         再补一条相邻的分节符会凭空多出一张空白页。遇到正文段落即停，
         保证只在本节边界范围内查找，不会把上一章的分节符误判成本章的。 */
      var foundSect = null;
      for (var back = prev; back; back = back.previousSibling) {
        if (back.nodeType !== 1) continue;
        if (back.localName === 'sectPr') { foundSect = back; break; }
        if (back.localName !== 'p') break;
        var bPr = childByNs(back, 'pPr', W_NS);
        if (bPr && childByNs(bPr, 'sectPr', W_NS)) { foundSect = childByNs(bPr, 'sectPr', W_NS); break; }
        if (!isBlankPara(back)) break;
      }
      if (foundSect) {
        /* 只有「分节符（下一页）」才同时满足分节和换页，直接跳过。
           oddPage/evenPage 也换页，保留作者的选择。 */
        var ft = childByNs(foundSect, 'type', W_NS);
        var ftv = ft ? (ft.getAttributeNS(W_NS, 'val') || '') : '';
        if (!ftv || ftv === 'nextPage' || ftv === 'oddPage' || ftv === 'evenPage') { kept++; continue; }
        /* 连续分节符不换页：就地升级成下一页。另补一条会在章前多出一个空节。 */
        if (ft) ft.setAttributeNS(W_NS, 'w:val', 'nextPage');
        else {
          ft = createW(doc, 'type');
          ft.setAttributeNS(W_NS, 'w:val', 'nextPage');
          insertInOrder(foundSect, ft, SECTPR_ORDER);
        }
        added++;
        continue;
      }

      /* 空段作分节符宿主。这里用 isBlankPara 而不是 isEmptySeparatorPara：
         后者把「带分页符的空段」排除在外，而原文档的章节正是靠这种空段换页的
         ——另起一个新空段会把它留在上一节末尾，分页符+分节符双重换页出空白页 */
      var host = (prev && prev.localName === 'p' && isBlankPara(prev)) ? prev : null;
      if (!host) { host = createW(doc, 'p'); body.insertBefore(host, head); }

      /* 分节符段落属于它前面那一节，所以往后找最近的 sectPr 作为版面蓝本 */
      var src = null;
      for (var sib = host.nextSibling; sib; sib = sib.nextSibling) {
        if (sib.nodeType !== 1) continue;
        if (sib.localName === 'sectPr') { src = sib; break; }
        if (sib.localName === 'p') {
          var spPr = childByNs(sib, 'pPr', W_NS);
          var sp = spPr && childByNs(spPr, 'sectPr', W_NS);
          if (sp) { src = sp; break; }
        }
      }
      var sect = src ? src.cloneNode(true) : createW(doc, 'sectPr');
      /* 页码必须接着上一节编下去：去掉「起始页码」，格式（罗马/阿拉伯）保留 */
      var pn = childByNs(sect, 'pgNumType', W_NS);
      if (pn) pn.removeAttributeNS(W_NS, 'w:start');

      var hpPr = childByNs(host, 'pPr', W_NS);
      if (!hpPr) { hpPr = createW(doc, 'pPr'); host.insertBefore(hpPr, host.firstChild); }
      var old = childByNs(hpPr, 'sectPr', W_NS);
      if (old) hpPr.removeChild(old);
      insertInOrder(hpPr, sect, PPR_ORDER);

      /* 分节符接手换页：去掉这个边界上原有的分页机制（宿主空段自己的、
         以及上一段末尾的），否则两套换页叠加会在章间多出一张空白页 */
      dropped += removePageBreaksFrom(host);
      var tail = host.previousSibling;
      while (tail && tail.nodeType !== 1) tail = tail.previousSibling;
      if (tail && tail.localName === 'p' && breakAtParaEnd(tail)) dropped += removePageBreaksFrom(tail);

      added++;
    }
    return { added: added, kept: kept, dropped: dropped };
  }

  function normalizeChapterBreaks(doc, paras, roles, frontBoundary) {
    var body = allByNs(doc, 'body', W_NS)[0];
    if (!body) return 0;
    /* 一级标题/目录标题/摘要/ABSTRACT 标题统一处理上下空行（五号字大小） */
    var heads = [];
    for (var i = 0; i < paras.length; i++) {
      if (roles[i] === 'h1' || roles[i] === 'tocHead' || roles[i] === 'absHead' || roles[i] === 'absEnHead') {
        heads.push(paras[i].p);
      }
    }
    var changed = 0;
    for (var h = 0; h < heads.length; h++) {
      var head = heads[h];
      if (isFrontElement(head, frontBoundary)) continue;
      changed += normalizeOneChapterBreak(doc, body, head);
    }
    return changed;
  }

  /* ---------- 自动目录域（参考样张：sdt 包裹的 TOC 域，可整体选择、可更新） ---------- */
  /* 构建 Word 自动目录域段落（sdt + TOC field，w:dirty 让 Word 打开时自动更新） */
  /* 构建 TOC 域指令 run（begin + instrText + separate） */
  function tocFieldBeginRun(doc) {
    var r = createW(doc, 'r');
    var f1 = createW(doc, 'fldChar');
    f1.setAttributeNS(W_NS, 'w:fldCharType', 'begin');
    f1.setAttributeNS(W_NS, 'w:dirty', 'true');   // 打开文档时自动更新目录
    r.appendChild(f1);
    var instr = createW(doc, 'instrText');
    instr.setAttributeNS('http://www.w3.org/XML/1998/namespace', 'xml:space', 'preserve');
    instr.textContent = ' TOC \\o "1-3" \\h \\u ';
    r.appendChild(instr);
    var f2 = createW(doc, 'fldChar');
    f2.setAttributeNS(W_NS, 'w:fldCharType', 'separate');
    r.appendChild(f2);
    return r;
  }

  /* 构建目录条目段落（toc 样式 + 标题文本 + tab + PAGEREF 页码域）。
     tocIds 为 styles.xml 中实际存在的 toc 样式 id（文档里的 id 常是 6/7/8 之类，
     硬编码 TOC1 会指向不存在的样式，Word 回落到正文格式） */
  function buildTocEntryPara(doc, level, text, bookmark, tocIds) {
    var p = createW(doc, 'p');
    var pPr = createW(doc, 'pPr');
    var ps = createW(doc, 'pStyle');
    ps.setAttributeNS(W_NS, 'w:val', (tocIds && tocIds[level + 1]) || ('TOC' + (level + 1)));
    pPr.appendChild(ps);
    var sp = createW(doc, 'spacing');
    sp.setAttributeNS(W_NS, 'w:line', '400');
    sp.setAttributeNS(W_NS, 'w:lineRule', 'exact');
    pPr.appendChild(sp);
    p.appendChild(pPr);

    var rt = createW(doc, 'r');
    var t = createW(doc, 't');
    t.textContent = text;
    rt.appendChild(t);
    p.appendChild(rt);

    var rtab = createW(doc, 'r');
    rtab.appendChild(createW(doc, 'tab'));
    p.appendChild(rtab);

    var rp = createW(doc, 'r');
    var f1 = createW(doc, 'fldChar');
    f1.setAttributeNS(W_NS, 'w:fldCharType', 'begin');
    rp.appendChild(f1);
    var instr = createW(doc, 'instrText');
    instr.setAttributeNS('http://www.w3.org/XML/1998/namespace', 'xml:space', 'preserve');
    instr.textContent = ' PAGEREF ' + bookmark + ' \\h ';
    rp.appendChild(instr);
    var f2 = createW(doc, 'fldChar');
    f2.setAttributeNS(W_NS, 'w:fldCharType', 'separate');
    rp.appendChild(f2);
    var f3 = createW(doc, 'fldChar');
    f3.setAttributeNS(W_NS, 'w:fldCharType', 'end');
    rp.appendChild(f3);
    p.appendChild(rp);
    return p;
  }

  /* 给标题段落加书签（供目录条目 PAGEREF 引用页码） */
  function addBookmarkTo(doc, p, id, name) {
    var bs = createW(doc, 'bookmarkStart');
    bs.setAttributeNS(W_NS, 'w:id', String(id));
    bs.setAttributeNS(W_NS, 'w:name', name);
    var firstChild = null;
    for (var i = 0; i < p.childNodes.length; i++) {
      var c = p.childNodes[i];
      if (c.nodeType === 1 && c.localName !== 'pPr') { firstChild = c; break; }
    }
    p.insertBefore(bs, firstChild);
    var be = createW(doc, 'bookmarkEnd');
    be.setAttributeNS(W_NS, 'w:id', String(id));
    p.appendChild(be);
  }

  /* 构建自动目录域 sdt：TOC 指令 + 预填充条目（标题文本 + 页码域），
     打开文档自动更新、可整体选择，格式用修正后的 toc 1/2/3 样式 */
  function buildTocFieldSdt(doc, headings, tocIds) {
    var sdt = createW(doc, 'sdt');
    var sdtPr = createW(doc, 'sdtPr');
    var dpo = createW(doc, 'docPartObj');
    var dpg = createW(doc, 'docPartGallery');
    dpg.setAttributeNS(W_NS, 'w:val', 'Table of Contents');
    dpo.appendChild(dpg);
    dpo.appendChild(createW(doc, 'docPartUnique'));
    sdtPr.appendChild(dpo);
    sdt.appendChild(sdtPr);

    var content = createW(doc, 'sdtContent');
    /* 首段：TOC 域开始 */
    var p0 = createW(doc, 'p');
    var pPr0 = createW(doc, 'pPr');
    var jc = createW(doc, 'jc');
    jc.setAttributeNS(W_NS, 'w:val', 'left');
    pPr0.appendChild(jc);
    p0.appendChild(pPr0);
    p0.appendChild(tocFieldBeginRun(doc));
    content.appendChild(p0);

    /* 条目段：标题文本 + tab + PAGEREF 页码域 */
    for (var i = 0; i < headings.length; i++) {
      content.appendChild(buildTocEntryPara(doc, headings[i].level, headings[i].text, headings[i].bookmark, tocIds));
    }

    /* 末段：TOC 域结束 */
    var pEnd = createW(doc, 'p');
    var rEnd = createW(doc, 'r');
    var fEnd = createW(doc, 'fldChar');
    fEnd.setAttributeNS(W_NS, 'w:fldCharType', 'end');
    rEnd.appendChild(fEnd);
    pEnd.appendChild(rEnd);
    content.appendChild(pEnd);

    sdt.appendChild(content);
    return sdt;
  }

  /* 用自动目录域替换手动目录条目：识别完成后直接生成完整目录
     （预填充条目 + 页码域），并应用之前定好的 toc 格式 */
  function replaceTocWithField(doc, paras, roles, tocIds) {
    var body = allByNs(doc, 'body', W_NS)[0];
    if (!body) return false;
    var tocHeadEl = null, tocEndEl = null;
    for (var i = 0; i < paras.length; i++) {
      if (roles[i] === 'tocHead') tocHeadEl = paras[i].p;
      /* 第一个一级标题定位时排除目录条目：Word 更新过的旧目录条目（HYPERLINK 域、
         无 tab 页码）可能被误识别为 h1，导致目录区边界错误、新旧目录叠加 */
      else if (roles[i] === 'h1' && !isTocEntry(paras[i], paras[i].text) && tocHeadEl && !tocEndEl) {
        tocEndEl = paras[i].p;
      }
    }
    if (!tocHeadEl || !tocEndEl) return false;

    /* 收集标题（一级/二级/三级）并加书签；
       只收集第一个一级标题（tocEndEl）之后的真实标题，跳过目录区内的旧目录条目
       （Word 更新过的目录条目可能被误识别为 h1/h2/h3，避免目录叠加） */
    var headings = [];
    var maxId = 0;
    var bsAll = allByNs(doc, 'bookmarkStart', W_NS);
    for (var b = 0; b < bsAll.length; b++) {
      var bid = parseInt(bsAll[b].getAttributeNS(W_NS, 'id') || '0', 10);
      if (!isNaN(bid) && bid > maxId) maxId = bid;
    }
    var collecting = false;
    for (var i2 = 0; i2 < paras.length; i2++) {
      var pEl = paras[i2].p;
      if (pEl === tocEndEl) collecting = true;   // 包含 tocEndEl 本身（绪论）
      if (!collecting) continue;
      /* 参考文献标题的角色是 refHead，但它在版面上与致谢/附录同属一级标题，
         目录里必须有它（附件8 样张：… 参考文献……50 / 致谢……51 / 附录A …52） */
      var role = roles[i2] === 'refHead' ? 'h1' : roles[i2];
      var lvl = role === 'h1' ? 0 : role === 'h2' ? 1 : role === 'h3' ? 2 : -1;
      if (lvl < 0) continue;
      var bm = '_Toc' + (headings.length + 1);
      addBookmarkTo(doc, pEl, ++maxId, bm);
      headings.push({ level: lvl, text: paraText(pEl).trim(), bookmark: bm });
    }
    if (headings.length === 0) return false;

    /* 删除目录标题后到分节段之间的元素（保留标题后的 1 个空段；
       分节段是目录节末尾与章节换页机制，其后内容（绪论前空段+绪论）不删除） */
    var keptOne = false, fldDepth = 0;
    var cur = tocHeadEl.nextSibling;
    while (cur && cur !== tocEndEl) {
      var nx = cur.nextSibling;
      var isSectP = cur.nodeType === 1 && cur.localName === 'p' &&
                    childByNs(childByNs(cur, 'pPr', W_NS), 'sectPr', W_NS);
      /* 分节段是目录节的换页机制，要留下；但若此刻正处在旧目录域内部（已经删掉了
         域的 begin 还没删到 end），说明旧目录被分节段劈成两半，必须继续删下去，
         否则后半截旧目录留在正文里、域的 end 变成孤儿（Word 会因此吞掉正文） */
      if (isSectP && fldDepth <= 0) break;
      var isSep = cur.nodeType === 1 && cur.localName === 'p' && isEmptySeparatorPara(cur);
      if (isSectP) { /* 保留分节段，只把其中的域字符计入深度 */ }
      else if (isSep && !keptOne) { keptOne = true; }
      else body.removeChild(cur);
      var fcs = wAll(cur, 'fldChar');
      for (var f = 0; f < fcs.length; f++) {
        var ft = fcs[f].getAttributeNS(W_NS, 'fldCharType') || '';
        if (ft === 'begin') fldDepth++;
        else if (ft === 'end') fldDepth--;
      }
      cur = nx;
    }
    /* 插入 TOC 域：位于标题后的空段之后、分节段之前 */
    var sdt = buildTocFieldSdt(doc, headings, tocIds);
    var anchor = tocHeadEl.nextSibling;
    var anchorIsSect = anchor && anchor.nodeType === 1 && anchor.localName === 'p' &&
                       childByNs(childByNs(anchor, 'pPr', W_NS), 'sectPr', W_NS);
    body.insertBefore(sdt, anchorIsSect ? anchor : (anchor ? anchor.nextSibling : tocEndEl));

    /* 兜底清理：目录区里还残留的旧目录条目（带 PAGEREF _TocN 域的段落）一并删掉。
       反复格式化、或输入文档的目录域本来就是坏的（例如半截目录留在正文里）时，
       这些段落会被当成普通文字排在正文前面，既难看又影响页码 */
    var c2 = tocHeadEl.nextSibling;
    while (c2 && c2 !== tocEndEl) {
      var n2 = c2.nextSibling;
      if (c2 !== sdt && c2.nodeType === 1) {
        var ins2 = allByNs(c2, 'instrText', W_NS);
        for (var q = 0; q < ins2.length; q++) {
          if (/^\s*PAGEREF\s+"?_Toc\d+/i.test(ins2[q].textContent || '')) { body.removeChild(c2); break; }
        }
      }
      c2 = n2;
    }
    return true;
  }

  /* 修正 / 补建 styles.xml 中 toc 1/2/3 样式：小四（12pt）、层级缩进
     （一级0/二级2字/三级4字），使 Word 自动更新目录后条目格式符合附件8。
     返回 {1:id,2:id,3:id}：目录条目 pStyle 必须引用文档里真实存在的样式 id，
     否则 Word 按正文格式渲染目录（文档中的 toc 样式 id 常为 6/7/8 之类）。 */
  async function ensureTocStyles(zip) {
    var entry = zip.file('word/styles.xml');
    if (!entry) return null;
    var stDoc = new DOMParser().parseFromString(await entry.async('string'), 'application/xml');
    var styles = stDoc.getElementsByTagNameNS(W_NS, 'style');
    var usedIds = {};
    for (var i = 0; i < styles.length; i++) {
      var sid = styles[i].getAttributeNS(W_NS, 'styleId') || '';
      if (sid) usedIds[sid] = true;
    }
    var byLevel = {};
    for (var j = 0; j < styles.length; j++) {
      var nm = childByNs(styles[j], 'name', W_NS);
      if (!nm) continue;
      var m = /^(?:toc|目录)\s*([123])$/i.exec((nm.getAttributeNS(W_NS, 'val') || '').trim());
      if (m && !byLevel[m[1]]) byLevel[m[1]] = styles[j];
    }
    /* 缺失的层级补建样式（引用不存在的样式 Word 会回落到正文格式） */
    for (var lv = 1; lv <= 3; lv++) {
      if (byLevel[lv]) continue;
      var newId = 'TOC' + lv, k = 1;
      while (usedIds[newId]) { k++; newId = 'TOC' + lv + k; }
      var newSt = createW(stDoc, 'style');
      newSt.setAttributeNS(W_NS, 'w:type', 'paragraph');
      newSt.setAttributeNS(W_NS, 'w:styleId', newId);
      var nEl = createW(stDoc, 'name');
      nEl.setAttributeNS(W_NS, 'w:val', 'toc ' + lv);
      newSt.appendChild(nEl);
      var bo = createW(stDoc, 'basedOn');
      bo.setAttributeNS(W_NS, 'w:val', 'Normal');
      newSt.appendChild(bo);
      var nx = createW(stDoc, 'next');
      nx.setAttributeNS(W_NS, 'w:val', 'Normal');
      newSt.appendChild(nx);
      stDoc.documentElement.appendChild(newSt);
      usedIds[newId] = true;
      byLevel[lv] = newSt;
    }

    var targets = { 1: 0, 2: 480, 3: 960 };
    var ids = {};
    for (var L = 1; L <= 3; L++) {
      var st = byLevel[L];
      if (!st) continue;
      ids[L] = st.getAttributeNS(W_NS, 'styleId');
      /* rPr：中文宋体、西文 Times New Roman、小四（sz=24 半磅） */
      var rPr = childByNs(st, 'rPr', W_NS);
      if (!rPr) { rPr = createW(stDoc, 'rPr'); st.appendChild(rPr); }
      var rf0 = childByNs(rPr, 'rFonts', W_NS);
      if (!rf0) { rf0 = createW(stDoc, 'rFonts'); rPr.insertBefore(rf0, rPr.firstChild); }
      rf0.setAttributeNS(W_NS, 'w:ascii', 'Times New Roman');
      rf0.setAttributeNS(W_NS, 'w:hAnsi', 'Times New Roman');
      rf0.setAttributeNS(W_NS, 'w:eastAsia', '宋体');
      var szCs0 = childByNs(rPr, 'szCs', W_NS);
      var sz0 = childByNs(rPr, 'sz', W_NS);
      if (!sz0) { sz0 = createW(stDoc, 'sz'); rPr.insertBefore(sz0, szCs0 || null); }
      sz0.setAttributeNS(W_NS, 'w:val', '24');
      if (!szCs0) { szCs0 = createW(stDoc, 'szCs'); rPr.appendChild(szCs0); }
      szCs0.setAttributeNS(W_NS, 'w:val', '24');
      /* pPr：层级缩进（toc1 无缩进 / toc2 2字 / toc3 4字，2字=480 twips） */
      var left = targets[L];
      var pPr = childByNs(st, 'pPr', W_NS);
      if (!pPr) { pPr = createW(stDoc, 'pPr'); st.insertBefore(pPr, rPr); }
      var ind = childByNs(pPr, 'ind', W_NS);
      if (left) {
        if (!ind) { ind = createW(stDoc, 'ind'); insertInOrder(pPr, ind, PPR_ORDER); }
        ind.setAttributeNS(W_NS, 'w:left', String(left));
      } else if (ind) {
        ind.removeAttributeNS(W_NS, 'w:left');
      }
    }
    zip.file('word/styles.xml', serialize(stDoc, 'word/styles.xml'));
    return ids;
  }

  /* ---------- 分节：页边距 / 纸张 / 页码 ---------- */
  function setPgSz(doc, sectPr, s) {
    var old = childByNs(sectPr, 'pgSz', W_NS);
    var landscape = old && Number(old.getAttributeNS(W_NS, 'w')) > Number(old.getAttributeNS(W_NS, 'h'));
    if (old) sectPr.removeChild(old);
    var pg = createW(doc, 'pgSz');
    if (landscape) { pg.setAttributeNS(W_NS, 'w:w', '16838'); pg.setAttributeNS(W_NS, 'w:h', '11906'); pg.setAttributeNS(W_NS, 'w:orient', 'landscape'); }
    else { pg.setAttributeNS(W_NS, 'w:w', '11906'); pg.setAttributeNS(W_NS, 'w:h', '16838'); }
    insertInOrder(sectPr, pg, SECTPR_ORDER);
  }

  function setPgMar(doc, sectPr, s) {
    var old = childByNs(sectPr, 'pgMar', W_NS);
    if (old) sectPr.removeChild(old);
    var mar = createW(doc, 'pgMar');
    mar.setAttributeNS(W_NS, 'w:top', String(Math.round(s.marginTop * 567)));
    mar.setAttributeNS(W_NS, 'w:bottom', String(Math.round(s.marginBottom * 567)));
    mar.setAttributeNS(W_NS, 'w:left', String(Math.round(s.marginLeft * 567)));
    mar.setAttributeNS(W_NS, 'w:right', String(Math.round(s.marginRight * 567)));
    mar.setAttributeNS(W_NS, 'w:header', '907');   // 页眉距 1.6cm
    mar.setAttributeNS(W_NS, 'w:footer', '1191');  // 页脚距 2.1cm
    mar.setAttributeNS(W_NS, 'w:gutter', '0');
    insertInOrder(sectPr, mar, SECTPR_ORDER);
  }

  function setDocGrid(doc, sectPr) {
    var old = childByNs(sectPr, 'docGrid', W_NS);
    if (old) sectPr.removeChild(old);
    var grid = createW(doc, 'docGrid');
    grid.setAttributeNS(W_NS, 'w:linePitch', '0'); // 关闭行网格约束，确保行距按设定生效
    insertInOrder(sectPr, grid, SECTPR_ORDER);
  }

  function addFooterRef(doc, sectPr, rid) {
    var olds = wAll(sectPr, 'footerReference');
    for (var i = 0; i < olds.length; i++) sectPr.removeChild(olds[i]);
    var fr = createW(doc, 'footerReference');
    fr.setAttributeNS(R_NS, 'r:id', rid);
    insertInOrder(sectPr, fr, SECTPR_ORDER);
  }

  function setSections(doc, s, footerRid, frontSectPrs) {
    var sectPrs = allByNs(doc, 'sectPr', W_NS);
    for (var i = 0; i < sectPrs.length; i++) {
      var sp = sectPrs[i];
      if (frontSectPrs.indexOf(sp) >= 0) continue;   // 前两页分节不修改
      setPgSz(doc, sp, s);
      setPgMar(doc, sp, s);
      setDocGrid(doc, sp);
      if (footerRid) addFooterRef(doc, sp, footerRid);
    }
    // 第一个未被跳过的分节加"首页不同"：封面不显示页码
    if (footerRid && sectPrs.length > 0 && frontSectPrs.indexOf(sectPrs[0]) < 0) {
      var first = sectPrs[0];
      if (!childByNs(first, 'titlePg', W_NS)) {
        var tp = createW(doc, 'titlePg');
        insertInOrder(first, tp, SECTPR_ORDER);
      }
    }
  }

  /* 分节页码：前置部分（摘要/目录）罗马数字，主体（绪论起）阿拉伯数字从 1 重新编页 */
  function setPgNumType(doc, sectPr, fmt, start) {
    var old = childByNs(sectPr, 'pgNumType', W_NS);
    if (old) sectPr.removeChild(old);
    var pn = createW(doc, 'pgNumType');
    pn.setAttributeNS(W_NS, 'w:fmt', fmt);
    if (start) pn.setAttributeNS(W_NS, 'w:start', String(start));
    insertInOrder(sectPr, pn, SECTPR_ORDER);
  }

  function applyPageNumbering(doc, paras, roles, settings, frontSectPrs) {
    if (!settings.pageNumber || settings.pageNumber === 'none') return;
    var body = allByNs(doc, 'body', W_NS)[0];
    if (!body) return;
    var firstH1 = null;
    for (var i = 0; i < roles.length; i++) {
      if (roles[i] === 'h1') { firstH1 = paras[i].p; break; }
    }
    if (!firstH1) {  // 无章节标题：全部统一阿拉伯数字
      var spsAll = allByNs(doc, 'sectPr', W_NS);
      for (var a = 0; a < spsAll.length; a++) {
        if (frontSectPrs.indexOf(spsAll[a]) >= 0) continue;
        setPgNumType(doc, spsAll[a], 'decimal', null);
      }
      return;
    }
    var before = true; // 当前节是否位于第一个一级标题之前（前置部分）
    var romanStarted = false; // 罗马数字是否已从摘要所在分节重新编页（封面不编页码）
    var bodyStarted = false; // 正文页码是否已从 1 开始（章节之间连续编页，不重复排序）
    for (var j = 0; j < body.childNodes.length; j++) {
      var c = body.childNodes[j];
      if (c.nodeType !== 1) continue;
      var sectPrInP = null;
      if (c.localName === 'p') {
        var pPr = childByNs(c, 'pPr', W_NS);
        if (pPr) sectPrInP = childByNs(pPr, 'sectPr', W_NS);
        if (c === firstH1) before = false;
      } else if (c.localName === 'sectPr') {
        if (frontSectPrs.indexOf(c) >= 0) continue;   // 前两页分节不修改
        var startV = null;
        if (before) { if (!romanStarted) { startV = 1; romanStarted = true; } }   // 摘要所在分节起罗马从 1 编页
        else { if (!bodyStarted) { startV = 1; bodyStarted = true; } }            // 绪论所在分节起阿拉伯从 1 编页
        setPgNumType(doc, c, before ? 'upperRoman' : 'decimal', startV);
        continue;
      }
      if (sectPrInP) {
        if (frontSectPrs.indexOf(sectPrInP) >= 0) continue;
        var startV2 = null;
        if (before) { if (!romanStarted) { startV2 = 1; romanStarted = true; } }
        else { if (!bodyStarted) { startV2 = 1; bodyStarted = true; } }
        setPgNumType(doc, sectPrInP, before ? 'upperRoman' : 'decimal', startV2);
      }
    }
  }

  /* ---------- 页码页脚 ---------- */
  function footerXml(jc) {
    var f = '<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/>' +
      '<w:sz w:val="18"/><w:szCs w:val="18"/>';
    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
      '<w:ftr xmlns:w="' + W_NS + '"><w:p><w:pPr><w:jc w:val="' + jc + '"/><w:rPr>' + f + '</w:rPr></w:pPr>' +
      '<w:r><w:rPr>' + f + '</w:rPr><w:fldChar w:fldCharType="begin"/></w:r>' +
      '<w:r><w:rPr>' + f + '</w:rPr><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r>' +
      '<w:r><w:rPr>' + f + '</w:rPr><w:fldChar w:fldCharType="end"/></w:r></w:p></w:ftr>';
  }

  function nextRid(relsDoc) {
    var max = 0;
    var rels = relsDoc.getElementsByTagName('Relationship');
    for (var i = 0; i < rels.length; i++) {
      var id = rels[i].getAttribute('Id') || '';
      var m = /rId(\d+)/.exec(id);
      if (m) max = Math.max(max, parseInt(m[1], 10));
    }
    return 'rId' + (max + 1);
  }

  async function addFooter(zip, doc, s) {
    var ctDoc = new DOMParser().parseFromString(await zip.file('[Content_Types].xml').async('string'), 'application/xml');
    var overrides = ctDoc.getElementsByTagName('Override');
    var existing = null;
    for (var i = 0; i < overrides.length; i++) {
      if (overrides[i].getAttribute('ContentType') === CT_FOOTER) { existing = overrides[i]; break; }
    }

    var target = existing ? existing.getAttribute('PartName').replace(/^\//, '') : null;
    if (!target) {
      var n = 1;
      while (zip.file('word/footer' + n + '.xml')) n++;
      target = 'word/footer' + n + '.xml';
      var ov = ctDoc.createElementNS(CT_NS, 'Override');
      ov.setAttribute('PartName', '/' + target);
      ov.setAttribute('ContentType', CT_FOOTER);
      ctDoc.documentElement.appendChild(ov);
      zip.file(target, footerXml(s.pageNumber === 'right' ? 'right' : 'center'));
    } else if (!zip.file(target)) {
      zip.file(target, footerXml(s.pageNumber === 'right' ? 'right' : 'center'));
    }
    zip.file('[Content_Types].xml', serialize(ctDoc, '[Content_Types].xml'));

    // relationships
    var relsEntry = zip.file('word/_rels/document.xml.rels');
    var relsDoc;
    if (relsEntry) {
      relsDoc = new DOMParser().parseFromString(await relsEntry.async('string'), 'application/xml');
    } else {
      relsDoc = new DOMParser().parseFromString(
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="' + PKG_REL_NS + '"></Relationships>',
        'application/xml');
    }
    var rels = relsDoc.getElementsByTagName('Relationship');
    var rid = null;
    for (var j = 0; j < rels.length; j++) {
      if (rels[j].getAttribute('Target') === target) { rid = rels[j].getAttribute('Id'); break; }
    }
    if (!rid) {
      rid = nextRid(relsDoc);
      var rel = relsDoc.createElementNS(PKG_REL_NS, 'Relationship');
      rel.setAttribute('Id', rid);
      rel.setAttribute('Type', REL_FOOTER);
      rel.setAttribute('Target', target.replace(/^word\//, '')); // 相对 word/ 解析
      relsDoc.documentElement.appendChild(rel);
    }
    zip.file('word/_rels/document.xml.rels', serialize(relsDoc, 'word/_rels/document.xml.rels'));
    return rid;
  }

  function hasFooterRef(doc) {
    var sps = allByNs(doc, 'sectPr', W_NS);
    for (var i = 0; i < sps.length; i++) {
      if (wAll(sps[i], 'footerReference').length > 0) return true;
    }
    return false;
  }

  /* ---------- 页眉：正文起各分节页眉 = 一级标题（章名） ---------- */
  /* 附件8：页眉为一级标题，五号宋体，页眉之下有一条下划线。
     用 STYLEREF 域自动取当前页所属章节的标题（改章名后页眉自动跟随，
     不必逐节写死文字）。前置部分（封面/声明/摘要/目录）的页眉保持原样。 */
  function escTxt(t) {
    return String(t == null ? '' : t).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  }

  /* STYLEREF 的样式名参数：中文版 Word 只认内置样式的显示名（标题 1），
     样式名含空格必须加引号，否则 Word 只取到空格前的部分而报「未定义样式」 */
  function stylerefName(styleName) {
    var m = /^heading\s*([1-9])$/i.exec(styleName || '');
    if (m) return '标题 ' + m[1];
    return styleName || '标题 1';
  }

  function headerXml(styleName, cachedText) {
    var f = '<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:eastAsia="宋体"/>' +
      '<w:sz w:val="21"/><w:szCs w:val="21"/>';
    var instr = ' STYLEREF "' + styleName + '" \\* MERGEFORMAT ';
    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
      '<w:hdr xmlns:w="' + W_NS + '"><w:p><w:pPr>' +
      '<w:pBdr><w:bottom w:val="single" w:sz="6" w:space="1" w:color="auto"/></w:pBdr>' +
      '<w:jc w:val="center"/><w:rPr>' + f + '</w:rPr></w:pPr>' +
      '<w:r><w:rPr>' + f + '</w:rPr><w:fldChar w:fldCharType="begin"/></w:r>' +
      '<w:r><w:rPr>' + f + '</w:rPr><w:instrText xml:space="preserve">' + instr + '</w:instrText></w:r>' +
      '<w:r><w:rPr>' + f + '</w:rPr><w:fldChar w:fldCharType="separate"/></w:r>' +
      '<w:r><w:rPr>' + f + '</w:rPr><w:t xml:space="preserve">' + escTxt(cachedText) + '</w:t></w:r>' +
      '<w:r><w:rPr>' + f + '</w:rPr><w:fldChar w:fldCharType="end"/></w:r>' +
      '</w:p></w:hdr>';
  }

  async function addHeaderPart(zip, styleName, cachedText) {
    var ctDoc = new DOMParser().parseFromString(await zip.file('[Content_Types].xml').async('string'), 'application/xml');
    var n = 1;
    while (zip.file('word/header' + n + '.xml')) n++;
    var target = 'word/header' + n + '.xml';
    var ov = ctDoc.createElementNS(CT_NS, 'Override');
    ov.setAttribute('PartName', '/' + target);
    ov.setAttribute('ContentType', CT_HEADER);
    ctDoc.documentElement.appendChild(ov);
    zip.file('[Content_Types].xml', serialize(ctDoc, '[Content_Types].xml'));
    zip.file(target, headerXml(styleName, cachedText));

    var relsEntry = zip.file('word/_rels/document.xml.rels');
    var relsDoc;
    if (relsEntry) {
      relsDoc = new DOMParser().parseFromString(await relsEntry.async('string'), 'application/xml');
    } else {
      relsDoc = new DOMParser().parseFromString(
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="' + PKG_REL_NS + '"></Relationships>',
        'application/xml');
    }
    var rid = nextRid(relsDoc);
    var rel = relsDoc.createElementNS(PKG_REL_NS, 'Relationship');
    rel.setAttribute('Id', rid);
    rel.setAttribute('Type', REL_HEADER);
    /* rsid 目标相对 word/ 解析，必须用文件名而非 word/xxx.xml（否则指向 word/word/xxx.xml） */
    rel.setAttribute('Target', target.replace(/^word\//, ''));
    relsDoc.documentElement.appendChild(rel);
    zip.file('word/_rels/document.xml.rels', serialize(relsDoc, 'word/_rels/document.xml.rels'));
    return rid;
  }

  /* 从「第一个一级标题所在分节」起（含该节及之后所有分节）页眉换成取章名的
     STYLEREF 页眉；前置部分分节（封面/声明/摘要/目录）不动 */
  async function applyBodyHeaders(zip, doc, paras, roles, frontSectPrs, styleName) {
    var body = allByNs(doc, 'body', W_NS)[0];
    if (!body) return false;
    var firstH1El = null, firstH1Text = '';
    for (var i = 0; i < roles.length; i++) {
      if (roles[i] === 'h1') { firstH1El = paras[i].p; firstH1Text = paras[i].text; break; }
    }
    if (!firstH1El) return false;

    // 第一个一级标题之后（含其所属分节）出现的所有 sectPr。
    // 注意不能只用 c === firstH1El 判断：Word 常把目录和正文包在 <w:sdt> 里，
    // 此时一级标题是 sdtContent 的子节点，body 的直接子节点里根本遇不到它，
    // seen 永远为 false，正文页眉就一个都设不上（表现为 bodyHeader:false）。
    function containsNode(anc, node) {
      for (var x = node; x; x = x.parentNode) if (x === anc) return true;
      return false;
    }
    var sectPrs = [], seen = false;
    for (var j = 0; j < body.childNodes.length; j++) {
      var c = body.childNodes[j];
      if (c.nodeType !== 1) continue;
      if (!seen && containsNode(c, firstH1El)) seen = true;
      if (!seen) continue;
      if (c.localName === 'sectPr') sectPrs.push(c);
      else if (c.localName === 'p') {
        var pPr = childByNs(c, 'pPr', W_NS);
        var sp = pPr && childByNs(pPr, 'sectPr', W_NS);
        if (sp) sectPrs.push(sp);
      }
    }
    sectPrs = sectPrs.filter(function (s) { return frontSectPrs.indexOf(s) < 0; });
    if (!sectPrs.length) return false;

    var rid;
    try { rid = await addHeaderPart(zip, styleName, firstH1Text); } catch (e) { return false; }

    for (var k = 0; k < sectPrs.length; k++) {
      var olds = wAll(sectPrs[k], 'headerReference');
      for (var o = olds.length - 1; o >= 0; o--) {
        if (olds[o].parentNode === sectPrs[k]) sectPrs[k].removeChild(olds[o]);
      }
      var hr = createW(doc, 'headerReference');
      hr.setAttributeNS(W_NS, 'w:type', 'default');
      hr.setAttributeNS(R_NS, 'r:id', rid);
      insertInOrder(sectPrs[k], hr, SECTPR_ORDER);
    }
    return true;
  }

  /* 确保一/二/三级标题段落带 Word 标题样式（纯手工加粗的标题在更新目录域后
     会掉出目录，页眉 STYLEREF 也取不到）。返回「一级标题」样式名。 */
  async function ensureHeadingStyles(zip, doc, paras, roles) {
    var entry = zip.file('word/styles.xml');
    if (!entry) return null;
    var stDoc = new DOMParser().parseFromString(await entry.async('string'), 'application/xml');
    var styles = allByNs(stDoc, 'style', W_NS);
    var byLevel = {}, nameByLevel = {}, used = {}, created = false;
    for (var i = 0; i < styles.length; i++) {
      var sid = styles[i].getAttributeNS(W_NS, 'styleId');
      if (sid) used[sid] = true;
      var nm = childByNs(styles[i], 'name', W_NS);
      if (!nm) continue;
      var name = (nm.getAttributeNS(W_NS, 'val') || '').trim();
      var m = /^(?:heading|标题)\s*([123])$/i.exec(name);
      if (!m) continue;
      var lvl = 'h' + m[1];
      if (byLevel[lvl]) continue;
      byLevel[lvl] = sid;
      nameByLevel[lvl] = name;
    }
    /* 缺哪级补哪级。很多文档（WPS 导出、手改过的论文）压根不带「标题 N」样式：
       而 updateFields 会让 Word 打开时刷新 TOC 域——没有标题样式的段落会被全部
       踢出目录，目录直接变空；页眉的 STYLEREF 也取不到章名。
       样式只声明 outlineLvl，字体字号交给段落上的直接格式，避免改版面。 */
    for (var lv = 1; lv <= 3; lv++) {
      var key = 'h' + lv;
      if (byLevel[key]) continue;
      var newId = 'Heading' + lv, k = 1;
      while (used[newId]) { k++; newId = 'Heading' + lv + '_' + k; }
      used[newId] = true;
      var st = createW(stDoc, 'style');
      st.setAttributeNS(W_NS, 'w:type', 'paragraph');
      st.setAttributeNS(W_NS, 'w:styleId', newId);
      var nEl = createW(stDoc, 'name');
      nEl.setAttributeNS(W_NS, 'w:val', 'heading ' + lv);
      st.appendChild(nEl);
      var bo = createW(stDoc, 'basedOn');
      bo.setAttributeNS(W_NS, 'w:val', 'Normal');
      st.appendChild(bo);
      var nx = createW(stDoc, 'next');
      nx.setAttributeNS(W_NS, 'w:val', 'Normal');
      st.appendChild(nx);
      var spPr = createW(stDoc, 'pPr');
      var ol = createW(stDoc, 'outlineLvl');
      ol.setAttributeNS(W_NS, 'w:val', String(lv - 1));
      spPr.appendChild(ol);
      st.appendChild(spPr);
      stDoc.documentElement.appendChild(st);
      byLevel[key] = newId;
      nameByLevel[key] = 'heading ' + lv;
      created = true;
    }
    if (created) zip.file('word/styles.xml', serialize(stDoc, 'word/styles.xml'));
    for (var j = 0; j < paras.length; j++) {
      /* 参考文献标题归为 refHead，但它与致谢/附录一样是一级标题：
         没有标题样式时页眉会错显上一章，更新目录也可能掉条目 */
      var role = roles[j] === 'refHead' ? 'h1' : roles[j];
      var id = byLevel[role];
      if (!id) continue;
      var pEl = paras[j].p;
      var pPr = childByNs(pEl, 'pPr', W_NS);
      if (pPr && childByNs(pPr, 'pStyle', W_NS)) continue; // 已有样式，不动
      if (!pPr) { pPr = createW(doc, 'pPr'); pEl.insertBefore(pPr, pEl.firstChild); }
      var ps = createW(doc, 'pStyle');
      ps.setAttributeNS(W_NS, 'w:val', id);
      pPr.insertBefore(ps, pPr.firstChild);
    }
    return nameByLevel.h1;
  }

  /* 打开文档时自动更新域（目录页码首次打开即刷新）：
     在 word/settings.xml 写入 <w:updateFields w:val="true"/>。
     schema 顺序中 updateFields 位于 <w:compat> 之前。 */
  async function enableUpdateFields(zip) {
    var entry = zip.file('word/settings.xml');
    if (!entry) return false;
    var xml = await entry.async('string');
    if (/<w:updateFields/.test(xml)) return true;
    var tag = '<w:updateFields w:val="true"/>';
    if (xml.indexOf('<w:compat') >= 0) xml = xml.replace('<w:compat', tag + '<w:compat');
    else if (xml.indexOf('<w:rsids') >= 0) xml = xml.replace('<w:rsids', tag + '<w:rsids');
    else xml = xml.replace('</w:settings>', tag + '</w:settings>');
    zip.file('word/settings.xml', xml);
    return true;
  }

  /* ---------- 导出文件名 ----------
     原名（去掉扩展名）+ 工具版本号：CYX_毕业设计.docx → CYX_毕业设计_v1.6.5.docx。
     只加版本号，不再缀「_格式化」——名字短，也一眼能看出是哪一版跑出来的。
     拿已经带版本号的文件再跑一遍时先去掉旧后缀，免得越拼越长
     （CYX_毕业设计_v1.6.2_v1.6.5.docx）；浏览器重名副本的「 (1)」也一并去掉。
     注意只认「_v数字」这种本工具写的后缀：原作者自己的 CYX_3.0.docx 不能动。 */
  function outputName(fileName) {
    var base = String(fileName || '论文').replace(/\.docx$/i, '');
    base = base.replace(/\s*\(\d+\)$/, '').replace(/[_\-\s]+v\d+(?:\.\d+)*$/i, '');
    return (base || '论文') + '_v' + (FormatTool.VERSION || '0') + '.docx';
  }

  /* ---------- 主入口 ---------- */
  async function formatDocx(data, settings, opts) {
    opts = opts || {};
    settings = Object.assign({}, DEFAULTS, settings || {});
    if (typeof JSZip === 'undefined') throw new Error('缺少 JSZip 库');
    if (typeof DOMParser === 'undefined' || typeof XMLSerializer === 'undefined') throw new Error('当前环境缺少 DOMParser / XMLSerializer');
    emptyNsNodes = []; // 每次格式化重新统计空命名空间节点

    var zip = await JSZip.loadAsync(data);
    var entry = zip.file('word/document.xml');
    if (!entry) throw new Error('不是有效的 .docx 文件（缺少 word/document.xml）');

    var doc = new DOMParser().parseFromString(await entry.async('string'), 'application/xml');
    var root = doc.documentElement;
    if (!root || root.localName !== 'document') throw new Error('文档 XML 解析失败');

    unwrapSdt(doc); // 解包目录 sdt，让目录条目参与识别与格式化

    var paras = collectParas(doc);
    var cls = classifyParas(paras);

    /* 一级标题改「1 绪论」体例（第1章 → 1；附录不编序号、图用图A1）。
       放在最前面：后面几道工序（删中西文空格、真题注）都在看最终文字 */
    if (settings.chapterNumber) cls.counts.renamedChapter = rewriteChapterHeadings(paras, cls.roles);

    /* 前置部分边界：第一个非 front 角色段落（通常即摘要标题）之前的
       元素一律不做任何修改；用元素引用作边界（normalize 会增删元素导致索引漂移） */
    var frontBoundary = null;
    for (var fi = 0; fi < paras.length; fi++) {
      if (cls.roles[fi] === 'front') continue;
      frontBoundary = paras[fi].p;
      break;
    }
    var frontSectPrs = collectFrontSectPrs(doc, frontBoundary);
    var bodyEl = allByNs(doc, 'body', W_NS)[0];

    applyFormatting(doc, paras, cls.roles, settings);

    /* 清除底色：突出显示 + 底纹（论文不应带灰底/彩底标记）。前置部分不动 */
    if (settings.clearHighlight) {
      cls.counts.clearedHighlight = clearHighlights(doc);
      cls.counts.clearedShading = clearShading(doc, frontBoundary);
    }

    /* 页眉/页脚页码域前的孤儿标点（页脚上的「、33」）。
       这是缺陷不是体例偏好，无条件清理，不给开关。 */
    try { cls.counts.strayPunctRemoved = await cleanFooterStrayPunct(zip); }
    catch (e) { cls.counts.strayPunctRemoved = 0; }

    /* 删除中西文之间的空格（标题/题注段首编号后的间隔保留） */
    if (settings.cjkSpace) cls.counts.joinedCjk = removeCjkLatinSpaces(doc, paras, cls.roles);

    /* 图片题注改用 Word 真题注（题注样式 + SEQ 域） */
    if (settings.wordCaption) {
      var capSt = buildStyles(settings).caption;
      var capStyleId = null;
      try { capStyleId = await ensureCaptionStyle(zip); } catch (e) { capStyleId = null; }
      var capRes = applyWordCaptions(doc, bodyEl, paras, cls.roles, frontBoundary, capSt, capStyleId);
      cls.counts.wordCaption = capRes.done;
      cls.counts.captionPlaceholder = capRes.placeholders;
      /* 图题改成「图4.1」后，正文里的「如图4-1所示」同步改，否则引用对不上 */
      cls.counts.figureRefFixed = normalizeFigureRefs(paras, cls.roles);
    }

    /* 每章末尾补「分节符（下一页）」，便于逐章单独设置页眉 */
    if (settings.chapterSection) {
      var secRes = ensureChapterSectionBreaks(doc, bodyEl, paras, cls.roles, frontBoundary);
      cls.counts.sectionAdded = secRes.added;
      cls.counts.sectionKept = secRes.kept;
      cls.counts.sectionBreakDropped = secRes.dropped;
    }

    if (settings.chapterPageBreak) normalizeChapterBreaks(doc, paras, cls.roles, frontBoundary); // 章节换页规整化（防空白页/空行被吞）
    if (settings.autoToc) {
      // 先修正/补建 toc 1/2/3 样式并取回真实样式 id，目录条目才能引用到它们
      var tocIds = null;
      try { tocIds = await ensureTocStyles(zip); } catch (e) { tocIds = null; }
      replaceTocWithField(doc, paras, cls.roles, tocIds); // 目录替换为自动目录域（可整体选择、可更新）
    }
    repositionCaptions(doc, frontBoundary); // 图题移到图片正下方、表题移到表格正上方
    singleSpaceImageParas(doc, frontBoundary); // 含图片的段落行距改单倍，避免图片被固定行距截断

    // 表格：仅保留顶线/底线 1.5 磅（清除其余线）；前两页表格不动
    if (settings.threeLineTable) {
      var tbls = wAll(doc, 'tbl');
      for (var t = 0; t < tbls.length; t++) {
        if (isFrontElement(tbls[t], frontBoundary)) continue;
        applyTableTopBottom(tbls[t], doc);
      }
    }
    /* 表格内文字：五号宋体 + Times New Roman（附件8）。必须单独走一道 ——
       表格段落不在 collectParas 的范围里，applyFormatting 够不着。封面表格不动 */
    if (settings.tableText) {
      var ttSt = { eastFont: settings.tableFont, latinFont: settings.latinFont, size: settings.tableSize };
      var tbls2 = wAll(doc, 'tbl');
      var ttDone = 0;
      for (var u = 0; u < tbls2.length; u++) {
        if (isFrontElement(tbls2[u], frontBoundary)) continue;
        ttDone += formatTableText(doc, tbls2[u], ttSt);
      }
      cls.counts.tableText = ttDone;
    }
    separateTablesAndImages(doc, frontBoundary); // 表格与图片紧邻时自动空一行，避免重叠

    var footerRid = null;
    if (settings.pageNumber !== 'none' && !hasFooterRef(doc)) {
      try { footerRid = await addFooter(zip, doc, settings); } catch (e) { footerRid = null; }
    }
    setSections(doc, settings, footerRid, frontSectPrs);
    if (settings.pageNumberSplit) applyPageNumbering(doc, paras, cls.roles, settings, frontSectPrs);

    /* 标题样式兜底：更新目录域、页眉 STYLEREF 都依赖「标题 N」样式 */
    var h1StyleName = null;
    try { h1StyleName = await ensureHeadingStyles(zip, doc, paras, cls.roles); } catch (e) { h1StyleName = null; }
    // 页眉：正文起各分节页眉 = 一级标题（章名）；前置部分页眉保持输入原样
    var headerSet = false;
    if (settings.bodyHeader && h1StyleName) {
      try {
        headerSet = await applyBodyHeaders(zip, doc, paras, cls.roles, frontSectPrs,
          stylerefName(settings.headerStyleName || h1StyleName));
      } catch (e) { headerSet = false; }
    }
    // 打开文档时自动更新域（目录页码/页眉章名首次打开即刷新）
    if (settings.updateFields) { try { await enableUpdateFields(zip); } catch (e) { /* ignore */ } }

    /* 导出前兜底：域字符必须配平，否则 Word 更新域时会吞掉正文 */
    var fieldsFixed = repairFields(doc);

    zip.file('word/document.xml', serialize(doc, 'word/document.xml'));

    var counts = cls.counts;
    /* 空命名空间节点：非 0 说明某个部件里有「无命名空间」元素，Word 会认成损坏 */
    counts.emptyNsNodes = emptyNsNodes.length;
    if (emptyNsNodes.length) console.warn('[格式助手] 序列化出现空命名空间节点，Word 可能报文件损坏：', emptyNsNodes.join(', '));
    counts.tables = wAll(doc, 'tbl').length;
    counts.formulas = allByNs(doc, 'oMath', M_NS).length;
    counts.images = Object.keys(zip.files).filter(function (f) { return /^word\/media\//.test(f); }).length;
    counts.paras = paras.length;
    counts.fieldsFixed = fieldsFixed;
    counts.footerAdded = !!footerRid;
    counts.bodyHeader = headerSet;

    var type = opts.format === 'nodebuffer' ? 'nodebuffer' : 'blob';
    var out = await zip.generateAsync({ type: type, mimeType: MIME_DOCX, compression: 'DEFLATE' });
    return { data: out, counts: counts, info: cls.info, settings: settings };
  }

  var FormatTool = { VERSION: '1.6.7', DEFAULTS: DEFAULTS, formatDocx: formatDocx,
    classifyParas: classifyParas, outputName: outputName };

  if (typeof module !== 'undefined' && module.exports) module.exports = FormatTool;
  else global.FormatTool = FormatTool;

})(typeof globalThis !== 'undefined' ? globalThis : this);
