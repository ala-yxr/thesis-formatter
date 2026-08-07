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
  var CT_NS = 'http://schemas.openxmlformats.org/package/2006/content-types';
  var PKG_REL_NS = 'http://schemas.openxmlformats.org/package/2006/relationships';
  var CT_FOOTER = 'application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml';
  var REL_FOOTER = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer';
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
    chapterPageBreak: true,        // 一级标题之间换页
    autoToc: true                  // 目录替换为 Word 自动目录域（可整体选择、可更新）
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
    h3:     /^[1-9]\d?\.\d{1,2}\.\d{1,2}\s/,
    h2:     /^[1-9]\d?\.\d{1,2}(?![\d.])\s*[^\d.]/,
    h2cn:   /^[一二三四五六七八九十]{1,3}[、.]/,
    cap:    /^(图|表|Figure|Table)\s*\d/,
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
  function allByNs(elm, name, ns) {
    return elm.getElementsByTagNameNS(ns, name);
  }
  function wAll(elm, name) { return allByNs(elm, name, W_NS); }
  function childByNs(elm, name, ns) {
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
  function serialize(doc) {
    // 解析时 xml 声明会作为 PI 节点保留，序列化前移除，避免双重声明
    while (doc.childNodes.length && doc.childNodes[0].nodeType === 7) {
      doc.removeChild(doc.childNodes[0]);
    }
    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
      new XMLSerializer().serializeToString(doc);
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

    /* 无目录标题（「目 录」）时的兜底识别：摘要/关键词结束之后、第一个一级标题
       之前，若存在 ≥2 个目录条目（tab+页码等特征），则视为目录章节 */
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

    // 摘要正文区域边界：无"关键词"行时延伸到英文摘要标题或第一个章节标题
    var firstH = -1;
    for (var h = 0; h < n; h++) {
      var ht = texts[h];
      if (ht && (RE.h1.test(ht) || RE.h1Word.test(ht) || (!hasChapterStyle && RE.h1Alt.test(ht)))) { firstH = h; break; }
    }
    var absBodyEnd = kwIdx >= 0 ? kwIdx : (absEnIdx >= 0 ? absEnIdx : (firstH >= 0 ? firstH : n));
    var absEnBodyEnd = kwEnIdx >= 0 ? kwEnIdx : (firstH >= 0 ? firstH : n);

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
      if (refIdx >= 0 && i2 > refIdx) {
        if (RE.ack.test(tx) || RE.app.test(tx)) { roles[i2] = 'h1'; inBack = true; }
        else roles[i2] = inBack ? 'body' : 'refItem';
        continue;
      }
      if (paras[i2].hasMath) { roles[i2] = 'formula'; continue; }
      if (!tx && paras[i2].hasDrawing) { roles[i2] = 'figure'; continue; }                  // 纯图片段：不动
      if (RE.h1.test(tx) || (!hasChapterStyle && RE.h1Alt.test(tx)) || RE.h1Word.test(tx)) { roles[i2] = 'h1'; continue; }
      /* 二/三级标题判定加长度守卫：正文长句（如「1.5倍…」「0.96英寸…」）不当作标题 */
      if (RE.h3.test(tx) && tx.length <= 40) { roles[i2] = 'h3'; continue; }
      if ((RE.h2.test(tx) || RE.h2cn.test(tx)) && tx.length <= 40) { roles[i2] = 'h2'; continue; }
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

  /* 构建目录条目段落（toc 样式 + 标题文本 + tab + PAGEREF 页码域） */
  function buildTocEntryPara(doc, level, text, bookmark) {
    var p = createW(doc, 'p');
    var pPr = createW(doc, 'pPr');
    var ps = createW(doc, 'pStyle');
    ps.setAttributeNS(W_NS, 'w:val', 'TOC' + (level + 1));
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
  function buildTocFieldSdt(doc, headings) {
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
      content.appendChild(buildTocEntryPara(doc, headings[i].level, headings[i].text, headings[i].bookmark));
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
  function replaceTocWithField(doc, paras, roles) {
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
      var role = roles[i2];
      var lvl = role === 'h1' ? 0 : role === 'h2' ? 1 : role === 'h3' ? 2 : -1;
      if (lvl < 0) continue;
      var bm = '_Toc' + (headings.length + 1);
      addBookmarkTo(doc, pEl, ++maxId, bm);
      headings.push({ level: lvl, text: paraText(pEl).trim(), bookmark: bm });
    }
    if (headings.length === 0) return false;

    /* 删除目录标题后到分节段之间的元素（保留标题后的 1 个空段；
       分节段是目录节末尾与章节换页机制，其后内容（绪论前空段+绪论）不删除） */
    var keptOne = false;
    var cur = tocHeadEl.nextSibling;
    while (cur && cur !== tocEndEl) {
      var nx = cur.nextSibling;
      var isSectP = cur.nodeType === 1 && cur.localName === 'p' &&
                    childByNs(childByNs(cur, 'pPr', W_NS), 'sectPr', W_NS);
      if (isSectP) break;
      var isSep = cur.nodeType === 1 && cur.localName === 'p' && isEmptySeparatorPara(cur);
      if (isSep && !keptOne) { keptOne = true; }
      else body.removeChild(cur);
      cur = nx;
    }
    /* 插入 TOC 域：位于标题后的空段之后、分节段之前 */
    var sdt = buildTocFieldSdt(doc, headings);
    var anchor = tocHeadEl.nextSibling;
    var anchorIsSect = anchor && anchor.nodeType === 1 && anchor.localName === 'p' &&
                       childByNs(childByNs(anchor, 'pPr', W_NS), 'sectPr', W_NS);
    body.insertBefore(sdt, anchorIsSect ? anchor : (anchor ? anchor.nextSibling : tocEndEl));
    return true;
  }

  /* 修正 styles.xml 中 toc 1/2/3 样式：小四（12pt）、层级缩进（一级0/二级2字/三级4字），
     使 Word 自动更新目录后条目格式符合附件8 */
  async function ensureTocStyles(zip) {
    var entry = zip.file('word/styles.xml');
    if (!entry) return;
    var stDoc = new DOMParser().parseFromString(await entry.async('string'), 'application/xml');
    var targets = { 'toc 1': 0, 'toc 2': 480, 'toc 3': 960 };
    var styles = stDoc.getElementsByTagNameNS(W_NS, 'style');
    for (var i = 0; i < styles.length; i++) {
      var st = styles[i];
      var nm = childByNs(st, 'name', W_NS);
      if (!nm) continue;
      var name = (nm.getAttributeNS(W_NS, 'val') || '').toLowerCase();
      if (!(name in targets)) continue;
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
      var left = targets[name];
      var pPr = childByNs(st, 'pPr', W_NS);
      var ind = pPr && childByNs(pPr, 'ind', W_NS);
      if (left) {
        if (!pPr) { pPr = createW(stDoc, 'pPr'); st.insertBefore(pPr, rPr); }
        if (!ind) { ind = createW(stDoc, 'ind'); insertInOrder(pPr, ind, PPR_ORDER); }
        ind.setAttributeNS(W_NS, 'w:left', String(left));
      } else if (ind) {
        ind.removeAttributeNS(W_NS, 'w:left');
      }
    }
    zip.file('word/styles.xml', serialize(stDoc));
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
      var ov = ctDoc.createElement('Override');
      ov.setAttribute('PartName', '/' + target);
      ov.setAttribute('ContentType', CT_FOOTER);
      ctDoc.documentElement.appendChild(ov);
      zip.file(target, footerXml(s.pageNumber === 'right' ? 'right' : 'center'));
    } else if (!zip.file(target)) {
      zip.file(target, footerXml(s.pageNumber === 'right' ? 'right' : 'center'));
    }
    zip.file('[Content_Types].xml', serialize(ctDoc));

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
      var rel = relsDoc.createElement('Relationship');
      rel.setAttribute('Id', rid);
      rel.setAttribute('Type', REL_FOOTER);
      rel.setAttribute('Target', target);
      relsDoc.documentElement.appendChild(rel);
    }
    zip.file('word/_rels/document.xml.rels', serialize(relsDoc));
    return rid;
  }

  function hasFooterRef(doc) {
    var sps = allByNs(doc, 'sectPr', W_NS);
    for (var i = 0; i < sps.length; i++) {
      if (wAll(sps[i], 'footerReference').length > 0) return true;
    }
    return false;
  }

  /* ---------- 主入口 ---------- */
  async function formatDocx(data, settings, opts) {
    opts = opts || {};
    settings = Object.assign({}, DEFAULTS, settings || {});
    if (typeof JSZip === 'undefined') throw new Error('缺少 JSZip 库');
    if (typeof DOMParser === 'undefined' || typeof XMLSerializer === 'undefined') throw new Error('当前环境缺少 DOMParser / XMLSerializer');

    var zip = await JSZip.loadAsync(data);
    var entry = zip.file('word/document.xml');
    if (!entry) throw new Error('不是有效的 .docx 文件（缺少 word/document.xml）');

    var doc = new DOMParser().parseFromString(await entry.async('string'), 'application/xml');
    var root = doc.documentElement;
    if (!root || root.localName !== 'document') throw new Error('文档 XML 解析失败');

    unwrapSdt(doc); // 解包目录 sdt，让目录条目参与识别与格式化

    var paras = collectParas(doc);
    var cls = classifyParas(paras);

    /* 前置部分边界：第一个非 front 角色段落（通常即摘要标题）之前的
       元素一律不做任何修改；用元素引用作边界（normalize 会增删元素导致索引漂移） */
    var frontBoundary = null;
    for (var fi = 0; fi < paras.length; fi++) {
      if (cls.roles[fi] === 'front') continue;
      frontBoundary = paras[fi].p;
      break;
    }
    var frontSectPrs = collectFrontSectPrs(doc, frontBoundary);

    applyFormatting(doc, paras, cls.roles, settings);
    if (settings.chapterPageBreak) normalizeChapterBreaks(doc, paras, cls.roles, frontBoundary); // 章节换页规整化（防空白页/空行被吞）
    if (settings.autoToc) {
      replaceTocWithField(doc, paras, cls.roles);   // 目录替换为自动目录域（可整体选择、可更新）
      await ensureTocStyles(zip);                   // 修正 toc 1/2/3 样式（小四+层级缩进）
    }
    // 页眉保持输入原样，不做任何修改
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
    separateTablesAndImages(doc, frontBoundary); // 表格与图片紧邻时自动空一行，避免重叠

    var footerRid = null;
    if (settings.pageNumber !== 'none' && !hasFooterRef(doc)) {
      try { footerRid = await addFooter(zip, doc, settings); } catch (e) { footerRid = null; }
    }
    setSections(doc, settings, footerRid, frontSectPrs);
    if (settings.pageNumberSplit) applyPageNumbering(doc, paras, cls.roles, settings, frontSectPrs);

    zip.file('word/document.xml', serialize(doc));

    var counts = cls.counts;
    counts.tables = wAll(doc, 'tbl').length;
    counts.formulas = allByNs(doc, 'oMath', M_NS).length;
    counts.images = Object.keys(zip.files).filter(function (f) { return /^word\/media\//.test(f); }).length;
    counts.paras = paras.length;
    counts.footerAdded = !!footerRid;

    var type = opts.format === 'nodebuffer' ? 'nodebuffer' : 'blob';
    var out = await zip.generateAsync({ type: type, mimeType: MIME_DOCX, compression: 'DEFLATE' });
    return { data: out, counts: counts, info: cls.info, settings: settings };
  }

  var FormatTool = { VERSION: '1.0.0', DEFAULTS: DEFAULTS, formatDocx: formatDocx, classifyParas: classifyParas };

  if (typeof module !== 'undefined' && module.exports) module.exports = FormatTool;
  else global.FormatTool = FormatTool;

})(typeof globalThis !== 'undefined' ? globalThis : this);
