/* ============================================================
 * 论文格式助手 —— 学校格式预设
 *
 * 在设置面板顶部的「学校要求」下拉框中选择学校，即可整组切换
 * 格式要求。新增学校只需在 SCHOOL_PRESETS 数组中追加一项：
 *   id       — 唯一标识（用于记忆上次选择）
 *   name     — 下拉框显示名称
 *   desc     — 切换后显示的要求简介
 *   settings — 对应设置面板中各字段的值（与 formatter.js DEFAULTS 同构，
 *              只列面板上可调的字段即可）
 * ============================================================ */
(function (global) {
  'use strict';

  var SCHOOL_PRESETS = [
    {
      id: 'cqgc',
      name: '重庆工程学院（附件8）',
      desc: '页边距 2.5cm / 行距固定 20 磅 / 宋体小四 + Times New Roman / 一级三号黑体居中、二级小三、三级四号右缩两字 / 图题表题五号宋体 / 表格仅保留顶底线 1.5 磅 / 目录小四宋体 2-4 级右缩两字 / 一级标题换页 / 分节页码（前置罗马、主体阿拉伯从 1 起）',
      settings: {
        marginTop: 2.5, marginBottom: 2.5, marginLeft: 2.5, marginRight: 2.5,
        bodyFont: '宋体', latinFont: 'Times New Roman', bodySize: 12,
        lineSpacingMode: 'exact', lineSpacing: 20,
        firstLineChars: 2,
        titleSize: 22, titleBold: false,
        h1Size: 16, h1Jc: 'center', h2Size: 15, h3Size: 14, headingBold: false,
        absHeadSize: 16, absBodySize: 12,
        refSize: 10.5, refHangingChars: 0,
        captionFont: '宋体', captionSize: 10.5,
        pageNumber: 'center', pageNumberSplit: true,
        threeLineTable: true, chapterPageBreak: true, autoToc: true
      }
    },
    {
      id: 'generic',
      name: '通用本科规范（示例）',
      desc: '常见通用格式：页边距 上3.0/下2.5/左3.0(装订)/右2.5cm / 1.5 倍行距 / 宋体小四 / 一级三号黑体居中 / 表格仅保留顶底线 1.5 磅 / 页码底端居中（不拆分罗马/阿拉伯）。选择后仍可按需手动微调各分项。',
      settings: {
        marginTop: 3.0, marginBottom: 2.5, marginLeft: 3.0, marginRight: 2.5,
        bodyFont: '宋体', latinFont: 'Times New Roman', bodySize: 12,
        lineSpacingMode: 'multiple', lineSpacing: 1.5,
        firstLineChars: 2,
        titleSize: 22, titleBold: false,
        h1Size: 16, h1Jc: 'center', h2Size: 15, h3Size: 14, headingBold: false,
        absHeadSize: 16, absBodySize: 12,
        refSize: 10.5, refHangingChars: 0,
        captionFont: '宋体', captionSize: 10.5,
        pageNumber: 'center', pageNumberSplit: false,
        threeLineTable: true, chapterPageBreak: true, autoToc: true
      }
    }
  ];

  global.SCHOOL_PRESETS = SCHOOL_PRESETS;
})(typeof globalThis !== 'undefined' ? globalThis : this);
