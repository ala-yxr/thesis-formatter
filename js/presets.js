/* ============================================================
 * 论文格式助手 —— 学校格式预设（工科 / 文科分组）
 *
 * 在设置面板顶部的「学校要求」下拉框中选择学校，即可整组切换
 * 格式要求。下拉框按当前学科类型（工科/文科）过滤，只显示
 * 该学科下的预设；两个学科的状态（设置/学校记忆/自定义预设）
 * 完全隔离，互不影响。
 *
 * 新增学校只需在 SCHOOL_PRESETS 数组中追加一项：
 *   id       — 唯一标识（用于记忆上次选择）
 *   name     — 下拉框显示名称
 *   mode     — 'gongke' 工科 | 'wenke' 文科（决定出现在哪个下拉框）
 *   desc     — 切换后显示的要求简介
 *   settings — 对应设置面板中各字段的值（与 formatter.js DEFAULTS 同构，
 *              只列面板上可调的字段即可）
 * ============================================================ */
(function (global) {
  'use strict';

  var SCHOOL_PRESETS = [
    {
      id: 'cqgc',
      mode: 'gongke',
      name: '重庆工程学院（附件8）',
      desc: '页边距 2.5cm / 行距固定 20 磅 / 宋体小四 + Times New Roman / 一级三号黑体居中、二级小三、三级四号右缩两字 / 一级标题「1 绪论」体例 / 图题表题五号宋体 / 三线表（顶底线 1.5 磅 + 栏目线 0.75 磅）/ 目录小四宋体 2-4 级右缩两字并自动更新页码 / 页眉自动显示章名 / 一级标题换页 / 分节页码（前置罗马、主体阿拉伯从 1 起）',
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
        threeLineTable: true, tableText: true, tableFont: '宋体', tableSize: 10.5,
        chapterPageBreak: true, autoToc: true,
        bodyHeader: true, updateFields: true, citeSuperscript: true,
        clearHighlight: true, cjkSpace: true, wordCaption: true, chapterSection: true,
        chapterNumber: true
      }
    },
    {
      id: 'generic',
      mode: 'wenke',
      name: '通用文科规范（示例）',
      desc: '常见文科通用格式：页边距 上3.0/下2.5/左3.0(装订)/右2.5cm / 1.5 倍行距 / 宋体小四 / 一级三号黑体居中 / 表格仅保留顶底线 1.5 磅 / 页码底端居中（不拆分罗马/阿拉伯）。选择后仍可按需手动微调各分项。',
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
        threeLineTable: true, tableText: true, tableFont: '宋体', tableSize: 10.5,
        chapterPageBreak: true, autoToc: true,
        bodyHeader: false, updateFields: true, citeSuperscript: true,
        clearHighlight: true, cjkSpace: true, wordCaption: true, chapterSection: true,
        chapterNumber: false   // 文科：一级标题用汉字序号（「一 员工绩效考核…」），章号体例不改
      }
    }
  ];

  /* 学科类型定义：入口选择页与顶栏徽标使用 */
  var MODES = {
    gongke: {
      label: '工科论文',
      sub: '理工科毕业设计 / 论文',
      defaultPresetId: 'cqgc',
      gateDesc: '适用于理工科类论文（如重庆工程学院《附件8》：行距固定 20 磅、三线表、分节页码等）'
    },
    wenke: {
      label: '文科论文',
      sub: '人文社科类论文',
      defaultPresetId: 'generic',
      gateDesc: '适用于人文社科类论文（通用文科规范：1.5 倍行距、装订边距等，可在设置面板调整）'
    }
  };

  global.SCHOOL_PRESETS = SCHOOL_PRESETS;
  global.MODES = MODES;
})(typeof globalThis !== 'undefined' ? globalThis : this);
