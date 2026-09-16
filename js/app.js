/* ============================================================
 * 论文格式助手 —— 界面逻辑
 * 依赖：libs/jszip.min.js, libs/mammoth.browser.min.js, js/formatter.js
 * ============================================================ */
(function () {
  'use strict';

  const $ = (sel) => document.querySelector(sel);
  const MODES = (typeof globalThis !== 'undefined' && globalThis.MODES) || {};
  const MODE_IDS = Object.keys(MODES);
  const LS_MODE_KEY = 'thesis-formatter-mode';

  /* ---------- 学科模式（工科/文科）完全隔离 ----------
   * 所有状态按学科分空间保存（key 加后缀），工科与文科互不读写：
   *   thesis-formatter-settings-gongke / -wenke      设置表单
   *   thesis-formatter-school-gongke / -wenke        记忆的学校
   *   thesis-formatter-custom-presets-gongke / -wenke 自定义学校
   * 入口选择页每次进入先选学科；顶栏徽标可随时切换 */
  let currentMode = 'gongke';
  const LS_KEY = 'thesis-formatter-settings';
  const LS_SCHOOL_KEY = 'thesis-formatter-school';

  /* 当前模式下的 storage key（隔离核心：key 按学科加后缀） */
  function modeKey(base) { return base + '-' + currentMode; }

  /* ---------- 学校格式预设 ----------
   * 内置：js/presets.js 中定义；自定义：用户上传要求文档后保存到 localStorage */
  const PRESETS = (typeof globalThis !== 'undefined' && globalThis.SCHOOL_PRESETS) || [];
  const LS_CUSTOM_KEY = 'thesis-formatter-custom-presets';
  function customPresets() {
    try { return JSON.parse(localStorage.getItem(modeKey(LS_CUSTOM_KEY))) || []; } catch (e) { return []; }
  }
  function saveCustomPresets(list) {
    try { localStorage.setItem(modeKey(LS_CUSTOM_KEY), JSON.stringify(list)); return true; }
    catch (e) {
      toast('保存失败：浏览器存储空间不足（要求文档较大），请删除部分预设或换更小的文档', true);
      return false;
    }
  }
  /* 当前学科下的内置预设（下拉框按学科过滤） */
  function modePresets() { return PRESETS.filter((p) => p.mode === currentMode); }
  function presetById(id) {
    return modePresets().find((p) => p.id === id) || customPresets().find((p) => p.id === id);
  }
  function defaultPresetId() {
    const def = MODES[currentMode] && MODES[currentMode].defaultPresetId;
    if (def && presetById(def)) return def;
    return modePresets().length ? modePresets()[0].id : null;
  }

  /* ---------- 旧版数据一次性迁移 ----------
   * v1.3 之前所有状态都在无后缀 key 下（即工科行为），首次运行新版本
   * 时整体搬进工科空间并删除旧 key，保证老用户工科状态不丢 */
  function migrateLegacy() {
    try {
      const legacy = ['thesis-formatter-settings', 'thesis-formatter-school', 'thesis-formatter-custom-presets'];
      const anyLegacy = legacy.some((k) => localStorage.getItem(k) !== null);
      if (!anyLegacy) return;
      if (currentMode === 'gongke') {
        legacy.forEach((k) => {
          const v = localStorage.getItem(k);
          if (v !== null && localStorage.getItem(modeKey(k)) === null) localStorage.setItem(modeKey(k), v);
          localStorage.removeItem(k);
        });
      } else {
        legacy.forEach((k) => localStorage.removeItem(k));
      }
    } catch (e) { /* ignore */ }
  }

  const S = {
    dropzone: $('#dropzone'), fileInput: $('#fileInput'),
    filebar: $('#filebar'), fileName: $('#fileName'), fileSize: $('#fileSize'),
    btnReupload: $('#btnReupload'),
    btnReset: $('#btnReset'), btnFormat: $('#btnFormat'),
    btnDownload: $('#btnDownload'), btnPreviewReformat: $('#btnPreviewReformat'),
    statsCard: $('#statsCard'), chips: $('#chips'), statsTime: $('#statsTime'), statsNotes: $('#statsNotes'),
    previewEmpty: $('#previewEmpty'), previewFrame: $('#previewFrame'), previewStatus: $('#previewStatus'),
    toast: $('#toast'), overlay: $('#overlay'), overlayText: $('#overlayText'),
    btnAddSchool: $('#btnAddSchool'), btnDelSchool: $('#btnDelSchool'), fileSchool: $('#fileSchool'),
    chkTemplateFront: $('#templateFront'), btnPickTemplate: $('#btnPickTemplate'),
    btnClearTemplate: $('#btnClearTemplate'), fileTemplate: $('#fileTemplate'),
    templateStatus: $('#templateStatus'),
    addSchoolModal: $('#addSchoolModal'), schoolNameInput: $('#schoolNameInput'),
    schoolDocName: $('#schoolDocName'), btnAddSchoolConfirm: $('#btnAddSchoolConfirm'),
    btnAddSchoolCancel: $('#btnAddSchoolCancel'),
    gate: $('#modeGate'), btnSwitchMode: $('#btnSwitchMode'), modeBadge: $('#modeBadge'),
  };

  let lastFile = null;      // 当前处理的文件
  let lastBlob = null;      // 当前格式化结果

  /* ---------- 设置读写 ---------- */
  const SETTINGS_FIELDS = [
    ['marginTop', 'number'], ['marginBottom', 'number'], ['marginLeft', 'number'], ['marginRight', 'number'],
    ['pageNumber', 'select'], ['bodyFont', 'select'], ['latinFont', 'select'], ['bodySize', 'number'],
    ['firstLineChars', 'number'], ['titleSize', 'number'], ['titleBold', 'check'],
    ['h1Size', 'number'], ['h1Jc', 'select'], ['h2Size', 'number'], ['h3Size', 'number'], ['headingBold', 'check'],
    ['absHeadSize', 'number'], ['absBodySize', 'number'], ['refSize', 'number'], ['refHangingChars', 'number'],
    ['captionFont', 'select'], ['captionSize', 'number'],
    ['tableFont', 'select'], ['tableSize', 'number'],
    ['pageNumberSplit', 'check'], ['threeLineTable', 'check'], ['tableText', 'check'], ['chapterPageBreak', 'check'],
    ['bodyHeader', 'check'], ['updateFields', 'check'], ['citeSuperscript', 'check'],
    ['clearHighlight', 'check'], ['cjkSpace', 'check'],
    ['wordCaption', 'check'], ['chapterSection', 'check'],
    ['chapterNumber', 'check'],
  ];

  // 行距采用「模式:数值」组合（如 exact:20 = 固定20磅，multiple:1.5 = 1.5倍）
  function currentSettings() {
    const s = {};
    SETTINGS_FIELDS.forEach(([id, type]) => {
      const el = $('#' + id);
      s[id] = type === 'check' ? el.checked : (type === 'number' ? parseFloat(el.value) : el.value);
    });
    const ls = $('#lineSpacing').value.split(':');
    s.lineSpacingMode = ls[0];
    s.lineSpacing = parseFloat(ls[1]);
    return s;
  }
  function applySettings(s) {
    SETTINGS_FIELDS.forEach(([id, type]) => {
      const el = $('#' + id);
      if (s[id] === undefined) return;
      if (type === 'check') el.checked = !!s[id];
      else el.value = String(s[id]);
    });
    if (s.lineSpacingMode && s.lineSpacing !== undefined) {
      $('#lineSpacing').value = s.lineSpacingMode + ':' + s.lineSpacing;
    }
  }
  function saveSettings() {
    try { localStorage.setItem(modeKey(LS_KEY), JSON.stringify(currentSettings())); } catch (e) { /* ignore */ }
  }

  /* ---------- 学科模式（工科/文科）切换 ---------- */
  const schoolSel = $('#schoolPreset');
  const schoolDesc = $('#schoolPresetDesc');

  // 顶栏徽标 + 设置面板提示显示当前学科
  function updateModeBadge() {
    const m = MODES[currentMode];
    if (!m) return;
    if (S.modeBadge) {
      S.modeBadge.textContent = m.label + '（' + m.sub + '）';
      S.modeBadge.dataset.mode = currentMode;
    }
    const hint = $('#modeHint');
    if (hint) hint.innerHTML = '当前模式：' + m.label + ' · 以下设置与自定义学校仅保存到' +
      (currentMode === 'gongke' ? '工科' : '文科') + '模式，不影响' + (currentMode === 'gongke' ? '文科' : '工科');
  }

  // 从 MODES 渲染入口选择页的两个学科卡片
  function renderGateOptions() {
    const wrap = document.getElementById('gateOptions');
    if (!wrap) return;
    wrap.innerHTML = MODE_IDS.map((id) => {
      const m = MODES[id];
      return '<button class="gate-option" data-mode="' + id + '" type="button">' +
        '<span class="gate-ico">' + (id === 'gongke' ? '⚙️' : '📖') + '</span>' +
        '<span class="gate-name">' + m.label + '</span>' +
        '<span class="gate-sub">' + (m.sub || '') + '</span>' +
        '<span class="gate-desc">' + (m.gateDesc || '') + '</span>' +
        '</button>';
    }).join('');
  }

  /* 进入某学科模式：重建下拉 → 应用该校预设（记忆的学校或默认）
     → 叠加该模式保存的手动设置 → 更新徽标 → 隐藏选择页 */
  function enterMode(mode) {
    if (!MODES[mode]) return;
    currentMode = mode;
    try { localStorage.setItem(LS_MODE_KEY, mode); } catch (e) { /* ignore */ }
    buildSchoolOptions();
    const savedSchool = localStorage.getItem(modeKey(LS_SCHOOL_KEY));
    if (savedSchool && presetById(savedSchool)) applyPreset(savedSchool, { silent: true });
    else {
      const defId = defaultPresetId();
      if (defId) applyPreset(defId, { silent: true });
    }
    const raw = localStorage.getItem(modeKey(LS_KEY));
    if (raw) applySettings(JSON.parse(raw));
    updateModeBadge();
    S.gate.hidden = true;
    if (lastFile) {
      toast('已进入「' + MODES[mode].label + '」模式，正在按该学科格式要求重新格式化…');
      formatFile(lastFile);
    }
  }

  function showGate() { S.gate.hidden = false; }

  function bindGate() {
    renderGateOptions();
    document.getElementById('gateOptions').addEventListener('click', (e) => {
      const btn = e.target.closest('.gate-option');
      if (btn && btn.dataset.mode) enterMode(btn.dataset.mode);
    });
    S.btnSwitchMode.addEventListener('click', showGate);
  }

  /* ---------- 学校要求切换 ---------- */
  function buildSchoolOptions() {
    const all = modePresets().concat(customPresets());
    if (!all.length) { schoolSel.closest('.row').hidden = true; return; }
    schoolSel.innerHTML = all.map((p) => `<option value="${p.id}">${p.name}</option>`).join('');
    updateDelSchoolBtn();
  }

  // 自定义预设（上传要求文档添加的）才允许删除
  function updateDelSchoolBtn() {
    const p = presetById(schoolSel.value);
    S.btnDelSchool.hidden = !(p && p.custom);
  }

  // 应用某套学校预设：整组写入设置表单；若已上传文档则自动按新要求重新格式化
  function applyPreset(id, opts) {
    const p = presetById(id);
    if (!p) return false;
    opts = opts || {};
    schoolSel.value = p.id;
    applySettings(p.settings);
    saveSettings();
    try { localStorage.setItem(modeKey(LS_SCHOOL_KEY), id); } catch (e) { /* ignore */ }
    schoolDesc.textContent = p.desc;
    updateDelSchoolBtn();
    if (opts.silent) return true;
    if (lastFile) {
      toast('已切换到「' + p.name + '」，正在按新格式要求重新格式化…');
      formatFile(lastFile);
    } else {
      toast('已切换到「' + p.name + '」格式要求');
    }
    return true;
  }

  /* ---------- 添加学校要求文档（自定义预设） ---------- */
  let pendingCustomDoc = null; // 待保存的要求文档 {name, dataUrl}

  function bindAddSchool() {
    // 上传要求文档
    S.btnAddSchool.addEventListener('click', () => S.fileSchool.click());
    S.fileSchool.addEventListener('change', () => {
      const f = S.fileSchool.files[0];
      S.fileSchool.value = '';
      if (!f) return;
      if (!/\.docx$/i.test(f.name)) { toast('请选择 .docx 格式的要求文档', true); return; }
      if (f.size > 2 * 1024 * 1024) { toast('要求文档超过 2MB，请压缩后重试', true); return; }
      showOverlay('正在读取要求文档…');
      const reader = new FileReader();
      reader.onload = () => {
        hideOverlay();
        pendingCustomDoc = { name: f.name, dataUrl: reader.result };
        S.schoolNameInput.value = f.name.replace(/\.docx$/i, '');
        S.schoolDocName.textContent = '要求文档：' + f.name + '（' + (f.size / 1024).toFixed(0) + ' KB）';
        S.addSchoolModal.hidden = false;
      };
      reader.onerror = () => { hideOverlay(); toast('读取文档失败', true); };
      reader.readAsDataURL(f);
    });

    // 保存：以当前设置面板的格式要求 + 要求文档，生成该校预设
    S.btnAddSchoolConfirm.addEventListener('click', () => {
      const name = S.schoolNameInput.value.trim();
      if (!name) { toast('请填写学校名称', true); return; }
      if (!pendingCustomDoc) return;
      const list = customPresets();
      list.push({
        id: 'custom-' + Date.now().toString(36) + Math.random().toString(36).slice(2, 6),
        name: name,
        mode: currentMode,   // 学科标签：工科建的自定义学校只出现在工科下拉
        desc: '自定义学校（' + MODES[currentMode].label + '）：以「' + pendingCustomDoc.name + '」为要求依据（' +
          new Date().toLocaleDateString('zh-CN') + ' 添加）',
        custom: true,
        doc: pendingCustomDoc.dataUrl,
        settings: currentSettings()
      });
      if (!saveCustomPresets(list)) return;
      pendingCustomDoc = null;
      S.addSchoolModal.hidden = true;
      buildSchoolOptions();
      applyPreset(list[list.length - 1].id);
    });

    S.btnAddSchoolCancel.addEventListener('click', () => {
      pendingCustomDoc = null;
      S.addSchoolModal.hidden = true;
    });

    // 删除自定义预设（删除后恢复为当前学科的内置默认预设）
    S.btnDelSchool.addEventListener('click', () => {
      const p = presetById(schoolSel.value);
      if (!p || !p.custom) return;
      saveCustomPresets(customPresets().filter((x) => x.id !== p.id));
      buildSchoolOptions();
      applyPreset(defaultPresetId(), { silent: true });
      toast('已删除「' + p.name + '」预设');
    });
  }

  /* ---------- 提示 ---------- */
  let toastTimer = null;
  function toast(msg, isErr) {
    S.toast.textContent = msg;
    S.toast.classList.toggle('err', !!isErr);
    S.toast.hidden = false;
    clearTimeout(toastTimer);
    toastTimer = setTimeout(() => { S.toast.hidden = true; }, 3200);
  }
  function showOverlay(text) { S.overlayText.textContent = text || '正在处理…'; S.overlay.hidden = false; }
  function hideOverlay() { S.overlay.hidden = true; }

  /* ---------- 结构识别结果渲染 ---------- */
  const CHIP_DEFS = [
    ['front', '封面/前部', 'ok'],
    ['title', '论文题目', 'ok'],
    ['absHead', '摘要', 'ok'],
    ['absBody', '摘要正文', null],
    ['kw', '关键词', 'ok'],
    ['enTitle', '英文题目', null],
    ['absEnHead', 'Abstract', 'ok'],
    ['absEnBody', '英文正文', null],
    ['kwEn', 'Keywords', null],
    ['tocHead', '目录', 'ok'],
    ['tocItem', '目录条目', null],
    ['h1', '一级标题', null],
    ['h2', '二级标题', null],
    ['h3', '三级标题', null],
    ['body', '正文', null],
    ['caption', '图表题注', null],
    ['formula', '公式', null],
    ['figure', '插图', null],
    ['refHead', '参考文献', 'ok'],
    ['refItem', '文献条目', null],
    ['tables', '表格', null],
    ['images', '图片', null],
  ];

  function renderStats(counts, info) {
    const chips = [];
    CHIP_DEFS.forEach(([key, label, flag]) => {
      const v = counts[key];
      if (!v) return;
      const cls = flag === 'ok' ? ' ok' : '';
      chips.push(`<span class="chip${cls}">${label} <b>${v}</b></span>`);
    });
    S.chips.innerHTML = chips.join('');
    S.statsTime.textContent = '共 ' + counts.paras + ' 段文本';
    S.statsCard.hidden = false;

    const notes = [];
    if (!info.absDetected) notes.push('未检测到「摘要」章节，全文已按标题/正文规则统一排版（封面也将一并调整）。');
    if (!info.tocDetected && counts.h1 > 0) notes.push('未检测到「目录」，如需目录请在 Word 中手动生成后更新域。');
    if (info.tocDetected) notes.push('目录已按新格式重排，建议在 Word 中右键目录选择「更新域」以获得带页码的目录。');
    if (!info.titleDetected) notes.push('未识别到独立论文题目行（摘要页可能不含题目），题目格式未套用。');
    if (!info.chapterStyle) notes.push('章节采用「1 绪论」式编号，已按一级标题处理。');
    if (counts.footerAdded) notes.push('已添加页码（封面首页不显示）。');
    if (counts.frontMatter) notes.push('已套用模板前置页：封面与声明页共 ' + counts.frontMatter + ' 个元素照搬自模板，本文档原封面区已替换。');
    if (counts.frontIgnoredHf && counts.frontIgnoredHf.length) notes.push('模板前置页的' + counts.frontIgnoredHf.join('、') + '未搬运（封面与声明页不编页码、不印页眉）。');
    if (counts.formula > 0) notes.push('检测到 ' + counts.formula + ' 处公式：原样保留不改动，预览中可能显示不完整，以 Word 为准。');
    S.statsNotes.innerHTML = notes.map((n) => `<span class="note">${n}</span>`).join('');
  }

  /* ---------- 预览 ---------- */
  // 压缩预览图片：超过 1000px 的缩放到 1000px 内并转 JPEG，避免超大 HTML 导致预览卡死
  function compressPreviewImage(dataUrl) {
    return new Promise((resolve) => {
      const img = new Image();
      img.onload = () => {
        try {
          const MAXW = 1000;
          if (img.width <= MAXW && img.height <= MAXW) { resolve(dataUrl); return; }
          const scale = Math.min(1, MAXW / img.width);
          const canvas = document.createElement('canvas');
          canvas.width = Math.round(img.width * scale);
          canvas.height = Math.round(img.height * scale);
          const ctx = canvas.getContext('2d');
          ctx.fillStyle = '#ffffff';
          ctx.fillRect(0, 0, canvas.width, canvas.height);
          ctx.drawImage(img, 0, 0, canvas.width, canvas.height);
          resolve(canvas.toDataURL('image/jpeg', 0.82));
        } catch (e) { resolve(dataUrl); }
      };
      img.onerror = () => resolve(dataUrl);
      img.src = dataUrl;
    });
  }

  let previewGen = 0; // 预览代次：防止旧预览覆盖新预览

  // 生成一次灰色占位图（预览省略的图片）
  function makePlaceholder() {
    try {
      const canvas = document.createElement('canvas');
      canvas.width = 320; canvas.height = 160;
      const ctx = canvas.getContext('2d');
      ctx.fillStyle = '#e8ecf2';
      ctx.fillRect(0, 0, 320, 160);
      ctx.fillStyle = '#8a97a8';
      ctx.font = '20px "Microsoft YaHei", sans-serif';
      ctx.textAlign = 'center';
      ctx.fillText('预览已省略此图片', 160, 88);
      return canvas.toDataURL('image/png');
    } catch (e) { return ''; }
  }

  async function renderPreview(blob, gen) {
    if (typeof mammoth === 'undefined') {
      toast('预览组件缺失（mammoth 未加载），可直接下载文档查看', true);
      return;
    }
    S.previewStatus.hidden = false;
    S.previewStatus.textContent = '正在生成预览…（图片较多时可能需要数秒，不影响下载）';
    try {
      const placeholder = makePlaceholder();
      let imgCount = 0;
      const result = await mammoth.convertToHtml(
        { arrayBuffer: await blob.arrayBuffer() },
        {
          convertImage: mammoth.images.imgElement((image) => {
            imgCount++;
            if (imgCount > 12) {
              // 预览仅内嵌前 12 张图，其余占位（下载的 Word 中图片完整保留）
              return Promise.resolve({ src: placeholder, alt: '预览已省略图片（Word 中完整保留）' });
            }
            return image.read('base64').then((buf) =>
              compressPreviewImage('data:' + image.contentType + ';base64,' + buf)
                .then((src) => ({ src: src }))
            );
          })
        }
      );
      if (gen !== previewGen) return; // 已有更新的预览，丢弃本次结果
      S.previewStatus.hidden = true;
      const style = '<style>' +
        'body{font-family:"宋体","SimSun","Microsoft YaHei",serif;margin:48px auto;max-width:720px;padding:0 24px;color:#222;}' +
        'p{line-height:1.5;text-indent:2em;margin:.6em 0;text-align:justify;}' +
        'h1,h2,h3{font-family:"黑体","SimHei","Microsoft YaHei",sans-serif;}' +
        'img{max-width:100%;height:auto;display:block;margin:8px auto;}' +
        'table{border-collapse:collapse;margin:10px auto;}' +
        'td,th{border:1px solid #999;padding:4px 10px;}' +
        '</style>';
      S.previewFrame.srcdoc = '<!DOCTYPE html><html lang="zh-CN"><head><meta charset="utf-8">' + style + '</head><body>' + result.value + '</body></html>';
      S.previewFrame.hidden = false;
      S.previewEmpty.hidden = true;
      if (gen === previewGen) toast('预览已生成，可下载 Word 文档');
    } catch (e) {
      console.error(e);
      S.previewStatus.hidden = true;
      if (gen === previewGen) toast('预览生成失败（不影响下载）：' + e.message, true);
    }
  }

  /* ---------- 模板前置页（封面 / 声明页照搬） ----------
   * 模板解析成 { name, front, on } 存 localStorage —— 含图片 base64（校徽一张
   * 一百多 KB），所以存不下时要明确告诉用户「本次能用，下次要重选」，不能装作存住了。
   * 不分学科：一个学生的封面就一份，工科文科共用。 */
  const LS_TPL_KEY = 'thesis-formatter-front-template';
  let templateCache = null;

  function loadTemplate() {
    try { return JSON.parse(localStorage.getItem(LS_TPL_KEY)) || null; }
    catch (e) { return null; }
  }
  function saveTemplate() {
    try {
      if (templateCache) localStorage.setItem(LS_TPL_KEY, JSON.stringify(templateCache));
      else localStorage.removeItem(LS_TPL_KEY);
    } catch (e) {
      toast('模板较大，浏览器存不下：本次仍可用，下次打开需要重新选择', true);
    }
  }
  function renderTemplateStatus() {
    const f = templateCache && templateCache.front;
    S.chkTemplateFront.disabled = !f;
    S.btnClearTemplate.hidden = !f;
    if (!f) {
      S.chkTemplateFront.checked = false;
      S.templateStatus.textContent = '尚未选择模板文档';
      return;
    }
    S.chkTemplateFront.checked = !!templateCache.on;
    const img = (f.media || []).length;
    S.templateStatus.textContent = '已记住「' + templateCache.name + '」：前置页 ' + f.els +
      ' 个元素' + (img ? '、' + img + ' 张图片' : '') + '（到「' + (f.boundary || '摘要') + '」为止）';
  }
  function bindTemplateFront() {
    S.btnPickTemplate.addEventListener('click', () => S.fileTemplate.click());
    S.fileTemplate.addEventListener('change', async () => {
      const f = S.fileTemplate.files[0];
      S.fileTemplate.value = '';
      if (!f) return;
      if (!/\.docx$/i.test(f.name)) { toast('模板需为 Word 文档（.docx）', true); return; }
      showOverlay('正在读取模板「' + f.name + '」…');
      try {
        const front = await FormatTool.extractFrontMatter(await f.arrayBuffer());
        templateCache = { name: f.name, front: front, on: true };
        saveTemplate();
        renderTemplateStatus();
        hideOverlay();
        const ig = front.ignoredParts || [];
        toast('模板前置页已记住 ✔' + (ig.length
          ? '（模板的' + ig.join('、') + '有内容，前置页不搬页眉页脚，已忽略）' : ''));
      } catch (e) {
        hideOverlay();
        toast(e.message || '模板读取失败，可能不是标准 .docx', true);
      }
    });
    S.chkTemplateFront.addEventListener('change', () => {
      if (!templateCache) return;
      templateCache.on = S.chkTemplateFront.checked;
      saveTemplate();
      renderTemplateStatus();
    });
    S.btnClearTemplate.addEventListener('click', () => {
      templateCache = null;
      saveTemplate();
      renderTemplateStatus();
      toast('已清除模板前置页');
    });
  }

  /* ---------- 格式化主流程 ---------- */
  async function formatFile(file) {
    showOverlay('正在格式化「' + file.name + '」…');
    try {
      const arrayBuffer = await file.arrayBuffer();
      const settings = currentSettings();
      const opts = { format: 'blob' };
      if (templateCache && templateCache.on && templateCache.front) {
        opts.templateFront = templateCache.front;
      }
      const result = await FormatTool.formatDocx(arrayBuffer, settings, opts);
      lastBlob = result.data;
      lastFile = file;

      renderStats(result.counts, result.info);
      S.btnDownload.disabled = false;
      S.btnPreviewReformat.disabled = false;

      // 格式化已完成：立即解除遮罩，下载立即可用；预览后台异步生成
      hideOverlay();
      toast('格式化完成 ✔ 可立即下载，预览正在生成…');
      const gen = ++previewGen;
      renderPreview(result.data, gen); // 不阻塞：预览在后台完成
    } catch (e) {
      console.error(e);
      hideOverlay();
      toast(e.message || '处理失败：文件可能已损坏或不是标准 .docx', true);
    }
  }

  /* ---------- 文件选择 ---------- */
  function acceptFile(file) {
    if (!file) return;
    if (!/\.docx$/i.test(file.name)) {
      toast('仅支持 Word 文档（.docx 格式）', true);
      return;
    }
    lastFile = file;
    S.fileName.textContent = file.name;
    S.fileSize.textContent = ((file.size / 1024 / 1024).toFixed(2)) + ' MB';
    S.filebar.hidden = false;
    S.btnFormat.disabled = false;
    S.dropzone.hidden = true;
    formatFile(file);
  }

  function resetFileUI() {
    S.dropzone.hidden = false;
    S.filebar.hidden = true;
    S.btnFormat.disabled = true;
  }

  /* ---------- 下载 ---------- */
  function download() {
    if (!lastBlob || !lastFile) return;
    /* 文件名 = 原名 + 工具版本号（见 FormatTool.outputName） */
    const name = FormatTool.outputName(lastFile.name);
    const url = URL.createObjectURL(lastBlob);
    const a = document.createElement('a');
    a.href = url;
    a.download = name;
    document.body.appendChild(a);
    a.click();
    a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 3000);
    toast('已开始下载「' + name + '」');
  }

  /* ---------- 事件绑定 ---------- */
  function bind() {
    // 上传
    S.dropzone.addEventListener('click', () => S.fileInput.click());
    S.dropzone.addEventListener('keydown', (e) => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); S.fileInput.click(); } });
    S.fileInput.addEventListener('change', () => { acceptFile(S.fileInput.files[0]); S.fileInput.value = ''; });
    S.btnReupload.addEventListener('click', resetFileUI);
    ['dragover', 'dragenter'].forEach((ev) =>
      S.dropzone.addEventListener(ev, (e) => { e.preventDefault(); S.dropzone.classList.add('dragover'); }));
    ['dragleave', 'drop'].forEach((ev) =>
      S.dropzone.addEventListener(ev, (e) => { e.preventDefault(); S.dropzone.classList.remove('dragover'); }));
    S.dropzone.addEventListener('drop', (e) => {
      const f = e.dataTransfer && e.dataTransfer.files && e.dataTransfer.files[0];
      if (f) acceptFile(f);
    });

    // 动作
    S.btnFormat.addEventListener('click', () => { if (lastFile) formatFile(lastFile); });
    S.btnPreviewReformat.addEventListener('click', () => { if (lastFile) formatFile(lastFile); });
    S.btnDownload.addEventListener('click', download);
    S.btnReset.addEventListener('click', () => {
      const defId = defaultPresetId();
      const def = defId && presetById(defId);
      if (def) {
        applyPreset(def.id, { silent: true });
        toast('已恢复为「' + def.name + '」默认格式要求');
      } else {
        applySettings(FormatTool.DEFAULTS);
        saveSettings();
        toast('已恢复为通用默认设置');
      }
    });
    schoolSel.addEventListener('change', () => applyPreset(schoolSel.value));
    SETTINGS_FIELDS.forEach(([id]) => {
      $('#' + id).addEventListener('change', saveSettings);
    });
  }

  /* ---------- 启动 ----------
   * 每次进入先迁移旧版数据（并入工科空间），再显示学科选择页；
   * 选择后 enterMode 才加载该学科的设置与学校记忆 */
  migrateLegacy();
  bind();
  bindAddSchool();
  bindTemplateFront();
  bindGate();
  templateCache = loadTemplate();   // 记忆的模板与学科无关，进来就恢复
  renderTemplateStatus();
  updateModeBadge();
  showGate();
  if (typeof FormatTool === 'undefined') {
    toast('核心引擎加载失败，请检查 js/formatter.js 是否存在', true);
  } else {
    const badge = $('#brandVer');
    if (badge) badge.textContent = 'v' + FormatTool.VERSION;
  }
})();
