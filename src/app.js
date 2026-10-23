
/* ============================================================
   试题转换台 Online — 核心逻辑
   解析试题、生成 Excel 全部在浏览器本地完成；
   只有"AI 出题"这一步会联网调用 DeepSeek 接口。
   ============================================================ */

(function(){
  'use strict';

  // ============================================================
  // 内置浏览器（钉钉/微信/企业微信等）检测
  // 这类 App 的"文件预览"通常是一个只读 WebView，会拦截
  // URL.createObjectURL + <a download> 触发的下载动作，
  // 点击按钮后 JS 正常执行但用户看不到任何反应。
  // 这里做两层探测：UA 关键词精确识别 + 下载未触发的兜底提示。
  // ============================================================
  const IN_APP_UA_PATTERNS = [
    { key:'DingTalk', label:'钉钉' },
    { key:'MicroMessenger', label:'微信' },
    { key:'wxwork', label:'企业微信' },
    { key:'Lark', label:'飞书' },
    { key:'QQ/', label:'QQ' },
  ];

  function detectInAppBrowser(){
    const ua = navigator.userAgent || '';
    for(const p of IN_APP_UA_PATTERNS){
      if(ua.indexOf(p.key) !== -1) return p.label;
    }
    // 兜底：被塞进 iframe 里展示，也常见于各类"文件预览"场景
    try{
      if(window.self !== window.top) return '当前页面的预览窗口';
    }catch(e){
      // 跨域访问 window.top 会抛异常，本身也是被 iframe 包裹的信号
      return '当前页面的预览窗口';
    }
    return null;
  }

  const inappBanner = document.getElementById('inappBanner');
  const inappBannerTitle = document.getElementById('inappBannerTitle');
  const inappBannerDismiss = document.getElementById('inappBannerDismiss');
  const downloadFallback = document.getElementById('downloadFallback');
  const downloadFallbackLinks = document.getElementById('downloadFallbackLinks');

  const inAppLabel = detectInAppBrowser();
  if(inAppLabel){
    inappBannerTitle.textContent = `检测到当前可能在${inAppLabel}内置窗口中打开`;
    inappBanner.classList.add('show');
  }
  inappBannerDismiss.addEventListener('click', () => {
    inappBanner.classList.remove('show');
  });

  // ---------- 状态 ----------
  const state = {
    // 每项统一结构：{source:'docx'|'paste'|'ai', file, name, status, paragraphs, questions, error, rawText}
    files: [],
    parsed: false,
  };

  // ---------- DOM ----------
  const dropzone = document.getElementById('dropzone');
  const fileInput = document.getElementById('fileInput');
  const fileListEl = document.getElementById('fileList');
  const parseBtn = document.getElementById('parseBtn');
  const previewArea = document.getElementById('previewArea');
  const exportPanel = document.getElementById('exportPanel');
  const exportBtn = document.getElementById('exportBtn');
  const logArea = document.getElementById('logArea');
  const cfgToggle = document.getElementById('cfgToggle');
  const cfgBody = document.getElementById('cfgBody');
  const cfgResetBtn = document.getElementById('cfgResetBtn');
  const cfgSaveBtn = document.getElementById('cfgSaveBtn');

  const pasteArea = document.getElementById('pasteArea');
  const pasteNameInput = document.getElementById('pasteNameInput');
  const pasteAddBtn = document.getElementById('pasteAddBtn');
  const pasteListEl = document.getElementById('pasteList');

  const CFG_IDS = [
    'cfg_single_kw','cfg_multi_kw','cfg_tf_kw','cfg_ans_kw',
    'cfg_qnum_re','cfg_opt_re','cfg_tf_true','cfg_tf_false',
    'cfg_score_single','cfg_score_multi','cfg_score_tf','cfg_score_total'
  ];
  const CFG_DEFAULTS = {};
  CFG_IDS.forEach(id => { CFG_DEFAULTS[id] = document.getElementById(id).value; });

  const LOCAL_KEY = 'exam_tool_cfg_v1';

  function log(msg){
    logArea.style.display = 'block';
    const t = new Date().toLocaleTimeString('zh-CN', {hour12:false});
    logArea.textContent += `[${t}] ${msg}\n`;
    logArea.scrollTop = logArea.scrollHeight;
  }

  function getCfg(){
    const c = {};
    CFG_IDS.forEach(id => { c[id] = document.getElementById(id).value.trim(); });
    return c;
  }

  function splitKw(s){
    return s.split(',').map(x=>x.trim()).filter(Boolean);
  }

  // 改了分值，导出前体检里的"总分"要跟着刷新
  ['cfg_score_single','cfg_score_multi','cfg_score_tf'].forEach(id => {
    document.getElementById(id).addEventListener('input', () => { updateIssueSummary(); });
  });

  // ---------- 配置面板 折叠/保存/恢复 ----------
  cfgToggle.addEventListener('click', () => {
    cfgToggle.classList.toggle('open');
    cfgBody.classList.toggle('open');
  });

  cfgResetBtn.addEventListener('click', () => {
    CFG_IDS.forEach(id => { document.getElementById(id).value = CFG_DEFAULTS[id]; });
    log('已恢复默认解析规则');
  });

  cfgSaveBtn.addEventListener('click', () => {
    try{
      localStorage.setItem(LOCAL_KEY, JSON.stringify(getCfg()));
      log('已将当前规则保存为本地默认（仅保存在本机浏览器中）');
      cfgSaveBtn.textContent = '已保存 ✓';
      setTimeout(()=>{ cfgSaveBtn.textContent = '保存为本地默认（下次打开自动生效）'; }, 1800);
    }catch(e){
      log('保存失败：' + e.message);
    }
  });

  (function loadSavedCfg(){
    try{
      const raw = localStorage.getItem(LOCAL_KEY);
      if(!raw) return;
      const saved = JSON.parse(raw);
      CFG_IDS.forEach(id => { if(saved[id] !== undefined) document.getElementById(id).value = saved[id]; });
    }catch(e){ /* ignore */ }
  })();

  // ---------- 输入方式 Tab 切换（AI 出题 / 上传 .docx / 粘贴文字） ----------
  const TAB_PAIRS = [['tabBtnAi','paneAi'], ['tabBtnFile','paneFile'], ['tabBtnPaste','panePaste']];
  TAB_PAIRS.forEach(([btnId]) => {
    document.getElementById(btnId).addEventListener('click', () => {
      TAB_PAIRS.forEach(([b2, p2]) => {
        document.getElementById(b2).classList.toggle('active', b2 === btnId);
        document.getElementById(p2).classList.toggle('active', b2 === btnId);
      });
    });
  });

  // ---------- 文件拖拽/选择（.docx） ----------
  dropzone.addEventListener('click', () => fileInput.click());
  dropzone.addEventListener('dragover', e => { e.preventDefault(); dropzone.classList.add('drag'); });
  dropzone.addEventListener('dragleave', () => dropzone.classList.remove('drag'));
  dropzone.addEventListener('drop', e => {
    e.preventDefault();
    dropzone.classList.remove('drag');
    handleFiles(e.dataTransfer.files);
  });
  fileInput.addEventListener('change', e => {
    handleFiles(e.target.files);
    fileInput.value = ''; // 允许重复选择同名文件
  });

  function handleFiles(fileListObj){
    const arr = Array.from(fileListObj).filter(f => /\.docx$/i.test(f.name));
    if(arr.length === 0){
      log('未检测到 .docx 文件，请确认文件格式');
      return;
    }
    arr.forEach(f => {
      state.files.push({ source:'docx', file:f, name:f.name, status:'pending', paragraphs:null, questions:null, error:null });
    });
    renderFileList();
    parseBtn.disabled = state.files.length === 0;
  }

  // ---------- 粘贴文字 加入列表 ----------
  let pasteCounter = 0;
  pasteAddBtn.addEventListener('click', () => {
    const text = pasteArea.value;
    if(!text || !text.trim()){
      log('文本框是空的，先粘贴一些试题文字再加入列表');
      return;
    }
    pasteCounter++;
    const customName = pasteNameInput.value.trim();
    const name = (customName ? customName : `粘贴试题${pasteCounter > 1 ? pasteCounter : ''}`) + '.txt';
    state.files.push({
      source:'paste', file:null, name, status:'pending',
      paragraphs:null, questions:null, error:null, rawText:text
    });
    pasteArea.value = '';
    pasteNameInput.value = '';
    renderFileList();
    parseBtn.disabled = state.files.length === 0;
    log(`已将粘贴内容加入待解析列表：${name}`);
  });

  function renderFileList(){
    fileListEl.innerHTML = '';
    state.files.forEach((item, idx) => {
      const row = document.createElement('div');
      row.className = 'file-row';
      const statusText = {
        pending:'待解析', parsing:'解析中…', ok:'已识别', warn:'部分识别', error:'解析失败'
      }[item.status] || item.status;
      const statusCls = {
        pending:'pending', parsing:'pending', ok:'ok', warn:'warn', error:'warn'
      }[item.status] || 'pending';
      const sourceTag = item.source === 'paste' ? '✎ ' : (item.source === 'ai' ? '🤖 ' : '📄 ');

      row.innerHTML = `
        <span class="fname">${sourceTag}${escapeHtml(item.name)}</span>
        <span class="fstat ${statusCls}">${statusText}</span>
        <span class="fremove" title="移除">✕</span>
      `;
      row.querySelector('.fremove').addEventListener('click', () => {
        state.files.splice(idx,1);
        renderFileList();
        parseBtn.disabled = state.files.length === 0;
        if(state.files.length === 0){
          previewArea.style.display = 'none';
          exportPanel.style.display = 'none';
        }
      });
      fileListEl.appendChild(row);
    });
  }

  function escapeHtml(s){
    return String(s).replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
  }

  function stripExt(name){
    return String(name).replace(/\.(docx|txt)$/i, '');
  }

  // ============================================================
  // DOCX 解析：提取段落纯文本序列
  // ============================================================
  async function extractParagraphs(file){
    const buf = await file.arrayBuffer();
    const zip = await JSZip.loadAsync(buf);
    const docXmlFile = zip.file('word/document.xml');
    if(!docXmlFile){
      throw new Error('未找到 word/document.xml，可能不是有效的 .docx 文件');
    }
    const xmlText = await docXmlFile.async('string');
    const parser = new DOMParser();
    const doc = parser.parseFromString(xmlText, 'application/xml');

    const errNode = doc.querySelector('parsererror');
    if(errNode){
      throw new Error('文档 XML 解析失败');
    }

    // 命名空间处理：用通配写法拿所有 w:p 段落（不同 Word 导出可能带不同前缀声明）
    const paragraphNodes = doc.getElementsByTagNameNS('http://schemas.openxmlformats.org/wordprocessingml/2006/main', 'p');

    const paragraphs = [];
    for(let i=0;i<paragraphNodes.length;i++){
      const p = paragraphNodes[i];
      const tNodes = p.getElementsByTagNameNS('http://schemas.openxmlformats.org/wordprocessingml/2006/main', 't');
      let text = '';
      for(let j=0;j<tNodes.length;j++){
        text += tNodes[j].textContent;
      }
      // 处理表格内文字：也可能落在 w:tbl 内部的 w:p，getElementsByTagNameNS 已经能一并抓到，因为按 tag 全局查找
      text = text.trim();
      if(text.length > 0){
        paragraphs.push(text);
      }
    }

    // 兜底：如果没抓到任何段落（例如极端命名空间写法），退化为正则扫描全部 <w:t>
    if(paragraphs.length === 0){
      const re = /<w:t[^>]*>([\s\S]*?)<\/w:t>/g;
      let m;
      let buf2 = '';
      let lastEnd = 0;
      while((m = re.exec(xmlText)) !== null){
        buf2 += decodeXmlEntities(m[1]);
      }
      if(buf2.trim()) paragraphs.push(buf2.trim());
    }

    return paragraphs;
  }

  function decodeXmlEntities(s){
    return s.replace(/&lt;/g,'<').replace(/&gt;/g,'>').replace(/&quot;/g,'"').replace(/&apos;/g,"'").replace(/&amp;/g,'&');
  }

  // ============================================================
  // 粘贴文字解析：按行切成"段落"数组，和 docx 提取结果同构，
  // 后续 parseDocxQuestions 逻辑完全复用，无需另写一套规则。
  // ============================================================
  function extractParagraphsFromText(text){
    return text
      .split(/\r\n|\r|\n/)
      .map(line => line.trim())
      .filter(line => line.length > 0);
  }


  // ============================================================
  // 题目识别引擎
  // ============================================================
  /*
    策略：不依赖分节大标题一定存在。
    先尝试用分节关键词把整篇切成 单选/多选/判断/答案 几个区域（如果能找到）；
    找不到分节标题时，退化为"逐段扫描 + 特征识别"：
      - 一行匹配 题号正则 => 视为新题开始（题干）
      - 紧随其后若干行匹配 选项正则(A. / B. ...) => 选择题的选项
      - 判断题：没有选项行，题干后不远处会在参考答案区标注 正确/错误
      - 单选 vs 多选：根据参考答案区该题号对应的答案字母数（1个=单选，≥2个=多选），
        如果没有参考答案区可用，则根据"每空最多几个选项被标记"等启发式回退为单选。
  */

  function buildQuestionNumRe(pattern){
    try{
      return new RegExp(pattern);
    }catch(e){
      return new RegExp('^\\s*(\\d{1,3})[\\.、．\\)\\s]\\s*(.+)$');
    }
  }
  function buildOptionRe(pattern){
    try{
      return new RegExp(pattern);
    }catch(e){
      return new RegExp('^\\s*([A-H])[\\.、．\\)\\s]\\s*(.+)$');
    }
  }

  // 把"选项识别正则"（形如 ^\s*([A-H])[\.、．\)\s]\s*(.+)$）转成一个
  // 不锚定行首/行尾、可以在一行中间反复命中的版本，用于探测
  // "A. xxx B. xxx C. xxx D. xxx" 这种四个选项挤在同一行里的写法。
  // 只去掉首尾的 ^ 和 (.+)$（避免贪婪吞掉后面的其他选项），
  // 转换失败时退化为一个通用的 "字母 + 分隔符" 探测正则。
  function buildInlineOptionScanRe(pattern){
    try{
      let p = pattern.trim();
      if(p.startsWith('^')) p = p.slice(1);
      if(p.startsWith('\\s*')) p = p.slice(3);
      // 把末尾的 \s*(.+)$ 替换成非贪婪、不含后续选项标记的捕获
      p = p.replace(/\\s\*\(\.\+\)\$$/, '\\s*');
      return new RegExp(p, 'g');
    }catch(e){
      return /([A-H])[\.、．\)]\s*/g;
    }
  }

  // 检测一行文字中是否包含"多个选项挤在一起"的情况，若是则拆分成
  // [{letter, text}, ...]；若一行只有 0 或 1 个选项标记，返回 null，
  // 交给原有的逐行 optRe 匹配流程处理（保持原有行为不变）。
  function splitInlineOptions(line, markerRe){
    const positions = [];
    let m;
    markerRe.lastIndex = 0;
    while((m = markerRe.exec(line)) !== null){
      positions.push({ letter: m[1].toUpperCase(), markerStart: m.index, contentStart: m.index + m[0].length });
      if(m[0].length === 0) markerRe.lastIndex++; // 防止零宽匹配死循环
    }
    if(positions.length < 2) return null; // 不足两个选项标记，不算"同行多选项"

    const items = [];
    for(let i=0;i<positions.length;i++){
      const start = positions[i].contentStart;
      const end = (i+1 < positions.length) ? positions[i+1].markerStart : line.length;
      const text = line.slice(start, end).trim();
      if(text.length > 0){
        items.push({ letter: positions[i].letter, text });
      }
    }
    return items.length >= 2 ? items : null;
  }

  function findSectionRanges(paragraphs, cfg){
    const singleKw = splitKw(cfg.cfg_single_kw);
    const multiKw = splitKw(cfg.cfg_multi_kw);
    const tfKw = splitKw(cfg.cfg_tf_kw);
    const ansKw = splitKw(cfg.cfg_ans_kw);

    // 分节标题的判定：标题一般出现在行首附近（允许"一、""二、"这类前缀），
    // 且关键词命中位置在行的前 12 个字符内 —— 用以避免误伤题干正文中偶然出现同名词语。
    // 不再用整行长度做门槛，因为标题后常跟"（共10题，每题4分…）"等较长说明文字。
    function headKeywordHit(line, kwList){
      for(const kw of kwList){
        const idx = line.indexOf(kw);
        if(idx >= 0 && idx <= 12) return true;
      }
      return false;
    }

    const markers = []; // {index, type}
    paragraphs.forEach((p, idx) => {
      if(headKeywordHit(p, singleKw)) markers.push({index:idx, type:'single'});
      else if(headKeywordHit(p, multiKw)) markers.push({index:idx, type:'multi'});
      else if(headKeywordHit(p, tfKw)) markers.push({index:idx, type:'tf'});
      else if(ansKw.some(kw => p.trim() === kw || (p.startsWith(kw) && p.length <= kw.length + 6))) markers.push({index:idx, type:'answer'});
    });

    return markers;
  }

  // 解析"参考答案"区块文字。答案区内部通常又分"一、单选 / 二、多选 / 三、判断"三个子块，
  // 且每个子块的题号都从 1 开始计数，因此必须按子块分别建立 map，不能用全局题号做 key，
  // 否则单选题 1 号答案会被多选题 1 号答案覆盖。
  // 返回 { single:{num:ans}, multi:{num:ans}, tf:{num:ans}, flat:{num:ans}(仅在未检测到子块时使用) }
  function parseAnswerBlock(lines, cfg){
    // 答案区的小标题常写得比正文分节标题更短（如只写"判断""单选""多选"），
    // 所以这里在用户配置的关键词之外，再补一组更短的写法，避免小标题认不出来而串号。
    const singleKw = splitKw(cfg.cfg_single_kw).concat(['单项选择', '单项', '单选']);
    const multiKw = splitKw(cfg.cfg_multi_kw).concat(['多项选择', '多项', '多选']);
    const tfKw = splitKw(cfg.cfg_tf_kw).concat(['判断', '是非']);

    // 返回命中的关键词末尾位置（用于把小标题后面同一行的答案继续解析），没命中返回 -1
    function headKeywordEnd(line, kwList){
      for(const kw of kwList){
        const idx = line.indexOf(kw);
        if(idx >= 0 && idx <= 12) return idx + kw.length;
      }
      return -1;
    }

    const result = { single:{}, multi:{}, tf:{}, flat:{} };
    // 答案写法：B / ABD / A,B,D / A、B、D / √ × / 正确 错误 …
    const re = /(\d{1,3})\s*[\.、．\)）:：]\s*((?:[A-H][,，、]?)+|[×√✓✔✗✘]|正确|错误|对|错|是|否)/g;

    let currentType = null;
    let sawAnySubHeader = false;

    lines.forEach(rawLine => {
      let line = rawLine;
      let end = -1;
      let type = null;
      if((end = headKeywordEnd(line, singleKw)) >= 0) type = 'single';
      else if((end = headKeywordEnd(line, multiKw)) >= 0) type = 'multi';
      else if((end = headKeywordEnd(line, tfKw)) >= 0) type = 'tf';

      if(type){
        currentType = type;
        sawAnySubHeader = true;
        line = line.slice(end); // 小标题后面如果同一行还跟着答案，继续解析
      }

      let m;
      re.lastIndex = 0;
      while((m = re.exec(line)) !== null){
        let val = m[2].replace(/[,，、]/g, '');
        if(val === '✓' || val === '✔') val = '√';
        if(val === '✗' || val === '✘') val = '×';
        result.flat[m[1]] = val; // 兜底
        if(currentType){
          result[currentType][m[1]] = val;
        }
      }
    });

    result.hasSubHeaders = sawAnySubHeader;
    return result;
  }

  function normalizeTF(raw, cfg){
    const trueKw = splitKw(cfg.cfg_tf_true);
    const falseKw = splitKw(cfg.cfg_tf_false);
    const v = (raw||'').trim();
    if(trueKw.includes(v)) return '正确';
    if(falseKw.includes(v)) return '错误';
    // 单字符 √ × 兜底
    if(v === '√') return '正确';
    if(v === '×' || v === 'x' || v === 'X') return '错误';
    return null;
  }

  function parseDocxQuestions(paragraphs, cfg){
    const qnumRe = buildQuestionNumRe(cfg.cfg_qnum_re);
    const optRe = buildOptionRe(cfg.cfg_opt_re);
    const inlineOptScanRe = buildInlineOptionScanRe(cfg.cfg_opt_re);

    const sections = findSectionRanges(paragraphs, cfg);

    // 找到"参考答案"分节起点（如果有），把它之后的内容单独作为答案区解析
    let answerStart = -1;
    for(const m of sections){ if(m.type === 'answer'){ answerStart = m.index; break; } }
    let bodyEnd = answerStart >= 0 ? answerStart : paragraphs.length;
    let answerLines = answerStart >= 0 ? paragraphs.slice(answerStart+1) : [];
    let answerMaps = parseAnswerBlock(answerLines, cfg);

    // 主体区域：逐段扫描，识别题目
    const rawQuestions = []; // {num, stem, options:[{letter,text}], sectionType}
    let currentSectionType = null; // 'single' | 'multi' | 'tf' | null(未知，靠特征判断)
    let current = null;

    const sectionAt = {};
    sections.forEach(m => { if(m.type !== 'answer' && m.index < bodyEnd) sectionAt[m.index] = m.type; });

    for(let i=0;i<bodyEnd;i++){
      const line = paragraphs[i];

      if(sectionAt[i]){
        currentSectionType = sectionAt[i];
        if(current){ rawQuestions.push(current); current = null; }
        continue; // 分节标题本身不作为内容
      }

      const qm = line.match(qnumRe);
      if(qm){
        // 新题开始
        if(current) rawQuestions.push(current);
        current = {
          num: qm[1],
          stem: qm[2].trim(),
          options: [],
          sectionType: currentSectionType
        };
        continue;
      }

      // 先检查这一行是不是"多个选项挤在同一行"（如 "A. xxx B. xxx C. xxx D. xxx"）
      const inlineOpts = current ? splitInlineOptions(line, inlineOptScanRe) : null;
      if(inlineOpts && current){
        inlineOpts.forEach(o => current.options.push(o));
        continue;
      }

      const om = current ? line.match(optRe) : null;
      if(om && current){
        current.options.push({ letter: om[1].toUpperCase(), text: om[2].trim() });
        continue;
      }

      // 既不是题号行也不是选项行：如果当前有正在构建的题目且尚无选项，
      // 很可能是题干换行延续，拼接到题干末尾
      if(current && current.options.length === 0){
        current.stem += line;
      }
      // 如果已经有选项了，忽略无法识别的行（可能是空行或说明文字）
    }
    if(current) rawQuestions.push(current);

    // 归类：优先信任题目被采集时所在的分节类型（sectionType）；
    // 只有在完全没有分节标题信息时，才退化为按选项数量/答案形态猜测。
    const result = { single:[], multi:[], tf:[], unclassified:[] };

    // 题干括号里直接写了答案，如"……（ B ）"或"……(ABD)"：抽出答案，并把括号清空
    function takeInlineAnswer(q){
      const m = q.stem.match(/[（(]\s*([A-H]{1,8}|[√×]|正确|错误)\s*[）)]/);
      if(!m) return null;
      q.stem = q.stem.replace(m[0], '（  ）');
      return m[1];
    }

    function lookupAnswer(q){
      let hit = null;
      if(!q.sectionType) hit = answerMaps.flat[q.num] || null;
      else{
        const byType = answerMaps[q.sectionType];
        if(byType && byType[q.num] !== undefined) hit = byType[q.num];
        // 子块答案区缺失该题号时，退化用 flat（可能不准，但聊胜于无）
        else hit = answerMaps.flat[q.num] || null;
      }
      return hit || takeInlineAnswer(q);
    }

    rawQuestions.forEach(q => {
      const ansRaw = lookupAnswer(q);
      const looksLikeTF = (q.options.length === 0);

      if(q.sectionType === 'tf' || (looksLikeTF && q.sectionType !== 'single' && q.sectionType !== 'multi')){
        const tfVal = normalizeTF(ansRaw, cfg);
        result.tf.push({
          num:q.num, stem:q.stem,
          answer: tfVal, answerRaw: ansRaw,
          flagged: tfVal === null
        });
        return;
      }

      if(q.options.length === 0){
        result.unclassified.push({ num:q.num, stem:q.stem, reason:'未检测到选项，且不匹配判断题特征' });
        return;
      }

      let isMulti;
      if(q.sectionType === 'single') isMulti = false;
      else if(q.sectionType === 'multi') isMulti = true;
      else if(ansRaw && /^[A-H]{2,}$/.test(ansRaw)) isMulti = true;
      else if(ansRaw && /^[A-H]$/.test(ansRaw)) isMulti = false;
      else isMulti = false; // 缺省按单选处理，用户可在预览中核对

      const target = isMulti ? result.multi : result.single;
      target.push({
        num:q.num,
        stem:q.stem,
        options:q.options,
        answer: ansRaw ? ansRaw.toUpperCase() : null,
        flagged: !ansRaw
      });
    });

    return result;
  }

  // ============================================================
  // 解析按钮
  // ============================================================
  parseBtn.addEventListener('click', async () => {
    if(state.files.length === 0) return;
    parseBtn.disabled = true;
    previewArea.style.display = 'block';
    previewArea.innerHTML = '';
    exportPanel.style.display = 'none';
    logArea.textContent = '';
    logArea.style.display = 'block';

    const cfg = getCfg();

    for(const item of state.files){
      if(item.source === 'ai') continue; // AI 生成的题已经是成品，不需要解析，也避免覆盖你的修改
      item.status = 'parsing';
      renderFileList();
      try{
        log(`开始解析：${item.name}`);
        const paragraphs = item.source === 'paste'
          ? extractParagraphsFromText(item.rawText)
          : await extractParagraphs(item.file);
        log(`  提取到 ${paragraphs.length} 个文本段落`);
        const questions = parseDocxQuestions(paragraphs, cfg);
        item.paragraphs = paragraphs;
        item.questions = questions;

        const total = questions.single.length + questions.multi.length + questions.tf.length;
        const flagged = [...questions.single, ...questions.multi, ...questions.tf].filter(q=>q.flagged).length;
        log(`  识别：单选 ${questions.single.length} 题 / 多选 ${questions.multi.length} 题 / 判断 ${questions.tf.length} 题` +
            (questions.unclassified.length ? ` / 未分类 ${questions.unclassified.length} 题` : ''));
        if(flagged > 0){
          log(`  ⚠ 其中 ${flagged} 题未能识别到明确答案，请在下方预览中核对`);
        }

        item.status = (flagged > 0 || questions.unclassified.length > 0) ? 'warn' : 'ok';
      }catch(e){
        item.status = 'error';
        item.error = e.message;
        log(`  ✕ 解析失败：${e.message}`);
      }
      renderFileList();
    }

    renderPreview();
    state.parsed = true;
    parseBtn.disabled = false;

    const anyOk = state.files.some(f => f.status === 'ok' || f.status === 'warn');
    exportPanel.style.display = anyOk ? 'block' : 'none';
  });

  // ============================================================
  // 预览 + 在线编辑
  // 每道题都可以直接改题干 / 选项 / 答案，也可以删除或新增，
  // 改动实时写回内存里的题目对象，导出时用的就是改完之后的内容。
  // ============================================================

  let uidCounter = 0;
  const LETTERS = 'ABCDEFGH';
  const MAX_OPT = { single:6, multi:8 };   // 与导入模板的列数一致（单选 A-F，多选 A-H）
  const KIND_NAME = { single:'单选题', multi:'多选题', tf:'判断题' };

  // 检查一道题有没有问题，返回问题文字数组（空数组 = 没问题）
  function validateQuestion(q, kind){
    const issues = [];
    if(!q.stem || !q.stem.trim()) issues.push('题干为空');
    if(kind === 'tf'){
      if(q.answer !== '正确' && q.answer !== '错误') issues.push('未选择答案');
      return issues;
    }
    const filled = new Set(q.options.filter(o => o.text && o.text.trim()).map(o => o.letter));
    if(filled.size < 2) issues.push('有效选项少于 2 个');
    const ans = (q.answer || '').split('').filter(Boolean);
    if(ans.length === 0) issues.push('未选择答案');
    else{
      if(kind === 'single' && ans.length !== 1) issues.push('单选题只能有 1 个答案');
      if(kind === 'multi' && ans.length < 2) issues.push('多选题至少要有 2 个答案');
      if(ans.some(l => !filled.has(l))) issues.push('答案对应的选项是空的');
    }
    return issues;
  }

  // 按"每题分值 × 题数"算总分，不再用写死的 100
  function calcScore(q, cfg){
    const s = Number(cfg.cfg_score_single) || 0;
    const m = Number(cfg.cfg_score_multi) || 0;
    const t = Number(cfg.cfg_score_tf) || 0;
    return q.single.length * s + q.multi.length * m + q.tf.length * t;
  }

  // 导出前体检：汇总所有试卷里的问题，显示在导出面板上方
  function updateIssueSummary(){
    const box = document.getElementById('issueSummary');
    if(!box) return;
    const cfg = getCfg();
    const lines = [];
    let totalQ = 0, verifyN = 0;
    state.files.forEach(item => {
      if(!item.questions) return;
      const q = item.questions;
      totalQ += q.single.length + q.multi.length + q.tf.length;
      ['single','multi','tf'].forEach(kind => {
        q[kind].forEach((qq, i) => {
          const iss = validateQuestion(qq, kind);
          if(iss.length) lines.push(`《${item.name.replace(/\.[^.]+$/,'')}》${KIND_NAME[kind]}第 ${i+1} 题：${iss.join('；')}`);
          if(qq.verify && !qq.verified) verifyN++;
        });
      });
      if(item.expect){
        ['single','multi','tf'].forEach(kind => {
          if(q[kind].length !== item.expect[kind]){
            lines.push(`《${item.name.replace(/\.[^.]+$/,'')}》${KIND_NAME[kind]}应有 ${item.expect[kind]} 题，现在是 ${q[kind].length} 题`);
          }
        });
      }
    });
    if(totalQ === 0){ box.style.display = 'none'; return; }
    box.style.display = 'block';
    box.className = lines.length ? 'issue-box bad' : 'issue-box good';
    const scoreTxt = state.files.filter(f => f.questions).map(f => `${f.name.replace(/\.[^.]+$/,'')}：总分 ${calcScore(f.questions, cfg)}`).join('；');
    let html = '';
    if(lines.length){
      html += `<b>⚠ 发现 ${lines.length} 处需要处理（导出前建议先改好）：</b><ul>` +
        lines.slice(0, 12).map(l => `<li>${escapeHtml(l)}</li>`).join('') +
        (lines.length > 12 ? `<li>……另有 ${lines.length - 12} 处</li>` : '') + '</ul>';
    }else{
      html += `<b>✓ 检查通过：</b>共 ${totalQ} 题，答案格式都正确。`;
    }
    if(verifyN) html += `<div class="verify-note">🔶 还有 ${verifyN} 题标了"需核实"：涉及具体数值或条款，请对照团标原文确认；确认后点题目右上角的"需核实"标签即可取消。</div>`;
    html += `<div class="score-note">${escapeHtml(scoreTxt)}</div>`;
    box.innerHTML = html;
  }

  function newBlankQuestion(kind){
    if(kind === 'tf') return { num:'', stem:'', answer:null, flagged:true };
    const opts = [];
    for(let i=0;i<4;i++) opts.push({ letter:LETTERS[i], text:'' });
    return { num:'', stem:'', options:opts, answer:null, flagged:true };
  }

  function renderPreview(){
    previewArea.innerHTML = '';
    state.files.forEach(item => {
      if(item.status === 'error'){
        const box = document.createElement('div');
        box.className = 'qtype-block';
        box.innerHTML = `<div class="qtype-head" style="color:#a5471f;">${escapeHtml(item.name)} — 解析失败：${escapeHtml(item.error||'')}</div>`;
        previewArea.appendChild(box);
        return;
      }
      if(!item.questions) return;

      const wrap = document.createElement('div');
      wrap.style.marginBottom = '28px';

      const titleRow = document.createElement('div');
      titleRow.className = 'paper-title-row';
      const title = document.createElement('div');
      title.className = 'paper-title';
      title.textContent = (item.source === 'ai' ? '🤖 ' : '📄 ') + item.name.replace(/\.[^.]+$/,'');
      titleRow.appendChild(title);
      const rm = document.createElement('button');
      rm.className = 'btn-ghost btn-small';
      rm.textContent = '移除此试卷';
      rm.addEventListener('click', () => {
        const i = state.files.indexOf(item);
        if(i >= 0) state.files.splice(i, 1);
        renderFileList();
        parseBtn.disabled = state.files.length === 0;
        renderPreview();
        const any = state.files.some(f => f.questions);
        exportPanel.style.display = any ? 'block' : 'none';
        if(!any) previewArea.style.display = 'none';
      });
      titleRow.appendChild(rm);
      wrap.appendChild(titleRow);

      const summary = document.createElement('div');
      summary.className = 'preview-summary';
      const q = item.questions;
      summary.innerHTML = [
        chip('单选题', q.single.length),
        chip('多选题', q.multi.length),
        chip('判断题', q.tf.length),
        chip('未分类', q.unclassified.length, true),
      ].join('');
      wrap.appendChild(summary);

      if(item.source === 'ai'){
        const tip = document.createElement('div');
        tip.className = 'ai-tip';
        tip.textContent = '以下题目由 AI 生成，请逐题核对。题干、选项、答案都可以直接修改，改完即生效。';
        wrap.appendChild(tip);
      }

      wrap.appendChild(renderQtypeBlock('单项选择题', q.single, 'single'));
      wrap.appendChild(renderQtypeBlock('多项选择题', q.multi, 'multi'));
      wrap.appendChild(renderQtypeBlock('判断题', q.tf, 'tf'));
      if(q.unclassified.length){
        wrap.appendChild(renderQtypeBlock('未能分类（不会导出，可自行核对原文件）', q.unclassified, 'unclassified'));
      }

      previewArea.appendChild(wrap);
    });
    updateIssueSummary();
  }

  function chip(label, n, isUnclassified){
    const cls = n === 0 ? 'count-chip zero' : 'count-chip';
    return `<span class="${cls}">${label} <b>${n}</b></span>`;
  }

  function renderQtypeBlock(title, arr, kind){
    const block = document.createElement('div');
    block.className = 'qtype-block';
    const head = document.createElement('div');
    head.className = 'qtype-head';
    const left = document.createElement('span');
    left.textContent = `${title}　${arr.length} 题`;
    head.appendChild(left);
    if(kind !== 'unclassified'){
      const add = document.createElement('button');
      add.className = 'btn-ghost btn-small';
      add.textContent = '＋ 添加一题';
      add.addEventListener('click', () => {
        arr.push(newBlankQuestion(kind));
        renderPreview();
      });
      head.appendChild(add);
    }
    block.appendChild(head);

    const body = document.createElement('div');
    body.className = 'qtype-body';
    if(arr.length === 0){
      body.innerHTML = `<div class="qrow" style="color:#b3ab98;">（无）</div>`;
    }else{
      arr.forEach((q, idx) => {
        if(kind === 'unclassified') body.appendChild(renderUnclassifiedRow(q));
        else body.appendChild(renderEditableRow(q, idx, arr, kind));
      });
    }
    block.appendChild(body);
    return block;
  }

  // 一道题的可编辑行（单选 / 多选 / 判断共用）
  function renderEditableRow(q, idx, arr, kind){
    const uid = 'q' + (++uidCounter);
    const row = document.createElement('div');
    row.className = 'qrow qedit';

    const top = document.createElement('div');
    top.className = 'qline';
    const no = document.createElement('span');
    no.className = 'qno';
    no.textContent = (idx + 1) + '.';
    top.appendChild(no);

    const stem = document.createElement('textarea');
    stem.className = 'qstem-edit';
    stem.rows = 2;
    stem.value = q.stem || '';
    stem.placeholder = '题干';
    stem.addEventListener('input', () => { q.stem = stem.value; refresh(); });
    top.appendChild(stem);

    const side = document.createElement('div');
    side.className = 'qside';
    if(q.verify){
      const vb = document.createElement('span');
      vb.className = 'verify-badge' + (q.verified ? ' done' : '');
      vb.textContent = q.verified ? '已核实' : '需核实';
      vb.title = '涉及具体数值或条款，请对照团标原文确认；点击切换';
      vb.addEventListener('click', () => {
        q.verified = !q.verified;
        vb.textContent = q.verified ? '已核实' : '需核实';
        vb.classList.toggle('done', !!q.verified);
        updateIssueSummary();
      });
      side.appendChild(vb);
    }
    const del = document.createElement('button');
    del.className = 'btn-ghost btn-small qdel';
    del.textContent = '删除';
    del.addEventListener('click', () => {
      arr.splice(idx, 1);
      renderPreview();
    });
    side.appendChild(del);
    top.appendChild(side);
    row.appendChild(top);

    const optsBox = document.createElement('div');
    optsBox.className = 'qopts';
    row.appendChild(optsBox);

    const issueEl = document.createElement('div');
    issueEl.className = 'qissue';
    row.appendChild(issueEl);

    function refresh(){
      const iss = validateQuestion(q, kind);
      q.flagged = iss.length > 0;
      row.classList.toggle('flag', q.flagged);
      issueEl.textContent = iss.length ? '⚠ ' + iss.join('；') : '';
      updateIssueSummary();
    }

    if(kind === 'tf'){
      const sel = document.createElement('select');
      sel.className = 'tf-select';
      [['', '— 选择答案 —'], ['正确', '正确'], ['错误', '错误']].forEach(([v, t]) => {
        const o = document.createElement('option');
        o.value = v; o.textContent = t;
        sel.appendChild(o);
      });
      sel.value = q.answer || '';
      sel.addEventListener('change', () => { q.answer = sel.value || null; refresh(); });
      optsBox.appendChild(sel);
    }else{
      // 选项不足 4 个时补空行，方便直接填写
      while(q.options.length < 4){
        const used = new Set(q.options.map(o => o.letter));
        const next = LETTERS.split('').find(l => !used.has(l));
        q.options.push({ letter:next, text:'' });
      }
      const drawOptions = () => {
        optsBox.innerHTML = '';
        const chosen = new Set((q.answer || '').split('').filter(Boolean));
        q.options.forEach(o => {
          const line = document.createElement('label');
          line.className = 'opt-row';
          const pick = document.createElement('input');
          pick.type = kind === 'multi' ? 'checkbox' : 'radio';
          pick.name = uid;
          pick.checked = chosen.has(o.letter);
          pick.title = '勾选 = 正确答案';
          pick.addEventListener('change', () => {
            const cur = new Set((q.answer || '').split('').filter(Boolean));
            if(kind === 'multi'){
              if(pick.checked) cur.add(o.letter); else cur.delete(o.letter);
            }else{
              cur.clear(); cur.add(o.letter);
            }
            q.answer = Array.from(cur).sort().join('') || null;
            refresh();
          });
          const lt = document.createElement('span');
          lt.className = 'opt-letter';
          lt.textContent = o.letter;
          const tx = document.createElement('input');
          tx.type = 'text';
          tx.className = 'opt-text';
          tx.value = o.text || '';
          tx.placeholder = '选项内容';
          tx.addEventListener('input', () => { o.text = tx.value; refresh(); });
          line.appendChild(pick); line.appendChild(lt); line.appendChild(tx);
          optsBox.appendChild(line);
        });
        if(q.options.length < MAX_OPT[kind]){
          const more = document.createElement('button');
          more.className = 'btn-ghost btn-small';
          more.textContent = '＋ 增加选项';
          more.addEventListener('click', () => {
            const used = new Set(q.options.map(o => o.letter));
            const next = LETTERS.split('').find(l => !used.has(l));
            q.options.push({ letter:next, text:'' });
            drawOptions();
          });
          optsBox.appendChild(more);
        }
      };
      drawOptions();
    }

    refresh();
    return row;
  }

  function renderUnclassifiedRow(q){
    const row = document.createElement('div');
    row.className = 'qrow flag';
    row.innerHTML = `
      <div class="qstem">${escapeHtml(q.num)}. ${escapeHtml(q.stem)}</div>
      <div class="qmeta"><span>${escapeHtml(q.reason)}</span></div>
    `;
    return row;
  }

  // ============================================================
  // AI 出题（接入 DeepSeek，这一步需要联网）
  // 你输入一个主题（可选：再贴上团标/教材原文），点一下就生成
  // 10 道单选 + 5 道多选 + 5 道判断，进入下方的预览编辑区，
  // 核对无误后再导出 Excel。
  // API Key 只由你在页面里填写，仅保存在这台设备的浏览器中，
  // 不会写进代码，也不会上传到仓库。
  // ============================================================

  const AI_STORE = 'exam_tool_ai_cfg_v1';
  const aiTopic = document.getElementById('aiTopic');
  const aiRef = document.getElementById('aiRef');
  const aiNSingle = document.getElementById('aiNSingle');
  const aiNMulti = document.getElementById('aiNMulti');
  const aiNTf = document.getElementById('aiNTf');
  const aiKey = document.getElementById('aiKey');
  const aiBase = document.getElementById('aiBase');
  const aiModel = document.getElementById('aiModel');
  const aiRemember = document.getElementById('aiRemember');
  const aiGenBtn = document.getElementById('aiGenBtn');
  const aiTestBtn = document.getElementById('aiTestBtn');
  const aiClearBtn = document.getElementById('aiClearBtn');
  const aiStatus = document.getElementById('aiStatus');
  const aiCfgToggle = document.getElementById('aiCfgToggle');
  const aiCfgBody = document.getElementById('aiCfgBody');

  aiCfgToggle.addEventListener('click', () => {
    aiCfgToggle.classList.toggle('open');
    aiCfgBody.classList.toggle('open');
  });

  (function loadAiCfg(){
    try{
      const raw = localStorage.getItem(AI_STORE);
      if(!raw) return;
      const c = JSON.parse(raw);
      if(c.key) aiKey.value = c.key;
      if(c.base) aiBase.value = c.base;
      if(c.model) aiModel.value = c.model;
      if(c.topic) aiTopic.value = c.topic;
      aiRemember.checked = c.remember !== false;
    }catch(e){ /* ignore */ }
  })();

  function saveAiCfg(){
    try{
      if(aiRemember.checked){
        localStorage.setItem(AI_STORE, JSON.stringify({
          key: aiKey.value.trim(), base: aiBase.value.trim(), model: aiModel.value.trim(),
          topic: aiTopic.value.trim(), remember: true
        }));
      }else{
        localStorage.setItem(AI_STORE, JSON.stringify({
          base: aiBase.value.trim(), model: aiModel.value.trim(), topic: aiTopic.value.trim(), remember: false
        }));
      }
    }catch(e){ /* ignore */ }
  }

  aiClearBtn.addEventListener('click', () => {
    aiKey.value = '';
    try{ localStorage.removeItem(AI_STORE); }catch(e){}
    setAiStatus('已清除本机保存的 Key 和设置。', 'ok');
  });

  function setAiStatus(msg, kind){
    aiStatus.className = 'ai-status' + (kind ? ' ' + kind : '');
    aiStatus.textContent = msg;
    aiStatus.style.display = msg ? 'block' : 'none';
  }

  function aiSettings(){
    return {
      key: aiKey.value.trim(),
      base: (aiBase.value.trim() || 'https://api.deepseek.com').replace(/\/+$/, ''),
      model: aiModel.value.trim() || 'deepseek-chat',
    };
  }

  // 把各种失败原因翻译成能看懂的话
  function explainAiError(e, status, bodyText){
    if(e && e.name === 'AbortError') return '等待超过 2.5 分钟，已取消。可以稍后重试，或把题量调小一点。';
    if(status === 401 || status === 403) return 'API Key 无效或没有权限（' + status + '）。请检查 Key 是否填对、是否被停用。';
    if(status === 402) return '账户余额不足（402）。请到 DeepSeek 开放平台充值后再试。';
    if(status === 429) return '请求太频繁或额度被限流（429）。等一两分钟再试。';
    if(status === 400 || status === 422) return '请求被拒绝（' + status + '）：' + (bodyText || '').slice(0, 160);
    if(status >= 500) return 'DeepSeek 服务暂时不可用（' + status + '），稍后重试。';
    if(status) return '调用失败（' + status + '）：' + (bodyText || '').slice(0, 160);
    return '连接失败：浏览器没能连上 ' + aiSettings().base + '。可能是网络不通，也可能是该接口不允许网页直接访问（跨域限制）。' +
      '如果换网络仍然失败，请告诉我，我改用中转方式。';
  }

  aiTestBtn.addEventListener('click', async () => {
    const st = aiSettings();
    if(!st.key){ setAiStatus('请先填写 API Key。', 'bad'); return; }
    setAiStatus('正在测试连接…');
    let status = 0, text = '';
    try{
      const res = await fetch(st.base + '/models', { headers:{ 'Authorization':'Bearer ' + st.key } });
      status = res.status; text = await res.text();
      if(res.ok){ saveAiCfg(); setAiStatus('✓ 连接成功，Key 有效。', 'ok'); return; }
    }catch(e){
      setAiStatus('✕ ' + explainAiError(e, 0, ''), 'bad'); return;
    }
    setAiStatus('✕ ' + explainAiError(null, status, text), 'bad');
  });

  function buildAiPrompt(topic, ref, ns, nm, nt){
    const refBlock = ref
      ? `【参考资料】\n${ref}\n\n出题必须严格依据上面的参考资料，资料里没有的内容不要出题；每题的正确答案要能在资料中找到依据。`
      : `没有提供参考资料。请依据国内公认的临床护理规范和通用知识出题，不要编造具体的标准条款编号或页码。`;
    return `你是医院护理培训考试的命题专家。请围绕下面的主题出一套考试题。

【主题】${topic}

${refBlock}

【题量】单选题 ${ns} 道、多选题 ${nm} 道、判断题 ${nt} 道。题目均为基础知识点。

【命题要求】
1. 覆盖面：题目要覆盖该主题的不同方面（例如：评估与适应证、操作要点、观察与记录、并发症与处理、健康宣教、感染防控、应急处置），不要集中在同一个知识点上，不要出现重复或近似的题。
2. 单选题：4 个选项，只有 1 个正确。错误选项要"似是而非"（常见的错误做法或相近概念），不能是一眼就能排除的荒谬选项；4 个选项长度和句式尽量接近，正确答案不要总是最长的；不要使用"以上都对""以上都不对""A 和 B"这类写法。
3. 多选题：4 或 5 个选项，正确项数在 2~4 之间，各题的正确项数要有变化（不能全是 3 个），其中至少有 2 题只有 2 个正确项；错误选项同样要似是而非。
4. 判断题：陈述句，对或错必须明确，不使用模糊措辞；错误的陈述应是常见的错误认知，不要只靠加"绝对""随意""禁止"之类的词来制造错误；对和错的数量大致均衡。
5. 所有题干的末尾都用中文括号"（  ）"留出作答位置。
6. verify 字段：如果这道题的正确答案取决于具体的数值、时间、角度、深度、频次、距离或具体条款，需要对照标准原文才能确认，则设为 true，否则设为 false。
7. 语言简洁、规范，每题自洽。

【输出格式】只输出一个 JSON 对象，不要输出任何解释，也不要使用 markdown 代码块。结构如下：
{
  "single": [ {"stem": "题干（  ）", "options": ["选项内容1", "选项内容2", "选项内容3", "选项内容4"], "answer": "B", "verify": false} ],
  "multi":  [ {"stem": "题干（  ）", "options": ["…", "…", "…", "…"], "answer": "ACD", "verify": false} ],
  "tf":     [ {"stem": "陈述句（  ）", "answer": "正确", "verify": false} ]
}
说明：options 里只写选项内容，不要带"A."这样的字母前缀；answer 对单选和多选写选项字母，对判断题写"正确"或"错误"；single 必须正好 ${ns} 个，multi 必须正好 ${nm} 个，tf 必须正好 ${nt} 个。`;
  }

  function shuffleInPlace(a){
    for(let i=a.length-1;i>0;i--){
      const j = Math.floor(Math.random() * (i+1));
      [a[i], a[j]] = [a[j], a[i]];
    }
    return a;
  }

  // 选项里带"以上""A和B"这类互相引用的写法时，不能打乱顺序
  function hasPositionalRef(opts){
    return opts.some(t => /以上|上述|全部|都对|都不|[A-D]\s*[和、与及,，]\s*[A-D]|均正确|均错误/.test(t));
  }

  function fixStem(stem, isTf){
    let t = String(stem || '').trim();
    if(!/[（(]\s*[）)]/.test(t)){
      if(!isTf) t = t.replace(/[？?：:]\s*$/, '');
      t += '（  ）';
    }
    return t;
  }

  function cleanOpt(s){
    return String(s == null ? '' : s).trim().replace(/^[A-Ha-h][\.、．\)）:：]\s*/, '');
  }

  function letterIdxs(ans, n){
    return String(ans == null ? '' : ans).toUpperCase().replace(/[^A-H]/g, '').split('')
      .map(l => LETTERS.indexOf(l)).filter(i => i >= 0 && i < n);
  }

  // 把模型返回的 JSON 整理成工具内部的题目结构，并把答案位置打散
  function normalizeAi(obj){
    const warn = [];
    const out = { single:[], multi:[], tf:[], unclassified:[] };

    // ---- 单选 ----
    const singles = [];
    (Array.isArray(obj.single) ? obj.single : []).forEach((x, i) => {
      const opts = (Array.isArray(x.options) ? x.options : []).map(cleanOpt).filter(Boolean).slice(0, 6);
      const idxs = letterIdxs(x.answer, opts.length);
      if(opts.length < 2 || idxs.length !== 1 || !String(x.stem || '').trim()){
        warn.push(`单选第 ${i+1} 题格式不完整，已跳过`);
        return;
      }
      singles.push({ stem: fixStem(x.stem, false), opts, ci: idxs[0], verify: !!x.verify });
    });
    // 让正确答案的位置均匀分布在 A/B/C/D
    const targets = [];
    for(let i=0;i<singles.length;i++) targets.push(i % 4);
    shuffleInPlace(targets);
    singles.forEach((q, i) => {
      const t = targets[i];
      if(!hasPositionalRef(q.opts) && t < q.opts.length && t !== q.ci){
        [q.opts[t], q.opts[q.ci]] = [q.opts[q.ci], q.opts[t]];
        q.ci = t;
      }
      out.single.push({
        num:String(i+1), stem:q.stem,
        options: q.opts.map((t2, k) => ({ letter:LETTERS[k], text:t2 })),
        answer: LETTERS[q.ci], flagged:false, verify:q.verify
      });
    });

    // ---- 多选 ----
    (Array.isArray(obj.multi) ? obj.multi : []).forEach((x, i) => {
      let opts = (Array.isArray(x.options) ? x.options : []).map(cleanOpt);
      const idxs = letterIdxs(x.answer, opts.length);
      const okOpts = opts.filter(Boolean);
      if(okOpts.length < 3 || idxs.length < 2 || !String(x.stem || '').trim()){
        warn.push(`多选第 ${i+1} 题格式不完整，已跳过`);
        return;
      }
      opts = opts.slice(0, 8);
      let order = opts.map((_, k) => k);
      if(!hasPositionalRef(opts)) shuffleInPlace(order);
      const answerSet = new Set(idxs);
      const newOpts = order.map(k => opts[k]);
      const newAns = order.map((k, pos) => answerSet.has(k) ? LETTERS[pos] : '').join('');
      out.multi.push({
        num:String(out.multi.length+1), stem: fixStem(x.stem, false),
        options: newOpts.map((t2, k) => ({ letter:LETTERS[k], text:t2 })).filter(o => o.text),
        answer: newAns, flagged:false, verify: !!x.verify
      });
    });

    // ---- 判断 ----
    (Array.isArray(obj.tf) ? obj.tf : []).forEach((x, i) => {
      const a = String(x.answer == null ? '' : x.answer).trim().toLowerCase();
      let v = null;
      if(['正确','对','√','true','是','t'].includes(a)) v = '正确';
      else if(['错误','错','×','x','false','否','f'].includes(a)) v = '错误';
      if(!v || !String(x.stem || '').trim()){
        warn.push(`判断第 ${i+1} 题格式不完整，已跳过`);
        return;
      }
      out.tf.push({ num:String(out.tf.length+1), stem: fixStem(x.stem, true), answer:v, flagged:false, verify: !!x.verify });
    });

    // 多选题里，选项被打乱后字母要重新按顺序排列，答案字母已同步
    return { questions: out, warn };
  }

  function parseAiJson(text){
    let t = String(text || '').trim();
    t = t.replace(/^```(?:json)?\s*/i, '').replace(/\s*```$/, '');
    try{ return JSON.parse(t); }catch(e){ /* 再试一次：截取第一个 { 到最后一个 } */ }
    const a = t.indexOf('{'), b = t.lastIndexOf('}');
    if(a >= 0 && b > a) return JSON.parse(t.slice(a, b + 1));
    throw new Error('模型返回的内容不是有效的 JSON');
  }

  let aiBusy = false;
  aiGenBtn.addEventListener('click', async () => {
    if(aiBusy) return;
    const topic = aiTopic.value.trim();
    const st = aiSettings();
    const ns = Math.max(0, Math.min(30, parseInt(aiNSingle.value, 10) || 0));
    const nm = Math.max(0, Math.min(30, parseInt(aiNMulti.value, 10) || 0));
    const nt = Math.max(0, Math.min(30, parseInt(aiNTf.value, 10) || 0));
    if(!topic){ setAiStatus('请先填写考试主题。', 'bad'); return; }
    if(!st.key){
      setAiStatus('请先展开"AI 接口设置"，填写你的 DeepSeek API Key。', 'bad');
      if(!aiCfgBody.classList.contains('open')) aiCfgToggle.click();
      return;
    }
    if(ns + nm + nt === 0){ setAiStatus('题量不能全是 0。', 'bad'); return; }

    aiBusy = true;
    aiGenBtn.disabled = true;
    saveAiCfg();
    const t0 = Date.now();
    const timer = setInterval(() => {
      setAiStatus(`正在生成试题…已用 ${Math.round((Date.now() - t0) / 1000)} 秒（通常 30~90 秒）`);
    }, 1000);
    setAiStatus('正在生成试题…');

    let status = 0, bodyText = '';
    const ctrl = new AbortController();
    const to = setTimeout(() => ctrl.abort(), 150000);
    try{
      const res = await fetch(st.base + '/chat/completions', {
        method:'POST',
        signal: ctrl.signal,
        headers:{ 'Content-Type':'application/json', 'Authorization':'Bearer ' + st.key },
        body: JSON.stringify({
          model: st.model,
          messages: [
            { role:'system', content:'你是严谨的医院护理培训考试命题专家，只输出 JSON。' },
            { role:'user', content: buildAiPrompt(topic, aiRef.value.trim(), ns, nm, nt) }
          ],
          temperature: 0.7,
          max_tokens: 7000,
          response_format: { type:'json_object' },
          stream: false
        })
      });
      status = res.status;
      bodyText = await res.text();
      if(!res.ok) throw new Error('HTTP ' + status);
      const data = JSON.parse(bodyText);
      const content = data.choices && data.choices[0] && data.choices[0].message && data.choices[0].message.content;
      const parsed = parseAiJson(content);
      const { questions, warn } = normalizeAi(parsed);

      let name = topic;
      if(!/试题|试卷|考试题|考核/.test(name)) name += '考试题';
      const item = {
        source:'ai', file:null, name: name + '.txt', status:'ok',
        paragraphs:null, questions, error:null,
        expect:{ single:ns, multi:nm, tf:nt }
      };
      state.files.push(item);
      renderFileList();
      parseBtn.disabled = false;
      previewArea.style.display = 'block';
      renderPreview();
      exportPanel.style.display = 'block';

      const c = questions;
      let msg = `✓ 已生成：单选 ${c.single.length} / 多选 ${c.multi.length} / 判断 ${c.tf.length}。请在下方逐题核对，可直接修改。`;
      if(warn.length) msg += '\n⚠ ' + warn.join('；');
      const tfT = c.tf.filter(q => q.answer === '正确').length;
      if(c.tf.length >= 4 && (tfT === 0 || tfT === c.tf.length)) msg += '\n⚠ 判断题答案全是同一个，建议改几题。';
      setAiStatus(msg, warn.length ? 'warn' : 'ok');
      previewArea.scrollIntoView({ behavior:'smooth', block:'start' });
    }catch(e){
      setAiStatus('✕ ' + explainAiError(e, status, bodyText), 'bad');
    }finally{
      clearTimeout(to);
      clearInterval(timer);
      aiBusy = false;
      aiGenBtn.disabled = false;
    }
  });

  // ============================================================
  // XLSX 生成：基于内嵌模板做“外科手术式”XML 替换，完整保留原模板格式
  // ============================================================

  const SHEET_FILE = {
    info: 'xl/worksheets/sheet1.xml',
    single: 'xl/worksheets/sheet2.xml',
    multi: 'xl/worksheets/sheet3.xml',
    tf: 'xl/worksheets/sheet4.xml',
  };

  function xmlEscape(s){
    return String(s == null ? '' : s)
      .replace(/&/g,'&amp;')
      .replace(/</g,'&lt;')
      .replace(/>/g,'&gt;')
      .replace(/"/g,'&quot;')
      .replace(/'/g,'&apos;');
  }

  function colLetter(n){
    // 1-indexed -> 'A','B',...
    let s = '';
    while(n > 0){
      const rem = (n - 1) % 26;
      s = String.fromCharCode(65 + rem) + s;
      n = Math.floor((n - 1) / 26);
    }
    return s;
  }

  // 从已有 sheet XML 中，提取"数据行样式表"：{colIndex(1-based): styleId}
  // 优先取第2行(第一条示例数据行)的样式；若不存在，退化为 <col> 定义的 style
  function extractBodyStyleMap(sheetXml, maxCols){
    const map = {};

    // 1) 尝试从 row r="2" 提取
    const rowMatch = sheetXml.match(/<row r="2"[^>]*>([\s\S]*?)<\/row>/);
    if(rowMatch){
      const cellRe = /<c r="([A-Z]+)2"(?:\s+s="(\d+)")?/g;
      let m;
      while((m = cellRe.exec(rowMatch[1])) !== null){
        const col = colToIndex(m[1]);
        if(m[2] !== undefined) map[col] = m[2];
      }
    }

    // 2) 用 <col> 定义补齐没取到的列
    const colDefRe = /<col min="(\d+)" max="(\d+)"[^>]*style="(\d+)"/g;
    let cm;
    while((cm = colDefRe.exec(sheetXml)) !== null){
      const minC = parseInt(cm[1],10), maxC = parseInt(cm[2],10), styleId = cm[3];
      for(let c=minC;c<=maxC && c<=maxCols;c++){
        if(map[c] === undefined) map[c] = styleId;
      }
    }

    return map;
  }

  function colToIndex(letters){
    let n = 0;
    for(let i=0;i<letters.length;i++){
      n = n*26 + (letters.charCodeAt(i) - 64);
    }
    return n;
  }

  // 构建一行 <row> XML，data: 数组，每项为字符串或 null；styleMap: {colIndex: styleId}
  function buildRowXml(rowIndex, dataArr, styleMap){
    let cells = '';
    for(let i=0;i<dataArr.length;i++){
      const val = dataArr[i];
      const colIdx = i+1;
      const ref = colLetter(colIdx) + rowIndex;
      const styleId = styleMap[colIdx];
      const sAttr = styleId !== undefined ? ` s="${styleId}"` : '';
      if(val === null || val === undefined || val === ''){
        cells += `<c r="${ref}"${sAttr}/>`;
      }else{
        cells += `<c r="${ref}"${sAttr} t="inlineStr"><is><t xml:space="preserve">${xmlEscape(val)}</t></is></c>`;
      }
    }
    return `<row r="${rowIndex}" spans="1:${dataArr.length}">${cells}</row>`;
  }

  // 替换 sheet XML 中的数据区（保留表头行1），写入 rows(二维数组)
  function patchSheetXml(sheetXml, rows, colCount){
    const styleMap = extractBodyStyleMap(sheetXml, colCount);

    // 提取 header row（row r="1"）原样保留
    const headerMatch = sheetXml.match(/<row r="1"[\s\S]*?<\/row>/);
    const headerXml = headerMatch ? headerMatch[0] : '';

    let bodyXml = '';
    rows.forEach((rowData, idx) => {
      bodyXml += buildRowXml(idx+2, rowData, styleMap);
    });

    const newSheetData = `<sheetData>${headerXml}${bodyXml}</sheetData>`;
    let newXml = sheetXml.replace(/<sheetData>[\s\S]*?<\/sheetData>/, newSheetData);

    // 更新 dimension
    const lastRow = rows.length + 1;
    const lastCol = colLetter(colCount);
    newXml = newXml.replace(/<dimension ref="[^"]*"\/>/, `<dimension ref="A1:${lastCol}${lastRow}"/>`);

    return newXml;
  }

  async function buildXlsxForFile(item, cfg){
    const bin = base64ToArrayBuffer(TEMPLATE_XLSX_BASE64);
    const zip = await JSZip.loadAsync(bin);

    const q = item.questions;
    const paperTitle = stripExt(item.name);

    // ---- 试卷信息 sheet ----
    let infoXml = await zip.file(SHEET_FILE.info).async('string');
    const infoStyleMap = extractBodyStyleMap(infoXml, 9);
    const infoRow = [
      paperTitle,          // A 试卷名称
      '',                  // B 试卷描述
      String(calcScore(q, cfg)), // C 试卷总分 = 各题型题数 × 每题分值
      cfg.cfg_score_single,// D 单选题分数
      cfg.cfg_score_multi, // E 多选题分数
      cfg.cfg_score_tf,    // F 是非题分数
      '', '', ''           // G H I 留空
    ];
    const infoHeaderMatch = infoXml.match(/<row r="1"[\s\S]*?<\/row>/);
    const infoHeaderXml = infoHeaderMatch ? infoHeaderMatch[0] : '';
    const infoBodyXml = buildRowXml(2, infoRow, infoStyleMap);
    infoXml = infoXml.replace(/<sheetData>[\s\S]*?<\/sheetData>/, `<sheetData>${infoHeaderXml}${infoBodyXml}</sheetData>`);
    infoXml = infoXml.replace(/<dimension ref="[^"]*"\/>/, `<dimension ref="A1:I2"/>`);
    zip.file(SHEET_FILE.info, infoXml);

    // ---- 单选题 sheet (13列: A题干 B-G选项A-F H答案 I思路 J-M职称/难易/能级/人员) ----
    let singleXml = await zip.file(SHEET_FILE.single).async('string');
    const singleRows = q.single.map(item2 => {
      const optByLetter = {};
      item2.options.forEach(o => { optByLetter[o.letter] = o.text; });
      return [
        item2.stem,
        optByLetter['A']||'', optByLetter['B']||'', optByLetter['C']||'',
        optByLetter['D']||'', optByLetter['E']||'', optByLetter['F']||'',
        item2.answer || '',
        '', '', '', '', ''
      ];
    });
    singleXml = patchSheetXml(singleXml, singleRows, 13);
    zip.file(SHEET_FILE.single, singleXml);

    // ---- 多选题 sheet (15列: A题干 B-I选项A-H J答案 K思路 L-O职称/难易/能级/人员) ----
    let multiXml = await zip.file(SHEET_FILE.multi).async('string');
    const multiRows = q.multi.map(item2 => {
      const optByLetter = {};
      item2.options.forEach(o => { optByLetter[o.letter] = o.text; });
      return [
        item2.stem,
        optByLetter['A']||'', optByLetter['B']||'', optByLetter['C']||'', optByLetter['D']||'',
        optByLetter['E']||'', optByLetter['F']||'', optByLetter['G']||'', optByLetter['H']||'',
        item2.answer || '',
        '', '', '', '', ''
      ];
    });
    multiXml = patchSheetXml(multiXml, multiRows, 15);
    zip.file(SHEET_FILE.multi, multiXml);

    // ---- 是非题 sheet (7列: A题干 B答案 C思路 D-G职称/难易/能级/人员) ----
    let tfXml = await zip.file(SHEET_FILE.tf).async('string');
    const tfRows = q.tf.map(item2 => [
      item2.stem,
      item2.answer || '',
      '', '', '', '', ''
    ]);
    tfXml = patchSheetXml(tfXml, tfRows, 7);
    zip.file(SHEET_FILE.tf, tfXml);

    const blob = await zip.generateAsync({ type:'blob', mimeType:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    return blob;
  }

  function base64ToArrayBuffer(base64){
    const binaryStr = atob(base64);
    const len = binaryStr.length;
    const bytes = new Uint8Array(len);
    for(let i=0;i<len;i++){
      bytes[i] = binaryStr.charCodeAt(i);
    }
    return bytes.buffer;
  }

  // ---------- 导出按钮 ----------
  let lastGeneratedItems = []; // 缓存最近一次生成的 {filename, blob}，供手动兜底按钮复用

  exportBtn.addEventListener('click', async () => {
    exportBtn.disabled = true;
    const origText = exportBtn.textContent;
    exportBtn.textContent = '生成中…';
    downloadFallback.classList.remove('show');
    downloadFallbackLinks.innerHTML = '';

    const cfg = getCfg();
    let successCount = 0;
    const fallbackItems = []; // {filename, blob}

    for(const item of state.files){
      if(item.status !== 'ok' && item.status !== 'warn') continue;
      try{
        log(`生成 Excel：${item.name}`);
        const blob = await buildXlsxForFile(item, cfg);
        const outName = stripExt(item.name) + '.xlsx';
        downloadBlob(blob, outName);
        fallbackItems.push({ filename: outName, blob });
        successCount++;
        log(`  ✓ 已生成：${outName}`);
      }catch(e){
        log(`  ✕ 生成失败：${item.name} — ${e.message}`);
      }
      // 浏览器多文件下载之间稍作间隔，避免被拦截
      await new Promise(r => setTimeout(r, 350));
    }

    exportBtn.disabled = false;
    exportBtn.textContent = origText;
    log(`完成，共生成 ${successCount} 个文件`);
    lastGeneratedItems = fallbackItems;

    // 如果检测到内置浏览器环境，主动显示可点击的兜底下载链接
    // （这些链接本身也是 blob URL，大概率在同一个受限环境里同样会被拦截，
    // 但改成"用户主动点击 <a>"而不是"JS 模拟点击"，在部分内置浏览器里
    // 仍能生效；即便还是不行，横幅里也已经给出了"切换到系统浏览器"的指引）
    if(fallbackItems.length > 0 && inAppLabel){
      renderFallbackLinks(fallbackItems);
    }
  });

  function renderFallbackLinks(items){
    downloadFallbackLinks.innerHTML = '';
    items.forEach(fi => {
      const url = URL.createObjectURL(fi.blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = fi.filename;
      a.textContent = `⬇ ${fi.filename}`;
      a.target = '_blank';
      downloadFallbackLinks.appendChild(a);
    });
    downloadFallback.classList.add('show');
  }

  const fallbackToggle = document.getElementById('fallbackToggle');
  fallbackToggle.addEventListener('click', () => {
    if(lastGeneratedItems.length === 0){
      log('还没有生成过文件，请先点击上方"生成并下载全部 Excel 文件"');
      return;
    }
    if(downloadFallback.classList.contains('show')){
      downloadFallback.classList.remove('show');
    }else{
      renderFallbackLinks(lastGeneratedItems);
    }
  });

  function downloadBlob(blob, filename){
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    setTimeout(() => URL.revokeObjectURL(url), 4000);
  }

})();
