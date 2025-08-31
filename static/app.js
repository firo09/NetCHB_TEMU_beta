// app.js

// —— 硬编码 Excel 密码 —— 
const EXCEL_PASSWORD = 'xU$&#3_*VB';

// ——— Port 映射（Port 字样输入=代码；Firms Code=firms）———
const PORT_MAP = {
  JFK: { code: '4701', firms: 'EAT5', state: 'NY' },
  LAX: { code: '2720', firms: 'WBH9', state: 'CA' },
  SFO: { code: '2801', firms: 'W0B3', state: 'CA' },
  ORD: { code: '3901', firms: 'HBT1', state: 'IL' },
  DFW: { code: '5501', firms: 'SE04', state: 'TX' },
  MIA: { code: '5206', firms: 'LEG0', state: 'FL' },
  ATL: { code: '1704', firms: 'L543', state: 'GA' },
  BOS: { code: '0417', firms: 'AAN5', state: 'MA' },
  SEA: { code: '3029', firms: 'WBU6', state: 'WA' },
};

// 识别当前页面客户（用于导出文件名）
const IS_SHEIN = (document.body && document.body.getAttribute('data-client') === 'shein');

// 根据客户选择不同的配置文件名（SHEIN 使用 *_shein.json）
const CONFIG_FILES = IS_SHEIN
  ? { rule: 'rule_shein.json', rule_consolidated: 'rule_consolidated_shein.json', hts: 'hts_shein.json', mid: 'mid_shein.json', pga: 'PGA_shein.json' }
  : { rule: 'rule.json',       rule_consolidated: 'rule_consolidated.json',       hts: 'hts.json',        mid: 'mid.json',        pga: 'PGA.json'       };

// 全局保存当前文件、默认 MAWB、预选项
let currentFile = null;
let currentDefaultMawb = '';
let selectedPortKey = '';
let selectedDateKey = ''; // '', 'today', 'tomorrow'

// 防缓存
const ts = Date.now();
const CONFIG_PATH = 'config';

// ==== 单位换算（值=1单位等于多少“基准单位”） ====
// 重量：以 kg 为基
const KG_FACTORS = {
  "mg": 1e-6, "cgm": 1e-5, "g": 1e-3, "ckg": 0.1, "kg": 1.0, "t": 1000.0,
  "car": 0.0002, "gr": 0.00006479891, "osg": 0.028349523125, "oz": 0.028349523125,
  "lb": 0.45359237, "kg cmsc": 1.0
};
// 计数：以 pcs 为基
const PCS_FACTORS = {
  "pcs": 1, "no.": 1, "prs.": 1, "doz.": 12, "doz. prs.": 12, "gross": 144
};
// 统一单位字符串
function normUnit(u){ return String(u||'').trim().toLowerCase(); }

// —— 只基于主列 HTS 计算 HTSQty / HTSQty2 ——
// 规则：重量单位 → 数量 = GrossWeight(kg) ÷ KG_FACTORS[unit]（三位小数）
//      计数单位 → 数量 = Piece(pcs)      ÷ PCS_FACTORS[unit]（三位小数）
//      配不到单位或数据缺失 → 空串
function attachHtsQty(row){
  const hts = String(row['HTS'] || '').replace(/\D/g, '');
  if (!hts) return;
  const cfg = (window.__unitConfig || {})[hts];
  if (!cfg) return;

  const kg  = Number(String(row['GrossWeight'] || '').replace(/,/g,'')) || 0;           // 已以 kg 计
  const pcs = Number(String(row['Manifest Qty Piece count'] || '').replace(/,/g,'')) || 0; // 已以 pcs 计

  function calc(uLabel){
    if (!uLabel) return '';
    const u = normUnit(uLabel);
    if (KG_FACTORS[u] !== undefined)  return (kg  / KG_FACTORS[u]).toFixed(3);
    if (PCS_FACTORS[u] !== undefined) return (pcs / PCS_FACTORS[u]).toFixed(3);
    return '';
  }

  const q1 = calc(cfg.Unit1 || cfg.unit1);
  const q2 = calc(cfg.Unit2 || cfg.unit2);
  if (q1) row['HTSQty']  = q1;
  if (q2) row['HTSQty2'] = q2;
}


// ==== 新增：是否 TEMU（留好多客户扩展位）====
const IS_TEMU = (typeof IS_SHEIN !== 'undefined') ? !IS_SHEIN : true;

// ==== 新增：合单判定所用的表头名（可被外部覆盖）====
// 优先级：window.__CONSOLIDATE_HEADER > <body data-consolidate-header="..."> > 默认值
const CONSOLIDATE_HEADER_KEY =
  (window.__CONSOLIDATE_HEADER) ||
  (document.body && document.body.getAttribute('data-consolidate-header')) ||
  'consignor_item_id';

// ==== 合单检测规范====
// 每个客户：指定只检查的 sheet 以及要匹配的表头名列表（可写别名）
const CONSOLIDATE_DETECT_SPEC = {
  temu:  { sheet: 'hawb', headers: ['consignor_item_id'] }
};

// 识别当前客户（优先 HTML 上 data-client，其次 IS_SHEIN 开关）
function getClientId() {
  const v = (document.body && document.body.getAttribute('data-client')) || '';
  if (v) return v.trim().toLowerCase();
  return (typeof IS_SHEIN !== 'undefined' && IS_SHEIN) ? 'shein' : 'temu';
}

// 归一化工具：忽略大小写、空格、下划线、短横线及其他符号
const canon = s => String(s ?? '')
  .toLowerCase()
  .replace(/[\s_-]+/g, '')
  .replace(/[^a-z0-9]/g, '');

// 解析“本次要检查的 sheet 与表头键”
// 允许通过 window.__CONSOLIDATE_SHEET / window.__CONSOLIDATE_HEADER（字符串或数组）覆盖
function resolveConsolidateDetectSettings() {
  const client = getClientId();
  const def = CONSOLIDATE_DETECT_SPEC[client] || { sheet: 'hawb', headers: ['consignor_item_id'] };

  // 覆盖（可选）
  const sheetOverride =
    (window.__CONSOLIDATE_SHEET) ||
    (document.body && document.body.getAttribute('data-consolidate-sheet'));

  let headersOverride = (window.__CONSOLIDATE_HEADER) ||
    (document.body && document.body.getAttribute('data-consolidate-header'));

  // 支持字符串或数组
  if (headersOverride && !Array.isArray(headersOverride)) {
    headersOverride = [String(headersOverride)];
  }

  const sheetKey = String(sheetOverride || def.sheet || 'hawb').trim().toLowerCase();
  const headerList = (headersOverride && headersOverride.length ? headersOverride : def.headers || ['consignor_item_id'])
    .map(h => String(h));

  return { sheetKey, headerList };
}

// ===== Shein 分组（可给 TEMU-合单复用）=====
// options.respectHouseAwb=true 时：先按 House AWB 捆绑，再以 998 上限装箱（FFD 贪心）
// 否则：均匀切块（保持你原来 SHEIN 的分组风格）
function applySheinGrouping(output, { respectHouseAwb = false } = {}) {
  const MAX_PER_GROUP = 998; // 如需调整，可改这里或外部注入 window.__SHEIN_GROUP_MAX
  const cap = Number.isFinite(window.__SHEIN_GROUP_MAX) ? Math.max(1, +window.__SHEIN_GROUP_MAX) : MAX_PER_GROUP;

  if (!Array.isArray(output) || output.length === 0) return;

  if (!respectHouseAwb) {
    // === 原来的均匀分组（每组≤cap，尽量均分） ===
    const total = output.length;
    const groups = Math.ceil(total / cap);
    if (groups <= 0) return;
    const base = Math.floor(total / groups);
    const extra = total % groups;
    const sizes = Array.from({ length: groups }, (_, i) => base + (i < extra ? 1 : 0));
    let idx = 0, gid = 1;
    for (const sz of sizes) {
      for (let k = 0; k < sz && idx < total; k++, idx++) {
        if (output[idx] && typeof output[idx] === 'object') {
          output[idx]['GroupIdentifier'] = gid;
        }
      }
      gid++;
    }
    return;
  }

  // === 合单模式：House AWB 相同必须同组（在 ≤cap 的前提下） ===
  // 1) 先按 House AWB 捆绑
  const keyName = 'House AWB';
  const bundlesMap = new Map(); // key -> index[]
  for (let i = 0; i < output.length; i++) {
    const row = output[i] || {};
    const key = String(row[keyName] ?? '').trim();
    const k = key || `__EMPTY__`; // 空值也单独成组，避免跨行混淆
    if (!bundlesMap.has(k)) bundlesMap.set(k, []);
    bundlesMap.get(k).push(i);
  }
  const bundles = Array.from(bundlesMap.entries()).map(([k, arr]) => ({ key: k, idxs: arr, size: arr.length }));

  // 2) 如果某个 House 超过上限，无法同时满足“≤cap 且同组”，此时仅能拆分（记录警告）
  const huge = bundles.filter(b => b.size > cap);
  if (huge.length) {
    console.warn('[Grouping] Some House AWB sizes exceed cap and must be split:', huge.map(h => ({ house:h.key, size:h.size })));
  }

  // 3) 按 size 降序做 First-Fit-Decreasing 装箱
  bundles.sort((a,b) => b.size - a.size);
  const groups = []; // {count:number, chunks: number[][]}
  for (const b of bundles) {
    if (b.size <= cap) {
      let placed = false;
      for (const g of groups) {
        if (g.count + b.size <= cap) {
          g.chunks.push(b.idxs);
          g.count += b.size;
          placed = true;
          break;
        }
      }
      if (!placed) groups.push({ count: b.size, chunks: [b.idxs] });
    } else {
      // 必须拆分：按 cap 切段，尽量连续
      for (let s = 0; s < b.size; s += cap) {
        const chunk = b.idxs.slice(s, s + cap);
        groups.push({ count: chunk.length, chunks: [chunk] });
      }
    }
  }

  // 4) 写入 GroupIdentifier
  let gid = 1;
  for (const g of groups) {
    for (const chunk of g.chunks) {
      for (const idx of chunk) {
        if (output[idx] && typeof output[idx] === 'object') {
          output[idx]['GroupIdentifier'] = gid;
        }
      }
    }
    gid++;
  }
}


// ==== 新增：区分“基础规则 / 合单规则 / 当前激活规则”====
let baseRuleConfig = [];
let ruleConsolidatedConfig = [];
let ruleConfig = [];               // 全局“当前激活”的规则（其余逻辑全部沿用它）
window.__isConsolidatedShipment = false;   // 供 UI 使用

// === Button loading helpers ===
function ensureSpinnerCss() {
  if (document.getElementById('btn-spinner-style')) return;
  const style = document.createElement('style');
  style.id = 'btn-spinner-style';
  style.textContent = `
    @keyframes spin{to{transform:rotate(360deg)}}
    .btn--loading{position:relative; pointer-events:none; opacity:.8}
    .btn--loading .btn__label{visibility:hidden}
    .btn--loading::after{
      content:""; position:absolute; top:50%; left:50%;
      width:16px; height:16px; margin:-8px 0 0 -8px; border-radius:50%;
      border:2px solid currentColor; border-top-color:transparent;
      animation:spin .6s linear infinite;
    }`;
  document.head.appendChild(style);
}
function lockButton(btn){
  ensureSpinnerCss();
  btn.classList.add('btn--loading');
  btn.disabled = true;
}
function unlockButton(btn){
  btn.classList.remove('btn--loading');
  btn.disabled = false;
}

// 入口页或 shein/temu 页都有这些元素（入口页不会加载 app.js）
const uploadBtn   = document.getElementById('upload-btn');
const fileInput   = document.getElementById('file-input');
const loadingMsg  = document.getElementById('loading-msg');
const continueBtn = document.getElementById('continue-btn');
const portSel     = document.getElementById('pref-port');
const dateSel     = document.getElementById('pref-date');
const generateBtn = document.getElementById('generate-btn');
(function initButtonLabels(){
  ['continue-btn', 'generate-btn'].forEach(id => {
    const btn = document.getElementById(id);
    if (btn && !btn.querySelector('.btn__label')) {
      const t = btn.textContent;
      btn.textContent = '';
      const span = document.createElement('span');
      span.className = 'btn__label';
      span.textContent = t;
      btn.appendChild(span);
    }
  });
})();
// 统一 TEMU 主题色（#0071bc）
(function applyTemuTheme(){
  if (typeof IS_SHEIN === 'undefined' || IS_SHEIN) return;
  const temuBlue = '#0071bc';
  const h1 = document.querySelector('h1');
  if (h1) h1.style.color = temuBlue;
  const gen = document.getElementById('generate-btn');
  if (gen) { gen.style.backgroundColor = temuBlue; gen.style.borderColor = temuBlue; }
  const prog = document.getElementById('progress');
  if (prog) { prog.style.backgroundColor = temuBlue; }
  const activeNav = document.querySelector('aside a[aria-current="page"]') || document.querySelector('aside a.scale-105');
  if (activeNav) { activeNav.style.background = temuBlue; activeNav.style.color = '#fff'; }
})();

let htsData = [], midData = [], pgaRules = [];

// === MID 替换：全局状态 ===
window.__midReplaceEnabled = false;
window.__midReplaceRowsSet = new Set(); // 存储“包含表头”的行号（Number），如 361、4330 等
window.__roundUpPgaQtyEnabled = false;  // 勾选“Auto-adjust…”时为 true

function __parseRowsInputToSet(inputText) {
  // 允许形式： "Rows: 361, 4330, 4676" 或随意空格/中英文逗号
  if (!inputText) return new Set();
  const nums = (inputText.match(/\d+/g) || []).map(n => parseInt(n, 10)).filter(n => Number.isFinite(n) && n > 0);
  return new Set(nums);
}


// ===== 自绘下拉：样式注入 + 构建 =====
(function injectSelectStyles(){
  if (document.getElementById('ui-select-styles')) return;
  const css = `
.ui-select{position:relative;width:100%}
.ui-select__btn{width:100%;border:1px solid #d1d5db;border-radius:12px;padding:10px 40px 10px 14px;background:#fff;
  box-shadow:0 1px 2px rgba(16,24,40,.05);line-height:1.2}
.ui-select__caret{position:absolute;right:12px;top:50%;transform:translateY(-50%);pointer-events:none;opacity:.6}
.ui-select__menu{position:absolute;z-index:50;left:0;top:calc(100% + 6px);width:100%;max-height:260px;overflow:auto;
  background:#fff;border:1px solid #e5e7eb;border-radius:12px;box-shadow:0 8px 24px rgba(16,24,40,.12);display:none}
.ui-select.open .ui-select__menu{display:block}
.ui-option{padding:10px 12px;cursor:pointer}
.ui-option:hover{background:#f3f4f6}
.ui-option[aria-selected="true"]{background:#eef2ff}
/* 文本输入框统一外观，和下拉按钮一致 */
.ui-input{width:100%;border:1px solid #d1d5db;border-radius:12px;padding:10px 14px;background:#fff;
  box-shadow:0 1px 2px rgba(16,24,40,.05);line-height:1.2;transition:box-shadow .15s,border-color .15s;outline:0}
.ui-input:focus{box-shadow:0 0 0 3px rgba(148,163,184,.25)};`
  const style = document.createElement('style'); style.id='ui-select-styles'; style.textContent = css;
  document.head.appendChild(style);
})();

function buildCustomSelect(sel, accent) {
  if (!sel || sel.dataset.uiBound) return;
  sel.dataset.uiBound = '1';
  sel.classList.add('hidden');

  const root = document.createElement('div');
  root.className = 'ui-select';
  sel.insertAdjacentElement('afterend', root);

  const btn = document.createElement('button');
  btn.type='button'; btn.className='ui-select__btn';
  btn.textContent = sel.options[sel.selectedIndex]?.text || '-- Select --';
  root.appendChild(btn);

  const caret = document.createElementNS('http://www.w3.org/2000/svg','svg');
  caret.setAttribute('viewBox','0 0 20 20'); caret.setAttribute('width','20'); caret.setAttribute('height','20');
  caret.classList.add('ui-select__caret');
  caret.innerHTML = '<path fill="currentColor" d="M5.3 7.3a1 1 0 0 1 1.4 0L10 10.6l3.3-3.3a1 1 0 1 1 1.4 1.4l-4 4a1 1 0 0 1-1.4 0l-4-4a1 1 0 0 1 0-1.4z"/>';
  root.appendChild(caret);

  const menu = document.createElement('div');
  menu.className = 'ui-select__menu';
  Array.from(sel.options).forEach(opt => {
    const item = document.createElement('div');
    item.className = 'ui-option';
    item.textContent = opt.text;
    item.dataset.value = opt.value;
    if (opt.selected) item.setAttribute('aria-selected','true');
    item.addEventListener('click', () => {
      sel.value = opt.value;
      sel.dispatchEvent(new Event('change', {bubbles:true}));
      btn.textContent = opt.text;
      menu.querySelectorAll('.ui-option[aria-selected="true"]').forEach(n => n.removeAttribute('aria-selected'));
      item.setAttribute('aria-selected','true');
      root.classList.remove('open');
    });
    menu.appendChild(item);
  });
  root.appendChild(menu);

  btn.addEventListener('click', () => root.classList.toggle('open'));
  document.addEventListener('click', (e)=>{ if(!root.contains(e.target)) root.classList.remove('open'); });

  const temuBlue = '#0071bc', sheinGreen = '#10b981';
  btn.addEventListener('focus', () => { btn.style.boxShadow = '0 0 0 3px rgba(148,163,184,.25)'; });
  btn.addEventListener('blur',  () => { btn.style.boxShadow = '0 1px 2px rgba(16,24,40,.05)'; });
}

function hexToRgba(hex, a){
  const m = hex.replace('#','');
  const bigint = parseInt(m.length===3? m.split('').map(x=>x+x).join(''): m, 16);
  const r = (bigint>>16)&255, g=(bigint>>8)&255, b=bigint&255;
  return `rgba(${r}, ${g}, ${b}, ${a})`;
}

function beautifyAllSelects(container, accent){
  container.querySelectorAll('select').forEach(sel => {
    if (sel.offsetParent !== null) {
      buildCustomSelect(sel, accent);
    }
  });
}

// 日期格式化
function formatDateByPattern(date, pattern) {
  const yyyy = date.getFullYear();
  const m = date.getMonth() + 1;
  const d = date.getDate();
  const pad = n => (n < 10 ? '0' + n : n);
  return pattern
    .replace(/yyyy/gi, yyyy)
    .replace(/mm/g, pad(m))
    .replace(/m/g, m)
    .replace(/dd/g, pad(d))
    .replace(/d/g, d);
}

// 只用于 mawb sheet 提取
function getValueFromMawbSheet(mawbSheetArr, colName) {
  if (!Array.isArray(mawbSheetArr) || mawbSheetArr.length < 2) return '';
  const header = mawbSheetArr[0] || [];
  const row    = mawbSheetArr[1] || [];
  const idx = header.findIndex(h => (h || '').toString().trim().toLowerCase() === colName.trim().toLowerCase());
  if (idx === -1) return '';
  return row[idx] || '';
}

function sanitize(label) { return label.replace(/[^\w]/g, '_'); }

function parseValue(val, parsing) {
  if (!parsing) return val;

  const p = String(parsing).trim();

  // === 新增：字母数字清洗 + 左/右截取（大小写不敏感） ===
  // abcRight(x): 先去除非字母数字，再取右数 x 个字符
  {
    const m = p.match(/^abcRight\(\s*(\d+)\s*\)$/i);
    if (m) {
      const n = parseInt(m[1], 10) || 0;
      const s = (val ?? '').toString().replace(/[^A-Za-z0-9]/g, '');
      if (n <= 0) return '';
      return s.slice(-n); // 不足 n 时返回全部
    }
  }
  // abcLeft(y): 先去除非字母数字，再取左数 y 个字符
  {
    const m = p.match(/^abcLeft\(\s*(\d+)\s*\)$/i);
    if (m) {
      const n = parseInt(m[1], 10) || 0;
      const s = (val ?? '').toString().replace(/[^A-Za-z0-9]/g, '');
      if (n <= 0) return '';
      return s.slice(0, n); // 不足 y 时返回全部
    }
  }

  // === 兼容你原来的 left(n) / right(n) 规则 ===
  const leftMatch  = p.match(/^left\((\d+)\)$/i);
  if (leftMatch)  return (val || '').toString().slice(0, parseInt(leftMatch[1], 10));

  const rightMatch = p.match(/^right\((\d+)\)$/i);
  if (rightMatch) return (val || '').toString().slice(-parseInt(rightMatch[1], 10));

  // 其它未识别：原样返回
  return val;
}



// —— 新增：自动识别标题行 + 关键词右侧取值（通用视图） —— 
function buildSheetView(arrAOA, expectedHeaders) {
  const norm = s => (s ?? '').toString().trim().toLowerCase();

  // 1) 自动识别标题行：命中预期列名最多者（相同则取更靠近数据区的一行）
  let headerRowIdx = -1, bestHit = -1;
  for (let r = 0; r < arrAOA.length; r++) {
    const row = (arrAOA[r] || []).map(norm);
    let hit = 0;
    for (const h of (expectedHeaders || [])) {
      const hh = norm(h);
      if (!hh) continue;
      if (row.includes(hh)) hit++;
    }
    if (hit > bestHit || (hit === bestHit && r > headerRowIdx)) {
      bestHit = hit; headerRowIdx = r;
    }
  }
  const pass = (bestHit >= 2) || (bestHit >= Math.ceil((expectedHeaders || []).length / 2));
  if (!pass) headerRowIdx = -1;

  const colMap = {};
  if (headerRowIdx >= 0) {
    const headerRow = arrAOA[headerRowIdx] || [];
    headerRow.forEach((name, idx) => { colMap[norm(name)] = idx; });
  }

  return {
    headerRowIdx,
    getByHeaderRow: (i, refName) => {
      if (headerRowIdx < 0) return '';
      const rowIdx = headerRowIdx + 1 + i;
      const colIdx = colMap[norm(refName)];
      if (colIdx == null) return '';
      const row = arrAOA[rowIdx] || [];
      return row[colIdx] ?? '';
    },
    getByKeywordRight: (keyword) => {
      if (!arrAOA || arrAOA.length === 0) return '';
      const lastRow = (headerRowIdx >= 0) ? headerRowIdx - 1 : (arrAOA.length - 1);
      const keyNorm = norm(keyword);
      for (let r = lastRow; r >= 0; r--) {
        const row = arrAOA[r] || [];
        for (let c = 0; c < row.length; c++) {
          const cell = row[c];
          if (norm(cell) === keyNorm) {
            for (let cc = c + 1; cc < row.length; cc++) {
              const v = row[cc];
              if (v !== '' && v !== undefined && v !== null) return v;
            }
            return row[c + 1] ?? '';
          }
        }
      }
      return '';
    }
  };
}
function buildRegex(fmt) {
  let regexStr = fmt.replace(/([.+?^=!:${}()|[\]\/\\])/g, '\\$1');
  regexStr = regexStr.replace(/y{4}/g, '\\d{4}');
  regexStr = regexStr.replace(/m{1,2}/gi, '\\d{1,2}');
  regexStr = regexStr.replace(/d{1,2}/gi, '\\d{1,2}');
  return new RegExp('^' + regexStr + '$');
}

// 加载配置（在 DOM 就绪后执行，确保元素存在；HTS/MID/PGA 缺失也不阻塞）
function __startConfigLoad() {
  const uploadBtn   = document.getElementById('upload-btn');
  const loadingMsg  = document.getElementById('loading-msg');
  if (!uploadBtn || !loadingMsg) return; // 非客户页

  loadingMsg.innerText = 'Loading configuration...';

  const safeFetch = (url, fallback) =>
    fetch(url).then(r => r.ok ? r.json() : fallback).catch(() => fallback);

  // 同时加载：rule + rule_consolidated + 其余数据
  Promise.all([
    safeFetch(`${CONFIG_PATH}/${CONFIG_FILES.rule}?ts=${ts}`, null),
    safeFetch(`${CONFIG_PATH}/${CONFIG_FILES.rule_consolidated}?ts=${ts}`, []),
    safeFetch(`${CONFIG_PATH}/${CONFIG_FILES.hts}?ts=${ts}`, []),
    safeFetch(`${CONFIG_PATH}/${CONFIG_FILES.mid}?ts=${ts}`, []),
    safeFetch(`${CONFIG_PATH}/${CONFIG_FILES.pga}?ts=${ts}`, []),
    safeFetch(`${CONFIG_PATH}/unit.json?ts=${ts}`, {})
  ]).then(([rule, ruleCons, hts, mid, pga, unitCfg]) => {
    if (!rule) throw new Error('rule config missing');

    baseRuleConfig         = rule || [];
    ruleConsolidatedConfig = Array.isArray(ruleCons) ? ruleCons : [];
    ruleConfig             = baseRuleConfig;   // 初始使用常规规则

    htsData = hts; midData = mid; pgaRules = pga;

    // 供后续计算 HTSQty / HTSQty2 使用（key=HTS，value={Unit1,Unit2}）
    window.__unitConfig = unitCfg || {};

    uploadBtn.disabled = false;
    uploadBtn.classList.remove('opacity-50');
    loadingMsg.innerText = '';
  }).catch(e => {
    console.error('Failed to load configs', e);
    loadingMsg.innerText = 'Failed to load configuration';
  });
}


// DOMContentLoaded 触发加载；若已就绪则立即加载
if (document.readyState === 'loading') {
  document.addEventListener('DOMContentLoaded', __startConfigLoad);
} else {
  __startConfigLoad();
}
// 选择文件
  if (uploadBtn)  uploadBtn.addEventListener('click', () => { try { fileInput.value = ''; } catch (_) {} fileInput.click(); });
  if (fileInput)  fileInput.addEventListener('change', () => {
    if (!fileInput.files.length) { alert('Please select a file'); return; }
    currentFile = fileInput.files[0];
    const base = currentFile.name.replace(/\.(xlsx|xls|csv)$/i, '');

    // 支持两种：11位；或 3位-8位（去掉连字符）
    let m = base.match(/(\d{11})$/);
    if (!m) {
      const m2 = base.match(/(\d{3})-(\d{8})$/);
      if (m2) currentDefaultMawb = m2[1] + m2[2];
      else    currentDefaultMawb = '';
    } else {
      currentDefaultMawb = m[1];
    }

    continueBtn && (continueBtn.disabled = false);

    // 上传成功提示（英文）
    if (uploadBtn) {
      if (uploadBtn.querySelector && uploadBtn.querySelector('span')) {
        uploadBtn.querySelector('span').textContent = '✅ File uploaded successfully';
      } else {
        uploadBtn.innerHTML = '✅ File uploaded successfully';
      }
    }
  // ✅ 新增这一行：让标题立刻变成 "123-45678901"
  window.updateTitleFromMAWB?.(currentDefaultMawb);
  });

  // 记录 Port/Date 选择
  portSel && portSel.addEventListener('change', () => selectedPortKey = portSel.value.trim());
  dateSel && dateSel.addEventListener('change', () => selectedDateKey = dateSel.value.trim());

// Continue：点击后先做合单检测与规则切换，完成后再进入表单页
if (continueBtn) {
  continueBtn.addEventListener('click', async (e) => {
    if (!currentFile) { alert('Please select a file'); return; }

    const btn = e.currentTarget;
    lockButton(btn);  // 禁用 + 隐文字 + 旋转加载

    try {
      // 只在代码里指定的 sheet/header 上判断：hawb + consignor_item_id（可在上方常量覆盖）
      const isCons = await detectConsolidatedShipmentFromFile(currentFile);
      window.__isConsolidatedShipment = !!isCons;

      // 命中则切到合单规则，否则走常规 rule
      ruleConfig = (isCons && Array.isArray(ruleConsolidatedConfig) && ruleConsolidatedConfig.length)
        ? ruleConsolidatedConfig
        : baseRuleConfig;

      // 检测完成后再切换页面并渲染（顶部会出现 Consolidated 横幅）
      document.getElementById('upload-section')?.classList.add('hidden');
      document.getElementById('form-section')?.classList.remove('hidden');
      renderForm(currentDefaultMawb, { portKey: selectedPortKey, dateKey: selectedDateKey });
    } catch (err) {
      console.warn('Continue flow failed:', err);
      alert('Failed to prepare the form. Please check the file and try again.');
    } finally {
      unlockButton(btn); // 切页后通常被隐藏，解锁无妨
    }
  });
}


// 只在“代码里指定的 sheet + 表头键”上判断是否 Consolidated
async function detectConsolidatedShipmentFromFile(file) {
  try {
    const buf = await file.arrayBuffer();
    const wb = XLSX.read(buf, { type: 'array', password: EXCEL_PASSWORD });

    const { sheetKey, headerList } = resolveConsolidateDetectSettings();
    const headerKeySet = new Set(headerList.map(canon));

    // 只检查目标 sheet；找不到就谨慎返回 false
    const sheetName = wb.SheetNames.find(n => String(n).trim().toLowerCase() === sheetKey);
    if (!sheetName) {
      console.info('[Consolidated] target sheet not found:', sheetKey);
      return false;
    }

    const ws  = wb.Sheets[sheetName];
    const aoa = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' }) || [];

    // 找“表头行 & 目标列”（任一 header 别名命中即可）
    let headerRowIdx = -1, colIdx = -1;
    for (let r = 0; r < aoa.length; r++) {
      const row = aoa[r] || [];
      const idx = row.findIndex(cell => headerKeySet.has(canon(cell)));
      if (idx !== -1) { headerRowIdx = r; colIdx = idx; break; }
    }
    if (headerRowIdx === -1 || colIdx === -1) {
      console.info('[Consolidated] header not found on sheet:', sheetKey, 'headers tried:', headerList);
      return false;
    }

    // 统计唯一值（忽略空白；整格内容为一个值）
    const uniques = new Set();
    for (let r = headerRowIdx + 1; r < aoa.length; r++) {
      const v = aoa[r]?.[colIdx];
      const s = String(v ?? '').trim();
      if (!s) continue;
      uniques.add(s);
      if (uniques.size > 1) {
        console.info('[Consolidated] unique>=2 on sheet=%s col=%d → TRUE', sheetName, colIdx);
        return true;
      }
    }
    const ok = uniques.size > 1;
    console.info('[Consolidated] uniqueCount=%d on sheet=%s col=%d → %s',
                 uniques.size, sheetName, colIdx, ok);
    return ok;
  } catch (err) {
    console.warn('Consolidated detection failed:', err);
    return false; // 失败不误判
  }
}

  // 生成下载
  generateBtn.addEventListener('click', () => {
    if (!currentFile) { alert('No file selected'); return; }
    generateAndDownload();
  });


// 渲染动态表单，并根据 Port/Date 做覆盖
function renderForm(defaultMawb, { portKey = '', dateKey = '' } = {}) {
  const formEl = document.getElementById('dynamic-form');

// ==== 新增：若判定为合单，在表单页顶部给出醒目的提醒 ====
(function maybeShowConsolidatedBanner() {
  const formEl = document.getElementById('dynamic-form');
  // 移除旧横幅（避免重复）
  const old = document.getElementById('consolidated-banner');
  if (old) old.remove();

  if (!window.__isConsolidatedShipment) return;

  const banner = document.createElement('div');
  banner.id = 'consolidated-banner';
  banner.className = 'col-span-2';
  banner.style.border = '1px solid #fbbf24';      // amber-400
  banner.style.background = '#fffbeb';            // amber-50
  banner.style.color = '#92400e';                 // amber-800
  banner.style.borderRadius = '12px';
  banner.style.padding = '12px 16px';
  banner.style.marginBottom = '12px';
  banner.style.boxShadow = '0 4px 16px rgba(146,64,14,0.08)';
  banner.style.fontWeight = '700';
  banner.style.display = 'flex';
  banner.style.alignItems = 'center';
  banner.style.gap = '10px';

  const icon = document.createElementNS('http://www.w3.org/2000/svg','svg');
  icon.setAttribute('viewBox','0 0 24 24');
  icon.setAttribute('width','22');
  icon.setAttribute('height','22');
  icon.innerHTML = '<path fill="currentColor" d="M1 21h22L12 2 1 21zm12-3h-2v-2h2v2zm0-4h-2V9h2v5z"/>';
  banner.appendChild(icon);

  const text = document.createElement('div');
  text.textContent = 'Consolidated Shipment';
  banner.appendChild(text);

  formEl.parentNode?.insertBefore(banner, formEl); // 插在表单网格上方，更显眼
})();

  // ===== 顶部：Replace MID with Manufacturer Name and Address 开关 + 输入框 =====
  // 容器
  const midBox = document.createElement('div');
  midBox.className = 'col-span-2 p-4 rounded-xl border border-gray-200 bg-gray-50 mt-2';
  midBox.style.marginTop = '4px';
  midBox.style.marginBottom = '0';

  // 开关行（左侧复选框风格）
  midBox.innerHTML = `
    <label class="inline-flex items-center gap-2 cursor-pointer">
      <input id="midreplace-toggle" type="checkbox" class="h-5 w-5 accent-blue-600">
      <span class="text-slate-700 font-semibold">Replace MID with Manufacturer Name and Address</span>
    </label>
  `;

  // 不额外设置 marginBottom，保持默认
  midBox.style.marginTop = '0';

  // 提示 + 文本框（默认隐藏）
  const inputWrap = document.createElement('div');
  inputWrap.className = 'mt-3 hidden';
  inputWrap.id = 'midreplace-input-wrap';

  // 提示行：Please enter the content in (parentheses) —— 括号文字标红（沿用错误红色）
  const tip = document.createElement('div');
  tip.className = 'text-sm text-slate-600 mb-1';
  tip.innerHTML = 'Please enter the content in <span style="color:red;">parentheses</span>';

  const input = document.createElement('input');
  input.type = 'text';
  input.id = 'midreplace-rows';
  input.className = 'ui-input w-full italic text-gray-500';
  input.placeholder = 'Rows: 361, 4330, 4676, 5475, 6870, 9340, 9354, 9377';

  inputWrap.appendChild(tip);
  inputWrap.appendChild(input);
  midBox.appendChild(inputWrap);

  // 事件：开关控制显示/隐藏（用 querySelector 取到新复选框）
  const switchInput = midBox.querySelector('#midreplace-toggle');
  switchInput.checked = !!window.__midReplaceEnabled;
  inputWrap.classList.toggle('hidden', !switchInput.checked);

  switchInput.addEventListener('change', () => {
    window.__midReplaceEnabled = switchInput.checked;
    inputWrap.classList.toggle('hidden', !switchInput.checked);
    window.__midReplaceRowsSet = __parseRowsInputToSet(
      document.getElementById('midreplace-rows')?.value || ''
    );
  });

  // 事件：开关控制显示/隐藏
  switchInput.addEventListener('change', () => {
    window.__midReplaceEnabled = switchInput.checked;
    inputWrap.classList.toggle('hidden', !switchInput.checked);
    // 每次切换时重算行集合
    window.__midReplaceRowsSet = __parseRowsInputToSet(document.getElementById('midreplace-rows')?.value || '');
  });

  // 事件：输入变化 → 解析行号集合
  input.addEventListener('input', () => {
    window.__midReplaceRowsSet = __parseRowsInputToSet(input.value || '');
    // placeholder 自然由浏览器处理：非空时不显示
    // 为了视觉统一，用户一旦输入，把斜体/灰色去掉，仅显示用户输入
    if ((input.value || '').trim().length > 0) {
      input.classList.remove('italic','text-gray-500');
    } else {
      input.classList.add('italic','text-gray-500');
    }
  });

  formEl.innerHTML = '';

  // 插到表单网格的最前面
  formEl.appendChild(midBox);

  // ===== 在 Consolidated Shipment 与 Replace MID 之间插入 Auto-adjust（无说明文字）=====
  try {
    // 只有当导出表里会含 FDAPRODUCTCODE 才显示此选项
    const hasFDACode = Array.isArray(ruleConfig) &&
                       ruleConfig.some(r => String(r.Column).trim() === 'FDAPRODUCTCODE');

    if (hasFDACode) {
      // 容器（与 MID 同风格）
      const pgaRow = document.createElement('div');
      pgaRow.id = 'row-pga-roundup';
      pgaRow.className = 'col-span-2 p-4 rounded-xl border border-gray-200 bg-gray-50 mt-2';
      pgaRow.style.marginTop = '4px';
      pgaRow.style.marginBottom = '0';

      // 左侧复选框 + 标题（无额外描述）
      pgaRow.innerHTML = `
        <label class="inline-flex items-center gap-2 cursor-pointer">
          <input id="chk-roundup-pga" type="checkbox" class="h-5 w-5 accent-blue-600">
          <span class="text-slate-700 font-semibold">Auto-adjust PGA QTY less than 1 to 1</span>
        </label>
      `;

      // 调整上下边距，避免和横幅/下方块产生大缝
      pgaRow.style.marginTop = '0';
      pgaRow.style.marginBottom = '-10px';   // 只管和 Replace MID 的间距

      // 放到 midBox 前面（让它居于 Consolidated 横幅与 MID 框之间）
      formEl.insertBefore(pgaRow, midBox);

      // 同步/监听勾选态
      const chkPga = pgaRow.querySelector('#chk-roundup-pga');
      chkPga.checked = !!window.__roundUpPgaQtyEnabled;
      chkPga.addEventListener('change', e => {
        window.__roundUpPgaQtyEnabled = !!e.target.checked;
      });
    }
  } catch (_) {}


  const labels = [];
  const primaryRuleFor = {};
  for (const r of ruleConfig) {
    if (r.Source.trim().toLowerCase() === 'user_input') {
      const lab = r.Label.trim();
      if (!labels.includes(lab)) { labels.push(lab); primaryRuleFor[lab] = r; }
    }
  }

  for (const label of labels) {
    const rule = primaryRuleFor[label];
    const id = sanitize(label);

    let defaultVal = '';
    if (rule.default_value?.startsWith('<from_filename:'))       defaultVal = defaultMawb;
    else if (label.toUpperCase() === 'MAWB')                      defaultVal = defaultMawb;
    else                                                          defaultVal = rule.default_value || '';

    const fmt = (rule.Format || '').trim();
    const placeholder = fmt || '';

    const wrapper = document.createElement('div');

    if ((rule.has_dropdown || '').trim().toUpperCase() === 'Y') {
      const opts = (rule.dropdown_options || '').split(',').map(o => o.trim()).filter(Boolean);
      wrapper.innerHTML = `
        <label for="${id}" class="font-semibold block mb-1">${label}</label>
        <select id="${id}" class="border rounded px-2 py-1 w-full">
          <option value="">--Select--</option>
          ${opts.map(o=>`<option value="${o}"${o===defaultVal?' selected':''}>${o}</option>`).join('')}
        </select>`;
    } else {
      wrapper.innerHTML = `
        <label for="${id}" class="font-semibold block mb-1">${label}</label>
        <input type="text" id="${id}" value="${defaultVal}"
               ${placeholder?`placeholder="${placeholder}"`:''}
               data-format="${fmt}" class="ui-input w-full placeholder-gray-400" />`;
    }

    formEl.appendChild(wrapper);
  }

  // —— 覆盖：Port ——（所有含“Port”的输入 + “Firms Code” + “State of Destination”）
  if (portKey && PORT_MAP[portKey]) {
    const { code, firms, state } = PORT_MAP[portKey];
    labels.forEach(lab => {
      if (lab.toLowerCase().includes('port')) {
        const el = document.getElementById(sanitize(lab));
        if (el && el.tagName === 'INPUT') el.value = code;
        if (el && el.tagName === 'SELECT') {
          el.value = code;
          el.dispatchEvent(new Event('change', { bubbles: true })); // 让自绘下拉按钮也更新文字
        }
      }
    });
    const firmsEl = document.getElementById(sanitize('Firms Code'));
    if (firmsEl) {
      firmsEl.value = firms;
      if (firmsEl.tagName === 'SELECT') {
        firmsEl.dispatchEvent(new Event('change', { bubbles: true }));
      }
    }

    const stateEl = document.getElementById(sanitize('State of Destination'));
    if (stateEl) {
      stateEl.value = state;
      if (stateEl.tagName === 'SELECT') {
        stateEl.dispatchEvent(new Event('change', { bubbles: true }));
      }
    }
  }

  // —— 覆盖：Date ——（所有 Label 含 “Date”的输入，格式 m/d/yyyy）
  if (['today', 'tomorrow', 'day_after_tomorrow'].includes(dateKey)) {
    const base = new Date();
    if (dateKey === 'tomorrow') base.setDate(base.getDate() + 1);
    if (dateKey === 'day_after_tomorrow') base.setDate(base.getDate() + 2);
    const text = formatDateByPattern(base, 'm/d/yyyy');

    labels.forEach(lab => {
      if (lab.toLowerCase().includes('date')) {
        const el = document.getElementById(sanitize(lab));
        if (el && el.tagName === 'INPUT') el.value = text;
        if (el && el.tagName === 'SELECT') {
          el.value = text;
          el.dispatchEvent(new Event('change', { bubbles: true })); // 同步更新自绘下拉
        }
      }
    });
  }

  // 同步美化：文本输入应用与下拉一致的外观
  formEl.querySelectorAll('input:not([type=hidden]):not([type=checkbox]):not([type=radio]), textarea')
    .forEach(el => { if(!el.classList.contains('ui-input')) el.classList.add('ui-input'); });
// 自绘美化第二步表单里的所有下拉
  beautifyAllSelects(formEl, IS_SHEIN ? 'shein' : 'temu');
  // 统一文本输入与下拉的外观
  if (typeof beautifyAllTextInputs === 'function') beautifyAllTextInputs(formEl);

  // ===== 动态标红：空白文本框 & 未选择的下拉（含自绘下拉） =====
  const NORMAL_BORDER = '#d1d5db';
  const ERROR_BORDER  = 'red';

  // 自绘下拉对应的可见按钮（select 后面紧跟的 .ui-select 下的 .ui-select__btn）
  function getCustomSelectButton(sel) {
    const uiRoot = sel.nextElementSibling;
    if (uiRoot && uiRoot.classList && uiRoot.classList.contains('ui-select')) {
      return uiRoot.querySelector('.ui-select__btn');
    }
    return null;
  }

  function setBorder(el, isError) {
    if (el.tagName === 'INPUT' || el.tagName === 'TEXTAREA') {
      el.style.borderColor = isError ? ERROR_BORDER : NORMAL_BORDER;
      return;
    }
    if (el.tagName === 'SELECT') {
      const btn = getCustomSelectButton(el);
      if (btn) btn.style.borderColor = isError ? ERROR_BORDER : NORMAL_BORDER;
      else     el.style.borderColor  = isError ? ERROR_BORDER : NORMAL_BORDER;
    }
  }

  function validateOne(el) {
    const isEmpty = (el.value || '').trim() === '';
    setBorder(el, isEmpty);
  }

  function validateAll() {
    formEl.querySelectorAll('input[type="text"], textarea, select').forEach(validateOne);
  }

  // 初始检查（有默认值的不会红；空的会红）
  validateAll();

  // 动态监听：清空→变红；填入/选择→恢复
  formEl.querySelectorAll('input[type="text"], textarea').forEach(el => {
    el.addEventListener('input', () => validateOne(el));
    el.addEventListener('change', () => validateOne(el)); // 兼容
  });
  formEl.querySelectorAll('select').forEach(sel => {
    sel.addEventListener('change', () => validateOne(sel));
  });
  // ===== 结束：动态标红 =====

}

// ====== 生成逻辑（略做调整：导出文件名根据客户变化） ======
async function generateAndDownload() {
  const formValues = {};
  document.querySelectorAll('#dynamic-form input, #dynamic-form select')
    .forEach(el => formValues[el.id] = el.value.trim());

  const buf = await currentFile.arrayBuffer();
  let wb;
  try {
    wb = XLSX.read(buf, { type: 'array', password: EXCEL_PASSWORD });
  } catch (err) {
    return alert('Failed to open encrypted file: ' + err.message);
  }

  const sheetData = {};
  let mawbSheetArr = [];
  for (const name of wb.SheetNames) {
    const key = name.trim().toLowerCase();
    const ws  = wb.Sheets[name];
    if (key === 'hawb') {
      const raw    = XLSX.utils.sheet_to_json(ws, { header:1, defval:'' });
      const header = raw[1] || [];
      const rows   = raw.slice(2).filter(r => r.some(c => c !== ''));
      sheetData['hawb'] = rows.map(rw => {
        const o = {}; header.forEach((h,i) => o[h] = rw[i] || ''); return o;
      });
    } else if (key === 'mawb') {
      mawbSheetArr = XLSX.utils.sheet_to_json(ws, { header:1, defval:'' });
      sheetData['mawb'] = XLSX.utils.sheet_to_json(ws, { defval:'' });
    } else {
      sheetData[key] = XLSX.utils.sheet_to_json(ws, { defval:'' });
    }
  }
  // —— 新增：构建每个 sheet 的 AOA 视图，用于自动识别标题行 & 非标题区关键词查找 ——
  const sheetAOA = {};
  for (const name of wb.SheetNames) {
    const key = name.trim().toLowerCase();
    const ws = wb.Sheets[name];
    sheetAOA[key] = XLSX.utils.sheet_to_json(ws, { header:1, defval:'' });
  }

  const expectedHeadersBySheet = {};
  for (const cfg of ruleConfig) {
    if ((cfg.Source || '').toString().trim().toLowerCase() !== 'user_upload') continue;
    const sk = (cfg.Sheet || '').toString().trim().toLowerCase();
    const lookup = (cfg.Lookup || 'header').toString().toLowerCase();
    if (lookup === 'keyword_right') continue;
    const ref = (cfg.Reference || '').toString().trim();
    if (!ref) continue;
    (expectedHeadersBySheet[sk] ||= new Set()).add(ref);
  }

  const sheetViews = {};
  Object.keys(sheetAOA).forEach(sk => {
    const expected = Array.from(expectedHeadersBySheet[sk] || []);
    sheetViews[sk] = buildSheetView(sheetAOA[sk], expected);
  });


// ===== Determine primary data rows from CONFIG (revised logic) =====
let mainCount = 0;
let primarySheetKey = null;
const sheetCandidates = new Set();

// Helper: check if a cell value should be treated as empty/placeholder
const isEmptyVal = (v) => {
  const s = String(v ?? '').trim().toLowerCase();
  if (!s) return true;                    // empty or whitespace
  if (/^[-–—]+$/.test(s)) return true;    // only dashes
  if (s === 'na' || s === 'n/a' || s === 'null') return true;
  if (/^0+(\.0+)?$/.test(s)) return true; // 0, 0.0, 000
  return false;
};

// Helper function to check if a row is effectively empty - looks only at key columns when provided
const isRowEffectivelyEmpty = (row, keyIdxs) => {
  if (!row || !Array.isArray(row) || row.length === 0) return true;
  const cells = (keyIdxs && keyIdxs.length) ? keyIdxs.map(i => row[i]) : row;
  return cells.every(isEmptyVal);
};

for (const cfg of ruleConfig) {
  const src = (cfg.Source || '').toString().trim().toLowerCase();
  if (src !== 'user_upload') continue;
  const lookup = (cfg.Lookup || 'header').toString().toLowerCase();
  if (lookup === 'keyword_right') continue;
  const sk = (cfg.Sheet || '').toString().trim().toLowerCase();
  if (!sk || sk === 'mawb') continue;
  sheetCandidates.add(sk);
}

if (sheetCandidates.size > 0) {
  sheetCandidates.forEach(sk => {
    const view = sheetViews[sk];
    const aoa = sheetAOA[sk] || [];
    if (view && view.headerRowIdx >= 0) {
      const dataRows = aoa.slice(view.headerRowIdx + 1);
      // Determine key columns for this sheet using expected header names
      const expected = Array.from(expectedHeadersBySheet[sk] || []);
      const headerRow = aoa[view.headerRowIdx] || [];
      const norm = (x) => String(x ?? '').trim().toLowerCase();
      const keyIdxs = expected
        .map(h => headerRow.findIndex(x => norm(x) === norm(h)))
        .filter(i => i >= 0);
      // Filter out rows where all key columns are effectively empty
      const nonEmptyRows = dataRows.filter(row => !isRowEffectivelyEmpty(row, keyIdxs));
      const count = nonEmptyRows.length;
      
      if (count > mainCount) {
        mainCount = count;
        primarySheetKey = sk;
        // remember key columns for primary sheet
        window.__primaryKeyIdxs = keyIdxs.slice();
      }
    }
  });
}

// 'main' only provides length for the primary loop
let main = [];
if (mainCount > 0) {
  main = new Array(mainCount).fill(0);
} else {
  alert('Could not find any data rows in the uploaded file.\n' +
        '- Please ensure at least one sheet has a valid header and data rows.\n' +
        '- The configuration must point to a header column in that sheet (Source=user_upload and Lookup!=keyword_right).');
  return;
}

  const output = [];
  const prog = document.getElementById('progress');
  const pt   = document.getElementById('progress-text');
  document.getElementById('progress-container').classList.remove('hidden');
  // reset Error_fix hits for this run
  window.__errorFixRows = [];


  for (let i = 0; i < main.length; i++) {
    const out = {};

    for (const cfg of ruleConfig) {
      const col = cfg.Column;
      const src = (cfg.Source || '').trim().toLowerCase();

      if (src === 'fixed') {
        out[col] = cfg.Value || '';
      } else if (src === 'user_upload') {
        const sk = (cfg.Sheet || '').trim().toLowerCase();
        const lookup = (cfg.Lookup || 'header').toString().toLowerCase();
        let value = '';
        if (sk === 'mawb' && cfg.Reference && (lookup === 'header' || lookup === 'singleton')) {
          value = getValueFromMawbSheet(mawbSheetArr, cfg.Reference);
        } else {
          const view = sheetViews[sk];
          if (view) {
            if (lookup === 'keyword_right') {
              value = view.getByKeywordRight(cfg.Reference);
            } else { // 默认 header
              value = view.getByHeaderRow(i, cfg.Reference);
            }
          } else {
            // 兜底：沿用旧逻辑
            const arr = sheetData[sk] || [];
            const row = arr[i] || {};
            const refLower = (cfg.Reference || '').toString().trim().toLowerCase();
            const key = Object.keys(row).find(k => k.toString().trim().toLowerCase() === refLower);
            value = key ? row[key] : '';
          }
        }
        out[col] = parseValue(value, cfg.Parsing);      } else if (src === 'user_input') {
        const label = cfg.Label?.trim() || '';
        let v = formValues[sanitize(label)] || '';
        out[col] = parseValue(v, cfg.Parsing);
      } else if (src === 'system') {
        const d = new Date();
        const fmt = (cfg.Format || '').trim();
        out[col] = fmt ? formatDateByPattern(d, fmt) : `${d.getFullYear()}/${d.getMonth() + 1}/${d.getDate()}`;
      }
    }

    // HTS 映射（按规则长度做前缀匹配，优先最长命中）
    (() => {
      const raw = (out.HTS || '').toString();
      // 只取数字，避免有空格、点或者横杠
      const rawDigits = (raw.match(/\d+/g) || []).join('');
      if (!rawDigits) return;
    
      let best = null;
      let bestLen = -1;
    
      for (const rule of htsData) {
        const ruleDigits = (rule.HTS || '').toString().replace(/\D/g, '');
        if (!ruleDigits) continue;
        const n = ruleDigits.length;
        // 只要“左 n 位”等于规则的 HTS，就算匹配
        if (rawDigits.slice(0, n) === ruleDigits) {
          if (n > bestLen) {
            best = rule;
            bestLen = n;
          }
        }
      }
    
      if (best) {
        Object.keys(best)
          .filter(k => /^HTS-\d+$/.test(k)) // 仅匹配 HTS-数字
          .sort((a, b) => parseInt(a.split('-')[1]) - parseInt(b.split('-')[1])) // 递增
          .forEach(k => { out[k] = best[k]; });
      }
    })();

    // MID 映射并清空（支持“Replace MID...” 开关的例外）
    (() => {
      const nm  = (out.ManufacturerName || '').trim();

      // 计算“导出文件行号（含表头）”：header=第1行，数据从第2行开始
      const exportRowNumber = i + 2;

      const keepNameAddrThisRow =
        !!window.__midReplaceEnabled &&
        window.__midReplaceRowsSet instanceof Set &&
        window.__midReplaceRowsSet.has(exportRowNumber);

      // 如果此行需要“保留名称与地址、删除 MID”，则不执行映射清空逻辑
      if (keepNameAddrThisRow) return;

      const hit = midData.find(r => nm.includes(r.ManufacturerName));
      if (hit) {
        out.ManufacturerCode = hit.ManufacturerCode || '';
        ['ManufacturerName','ManufacturerStreetAddress','ManufacturerCity','ManufacturerPostalCode','ManufacturerCountry']
          .forEach(f => out[f] = '');
      }
    })();


    // PGA（按 FDAPRODUCTCODE 最长前缀匹配；支持 Anything_else + Description_contain；Error_fix 记录原始值）
    (() => {
      const rawCode = (out.FDAPRODUCTCODE || '').toString();
      if (!rawCode) return;

      // 应用规则并在发生 Error_fix 时记录原值+高亮行信息
      const tryApplyRule = (rule) => {
        if (!rule) return false;

        const orig = out.FDAPRODUCTCODE;

        // 1) Error_fix：前缀替换（长度 = Error_fix 长度）
        if (rule.Error_fix) {
          const fix = String(rule.Error_fix);
          const n = fix.length;
          out.FDAPRODUCTCODE = fix + rawCode.slice(n);

          if (orig && orig !== out.FDAPRODUCTCODE) {
            out.Original_ProductCode = orig;
            (window.__errorFixRows ||= []).push({
              row: i,                              // 当前输出行索引（从 0 开始）
              codeBefore: orig,                        // FDAPRODUCTCODE before Error_fix
              codeAfter: out.FDAPRODUCTCODE,           // FDAPRODUCTCODE after Error_fix
              descHeader: rule.Description_contain || '' // 用于后续找“对应 header 的输出列”
            });
          }
        }

        // 2) Delete_code 优先级最高
        if (rule.Delete_code === 'Y') {
          out.FDAPRODUCTCODE = '';
          Object.keys(rule).forEach(k => {
            if (!['FDAPRODUCTCODE','Delete_code','Error_fix','Description_contain'].includes(k)) {
              out[k] = '';
            }
          });
          return true;
        }

        // 3) 复制其它字段（跳过标识字段；仅写有值的项）
        Object.entries(rule).forEach(([k, v]) => {
          if (['FDAPRODUCTCODE','Delete_code','Error_fix','Description_contain'].includes(k)) return;
          if (v !== undefined && v !== null && v !== '') out[k] = v;
        });

        return true;
      };

      // A) 先做“最长前缀”常规匹配（显式排除 Anything_else）
      let best = null, bestLen = -1;
      for (const rule of pgaRules) {
        const key = (rule.FDAPRODUCTCODE || '').toString();
        if (!key || key === 'Anything_else') continue;
        const n = key.length;
        if (rawCode.slice(0, n) === key && n > bestLen) {
          best = rule; bestLen = n;
        }
      }
      if (best) { tryApplyRule(best); return; }

      // B) 兜底：匹配 FDAPRODUCTCODE = "Anything_else"
      // 支持 Description_contain: "Header: keyword, keyword"
      for (const rule of pgaRules.filter(r => (r.FDAPRODUCTCODE || '') === 'Anything_else')) {
        let pass = true;
      //Anything_else支持多个关键词，Header: keyword, keyword
        if (rule.Description_contain) {
          const [headerPart, keywordPart] = rule.Description_contain.split(':');
          if (headerPart && keywordPart) {
            const header = headerPart.trim();
            const keywords = keywordPart.split(',')
                                        .map(s => s.trim().toLowerCase())
                                        .filter(Boolean);
            const view = sheetViews[primarySheetKey];
            const cellVal = (view?.getByHeaderRow(i, header) || '').toLowerCase();
            pass = keywords.some(k => cellVal.includes(k));
          } else {
            pass = false;
          }
        }
          
        if (pass) { tryApplyRule(rule); return; }
      }
    })();

    // —— 校验 ManufacturerPostalCode（所有客户适用） ——
    // 规则：若包含任何字母，直接替换为 528000；否则若数字长度 < 6（且存在数字）或全部为 0，则替换为 528000
    {
      const raw = (out.ManufacturerPostalCode ?? '').toString();

      // 新增：若包含任何字母，直接替换为 528000
      const hasLetter = /[A-Za-z]/.test(raw); // 如需支持全部 Unicode 字母，可改为 /\p{L}/u
      if (hasLetter) {
        out.ManufacturerPostalCode = '528000';
      } else {
        const digits = (raw.match(/\d+/g) || []).join('');  // 只取数字
        const allZero = /^0+$/.test(digits);
        const tooShort = (digits.length > 0 && digits.length < 6);
        if (tooShort || allZero) {
          out.ManufacturerPostalCode = '528000';
        }
      }
    }

    // ===== 最终收口：根据开关与用户指定的“包含表头的行号”决定清空哪几列 =====
    (function finalizeManufacturerColumns() {
      const exportRowNumber = i + 2; // 1=表头，2起为数据
      const five = ['ManufacturerName','ManufacturerStreetAddress','ManufacturerCity','ManufacturerPostalCode','ManufacturerCountry'];

      if (window.__midReplaceEnabled) {
        const inKeepSet = window.__midReplaceRowsSet instanceof Set && window.__midReplaceRowsSet.has(exportRowNumber);
        if (inKeepSet) {
          // 这些行：删除 MID（ManufacturerCode），保留五列
          out.ManufacturerCode = '';
        } else {
          // 其他行：删除五列，保留/映射到的 MID
          five.forEach(f => { out[f] = ''; });
        }
      } else {
        // 开关关闭：始终删除五列
        five.forEach(f => { out[f] = ''; });
      }
    })();


    // —— 基于主驱动 sheet 的真实行判空，整行仅空白或0就跳过（避免尾部空壳行） ——
    (function skipIfPrimaryRowEmpty() {
      const view = sheetViews[primarySheetKey];
      if (!view || view.headerRowIdx < 0) return; // 没识别到主表就不拦

      const aoa = sheetAOA[primarySheetKey] || [];
      const row = aoa[view.headerRowIdx + 1 + i] || [];   // 定位真实数据行

      // evaluate emptiness based on primary key columns if available
      const keyIdxs = (window.__primaryKeyIdxs && window.__primaryKeyIdxs.length)
        ? window.__primaryKeyIdxs : null;
      const isEmpty = isRowEffectivelyEmpty(row, keyIdxs);

      if (isEmpty) {
        window.__skipRow = true;
      }
    })();
    if (window.__skipRow) { window.__skipRow = false; continue; }

    // 基于主列 HTS + GrossWeight/Manifest Qty Piece count，计算两个数量（没列名时不会写出）
    attachHtsQty(out);
    
    // —— 新增：单位转换兜底 ——
    // 若 HTSQty 为空（比如该 HTS 在 unit.json 没匹配到单位），则用 PCS 填充
    (function fillHtsQtyWithPcsIfEmpty(){
      const isBlank = v => v == null || String(v).trim() === '';
      if (isBlank(out.HTSQty)) {
        const raw = out['Manifest Qty Piece count'];
        const pcsNum = Number(String(raw ?? '').replace(/,/g, '')); // 转成数字
        if (!Number.isNaN(pcsNum)) {
          out.HTSQty = pcsNum.toFixed(3); // 保持三位小数格式，和 attachHtsQty 计算一致
        }
      }
    })();

    // 若勾选“Auto-adjust PGA QTY...”：FDAPRODUCTCODE 有值且 HTSQty < 1 的行，把 HTSQty 调为 1
    if (window.__roundUpPgaQtyEnabled) {
      const hasFda = (out.FDAPRODUCTCODE ?? '') !== '';
      const qtyNum = Number(String(out.HTSQty ?? '').replace(/,/g, ''));
      if (hasFda && isFinite(qtyNum) && qtyNum < 1) out.HTSQty = 1;
    }
    
    output.push(out);

    if ((i + 1) % 20 === 0 || i === main.length - 1) {
      const pct = Math.round(((i + 1) / main.length) * 100);
      pt.innerText = `${pct}%`; prog.style.width = `${pct}%`;
      await new Promise(r => setTimeout(r, 0));
    }
  }


  // === SHEIN: 生成 GroupIdentifier（每组≤998，尽量均分，组号从1开始） ===
  // === 分组：SHEIN 或 TEMU-合单都应用 ===
  {
    const needGrouping = IS_SHEIN || window.__isConsolidatedShipment;
    if (needGrouping) {
      // TEMU-合单 → House 约束；SHEIN（非合单）→ 均匀分组
      applySheinGrouping(output, { respectHouseAwb: !!window.__isConsolidatedShipment });
    }
  }

  // 写回
  const header = ruleConfig.map(r => r.Column);

  // --- 动态把所有出现过的 HTS-* 列插到 HTS 右侧 ---
  const htsIndex = header.indexOf('HTS');
  if (htsIndex !== -1) {
    const htsDynamicCols = Array.from(
      new Set(
        output.flatMap(o => Object.keys(o || {}).filter(k => /^HTS-\d+$/.test(k)))
      )
    ).sort((a, b) => parseInt(a.split('-')[1]) - parseInt(b.split('-')[1])); // 递增

    // 从后往前插，避免索引位移
    htsDynamicCols.slice().reverse().forEach(col => {
      if (!header.includes(col)) header.splice(htsIndex + 1, 0, col);
    });
}


  // 如果任意一行存在 Original_ProductCode，则把该列插在 FDAPRODUCTCODE 右侧（若 FDAPRODUCTCODE 不在表头则追加到末尾）
  if (output.some(o => o && typeof o === 'object' && o.Original_ProductCode) && !header.includes('Original_ProductCode')) {
    const idxFDA = header.indexOf('FDAPRODUCTCODE');
    if (idxFDA >= 0) header.splice(idxFDA + 1, 0, 'Original_ProductCode');
    else header.push('Original_ProductCode');
  }

  // SHEIN 或 TEMU-合单都需要导出 GroupIdentifier
  if ((IS_SHEIN || window.__isConsolidatedShipment) && !header.includes('GroupIdentifier')) {
    header.push('GroupIdentifier');
}

  // === 最后一步（导出前）：Consolidated 模式下把 "Manifest Qty Piece count" 全部置为 1 ===
  if (window.__isConsolidatedShipment) {
    const targetCol = 'Manifest Qty Piece count';
    // 只有当最终导出的表头里包含该列时才执行覆盖
    if (header.includes(targetCol)) {
      for (const row of output) {
        if (row && typeof row === 'object') {
          row[targetCol] = 1;
        }
      }
    }
  }


  // 然后再生成 aoa 与工作表
  const aoa = [header].concat(output.map(o => header.map(c => (o[c] ?? ''))));
  const ws2 = XLSX.utils.aoa_to_sheet(aoa);

  // === 日期/时间列（按配置 Format + 列名定位，适配动态表头顺序变动）===
  (function applyDateTimeFormatsByHeader() {
    // 1) 从配置里收集需要设定为日期/时间的列及其格式
    //    规则：Format 同时包含 y / m / d（大小写不敏感）则视为日期/时间列
    const colFmtMap = {};
    ruleConfig.forEach(r => {
      const fmt = String(r.Format || '');
      const f = fmt.toLowerCase();
      if (f && /y/.test(f) && /m/.test(f) && /d/.test(f)) {
        colFmtMap[r.Column] = fmt; // 保留原格式串，如 "m/d/yyyy" 或 "m/d/yyyy h:mm"
      }
    });

    // 没有需要处理的列就直接返回
    const names = Object.keys(colFmtMap);
    if (!names.length) return;

    // 2) 依据“列名”在 header 里找出当前的列索引（适配 HTS-* 动态插列后的新顺序）
    const targets = names
      .map(name => ({ name, fmt: colFmtMap[name], idx: header.indexOf(name) }))
      .filter(x => x.idx >= 0);

    if (!targets.length) return;

    // 3) 逐行写入单元格类型与格式
    for (let r = 1; r < aoa.length; r++) {
      for (const t of targets) {
        const c = t.idx;
        const colLetter = XLSX.utils.encode_col(c);
        const cellRef = colLetter + (r + 1);
        const raw = aoa[r][c];

        if (raw == null || raw === '') continue;

        // 尝试解析为日期（兼容 "m/d/yyyy"、"m-d-yyyy"、以及已有的 Date 对象）
        let d = (raw instanceof Date) ? raw : new Date(raw);
        if (isNaN(d.getTime())) {
          // 简单兜底：匹配 m/d/yyyy 或 m-d-yyyy
          const m = String(raw).match(/(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})(?:[ T](\d{1,2}):(\d{2})(?::(\d{2}))?)?/);
          if (m) {
            const mm = parseInt(m[1], 10) - 1;
            const dd = parseInt(m[2], 10);
            const yyyy = parseInt(m[3], 10);
            const hh = m[4] ? parseInt(m[4], 10) : 0;
            const min = m[5] ? parseInt(m[5], 10) : 0;
            const ss = m[6] ? parseInt(m[6], 10) : 0;
            d = new Date(yyyy, mm, dd, hh, min, ss);
          }
        }
        if (isNaN(d.getTime())) continue;

        // 写入 Excel 单元格为日期/时间，并使用配置中的格式串
        if (!ws2[cellRef]) ws2[cellRef] = {};
        ws2[cellRef].t = 'd';
        ws2[cellRef].v = d;
        ws2[cellRef].z = String(t.fmt).replace(/Y/g, 'y').replace(/D/g, 'd'); // 兼容大小写
        // 同步回 AOA（可选）
        aoa[r][c] = d;
      }
    }
  })();


// === 新：FDAARRIVALTIME 列强制为“Time”格式（h:mm） ===
(function formatFDAArrivalTimeAsTime() {
  const colIndex = header.indexOf('FDAARRIVALTIME');
  if (colIndex === -1) return;

  for (let r = 1; r < aoa.length; r++) {
    const raw = aoa[r][colIndex];
    if (raw == null || raw === '') continue;

    const str = String(raw).trim();
    // 兼容 "H:MM"、"HH:MM"、可选秒、以及 AM/PM
    const m = str.match(/^(\d{1,2}):(\d{2})(?::(\d{2}))?\s*(AM|PM|am|pm)?$/);
    if (!m) continue;

    let h = parseInt(m[1], 10);
    let mm = parseInt(m[2], 10);
    let ss = parseInt(m[3] || '0', 10);
    const ap = (m[4] || '').toLowerCase();

    if (ap === 'pm' && h < 12) h += 12;
    if (ap === 'am' && h === 12) h = 0;

    // Excel 时间序列（一天 = 1）
    const num = (h * 3600 + mm * 60 + ss) / 86400;

    const cellRef = XLSX.utils.encode_col(colIndex) + (r + 1);
    if (!ws2[cellRef]) ws2[cellRef] = {};
    ws2[cellRef].t = 'n';
    ws2[cellRef].v = num;
    ws2[cellRef].z = 'h:mm';

    // 同步 AOA
    aoa[r][colIndex] = num;
  }
})();

  // ==== 把 HTSValue / GrossWeight 列写成数值；仅 HTSValue 保留两位小数 ====
  // 扩展性：往 numericHeaders 里加更多表头名即可复用
  (function formatNumericColumns() {
    const numericHeaders = new Set(['HTSValue', 'GrossWeight', 'HTSQty', 'HTSQty2']);

    // 工具：从原始字符串里推断小数位数（尽量保留原精度）
    function inferDecimals(raw) {
      if (raw === undefined || raw === null) return 0;
      // 优先使用原始字符串推断（可以保留诸如 12.340 的尾零）
      if (typeof raw === 'string') {
        const s = raw.replace(/,/g, '').trim();
        const m = s.match(/^[+-]?\d+(?:\.(\d+))?$/);
        return m && m[1] ? m[1].length : 0;
      }
      // 若原本就是数值，只能用字符串化结果推断（可能无法区分尾零）
      const s = String(raw);
      // 科学计数法时不强行设定位数，交给 Excel 自己显示
      if (/e|E/.test(s)) return 0;
      const m = s.match(/^[+-]?\d+(?:\.(\d+))?$/);
      return m && m[1] ? m[1].length : 0;
    }

    // 工具：根据小数位生成格式串（避免显示为 General）
    function fmtFromDecimals(dec) {
      return dec > 0 ? '0.' + '0'.repeat(dec) : '0';
    }

    for (let col = 0; col < header.length; col++) {
      const headerName = header[col];
      if (!numericHeaders.has(headerName)) continue;

      const isHTS = headerName === 'HTSValue';

      for (let r = 1; r < aoa.length; r++) {
        const rowIdx = r + 1; // Excel 行号（含表头）
        const c = XLSX.utils.encode_col(col);
        const cellRef = c + rowIdx;

        let raw = aoa[r][col];
        if (raw === undefined || raw === null || raw === '') continue;

        // 先推测原始小数位（在转 Number 之前）
        const decPlaces = isHTS ? 2 : inferDecimals(raw);

        // 去除逗号并转数值
        if (typeof raw === 'string') raw = raw.replace(/,/g, '').trim();
        let num = Number(raw);
        if (!isFinite(num)) continue;

        // HTSValue 强制两位小数（数值型）并固定显示格式 0.00
        // 其他列不改数值，仅按原小数位设置显示格式
        let zfmt;
        if (isHTS) {
          num = Math.round(num * 100) / 100;
          zfmt = '0.00';
        } else {
          zfmt = fmtFromDecimals(decPlaces); // 如 0 / 0.0 / 0.000 ...
        }

        // 写回 sheet 单元格（确保是数值 t:'n'），并设置 z 避免 General
        if (!ws2[cellRef]) {
          ws2[cellRef] = { t: 'n', v: num, z: zfmt };
        } else {
          ws2[cellRef].t = 'n';
          ws2[cellRef].v = num;
          ws2[cellRef].z = zfmt;
        }

        // 同步回 AOA（可选）
        aoa[r][col] = num;
      }
    }
  })();


  const wb2 = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb2, ws2, 'Sheet1');

  const mawbOrig = (document.getElementById(sanitize('MAWB'))?.value || currentDefaultMawb || '').trim();
  const tag = IS_SHEIN ? 'SHEIN' : 'TEMU';

  // ==== Show red warning banner in the page if any Error_fix applied (EN only) ====
  (function showErrorFixBanner(){
    const pc = document.getElementById('progress-container');
    if (!pc) return;

    // remove old banner if exists
    const old = document.getElementById('error-fix-banner');
    if (old) old.remove();

    const hits = Array.isArray(window.__errorFixRows) ? window.__errorFixRows : [];
    if (!hits.length) return;

    // build banner
    const banner = document.createElement('div');
    banner.id = 'error-fix-banner';
    banner.style.color = '#b91c1c';          // red-700
    banner.style.fontWeight = '700';
    banner.style.marginTop = '10px';
    banner.style.lineHeight = '1.5';

    const summary = document.createElement('div');
    summary.textContent = `⚠ ${hits.length} row(s) had FDAPRODUCTCODE automatically modified. Please review, then delete column Original_ProductCode.`;
    banner.appendChild(summary);

    const details = document.createElement('div');
    details.style.fontWeight = '600';
    details.style.marginTop = '6px';
    details.innerHTML = hits.map(h =>
      `Row ${h.row + 1}: ${h.codeBefore || '(empty)'} → ${h.codeAfter || '(empty)'}`
    ).join('<br>');
    banner.appendChild(details);

    pc.appendChild(banner);

    // avoid sticking to the very top edge when scrolling
    banner.style.scrollMarginTop = '120px';

    // smooth-scroll after the DOM has painted
    setTimeout(() => {
      try { banner.scrollIntoView({ behavior: 'smooth', block: 'center' }); } catch(_) {}
    }, 100);
  })();


  let outFileName = `${mawbOrig}_NETChb_${tag}.xlsx`;

  // 如果开启了 Replace MID 开关，加前缀 Fixed_
  if (window.__midReplaceEnabled) {
    outFileName = `Fixed_${outFileName}`;
  }

  // 如果勾选了“Auto-adjust PGA QTY…”，文件名加后缀 _RoundedUp
  if (window.__roundUpPgaQtyEnabled) {
    outFileName = outFileName.replace(/\.xlsx$/i, '_RoundedUp.xlsx');
  }

  XLSX.writeFile(wb2, outFileName);

}


/** 统一美化文本/日期/数字输入框，使其与自绘下拉(btn)完全一致的半径/边框/阴影 */
function beautifyAllTextInputs(container){
  if (!container) return;
  const inputs = container.querySelectorAll('input:not([type=hidden]):not([type=checkbox]):not([type=radio]), textarea');
  inputs.forEach(el => {
    if (el.dataset.uiTxt === '1') return;
    el.dataset.uiTxt = '1';
    // 基础外观（与 .ui-select__btn 对齐）
    el.style.border = '1px solid #d1d5db';             // slate-300
    el.style.borderRadius = '12px';                    // 圆角与下拉一致
    el.style.padding = '10px 14px';                    // 与下拉近似（下拉右侧有箭头多 26px）
    el.style.background = '#fff';
    el.style.boxShadow = '0 1px 2px rgba(16,24,40,.05)';
    el.style.transition = 'box-shadow .15s, border-color .15s';
    el.style.outline = 'none';
    el.addEventListener('focus', () => {
      el.style.boxShadow = '0 0 0 3px rgba(148,163,184,.25)';   // 与下拉聚焦外光一致
      el.style.borderColor = '#d1d5db';
    });
    el.addEventListener('blur',  () => {
      el.style.boxShadow = '0 1px 2px rgba(16,24,40,.05)';
      // 如果值为空，保持红色；否则恢复灰色
      if ((el.value || '').trim() === '') {
        el.style.borderColor = 'red';
      } else {
        el.style.borderColor = '#d1d5db';
      }
    });
  });
}

document.addEventListener('DOMContentLoaded', () => {
  const sections = [document.getElementById('dynamic-form'), document.getElementById('upload-section')];
  sections.forEach(sec => { if (sec) { try { beautifyAllTextInputs(sec); } catch(e){} } });
});


// Apply input skin globally on load
document.addEventListener('DOMContentLoaded', () => {
  const accent = IS_SHEIN ? 'emerald' : 'blue';
  beautifyAllTextInputs(document);
  const df = document.getElementById('dynamic-form');
  if (df && typeof observeNewInputs === 'function') observeNewInputs(df, 'neutral');
});

