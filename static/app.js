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
const IS_SHEIN = (document.body.getAttribute('data-client') === 'shein');

// 根据客户选择不同的配置文件名（SHEIN 使用 *_shein.json）
const CONFIG_FILES = IS_SHEIN
  ? { rule: 'rule_shein.json', hts: 'hts_shein.json', mid: 'mid_shein.json', pga: 'PGA_shein.json' }
  : { rule: 'rule.json',       hts: 'hts.json',        mid: 'mid.json',        pga: 'PGA.json'       };

// 全局保存当前文件、默认 MAWB、预选项
let currentFile = null;
let currentDefaultMawb = '';
let selectedPortKey = '';
let selectedDateKey = ''; // '', 'today', 'tomorrow'

// 防缓存
const ts = Date.now();
const CONFIG_PATH = 'config';

// 入口页或 shein/temu 页都有这些元素（入口页不会加载 app.js）
const uploadBtn   = document.getElementById('upload-btn');
const fileInput   = document.getElementById('file-input');
const loadingMsg  = document.getElementById('loading-msg');
const continueBtn = document.getElementById('continue-btn');
const portSel     = document.getElementById('pref-port');
const dateSel     = document.getElementById('pref-date');
const generateBtn = document.getElementById('generate-btn');
// 统一 TEMU 主题色（#0071bc）
(function applyTemuTheme(){
  if (typeof IS_SHEIN === 'undefined' || IS_SHEIN) return;
  const temuBlue = '#0071bc';
  const h1 = document.querySelector('h1');
  if (h1) h1.style.color = temuBlue;
  const cont = document.getElementById('continue-btn');
  if (cont) { cont.style.backgroundColor = temuBlue; cont.style.borderColor = temuBlue; }
  const gen = document.getElementById('generate-btn');
  if (gen) { gen.style.backgroundColor = temuBlue; gen.style.borderColor = temuBlue; }
  const activeNav = document.querySelector('aside a[aria-current="page"]') || document.querySelector('aside a.scale-105');
  if (activeNav) { activeNav.style.background = temuBlue; activeNav.style.color = '#fff'; }
})();

let ruleConfig = [], htsData = [], midData = [], pgaRules = [];


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
  if (!parsing || parsing.toLowerCase() === 'raw') return val;
  const leftMatch = parsing.match(/^left\((\d+)\)$/i);
  if (leftMatch) return (val || '').toString().slice(0, parseInt(leftMatch[1], 10));
  const rightMatch = parsing.match(/^right\((\d+)\)$/i);
  if (rightMatch) return (val || '').toString().slice(-parseInt(rightMatch[1], 10));
  return val;
}

function buildRegex(fmt) {
  let regexStr = fmt.replace(/([.+?^=!:${}()|[\]\/\\])/g, '\\$1');
  regexStr = regexStr.replace(/y{4}/g, '\\d{4}');
  regexStr = regexStr.replace(/m{1,2}/gi, '\\d{1,2}');
  regexStr = regexStr.replace(/d{1,2}/gi, '\\d{1,2}');
  return new RegExp('^' + regexStr + '$');
}

// 加载配置（仅客户页面会有 uploadBtn）
if (uploadBtn) {
  Promise.all([
    fetch(`${CONFIG_PATH}/${CONFIG_FILES.rule}?ts=${ts}`).then(r => r.json()),
    fetch(`${CONFIG_PATH}/${CONFIG_FILES.hts}?ts=${ts}`).then(r => r.json()),
    fetch(`${CONFIG_PATH}/${CONFIG_FILES.mid}?ts=${ts}`).then(r => r.json()),
    fetch(`${CONFIG_PATH}/${CONFIG_FILES.pga}?ts=${ts}`).then(r => r.json())
  ]).then(([rule, hts, mid, pga]) => {
    ruleConfig = rule; htsData = hts; midData = mid; pgaRules = pga;
    uploadBtn.disabled = false;
    uploadBtn.classList.remove('opacity-50');
    loadingMsg.innerText = '';
  }).catch(e => {
    console.error('Failed to load configs', e);
    loadingMsg.innerText = 'Failed to load configuration';
  });

  // 选择文件
  uploadBtn.addEventListener('click', () => { try { fileInput.value = ''; } catch (_) {} fileInput.click(); });
  fileInput.addEventListener('change', () => {
    if (!fileInput.files.length) { alert('Please select a file'); return; }
    currentFile = fileInput.files[0];
    const base = currentFile.name.replace(/\.(xlsx|xls|csv)$/i, '');
    const m = base.match(/(\d{11})$/);
    currentDefaultMawb = m ? m[1] : '';
    continueBtn && (continueBtn.disabled = false);

    // 上传成功提示（英文）
    if (uploadBtn) {
      if (uploadBtn.querySelector && uploadBtn.querySelector('span')) {
        uploadBtn.querySelector('span').textContent = '✅ File uploaded successfully';
      } else {
        uploadBtn.innerHTML = '✅ File uploaded successfully';
      }
    }
  });

  // 记录 Port/Date 选择
  portSel && portSel.addEventListener('change', () => selectedPortKey = portSel.value.trim());
  dateSel && dateSel.addEventListener('change', () => selectedDateKey = dateSel.value.trim());

  // Continue：进入表单页并渲染
  continueBtn && continueBtn.addEventListener('click', () => {
    if (!currentFile) { alert('Please select a file'); return; }
    document.getElementById('upload-section').classList.add('hidden');
    document.getElementById('form-section').classList.remove('hidden');
    renderForm(currentDefaultMawb, { portKey: selectedPortKey, dateKey: selectedDateKey });
  });

  // 生成下载
  generateBtn.addEventListener('click', () => {
    if (!currentFile) { alert('No file selected'); return; }
    generateAndDownload();
  });
}

// 渲染动态表单，并根据 Port/Date 做覆盖
function renderForm(defaultMawb, { portKey = '', dateKey = '' } = {}) {
  const formEl = document.getElementById('dynamic-form');
  formEl.innerHTML = '';

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
        if (el && el.tagName === 'SELECT') el.value = code;
      }
    });
    const firmsEl = document.getElementById(sanitize('Firms Code'));
    if (firmsEl) firmsEl.value = firms;

    const stateEl = document.getElementById(sanitize('State of Destination'));
    if (stateEl) stateEl.value = state;
  }

  // —— 覆盖：Date ——（所有 Label 含 “Date”的输入，格式 m/d/yyyy）
  if (dateKey === 'today' || dateKey === 'tomorrow') {
    const base = new Date();
    if (dateKey === 'tomorrow') base.setDate(base.getDate() + 1);
    const text = formatDateByPattern(base, 'm/d/yyyy');

    labels.forEach(lab => {
      if (lab.toLowerCase().includes('date')) {
        const el = document.getElementById(sanitize(lab));
        if (el && (el.tagName === 'INPUT' || el.tagName === 'SELECT')) el.value = text;
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
}

// ====== 下面开始是你原先的生成逻辑（略做调整：导出文件名根据客户变化） ======
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

  const main = sheetData['hawb'] || [];
  const output = [];
  const prog = document.getElementById('progress');
  const pt   = document.getElementById('progress-text');
  document.getElementById('progress-container').classList.remove('hidden');

  for (let i = 0; i < main.length; i++) {
    const out = {};

    for (const cfg of ruleConfig) {
      const col = cfg.Column;
      const src = (cfg.Source || '').trim().toLowerCase();

      if (src === 'fixed') {
        out[col] = cfg.Value || '';
      } else if (src === 'user_upload') {
        const sk = (cfg.Sheet || '').trim().toLowerCase();
        let value;
        if (sk === 'mawb' && cfg.Reference) value = getValueFromMawbSheet(mawbSheetArr, cfg.Reference);
        else {
          const arr = sheetData[sk] || [];
          const row = sk === 'mawb' ? (arr[0] || {}) : (arr[i] || {});
          value = row[cfg.Reference] || '';
        }
        out[col] = parseValue(value, cfg.Parsing);
      } else if (src === 'user_input') {
        const label = cfg.Label?.trim() || '';
        let v = formValues[sanitize(label)] || '';
        out[col] = parseValue(v, cfg.Parsing);
      } else if (src === 'system') {
        const d = new Date();
        const fmt = (cfg.Format || '').trim();
        out[col] = fmt ? formatDateByPattern(d, fmt) : `${d.getFullYear()}/${d.getMonth() + 1}/${d.getDate()}`;
      }
    }

    // HTS 映射
    (() => {
      const raw  = (out.HTS || '').toString();
      const digs = (raw.match(/\d+/g) || []).join('').slice(0, 8);
      const hit  = htsData.find(r => r.HTS === digs);
      if (hit) ['HTS-1','HTS-2','HTS-3','HTS-4','HTS-5'].forEach(c => out[c] = hit[c] || '');
    })();

    // MID 映射并清空
    (() => {
      const nm  = (out.ManufacturerName || '').trim();
      const hit = midData.find(r => nm.includes(r.ManufacturerName));
      if (hit) {
        out.ManufacturerCode = hit.ManufacturerCode || '';
        ['ManufacturerName','ManufacturerStreetAddress','ManufacturerCity','ManufacturerPostalCode','ManufacturerCountry']
          .forEach(f => out[f] = '');
      }
    })();

    // PGA
    (() => {
      const code = (out.FDAPRODUCTCODE || '').toString().slice(0, 2);
      for (const rule of pgaRules) {
        if (code === rule.FDAPRODUCTCODE) {
          if (rule.Delete_code === 'Y') {
            out.FDAPRODUCTCODE = '';
            Object.keys(rule).forEach(k => { if (k!=='FDAPRODUCTCODE' && k!=='Delete_code') out[k]=''; });
          } else {
            Object.entries(rule).forEach(([k, v]) => { if (k!=='FDAPRODUCTCODE' && k!=='Delete_code' && v) out[k]=v; });
          }
          break;
        }
      }
    })();

    // —— SHEIN 专属：校验 ManufacturerPostalCode ——
    // 规则：仅在 IS_SHEIN 为真时检查；若数字长度 < 6（且存在数字）或全部为 0，则替换为 528000
    if (IS_SHEIN) {
      const raw = (out.ManufacturerPostalCode ?? '').toString();
      const digits = (raw.match(/\d+/g) || []).join('');  // 只取数字
      const allZero = /^0+$/.test(digits);
      const tooShort = (digits.length > 0 && digits.length < 6);
      if (tooShort || allZero) {
        out.ManufacturerPostalCode = '528000';
      }
    }
    
    output.push(out);

    if ((i + 1) % 20 === 0 || i === main.length - 1) {
      const pct = Math.round(((i + 1) / main.length) * 100);
      pt.innerText = `${pct}%`; prog.style.width = `${pct}%`;
      await new Promise(r => setTimeout(r, 0));
    }
  }


  // === SHEIN: 生成 GroupIdentifier（每组≤998，尽量均分，组号从1开始） ===
  if (IS_SHEIN) {
    const MAX_PER_GROUP = 998;
    const total = output.length;
    if (total > 0) {
      const groups = Math.ceil(total / MAX_PER_GROUP);
      const base = Math.floor(total / groups);
      const extra = total % groups; // 前 extra 组分配 base+1，其余 base
      const sizes = Array.from({ length: groups }, (_, i) => base + (i < extra ? 1 : 0));
      let idx = 0;
      let gid = 1;
      for (const sz of sizes) {
        for (let k = 0; k < sz && idx < total; k++, idx++) {
          if (typeof output[idx] === 'object' && output[idx] !== null) {
            output[idx]['GroupIdentifier'] = gid;
          }
        }
        gid++;
      }
    }
  }

  // 写回
  const header = ruleConfig.map(r => r.Column);
  if (IS_SHEIN && !header.includes('GroupIdentifier')) header.push('GroupIdentifier');
  const aoa = [header].concat(output.map(o => header.map(c => o[c] || '')));
  const ws2 = XLSX.utils.aoa_to_sheet(aoa);

  // 日期列推断（维持你原逻辑）
  const dateCols = ruleConfig
    .map((r, idx) => ({ idx, fmt: (r.Format || '').toLowerCase() }))
    .filter(r => r.fmt.includes('yyyy') && r.fmt.includes('m') && r.fmt.includes('d'))
    .map(r => r.idx);

  for (let r = 1; r < aoa.length; r++) {
    for (const c of dateCols) {
      const colLetter = XLSX.utils.encode_col(c);
      const cellRef = colLetter + (r + 1);
      const val = aoa[r][c];
      if (!val) continue;
      let d = new Date(val);
      if (isNaN(d.getTime())) {
        const m = String(val).match(/(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})/);
        if (m) d = new Date(m[3], m[1] - 1, m[2]);
      }
      if (!isNaN(d.getTime())) {
        ws2[cellRef].t = 'd';
        ws2[cellRef].z = 'm/d/yyyy';
        ws2[cellRef].v = d;
      }
    }
  }

  const wb2 = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb2, ws2, 'Sheet1');

  const mawbOrig = (document.getElementById(sanitize('MAWB'))?.value || currentDefaultMawb || '').trim();
  const tag = IS_SHEIN ? 'SHEIN' : 'TEMU';
  XLSX.writeFile(wb2, `${mawbOrig}_NETChb_${tag}.xlsx`);
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
      el.style.borderColor = '#d1d5db';
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
  if (df) observeNewInputs(df, 'neutral');
});


