/* validator.js
 * 说明：所有对用户可见的文本/报错均为英文；仅注释使用中文。
 */

const $ = (id) => document.getElementById(id);

const state = {
  wb: null,
  sheetNames: [],
  sheetIndex: 0,
  aoa: [],            // 当前 sheet 的 AOA
  headerCandidates: [],
  headerRow: null,    // 1-based
  headerCols: [],     // [{idx, text, score}]
  chosenColIdx: null, // 0-based
  busy: false,        // 防止多次点击
};

// 关键词计分（不区分大小写，连续匹配计数）——用于排序，不对用户展示
const KEYWORDS = ["fda", "product", "code"];
function keywordScore(s) {
  const t = String(s || "").toLowerCase();
  let score = 0;
  for (const kw of KEYWORDS) {
    const re = new RegExp(kw, "g");
    const m = t.match(re);
    if (m) score += m.length;
  }
  return score;
}

// 本地合法性正则（大小写不敏感：调用处统一转大写）
// 规则：前两位数字，第3位字母，第4/5位不得为数字（可为字母或 -），后两位数字
const CODE_RE = /^\d{2}[A-Za-z][A-Za-z-]{2}\d{2}$/;

// —— 统一门控：任何时候都用它来决定按钮是否可点 —— //
function gateGenerate() {
  const ok =
    !!state.wb &&                                   // 已有工作簿
    Array.isArray(state.aoa) && state.aoa.length > 0 && // 已解析出数据
    state.headerRow != null &&                      // 已确定表头行
    state.chosenColIdx != null &&                   // 已选择列
    !state.busy;                                    // 不在忙碌中
  $("btn-generate").disabled = !ok;
}

// —— 上传区：新增“上传中”态（点击/拖拽一开始就切换） —— //
function renderDropZoneUploading(filename = "") {
  const dz = $("drop-zone");
  dz.classList.remove("bg-emerald-50","border-emerald-400","border-slate-300");
  dz.classList.add("bg-slate-50","border-slate-400");
  dz.innerHTML = `
    <div class="flex items-center justify-center gap-2">
      <svg class="animate-spin h-4 w-4 text-slate-600" viewBox="0 0 24 24">
        <circle class="opacity-25" cx="12" cy="12" r="10" stroke="currentColor" stroke-width="3" fill="none"></circle>
        <path class="opacity-75" fill="currentColor" d="M4 12a8 8 0 018-8v3a5 5 0 00-5 5H4z"></path>
      </svg>
      <span class="text-slate-700 font-medium">Uploading${filename ? `: ${filename}` : ""}…</span>
    </div>
    <div class="mt-1 text-xs text-slate-500">Please wait…</div>
  `;
}

// Excel 列号 -> 字母
function colLetter(n) {
  let s = "", x = n + 1;
  while (x > 0) { x--; s = String.fromCharCode(65 + (x % 26)) + s; x = Math.floor(x/26); }
  return s;
}

// 读取 AOA
function loadAOA(sheetName) {
  const ws = state.wb.Sheets[sheetName];
  const opts = { header: 1, blankrows: false, raw: true, defval: "" };
  state.aoa = XLSX.utils.sheet_to_json(ws, opts);
}

// 解析候选表头行（启发式）
function analyzeHeaderCandidates() {
  const rows = state.aoa;
  if (!rows.length) return [];
  const width = rows.reduce((mx, r) => Math.max(mx, r.length), 0);

  const cands = [];
  const upTo = Math.min(rows.length, 100); // 只看前100行
  for (let i = 0; i < upTo; i++) {
    const row = rows[i].slice(0, width);
    const nonEmptyIdx = row.map((v, idx) => (v!=="" ? idx : -1)).filter(x=>x>=0);
    if (nonEmptyIdx.length < 2) continue;

    // 连续性（首非空到尾非空之间无空）
    const start = nonEmptyIdx[0], end = nonEmptyIdx[nonEmptyIdx.length-1];
    let contiguous = true;
    for (let c = start; c <= end; c++) { if (row[c] === "") { contiguous = false; break; } }
    if (!contiguous) continue;

    // 字符串占比
    const strCnt = row.slice(start, end+1).filter(v => typeof v === "string").length;
    const ratio = strCnt / (end - start + 1);
    if (ratio < 0.6) continue;

    // 下一行的“数值感”
    let nextScore = 0;
    if (i + 1 < rows.length) {
      const next = rows[i+1].slice(start, end+1);
      const numCnt = next.filter(v => typeof v === "number" || /^[0-9]+(\.[0-9]+)?$/.test(String(v))).length;
      nextScore = numCnt;
    }

    // 关键词权重
    const kwScore = row.reduce((s, v) => s + keywordScore(v), 0);
    const total = kwScore * 3 + nextScore;
    cands.push({ row1: i+1, score: total });
  }

  cands.sort((a, b) => b.score - a.score || b.row1 - a.row1);
  const unique = [];
  const seen = new Set();
  for (const c of cands) if (!seen.has(c.row1)) { unique.push(c); seen.add(c.row1); }
  return unique.slice(0, 20).map(c => c.row1).sort((a,b)=>a-b);
}

// 依据 header 行生成列清单（关键词命中越多越靠前，其余从左到右）
function buildHeaderCols() {
  const r = state.headerRow - 1;
  const row = state.aoa[r] || [];
  const cols = row.map((v, idx) => ({ idx, text: String(v||""), score: keywordScore(v) }));
  cols.sort((a, b) => b.score - a.score || a.idx - b.idx);
  state.headerCols = cols;
}

// 进度条工具（显示/隐藏/设置/温和自增）
const progress = (() => {
  const box = $("progress");
  const bar = $("progress-bar");
  const label = $("progress-label");
  const percent = $("progress-percent");
  let timer = null;
  let cur = 0;

  function show(text = "Preparing…", p = 0) {
    box.classList.remove("hidden");
    set(p, text);
  }
  function hide() {
    box.classList.add("hidden");
    stop();
    set(0, "Preparing…");
  }
  function set(p, text) {
    cur = Math.max(0, Math.min(100, p));
    bar.style.width = cur + "%";
    percent.textContent = Math.round(cur) + "%";
    if (text) label.textContent = text;
  }
  function trickle(target = 70, text = "Validating on server…", step = 1, ms = 300) {
    stop();
    label.textContent = text;
    timer = setInterval(() => {
      if (cur + step >= target) { set(target, text); stop(); return; }
      set(cur + step, text);
    }, ms);
  }
  function stop() { if (timer) { clearInterval(timer); timer = null; } }

  return { show, hide, set, trickle, stop };
})();

// 设置按钮忙碌态（仅禁用，不改文字，不改变位置）
function setBusy(flag) {
  const btn = $("btn-generate");
  state.busy = !!flag;
  btn.disabled = state.busy || btn.disabled; // 保持禁用优先
}

// —— 上传区：默认与成功态文案 —— //
function renderDropZoneDefault() {
  const dz = $("drop-zone");
  dz.classList.remove("bg-emerald-50","border-emerald-400");
  dz.classList.add("border-slate-300");
  dz.innerHTML = `
    <div class="text-lg font-semibold mb-1">Click or drop an Excel file</div>
    <div class="text-xs text-slate-500">.xls or .xlsx</div>
  `;
}
function renderDropZoneSuccess(filename) {
  const dz = $("drop-zone");
  dz.classList.remove("border-slate-300");
  dz.classList.add("bg-emerald-50","border-emerald-400");
  dz.innerHTML = `
    <div class="flex items-center justify-center gap-2">
      <span class="inline-flex items-center justify-center w-5 h-5 rounded-full bg-emerald-500 text-white text-xs">✓</span>
      <span class="text-emerald-700 font-semibold">Uploaded:</span>
      <span class="text-emerald-700 truncate max-w-[280px]">${filename}</span>
    </div>
    <div class="mt-1 text-xs text-emerald-700">Choose a sheet from the dropdown, then click “Continue”.</div>
  `;
}

// 在选中列右侧插入 Result，并下载
function insertResultAndDownload(resultMap) {
  const sheetName = state.sheetNames[state.sheetIndex] || "Sheet1";
  // 克隆原 sheet（避免直接改 state.wb）
  const wsOrig = state.wb.Sheets[sheetName];
  const ws = JSON.parse(JSON.stringify(wsOrig)); // 简易深拷贝

  // 计算插入列（0-based）：选中列的右侧
  const insertC = state.chosenColIdx + 1;
  const headR = state.headerRow - 1;

  // 解析并扩展 !ref（右边列数 +1）
  const range = XLSX.utils.decode_range(ws["!ref"]);
  range.e.c += 1;
  ws["!ref"] = XLSX.utils.encode_range(range);

  // 1) 先把所有 c >= insertC 的单元格整体右移 1 列
  const keys = Object.keys(ws).filter(k => k[0] !== "!");
  // 为避免覆盖，从右往左移动
  keys.sort((a, b) => {
    const ca = XLSX.utils.decode_cell(a), cb = XLSX.utils.decode_cell(b);
    if (ca.r !== cb.r) return cb.r - ca.r;
    return cb.c - ca.c;
  });
  for (const k of keys) {
    const addr = XLSX.utils.decode_cell(k);
    if (addr.c >= insertC) {
      const nk = XLSX.utils.encode_cell({ r: addr.r, c: addr.c + 1 });
      ws[nk] = ws[k];
      delete ws[k];
    }
  }

  // 2) 写入表头 "Result"
  ws[XLSX.utils.encode_cell({ r: headR, c: insertC })] = { t: "s", v: "Result" };

  // 3) 从数据区写入结果（字符串），空白保持空
  for (let r = headR + 1; r <= range.e.r; r++) {
    const src = XLSX.utils.encode_cell({ r, c: state.chosenColIdx });
    let raw = ws[src]?.v;
    // 读取展示文本优先（若有），否则读 v
    if (ws[src] && typeof ws[src].w === "string") raw = ws[src].w;
    const val = String(raw ?? "").trim().toUpperCase();
    let out = "";
    if (val) {
      if (CODE_RE.test(val)) out = resultMap[val] || "";
      else out = "Invalid";
    }
    const dst = XLSX.utils.encode_cell({ r, c: insertC });
    ws[dst] = { t: "s", v: out };
  }

  // 4) 生成新工作簿并下载（保留了原单元格的 number format 等）
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, sheetName);
  const now = new Date();
  const pad = n => String(n).padStart(2, "0");
  const stamp = `${now.getFullYear()}-${pad(now.getMonth()+1)}-${pad(now.getDate())}_${pad(now.getHours())}${pad(now.getMinutes())}${pad(now.getSeconds())}`;
  const fname = `validated_${sheetName}_${stamp}.xlsx`;
  const wbout = XLSX.write(wb, { bookType: "xlsx", type: "array" });
  saveAs(new Blob([wbout], { type: "application/octet-stream" }), fname);
}

// 初始化与事件绑定
(function init() {
  const fileInput = $("file-input");
  const dropZone  = $("drop-zone");
  const sheetSelect = $("sheet-select");
  const btnContinue = $("btn-continue");
  const headerRowSelect = $("header-row-select");
  const chkCustom = $("custom-row-check");
  const txtCustom = $("custom-row-input");
  const errCustom = $("custom-row-error");
  const colSelect = $("column-select");

  // 上传动作开始前的“立即上锁”，避免短暂空窗期
  function lockBeforeUpload(filename = "") {
    // 先把 UI 全部锁住（同步执行，毫秒级生效）
    $("sheet-select").disabled = true;
    $("btn-continue").disabled = true;
    $("header-row-select").disabled = true;
    $("column-select").disabled = true;

    // 清除依赖选择的关键信息，确保按钮立刻禁用
    state.headerRow = null;
    state.chosenColIdx = null;
    state.aoa = [];
    gateGenerate(); // 统一门控：确保 Generate & Download 立刻禁用

    // 上传区切到“上传中”视觉
    renderDropZoneUploading(filename);
  }

  // 统一处理文件（点击或拖拽都会调用它）
  async function handleFile(file) {
    if (!file) return;
    try {
      const data = await file.arrayBuffer();
      state.wb = XLSX.read(data, { type: "array" });
    } catch (e) {
      alert("Failed to read the file. Please try another Excel file.");
      renderDropZoneDefault(); // 失败则回退视觉
      gateGenerate();
      return;
    }
    // 填充 sheet 下拉
    state.sheetNames = state.wb.SheetNames || [];
    sheetSelect.innerHTML = "";
    state.sheetNames.forEach((name, idx) => {
      const opt = document.createElement("option");
      opt.value = String(idx); opt.textContent = name;
      sheetSelect.appendChild(opt);
    });
    sheetSelect.disabled = state.sheetNames.length === 0;
    btnContinue.disabled = state.sheetNames.length === 0;

    // ☆ 保险：即使解析完成，也不要点亮生成按钮（要等 Continue + 列选择）
    state.headerRow = null;
    state.chosenColIdx = null;
    gateGenerate();

    // 上传成功的视觉反馈
    renderDropZoneSuccess(file.name);
  }

  // Drop zone 点击触发本地选择
  dropZone.addEventListener("click", () => fileInput.click());

  // 拖拽高亮+允许放置（必须阻止默认行为）
  ["dragenter","dragover"].forEach(ev => {
    dropZone.addEventListener(ev, (e) => {
      e.preventDefault();
      dropZone.classList.add("ring-2","ring-slate-300");
    });
  });
  ["dragleave","drop"].forEach(ev => {
    dropZone.addEventListener(ev, (e) => {
      e.preventDefault();
      dropZone.classList.remove("ring-2","ring-slate-300");
    });
  });

  // 点击选择文件
  fileInput.addEventListener("change", async (e) => {
    const f = e.target.files?.[0];
    if (!f) return;
    // ☆ 关键：先上锁（同步禁用按钮），再去异步读取
    lockBeforeUpload(f.name);
    resetAll(false);            // 不清 file-input 自身，但会清状态并再 gate 一次
    await handleFile(f);
  });

  // 拖拽放下文件
  dropZone.addEventListener("drop", async (e) => {
    const f = e.dataTransfer?.files?.[0];
    if (!f) return;
    const ok = /\.xls[x]?$/i.test(f.name);
    if (!ok) {
      alert("Only .xls or .xlsx files are supported.");
      renderDropZoneDefault();
      gateGenerate();
      return;
    }
    // ☆ 关键：先上锁（同步禁用按钮），再去异步读取
    lockBeforeUpload(f.name);
    resetAll(true);             // 清状态
    await handleFile(f);
  });

  sheetSelect.addEventListener("change", () => {
    state.sheetIndex = Number(sheetSelect.value || 0);
    gateGenerate();
  });

  btnContinue.addEventListener("click", () => {
    state.sheetIndex = Number(sheetSelect.value || 0);
    loadAOA(state.sheetNames[state.sheetIndex]);

    state.headerCandidates = analyzeHeaderCandidates();
    headerRowSelect.innerHTML = "";
    state.headerCandidates.forEach((row1) => {
      const opt = document.createElement("option");
      opt.value = String(row1); opt.textContent = `Row ${row1}`;
      headerRowSelect.appendChild(opt);
    });

    if (state.headerCandidates.length) {
      const def = Math.max(...state.headerCandidates);
      headerRowSelect.value = String(def);
      state.headerRow = def;
      headerRowSelect.disabled = false;
      chkCustom.checked = false;
      txtCustom.disabled = true;
      txtCustom.classList.remove("error");
      errCustom.classList.add("hidden");
      buildHeaderCols();
      fillColumnSelect();
    } else {
      headerRowSelect.disabled = true;
      colSelect.disabled = true;
      alert("No header candidates detected. Please enable 'Use custom row number' and enter the row index.");
      chkCustom.checked = true;
      txtCustom.disabled = false;
    }
    gateGenerate(); // Continue 后再次门控
  });

  chkCustom.addEventListener("change", () => {
    if (chkCustom.checked) {
      headerRowSelect.disabled = true;
      txtCustom.disabled = false;
      txtCustom.focus();
    } else {
      txtCustom.disabled = true;
      txtCustom.value = "";
      txtCustom.classList.remove("error");
      errCustom.classList.add("hidden");
      headerRowSelect.disabled = false;
      state.headerRow = Number(headerRowSelect.value || 1);
      buildHeaderCols(); fillColumnSelect();
    }
    gateGenerate();
  });

  txtCustom.addEventListener("input", () => {
    const ok = /^\d+$/.test(txtCustom.value.trim());
    if (!ok) {
      txtCustom.classList.add("error");
      errCustom.classList.remove("hidden");
      state.headerRow = null;
      $("column-select").disabled = true;
      $("btn-generate").disabled = true;
      gateGenerate();
      return;
    }
    txtCustom.classList.remove("error");
    errCustom.classList.add("hidden");
    state.headerRow = Number(txtCustom.value.trim());
    buildHeaderCols(); fillColumnSelect();
    gateGenerate();
  });

  headerRowSelect.addEventListener("change", () => {
    state.headerRow = Number(headerRowSelect.value || 1);
    buildHeaderCols(); fillColumnSelect();
    gateGenerate();
  });

  $("column-select").addEventListener("change", () => {
    const v = $("column-select").value;
    if (v === "" || v == null) {
      state.chosenColIdx = null;
    } else {
      state.chosenColIdx = Number(v);
    }
    gateGenerate(); // 只有选择有效列后才会放开按钮
  });

  $("btn-generate").addEventListener("click", onGenerate);

  function fillColumnSelect() {
    const colSelect = $("column-select");
    colSelect.innerHTML = "";

    // 没有有效的 header 行时，禁用下拉和按钮
    if (!state.headerRow) {
      colSelect.disabled = true;
      state.chosenColIdx = null;
      $("btn-generate").disabled = true;
      gateGenerate();
      return;
    }

    // 先按规则构建并排序（已按命中数从高到低、同分按列序）
    buildHeaderCols();

    // —— 不再添加“Select a column…”占位项，改为直接渲染选项 —— //
    state.headerCols.forEach(({ idx, text }) => {
      const opt = document.createElement("option");
      opt.value = String(idx);
      opt.textContent = text ? text : `Column ${colLetter(idx)}`;
      colSelect.appendChild(opt);
    });

    colSelect.disabled = false;

    // —— 关键：自动选择命中最多的那一列（state.headerCols[0]）—— //
    if (state.headerCols.length > 0) {
      const best = state.headerCols[0];         // 已按 score 排好
      colSelect.value = String(best.idx);
      state.chosenColIdx = best.idx;
      $("btn-generate").disabled = false;       // 因为已经“有选择”了
    } else {
      // 极端情况：header 行为空
      state.chosenColIdx = null;
      $("btn-generate").disabled = true;
    }

    gateGenerate();  // 统一门控，最终裁决按钮可用性
  }

  function resetAll(clearFile=false) {
    if (clearFile) $("file-input").value = "";
    $("sheet-select").innerHTML = ""; $("sheet-select").disabled = true;
    $("btn-continue").disabled = true;
    $("header-row-select").innerHTML = ""; $("header-row-select").disabled = true;
    $("custom-row-check").checked = false;
    $("custom-row-input").value = ""; $("custom-row-input").disabled = true; $("custom-row-input").classList.remove("error");
    $("custom-row-error").classList.add("hidden");
    $("column-select").innerHTML = ""; $("column-select").disabled = true;
    $("btn-generate").disabled = true;

    state.wb = null; state.sheetNames = []; state.sheetIndex = 0;
    state.aoa = []; state.headerCandidates = []; state.headerRow = null;
    state.headerCols = []; state.chosenColIdx = null;

    // 复位 busy/进度条/上传区
    setBusy(false);
    progress.hide();
    renderDropZoneDefault();
    gateGenerate();
  }
})();

// 生成并下载：前端去重->调用后端->插入Result->下载（带进度条，按钮不改变文案）
async function onGenerate() {
  if (state.busy) return;        // 防止连点
  const rHead = state.headerRow - 1;
  const col = state.chosenColIdx;
  if (rHead < 0 || col == null) return;

  try {
    setBusy(true);
    gateGenerate();
    progress.show("Preparing…", 10);

    // Step 1: 收集列数据
    const values = [];
    for (let i = rHead + 1; i < state.aoa.length; i++) {
      const raw = String( (state.aoa[i][col] ?? "") ).trim();
      values.push(raw);
    }
    progress.set(20, "Scanning column…");

    // Step 2: 本地初筛 + 去重
    const uniq = new Set();
    for (const v of values) {
      const s = v.toUpperCase();
      if (s && CODE_RE.test(s)) uniq.add(s);
    }
    progress.set(30, `Found ${uniq.size} unique code(s)`);

    // Step 3: 请求后端（期间缓慢自增到 70%）
    progress.trickle(70, "Validating on server…", 1, 250);
    const payload = { codes: Array.from(uniq) };
    const resp = await fetch("/api/validate-codes", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payload),
    }).catch(() => null);

    if (!resp || !resp.ok) {
      progress.hide();
      setBusy(false);
      gateGenerate();
      alert("Server returned an error. Please try again later.");
      return;
    }
    const data = await resp.json();
    progress.set(85, "Writing results…");

    // Step 4: 写回并导出
    const map = data.results || {};
    insertResultAndDownload(map);

    // Step 5: 完成
    progress.set(100, "Download started");
    setTimeout(() => { progress.hide(); setBusy(false); gateGenerate(); }, 1200);

  } catch (e) {
    progress.hide();
    setBusy(false);
    gateGenerate();
    alert("Unexpected error. Please try again.");
  }
}
