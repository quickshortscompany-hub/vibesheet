/* Vibe Sheet — v1.1
 * Select cells → describe → preview → accept → undo.
 * AI proposes structured operations; nothing touches the workbook until Accept.
 */
"use strict";

/* ---------------- constants ---------------- */
const VERSION = "1.1.1";
const DEFAULT_API = "https://vibesheet-api.quickshortscompany.workers.dev";   // used if config.js is missing, old or cached
const CONFIG = window.VIBESHEET_CONFIG || {};
if (!CONFIG.apiUrl || /YOUR-SUBDOMAIN/.test(CONFIG.apiUrl)) CONFIG.apiUrl = DEFAULT_API;
const PROXY = /YOUR-SUBDOMAIN|^\s*$/.test(CONFIG.apiUrl || "") ? "" : String(CONFIG.apiUrl).trim().replace(/\/+$/, "");
const API_URL = "https://api.anthropic.com/v1/messages";
const MODELS_URL = "https://api.anthropic.com/v1/models?limit=100";
const SEL_MAX_ROWS = 80, SEL_MAX_COLS = 26;     // selection cells sent to AI
const OVR_MAX_ROWS = 25, OVR_MAX_COLS = 20;     // sheet overview sent to AI
const SNAP_LIMIT = 25000;                       // max cells snapshotted for undo
const CHECK_LIMIT = 20000;                      // max cells checked for "outside selection"
const MODIFYING = new Set(["write","format","clear","sort","removeDuplicates","findReplace","merge","unmerge",
  "conditionalFormat","dataValidation","table","insert","delete"]);
const STRUCTURAL = new Set(["insert","delete"]);

/* ---------------- state ---------------- */
const state = {
  inExcel: false,
  image: null,          // {base64, mediaType, dataUrl}
  req: null,            // {prompt, ctx}
  pending: null,        // {summary, message, items:[{op, ...preview}]}
  history: [],          // [{id, time, prompt, summary, titles, steps, undone, records}]
  chat: [],             // recent [{prompt, summary}]
  busy: false,
  lastError: "",        // for "Report this problem"
  lastPrompt: "",
};

/* ---------------- tiny helpers ---------------- */
const $ = (id) => document.getElementById(id);
const esc = (s) => String(s ?? "").replace(/[&<>"']/g, (c) => ({ "&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;" }[c]));
const store = {
  get(k) { try { return localStorage.getItem("vibesheet." + k); } catch { return store._m[k] ?? null; } },
  set(k, v) { try { localStorage.setItem("vibesheet." + k, v); } catch { store._m[k] = v; } },
  _m: {},
};
function colToNum(c) { let n = 0; for (const ch of c.toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64); return n - 1; }
function numToCol(n) { let s = ""; n++; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; }
function cellAddr(r, c) { return numToCol(c) + (r + 1); }
function splitSheet(addr) {
  addr = String(addr || "").trim();
  const i = addr.lastIndexOf("!");
  if (i < 0) return { sheet: null, ref: addr };
  let s = addr.slice(0, i);
  if (s.startsWith("'") && s.endsWith("'")) s = s.slice(1, -1).replace(/''/g, "'");
  return { sheet: s, ref: addr.slice(i + 1) };
}
function stripSheet(addr) { return splitSheet(addr).ref; }
function shortVal(v, n = 60) { const s = v === null || v === undefined ? "" : String(v); return s.length > n ? s.slice(0, n - 1) + "…" : s; }
function norm2D(d) {
  if (d === null || d === undefined) return [[""]];
  if (!Array.isArray(d)) return [[d]];
  if (!d.length) return [[""]];
  if (!Array.isArray(d[0])) d = [d];
  const w = Math.max(...d.map((r) => (Array.isArray(r) ? r.length : 1)));
  return d.map((r) => { r = Array.isArray(r) ? r.slice() : [r]; while (r.length < w) r.push(null); return r.map(cleanCell); });
}
function cleanCell(v) {
  if (v === null || v === undefined) return null;
  if (typeof v === "number" || typeof v === "boolean" || typeof v === "string") return v;
  return JSON.stringify(v);
}
function cleanName(n) {
  let s = String(n || "Name").trim().replace(/[^A-Za-z0-9_.\\]/g, "_").slice(0, 255);
  if (!/^[A-Za-z_\\]/.test(s)) s = "_" + s;
  if (/^[A-Za-z]{1,3}\d+$/.test(s) || /^[rRcC]$/.test(s) || /^[rR]\d*[cC]\d*$/.test(s)) s += "_";
  return s;
}
function deviceId() {
  let d = store.get("device");
  if (!d) { d = (crypto.randomUUID ? crypto.randomUUID() : Date.now().toString(36) + "-" + Math.random().toString(36).slice(2, 12)); store.set("device", d); }
  return d;
}
function useOwnKey() { return store.get("useOwnKey") === "1" && !!store.get("apiKey"); }
function aiReady() { return useOwnKey() || !!PROXY; }
function fill2D(rows, cols, v) { return Array.from({ length: rows }, () => Array(cols).fill(v)); }
function cap(s) { s = String(s || ""); return s.charAt(0).toUpperCase() + s.slice(1).toLowerCase(); }
function colOffset(col, startColIndex) {
  if (typeof col === "number") return col;
  if (/^\d+$/.test(String(col))) return parseInt(col, 10);
  return colToNum(String(col).replace(/[^A-Za-z]/g, "")) - startColIndex;
}
function toCfFormula(v) {
  if (v === null || v === undefined) return undefined;
  if (typeof v === "number") return String(v);
  v = String(v);
  if (v.startsWith("=") || /^-?\d+(\.\d+)?$/.test(v)) return v;
  return '="' + v.replace(/"/g, '""') + '"';
}

/* ---------------- UI wiring ---------------- */
function initUI() {
  document.querySelectorAll(".tab").forEach((b) => b.addEventListener("click", () => showTab(b.dataset.tab)));
  document.querySelectorAll("[data-goto]").forEach((a) => a.addEventListener("click", (e) => { e.preventDefault(); showTab(a.dataset.goto); }));
  document.querySelectorAll(".chip").forEach((c) => c.addEventListener("click", () => { $("prompt").value = c.dataset.p; send(); }));
  $("send").addEventListener("click", send);
  $("prompt").addEventListener("keydown", (e) => { if (e.key === "Enter" && (e.ctrlKey || e.metaKey)) { e.preventDefault(); send(); } });
  $("accept").addEventListener("click", accept);
  $("reject").addEventListener("click", () => { state.pending = null; $("result").classList.add("hidden"); setStatus(""); });
  $("undoLast").addEventListener("click", undoLast);
  $("imgInput").addEventListener("change", (e) => { const f = e.target.files[0]; if (f) attachImage(f); e.target.value = ""; });
  $("imgRemove").addEventListener("click", clearImage);
  document.addEventListener("paste", (e) => {
    const items = e.clipboardData && e.clipboardData.items; if (!items) return;
    for (const it of items) if (it.kind === "file" && it.type.startsWith("image/")) { e.preventDefault(); attachImage(it.getAsFile()); return; }
  });
  const drop = $("drop");
  ["dragenter", "dragover"].forEach((t) => drop.addEventListener(t, (e) => { e.preventDefault(); drop.classList.add("drag"); }));
  ["dragleave", "drop"].forEach((t) => drop.addEventListener(t, (e) => { e.preventDefault(); drop.classList.remove("drag"); }));
  drop.addEventListener("drop", (e) => { const f = e.dataTransfer.files[0]; if (f && f.type.startsWith("image/")) attachImage(f); });

  // help examples: fill the prompt, let the user select cells and press Go
  document.querySelectorAll(".example").forEach((b) => b.addEventListener("click", () => { $("prompt").value = b.dataset.p; showTab("prompt"); $("prompt").focus(); }));
  // feedback
  $("fbSend").addEventListener("click", sendFeedback);
  $("status").addEventListener("click", (e) => { if (e.target.id === "reportErr") { e.preventDefault(); openReport(); } });

  // settings
  $("ver").textContent = VERSION;
  $("useOwnKey").checked = store.get("useOwnKey") === "1";
  $("apiKey").value = store.get("apiKey") || "";
  const savedModel = store.get("model");
  if (savedModel) $("modelSelect").innerHTML = `<option value="${esc(savedModel)}">${esc(savedModel)}</option>`;
  $("loadModels").addEventListener("click", loadModels);
  $("saveSettings").addEventListener("click", saveSettings);
  refreshKeyNotice();
  renderAiMode();
  refreshQuota();
  renderHistory();
}
function showTab(name) {
  document.querySelectorAll(".tab").forEach((b) => b.classList.toggle("active", b.dataset.tab === name));
  ["prompt", "history", "help", "feedback", "settings"].forEach((t) => $("tab-" + t).classList.toggle("hidden", t !== name));
}
function refreshKeyNotice() { $("noKey").classList.toggle("hidden", aiReady()); }
function renderAiMode() {
  const m = $("aiMode");
  if (useOwnKey()) { m.className = "status"; m.innerHTML = "Using <b>your own Anthropic key</b>."; }
  else if (PROXY) { m.className = "status ok"; m.innerHTML = "✓ Using <b>Vibe Sheet's free AI</b> — nothing to set up." + (state.quota ? `<br><span class="small">${state.quota.remaining} of ${state.quota.limit} prompts left today.</span>` : ""); }
  else { m.className = "status err"; m.innerHTML = "The free AI isn't configured in this copy (config.js). Use your own key below."; $("ownKeyBox").open = true; }
}
async function refreshQuota() {
  if (!PROXY || useOwnKey()) { $("quota").textContent = ""; return; }
  try {
    const r = await fetch(`${PROXY}/v1/status?device=${encodeURIComponent(deviceId())}`);
    const j = await r.json(); if (r.ok) setQuota(j.remaining, j.limit);
  } catch { /* offline: ignore */ }
}
function setQuota(remaining, limit) {
  if (remaining == null || limit == null) return;
  state.quota = { remaining, limit };
  $("quota").textContent = `${remaining}/${limit} free prompts left today`;
  renderAiMode();
}
function setStatus(html, kind = "") {
  const s = $("status");
  if (!html) { s.classList.add("hidden"); return; }
  s.className = "status " + kind; s.innerHTML = html;
  if (kind === "err") {
    state.lastError = s.textContent;
    s.insertAdjacentHTML("beforeend", ` <br><a href="#" id="reportErr">Report this problem →</a>`);
  }
}
function openReport() {
  showTab("feedback");
  document.querySelector('input[name="fbType"][value="bug"]').checked = true;
  $("fbContext").checked = true;
  if (!$("fbMessage").value.trim()) $("fbMessage").value = "It went wrong when I tried: " + (state.lastPrompt || "(describe what you did)") + "\n\n";
  $("fbMessage").focus();
}
async function sendFeedback() {
  const st = $("fbStatus");
  const message = $("fbMessage").value.trim();
  if (!message) { st.className = "status err"; st.textContent = "Please write a message first."; return; }
  const type = (document.querySelector('input[name="fbType"]:checked') || {}).value || "other";
  const context = $("fbContext").checked ? `Last request: ${state.lastPrompt || "-"}\nLast error: ${state.lastError || "-"}` : "";
  const payload = { device: deviceId(), type, message, email: $("fbEmail").value.trim(), context, version: VERSION, host: state.inExcel ? "excel" : "browser" };
  if (!PROXY) {
    st.className = "status err";
    st.innerHTML = CONFIG.feedbackEmail ? `Feedback service not set up. Please email <a href="mailto:${esc(CONFIG.feedbackEmail)}?subject=Vibe%20Sheet%20${esc(type)}&body=${encodeURIComponent(message + "\n\n" + context)}">${esc(CONFIG.feedbackEmail)}</a>.` : "Feedback service not set up in this copy.";
    return;
  }
  $("fbSend").disabled = true; st.className = "status"; st.innerHTML = '<span class="spinner"></span>Sending…';
  try {
    const r = await fetch(`${PROXY}/v1/feedback`, { method: "POST", headers: { "content-type": "application/json" }, body: JSON.stringify(payload) });
    const j = await r.json().catch(() => ({}));
    if (!r.ok) throw new Error(j.error || `Couldn't send (${r.status}).`);
    st.className = "status ok"; st.textContent = "Thanks! Your feedback was sent.";
    $("fbMessage").value = "";
  } catch (e) { st.className = "status err"; st.textContent = e.message; }
  finally { $("fbSend").disabled = false; }
}
function setBusy(b, msg) {
  state.busy = b; $("send").disabled = b; $("accept").disabled = b;
  if (b) setStatus(`<span class="spinner"></span>${esc(msg || "Thinking…")}`);
}

/* ---------------- settings ---------------- */
async function loadModels() {
  const key = cleanKey($("apiKey").value); $("apiKey").value = key;
  const st = $("settingsStatus");
  if (!key) { st.className = "status err"; st.textContent = "Enter your API key first."; return; }
  st.className = "status"; st.innerHTML = '<span class="spinner"></span>Loading models…';
  try {
    const list = await fetchModels(key);
    const cur = store.get("model");
    $("modelSelect").innerHTML = list.map((m) => `<option value="${esc(m.id)}">${esc(m.display_name || m.id)}</option>`).join("");
    $("modelSelect").value = cur && list.some((m) => m.id === cur) ? cur : pickDefaultModel(list);
    st.className = "status ok"; st.textContent = `Loaded ${list.length} models. Press Save.`;
  } catch (e) { st.className = "status err"; st.textContent = e.message; }
}
async function fetchModels(key) {
  const h = apiHeaders(key); delete h["content-type"];
  const r = await fetch(MODELS_URL, { headers: h });
  if (!r.ok) {
    let detail = ""; try { const j = await r.json(); detail = (j.error && j.error.message) || JSON.stringify(j); } catch { /* ignore */ }
    if (r.status === 401) throw new Error("API key rejected (401). Copy the key again from console.anthropic.com." + (detail ? " — " + detail : ""));
    throw new Error(`Could not load models (${r.status}): ${detail || "no details"}. You can type a model ID below instead.`);
  }
  const j = await r.json();
  return j.data || [];
}
function pickDefaultModel(list) {
  const s = list.find((m) => /sonnet/i.test(m.id)) || list[0];
  return s ? s.id : "";
}
function saveSettings() {
  store.set("apiKey", cleanKey($("apiKey").value));
  const typed = ($("modelManual").value || "").trim();
  if (typed) store.set("model", typed); else if ($("modelSelect").value) store.set("model", $("modelSelect").value);
  store.set("useOwnKey", $("useOwnKey").checked ? "1" : "0");
  const st = $("settingsStatus"); st.className = "status ok"; st.textContent = "Saved.";
  if ($("useOwnKey").checked && !store.get("apiKey")) { st.className = "status err"; st.textContent = "Saved, but add a key to use your own."; }
  refreshKeyNotice(); renderAiMode(); refreshQuota();
}
function cleanKey(k) { return String(k || "").replace(/[^\x21-\x7E]/g, ""); }   // strips spaces, line breaks, hidden characters
function apiHeaders(key) {
  return { "content-type": "application/json", "x-api-key": key, "anthropic-version": "2023-06-01",
    "anthropic-dangerous-direct-browser-access": "true" };
}

/* ---------------- images ---------------- */
function attachImage(file) {
  const fr = new FileReader();
  fr.onload = () => {
    const img = new Image();
    img.onload = () => {
      const max = 1568, s = Math.min(1, max / Math.max(img.width, img.height));
      const cv = document.createElement("canvas");
      cv.width = Math.round(img.width * s); cv.height = Math.round(img.height * s);
      const g = cv.getContext("2d"); g.fillStyle = "#fff"; g.fillRect(0, 0, cv.width, cv.height); g.drawImage(img, 0, 0, cv.width, cv.height);
      let mediaType = file.type === "image/png" ? "image/png" : "image/jpeg";
      let url = cv.toDataURL(mediaType, 0.9);
      if (url.length > 4_500_000) { mediaType = "image/jpeg"; url = cv.toDataURL(mediaType, 0.8); }
      state.image = { dataUrl: url, mediaType, base64: url.split(",")[1] };
      $("imgThumb").src = url; $("imgWrap").classList.remove("hidden");
    };
    img.onerror = () => setStatus("Couldn't read that image.", "err");
    img.src = fr.result;
  };
  fr.readAsDataURL(file);
}
function clearImage() { state.image = null; $("imgWrap").classList.add("hidden"); $("imgThumb").removeAttribute("src"); }

/* ---------------- selection tracking ---------------- */
let selTimer = null;
function onSelectionChanged() { clearTimeout(selTimer); selTimer = setTimeout(refreshSelection, 150); }
async function refreshSelection() {
  if (!state.inExcel) return;
  try {
    await Excel.run(async (ctx) => {
      const r = ctx.workbook.getSelectedRange();
      r.load("address,rowCount,columnCount,rowIndex,columnIndex"); r.worksheet.load("name");
      await ctx.sync();
      $("selAddr").textContent = r.address;
      $("selSize").textContent = `${r.rowCount} × ${r.columnCount}`;
      showCellPrompt(r.worksheet.name, r.rowIndex, r.columnIndex);
    });
  } catch { /* ignore */ }
}

/* prompts remembered per cell, saved inside the workbook */
function getCellPrompts() { try { return Office.context.document.settings.get("vibesheet.cellPrompts") || []; } catch { return []; } }
function setCellPrompts(list) {
  try { Office.context.document.settings.set("vibesheet.cellPrompts", list.slice(-300)); Office.context.document.settings.saveAsync(); } catch { /* ignore */ }
}
function showCellPrompt(sheet, r, c) {
  const list = getCellPrompts();
  for (let i = list.length - 1; i >= 0; i--) {
    const p = list[i];
    if (p.sheet === sheet && r >= p.r1 && r <= p.r2 && c >= p.c1 && c <= p.c2) {
      $("cellPrompt").innerHTML = `✨ Made by: “${esc(shortVal(p.prompt, 140))}” <span class="muted">· ${esc(new Date(p.time).toLocaleString())}</span>`;
      $("cellPrompt").classList.remove("hidden"); return;
    }
  }
  $("cellPrompt").classList.add("hidden");
}

/* ---------------- gather workbook context ---------------- */
async function gatherContext() {
  return Excel.run(async (ctx) => {
    const wb = ctx.workbook;
    const sel = wb.getSelectedRange();
    sel.load("address,rowIndex,columnIndex,rowCount,columnCount");
    const ws = wb.worksheets.getActiveWorksheet(); ws.load("name");
    const sheets = wb.worksheets; sheets.load("items/name");
    const used = ws.getUsedRangeOrNullObject(true); used.load("address,rowIndex,columnIndex,rowCount,columnCount");
    const tables = ws.tables; tables.load("items/name");
    const names = wb.names; names.load("items/name,items/formula,items/visible");
    await ctx.sync();
    const nameList = names.items.filter((n) => n.visible !== false).map((n) => ({ name: n.name, ref: String(n.formula || "").replace(/^=/, "") }));
    const nameMap = {};   // "A1" on active sheet -> name
    for (const n of nameList) {
      const { sheet, ref } = splitSheet(n.ref);
      if ((sheet || ws.name) === ws.name && ref) nameMap[ref.replace(/\$/g, "").split(":")[0].toUpperCase()] = n.name;
    }

    const out = {
      sheet: ws.name, sheets: sheets.items.map((s) => s.name), tables: tables.items.map((t) => t.name), names: nameList,
      sel: { address: sel.address, r1: sel.rowIndex, c1: sel.columnIndex, r2: sel.rowIndex + sel.rowCount - 1, c2: sel.columnIndex + sel.columnCount - 1, rows: sel.rowCount, cols: sel.columnCount },
      used: used.isNullObject ? null : used.address, selText: "(empty)", overview: "(sheet is empty)", selTruncated: false,
    };
    if (!used.isNullObject) {
      const inter = sel.getIntersectionOrNullObject(used);
      inter.load("rowIndex,columnIndex,rowCount,columnCount");
      await ctx.sync();
      if (!inter.isNullObject) {
        const rr = Math.min(inter.rowCount, SEL_MAX_ROWS), cc = Math.min(inter.columnCount, SEL_MAX_COLS);
        const capR = ws.getRangeByIndexes(inter.rowIndex, inter.columnIndex, rr, cc);
        capR.load("formulas,values,numberFormat");
        await ctx.sync();
        out.selText = gridText(inter.rowIndex, inter.columnIndex, capR.formulas, capR.values, capR.numberFormat, nameMap);
        out.selTruncated = rr < inter.rowCount || cc < inter.columnCount;
      }
      const rr = Math.min(used.rowCount, OVR_MAX_ROWS), cc = Math.min(used.columnCount, OVR_MAX_COLS);
      const top = ws.getRangeByIndexes(used.rowIndex, used.columnIndex, rr, cc);
      top.load("formulas,values,numberFormat");
      await ctx.sync();
      out.overview = gridText(used.rowIndex, used.columnIndex, top.formulas, top.values, top.numberFormat, nameMap);
    }
    return out;
  });
}
function gridText(r0, c0, formulas, values, fmts, nameMap = {}) {
  const lines = [];
  for (let r = 0; r < formulas.length; r++) {
    const cells = [];
    for (let c = 0; c < formulas[r].length; c++) {
      const f = formulas[r][c], v = values[r][c], nf = fmts ? fmts[r][c] : "General";
      if ((f === "" || f === null) && (v === "" || v === null)) continue;
      let s = typeof f === "string" && f.startsWith("=") ? `${f} → ${shortVal(v, 40)}` : shortVal(v, 80);
      if (nf && nf !== "General") s += ` [fmt ${nf}]`;
      const a = cellAddr(r0 + r, c0 + c);
      if (nameMap[a]) s += ` [named ${nameMap[a]}]`;
      cells.push(`${a}: ${s}`);
    }
    if (cells.length) lines.push(cells.join(" | "));
  }
  return lines.length ? lines.join("\n") : "(empty)";
}

/* ---------------- AI ---------------- */
const OP_TYPES = ["write","format","clear","sort","removeDuplicates","findReplace","merge","unmerge","conditionalFormat",
  "dataValidation","chart","table","pivot","freeze","insert","delete","newSheet","name"];
const TOOL = {
  name: "apply_sheet_changes",
  description: "Propose precise changes to the user's Excel workbook. The user previews and accepts them.",
  input_schema: {
    type: "object",
    properties: {
      summary: { type: "string", description: "One short plain-English sentence: what will change." },
      message: { type: "string", description: "Optional explanation/answer for the user (for questions, error diagnosis, notes)." },
      operations: { type: "array", items: { type: "object", properties: { type: { type: "string", enum: OP_TYPES } }, required: ["type"], additionalProperties: true } },
    },
    required: ["summary", "operations"],
  },
};
function systemPrompt() {
  return `You are Vibe Sheet, an AI working inside Microsoft Excel. The user selects cells and describes what they want in plain language (any wording, typos are fine). You reply ONLY by calling apply_sheet_changes with precise operations. The user previews every operation and accepts or rejects it.

PRINCIPLES
1. Precision and control. Change only what was asked, only inside the user's selection. You may also fill EMPTY cells next to the selection when the result needs more room (e.g. one cell selected and they ask for a table, or a helper column). Never rebuild or rewrite the sheet. In write operations include only cells that actually change; use several small write ops for separate blocks; use null to skip a cell inside a block.
2. Prefer live Excel formulas over hard-coded results whenever the answer comes from other cells (=SUM(B2:B10), not 450) so the user can audit it. Hard values only for constants, in-place data cleaning, transcriptions, or when asked.
3. Formulas: English function names, comma separators, A1 references. For a formula down a column, write each row's formula with correct relative references. Use XLOOKUP/FILTER/UNIQUE/TEXTSPLIT/LET when helpful; IFERROR only when the user wants clean output.
4. Formatting (colours, bold, borders, number formats, alignment, widths, highlight rules) → format / conditionalFormat ops; never rewrite values for a formatting request.
5. Images: a table → transcribe into cells (numbers as numbers, keep headers). A maths/engineering/finance formula → create labelled input cells (label in one column, value next to it, units in the label; use values from the image or leave the value blank) plus a clearly labelled result cell with the Excel formula referencing those inputs. A screenshot of an Excel problem → diagnose and fix.
6. Questions/explanations ("what does this do", "why #N/A") → answer in "message" (short, plain, use '-' bullets), and include fix operations if something should change. If nothing should change, operations = [].
7. Data cleaning (trim, case, split, numbers stored as text, dates) → write cleaned values over the original cells unless the user asks for new columns. Dates: match the sheet's existing style; if unclear use dd/mm/yyyy. Write dates as =DATE(y,m,d) or "YYYY-MM-DD" text plus a numberFormat.
8. Never use whole-column or whole-row ranges (A:AO, 1:1000) for format, clear, sort or other changes; use the actual data area from the used range (e.g. A1:AO120). Addresses are A1 style, optionally with sheet: 'Sheet Name'!A1:C5. No sheet name = the active sheet. Header names in pivots must match the header cells exactly.
9. Named cells. Existing names are listed and marked [named X] next to their cells: use the names in formulas instead of addresses (=Load*Span^2/8, not =B3*B2^2/8). "Call this Span" / "name these cells Rates" → a name op. Whenever you create labelled input cells (calcs, formulas from images), also add a name op for each input and for the result, and write the formula using those names. Names: letters, digits, underscores, no spaces, start with a letter, never look like a cell address (not A1, XY12) and never the single letters c, C, r, R — prefer short words (Span, Load_w, Moment).
10. Reply in the user's language. "summary" = one short sentence.
Today's date: ${new Date().toISOString().slice(0, 10)}.

OPERATIONS (field "type" + fields):
- write {start:"B2", data:[[row1...],[row2...]]}  strings starting with "=" are formulas; null = leave cell unchanged
- format {range, bold, italic, underline, strikethrough, fontColor:"#hex", fontSize, fontName, fill:"#hex"|"none", numberFormat:"#,##0.00"|"0%"|"$#,##0"|"dd/mm/yyyy"|..., hAlign:left|center|right, vAlign:top|center|bottom, wrap, border:all|outline|bottom|none, borderColor, borderWeight:thin|medium|thick, columnWidth (points), rowHeight, autofit:true}
- clear {range, what:contents|formats|all}
- sort {range, hasHeaders, keys:[{column:"C", ascending:true}]}
- removeDuplicates {range, hasHeaders, columns:["A","B"]}  (omit columns = whole row)
- findReplace {range, find, replace, matchCase, wholeCell}
- merge {range, across} | unmerge {range}
- conditionalFormat {range, rule:greaterThan|lessThan|between|equalTo|notEqualTo|greaterOrEqual|lessOrEqual|textContains|duplicates|unique|top|bottom|colorScale|dataBar|formula, value1, value2, text, count, formula:"=$C2>100" (relative to range's top-left), fontColor, fill, bold}
- dataValidation {range, list:["Yes","No"]} | {range, number:{min,max,whole:true}} | {range, date:{min:"2026-01-01", max:"2026-12-31"}}, errorMessage
- chart {range, chartType:column|stackedColumn|bar|line|lineMarkers|pie|doughnut|scatter|area, title, xTitle, yTitle, seriesBy:auto|columns|rows, position:"H2"}
- table {range, hasHeaders, style:"TableStyleMedium2", name}
- pivot {source, destination:"newSheet"|"H2", rows:["Header"], columns:[], values:[{field:"Sales", summarizeBy:sum|count|average|max|min}], filters:[]}
- freeze {rows, columns}   (0,0 = unfreeze)
- insert {range:"5:6"|"C:C"|"B2:C3", shift:down|right}
- delete {range, shift:up|left}
- newSheet {name}
- name {range:"B2", name:"Span"}   creates (or replaces) an Excel named range; it can also name a block, e.g. {range:"B2:B20", name:"Rates"}`;
}
function buildUserText(prompt, c, allowOutside) {
  const recent = state.chat.slice(-4).map((x, i) => `${i + 1}. "${shortVal(x.prompt, 120)}" → ${shortVal(x.summary, 120)}`).join("\n");
  return [
    `Workbook sheets: ${c.sheets.join(", ")}`,
    `Active sheet: ${c.sheet}${c.used ? ` (used range ${stripSheet(c.used)})` : " (empty)"}${c.tables.length ? ` · tables: ${c.tables.join(", ")}` : ""}`,
    c.names && c.names.length ? `Named cells: ${c.names.slice(0, 60).map((n) => `${n.name} = ${n.ref}`).join(", ")}` : "",
    `USER SELECTION: ${c.sel.address} (${c.sel.rows} rows × ${c.sel.cols} cols)`,
    `Selection contents${c.selTruncated ? " (truncated)" : ""}:\n${c.selText}`,
    `Sheet overview (top-left of used range):\n${c.overview}`,
    recent ? `Recent requests this session:\n${recent}` : "",
    allowOutside ? "The user ALLOWS edits outside the selection for this request." : "Edits outside the selection are NOT allowed except filling empty neighbouring cells.",
    `REQUEST: ${prompt}`,
  ].filter(Boolean).join("\n\n");
}
async function callClaude(userText, image) {
  if (!useOwnKey()) {
    if (!PROXY) throw new Error("The AI isn't connected. Open Settings.");
    let res;
    try {
      res = await fetch(`${PROXY}/v1/vibe`, { method: "POST", headers: { "content-type": "application/json" },
        body: JSON.stringify({ device: deviceId(), text: userText, image: image ? { mediaType: image.mediaType, base64: image.base64 } : null, version: VERSION }) });
    } catch { throw new Error("Couldn't reach the Vibe Sheet AI service. Check your internet connection."); }
    const j = await res.json().catch(() => ({}));
    if (j.remaining != null) setQuota(j.remaining, j.limit);
    if (!res.ok) throw new Error(j.error || `AI service error ${res.status}.`);
    return { ...j.result, usage: j.usage, model: j.model };
  }
  const key = store.get("apiKey");
  if (!key) throw new Error("Add your API key in Settings first.");
  let model = store.get("model");
  if (!model) { model = pickDefaultModel(await fetchModels(key)); if (!model) throw new Error("No models available for this key."); store.set("model", model); }
  const content = [];
  if (image) content.push({ type: "image", source: { type: "base64", media_type: image.mediaType, data: image.base64 } });
  content.push({ type: "text", text: userText });
  const res = await fetch(API_URL, {
    method: "POST", headers: apiHeaders(key),
    body: JSON.stringify({ model, max_tokens: 8192, system: systemPrompt(), tools: [TOOL], tool_choice: { type: "tool", name: TOOL.name }, messages: [{ role: "user", content }] }),
  });
  if (!res.ok) {
    let detail = ""; try { detail = (await res.json()).error.message; } catch { /* ignore */ }
    if (res.status === 401) throw new Error("API key rejected. Check it in Settings.");
    if (res.status === 429 || res.status === 529) throw new Error("The AI is busy or you hit a rate limit — try again in a moment.");
    if (res.status === 404) throw new Error(`Model "${model}" not found. Open Settings → Load → Save.`);
    throw new Error(`AI error ${res.status}: ${detail}`);
  }
  const j = await res.json();
  if (j.stop_reason === "max_tokens") throw new Error("That response was too big. Select a smaller range or split the request.");
  const tu = (j.content || []).find((b) => b.type === "tool_use");
  if (!tu) throw new Error("The AI didn't return any changes. Try rephrasing.");
  return { ...tu.input, usage: j.usage, model };
}

/* ---------------- send → preview ---------------- */
async function send() {
  if (state.busy) return;
  const prompt = $("prompt").value.trim();
  if (!prompt && !state.image) { $("prompt").focus(); return; }
  if (!state.inExcel) { setStatus("Open this panel inside Excel to use it.", "err"); return; }
  if (!aiReady()) { refreshKeyNotice(); showTab("settings"); return; }
  state.lastPrompt = prompt || "(image)";
  $("result").classList.add("hidden"); $("applied").classList.add("hidden");
  setBusy(true, "Reading your selection…");
  try {
    const c = await gatherContext();
    const allowOutside = $("allowOutside").checked;
    state.req = { prompt: prompt || "(image)", ctx: c, allowOutside };
    setBusy(true, state.image ? "Reading the image and thinking…" : "Thinking…");
    const ai = await callClaude(buildUserText(prompt || "Use the attached image.", c, allowOutside), state.image);
    setBusy(true, "Preparing preview…");
    const items = await buildPreview(Array.isArray(ai.operations) ? ai.operations : [], c, allowOutside);
    state.pending = { summary: ai.summary || "", message: ai.message || "", items, usage: ai.usage, model: ai.model };
    renderPending();
    setStatus("");
  } catch (e) {
    setStatus(esc(e.message || String(e)), "err");
  } finally { setBusy(false); if (!$("status").classList.contains("err") && !state.busy) setStatus(""); }
}

/* Work out what each op touches, whether it's outside the selection, and the before/after diff. */
async function buildPreview(ops, c, allowOutside) {
  const defSheet = c.sheet;
  const newSheets = new Set(ops.filter((o) => o.type === "newSheet" && o.name).map((o) => o.name));
  const items = [];
  await Excel.run(async (ctx) => {
    for (const op of ops) {
      const it = { op, title: opTitle(op), desc: opDesc(op), checked: true, badges: [], diffs: [], target: null };
      items.push(it);
      if (!OP_TYPES.includes(op.type)) { it.checked = false; it.badges.push(["warn", "unknown action"]); continue; }
      if (op.type === "name") {
        const nm = cleanName(op.name);
        if ((c.names || []).some((n) => n.name.toLowerCase() === nm.toLowerCase())) it.badges.push(["warn", "replaces existing name"]);
        if (nm !== String(op.name || "")) it.badges.push(["info", "name adjusted for Excel"]);
        continue;
      }
      let spec = null;
      if (op.type === "write") spec = op.start;
      else if (op.type === "pivot") spec = op.destination && op.destination !== "newSheet" ? op.destination : null;
      else if (MODIFYING.has(op.type)) spec = op.range;
      if (!spec) continue;
      const { sheet } = splitSheet(spec);
      if (sheet && newSheets.has(sheet)) { it.badges.push(["info", "on new sheet"]); if (op.type === "write") it.diffs = writeDiffNew(op); continue; }
      try {
        let rng;
        if (op.type === "write") { const d = norm2D(op.data); rng = getRange(ctx, op.start, defSheet).getResizedRange(d.length - 1, d[0].length - 1); }
        else rng = getRange(ctx, spec, defSheet);
        rng.load("address,rowIndex,columnIndex,rowCount,columnCount,cellCount"); rng.worksheet.load("name");
        await ctx.sync();
        const t = { sheet: rng.worksheet.name, address: rng.address, r1: rng.rowIndex, c1: rng.columnIndex, r2: rng.rowIndex + rng.rowCount - 1, c2: rng.columnIndex + rng.columnCount - 1, cells: rng.cellCount };
        it.target = t;
        const s = c.sel;
        const inside = t.sheet === c.sheet && t.r1 >= s.r1 && t.r2 <= s.r2 && t.c1 >= s.c1 && t.c2 <= s.c2;
        let needVals = op.type === "write" || (!inside && !STRUCTURAL.has(op.type));
        if (needVals && t.cells <= CHECK_LIMIT) { rng.load("formulas"); await ctx.sync(); }
        else needVals = false;
        if (op.type === "write") it.diffs = writeDiff(op, t, rng.formulas);
        if (!inside) {
          let nonEmpty = 0;
          if (STRUCTURAL.has(op.type) || !needVals) nonEmpty = -1;
          else rng.formulas.forEach((row, r) => row.forEach((f, cc) => {
            const R = t.r1 + r, C = t.c1 + cc;
            const out = t.sheet !== c.sheet || R < s.r1 || R > s.r2 || C < s.c1 || C > s.c2;
            if (out && f !== "" && f !== null) nonEmpty++;
          }));
          if (nonEmpty === 0) it.badges.push(["info", "extends into empty cells"]);
          else {
            it.badges.push(["warn", nonEmpty > 0 ? `outside selection · touches ${nonEmpty} filled cell${nonEmpty > 1 ? "s" : ""}` : "outside your selection"]);
            it.checked = allowOutside;
          }
        }
        if (op.type === "write" && !it.diffs.length) { it.checked = false; it.badges.push(["info", "no change"]); }
      } catch (e) {
        it.checked = false; it.badges.push(["warn", "invalid: " + shortVal(e.message, 60)]);
      }
    }
  });
  return items;
}
function writeDiff(op, t, oldF) {
  const d = norm2D(op.data), diffs = [];
  for (let r = 0; r < d.length; r++) for (let c = 0; c < d[r].length; c++) {
    const nv = d[r][c]; if (nv === null) continue;
    const ov = oldF ? oldF[r][c] : "";
    if (String(ov) !== String(nv)) diffs.push({ addr: cellAddr(t.r1 + r, t.c1 + c), old: ov, new: nv });
  }
  return diffs;
}
function writeDiffNew(op) {
  const d = norm2D(op.data), diffs = []; const { ref } = splitSheet(op.start);
  const m = /^\$?([A-Z]+)\$?(\d+)/i.exec(ref); const r0 = m ? +m[2] - 1 : 0, c0 = m ? colToNum(m[1]) : 0;
  d.forEach((row, r) => row.forEach((v, c) => { if (v !== null && v !== "") diffs.push({ addr: cellAddr(r0 + r, c0 + c), old: "", new: v }); }));
  return diffs;
}
function opTitle(op) {
  const a = (x) => stripSheet(x || "");
  switch (op.type) {
    case "write": return `Write cells from ${a(op.start)}`;
    case "format": return `Format ${a(op.range)}`;
    case "clear": return `Clear ${op.what || "contents"} in ${a(op.range)}`;
    case "sort": return `Sort ${a(op.range)}`;
    case "removeDuplicates": return `Remove duplicates in ${a(op.range)}`;
    case "findReplace": return `Replace “${shortVal(op.find, 20)}” → “${shortVal(op.replace, 20)}” in ${a(op.range)}`;
    case "merge": return `Merge ${a(op.range)}`;
    case "unmerge": return `Unmerge ${a(op.range)}`;
    case "conditionalFormat": return `Highlight rule on ${a(op.range)}`;
    case "dataValidation": return `Dropdown / validation on ${a(op.range)}`;
    case "chart": return `${cap(op.chartType || "column")} chart of ${a(op.range)}`;
    case "table": return `Format ${a(op.range)} as table`;
    case "pivot": return `Pivot table from ${a(op.source)}`;
    case "freeze": return (op.rows || op.columns) ? `Freeze ${op.rows || 0} row(s), ${op.columns || 0} column(s)` : "Unfreeze panes";
    case "insert": return `Insert ${a(op.range)}`;
    case "delete": return `Delete ${a(op.range)}`;
    case "newSheet": return `New sheet “${op.name}”`;
    case "name": return `Name ${stripSheet(op.range || "")} “${cleanName(op.name)}”`;
    default: return op.type;
  }
}
function opDesc(op) {
  switch (op.type) {
    case "format": {
      const p = [];
      if (op.bold) p.push("bold"); if (op.bold === false) p.push("not bold"); if (op.italic) p.push("italic"); if (op.underline) p.push("underline");
      if (op.strikethrough) p.push("strikethrough"); if (op.fontColor) p.push(`text ${op.fontColor}`); if (op.fill) p.push(`fill ${op.fill}`);
      if (op.fontSize) p.push(`${op.fontSize}pt`); if (op.fontName) p.push(op.fontName); if (op.numberFormat) p.push(`format ${op.numberFormat}`);
      if (op.hAlign) p.push(`align ${op.hAlign}`); if (op.wrap) p.push("wrap text"); if (op.border) p.push(`${op.border} borders`);
      if (op.columnWidth) p.push(`col width ${op.columnWidth}`); if (op.autofit) p.push("autofit columns");
      return p.join(", ");
    }
    case "sort": return (op.keys || []).map((k) => `${k.column} ${k.ascending === false ? "Z→A / high→low" : "A→Z / low→high"}`).join(", ");
    case "conditionalFormat": return [op.rule, op.value1, op.value2, op.text, op.formula, op.fill && `fill ${op.fill}`, op.fontColor && `text ${op.fontColor}`].filter((x) => x !== undefined && x !== null && x !== "").join(" · ");
    case "dataValidation": return op.list ? `List: ${op.list.join(", ")}` : op.number ? `Number ${op.number.min ?? ""}–${op.number.max ?? ""}` : op.date ? `Date ${op.date.min ?? ""}–${op.date.max ?? ""}` : "";
    case "pivot": return [`rows: ${(op.rows || []).join(", ") || "—"}`, op.columns && op.columns.length ? `cols: ${op.columns.join(", ")}` : "", `values: ${(op.values || []).map((v) => `${v.summarizeBy || "sum"} of ${v.field}`).join(", ")}`, `→ ${op.destination || "newSheet"}`].filter(Boolean).join(" · ");
    case "chart": return [op.title, op.position && `at ${op.position}`].filter(Boolean).join(" · ");
    case "removeDuplicates": return op.columns ? `by ${op.columns.join(", ")}` : "whole rows";
    default: return "";
  }
}
function renderPending() {
  const p = state.pending;
  $("resSummary").textContent = p.summary || (p.items.length ? "Proposed changes" : "Answer");
  $("resMessage").textContent = p.message || "";
  $("resMessage").classList.toggle("hidden", !p.message);
  $("ops").innerHTML = p.items.map((it, i) => {
    const badges = it.badges.map(([k, t]) => `<span class="badge ${k}">${esc(t)}</span>`).join("");
    let diff = "";
    if (it.diffs.length) {
      const rows = it.diffs.slice(0, 30).map((d) => `<tr><td class="addr">${esc(d.addr)}</td><td class="${d.old === "" ? "" : "old"}" title="${esc(d.old)}">${esc(shortVal(d.old, 40))}</td><td class="new" title="${esc(d.new)}">${esc(shortVal(d.new, 40))}</td></tr>`).join("");
      const more = it.diffs.length > 30 ? `<tr><td colspan="3" class="muted">+ ${it.diffs.length - 30} more cells</td></tr>` : "";
      diff = `<table class="diff"><tr><th class="addr">Cell</th><th>Before</th><th>After</th></tr>${rows}${more}</table>`;
    }
    const count = it.diffs.length ? ` <span class="muted">(${it.diffs.length} cell${it.diffs.length > 1 ? "s" : ""})</span>` : "";
    return `<div class="op"><label class="op-head"><input type="checkbox" data-i="${i}" ${it.checked ? "checked" : ""}/>
      <div><div class="op-title">${esc(it.title)}${count}${badges}</div>${it.desc ? `<div class="op-desc">${esc(it.desc)}</div>` : ""}</div></label>${diff}</div>`;
  }).join("");
  $("ops").querySelectorAll("input[type=checkbox]").forEach((cb) => cb.addEventListener("change", () => { p.items[+cb.dataset.i].checked = cb.checked; }));
  $("resActions").classList.toggle("hidden", !p.items.length);
  const u = p.usage; $("usage").textContent = u ? `${p.model} · ${u.input_tokens} in / ${u.output_tokens} out tokens` : "";
  $("result").classList.remove("hidden");
  state.chat.push({ prompt: state.req.prompt, summary: p.summary || p.message });
  state.chat = state.chat.slice(-6);
}

/* ---------------- accept → apply ---------------- */
async function accept() {
  const p = state.pending; if (!p || state.busy) return;
  const rank = (t) => (t === "newSheet" ? 0 : t === "name" ? 1 : 2);   // sheets, then names, then the rest
  const chosen = p.items.filter((it) => it.checked).sort((a, b) => rank(a.op.type) - rank(b.op.type));
  if (!chosen.length) { setStatus("Nothing selected to apply.", "err"); return; }
  setBusy(true, "Applying…");
  const entry = { id: Date.now(), time: Date.now(), prompt: state.req.prompt, summary: p.summary, titles: [], steps: [], undone: false, records: [] };
  const errors = [];
  for (const it of chosen) {
    try {
      const res = await applyOp(it.op, state.req.ctx.sheet);
      entry.steps.push(...res.steps); entry.titles.push(it.title);
      if (res.noUndo) errors.push(`“${it.title}” applied, but it's too large to undo.`);
      if (it.target && MODIFYING.has(it.op.type)) entry.records.push({ sheet: it.target.sheet, r1: it.target.r1, c1: it.target.c1, r2: Math.min(it.target.r2, it.target.r1 + 5000), c2: it.target.c2, prompt: entry.prompt, time: entry.time, entry: entry.id });
    } catch (e) { errors.push(`“${it.title}” failed: ${/payload size/i.test(e.message) ? "that range is too big for Excel to change in one go. Select just your data (not whole columns) and try again." : e.message}`); }
  }
  setBusy(false);
  if (entry.titles.length) {
    state.history.unshift(entry);
    setCellPrompts(getCellPrompts().concat(entry.records));
    $("appliedText").textContent = `✓ Applied ${entry.titles.length} change${entry.titles.length > 1 ? "s" : ""}`;
    $("applied").classList.remove("hidden");
    renderHistory();
    clearImage(); $("prompt").value = "";
  }
  state.pending = null; $("result").classList.add("hidden");
  setStatus(errors.length ? errors.map(esc).join("<br>") : "", errors.length ? "err" : "");
  refreshSelection();
}

function getRange(ctx, addr, defSheet) {
  const { sheet, ref } = splitSheet(addr);
  const ws = ctx.workbook.worksheets.getItem(sheet || defSheet);
  return ws.getRange(ref.replace(/\$/g, ""));
}

/* Snapshot values + formats so any cell change can be undone exactly. */
async function snapshot(ctx, rng) {
  rng.load("address,cellCount,rowCount,columnCount"); await ctx.sync();
  let target = rng;
  if (rng.cellCount > SNAP_LIMIT) {
    const used = rng.worksheet.getUsedRangeOrNullObject(true);
    used.load("address"); await ctx.sync();
    if (used.isNullObject) return { empty: true };
    target = rng.getIntersectionOrNullObject(used);
    target.load("address,cellCount"); await ctx.sync();
    if (target.isNullObject) return { empty: true };
    if (target.cellCount > SNAP_LIMIT) return null;
  }
  target.load("address,formulas,numberFormat");
  const props = target.getCellProperties({ format: { font: { bold: true, italic: true, color: true, size: true, name: true, underline: true, strikethrough: true },
    fill: { color: true, pattern: true }, horizontalAlignment: true, verticalAlignment: true, wrapText: true, borders: { color: true, style: true, weight: true } } });
  const cols = target.getColumnProperties({ format: { columnWidth: true } });
  const rows = target.getRowProperties({ format: { rowHeight: true } });
  await ctx.sync();
  return { address: target.address, formulas: target.formulas, numberFormat: target.numberFormat, props: props.value, cols: cols.value, rows: rows.value };
}
async function restore(ctx, snap) {
  if (!snap || snap.empty) return;
  const r = getRange(ctx, snap.address);
  r.formulas = snap.formulas;
  r.numberFormat = snap.numberFormat;
  r.format.fill.clear();
  const clean = (withBorders) => snap.props.map((row) => row.map((p) => {
    const f = { ...(p.format || {}) };
    if (f.fill && f.fill.pattern === "None") delete f.fill; else if (f.fill) f.fill = { color: f.fill.color };
    if (!withBorders) delete f.borders;
    if (f.font) { const fo = { ...f.font }; Object.keys(fo).forEach((k) => fo[k] === null && delete fo[k]); f.font = fo; }
    return { format: f };
  }));
  await ctx.sync();
  try { r.setCellProperties(clean(true)); await ctx.sync(); }
  catch { r.setCellProperties(clean(false)); await ctx.sync(); }
  try {
    r.setColumnProperties(snap.cols.map((c) => ({ format: { columnWidth: c.format.columnWidth } })));
    r.setRowProperties(snap.rows.map((c) => ({ format: { rowHeight: c.format.rowHeight } })));
    await ctx.sync();
  } catch { /* widths are best-effort */ }
}

async function applyOp(op, defSheet) {
  return Excel.run(async (ctx) => {
    const wb = ctx.workbook, steps = [];
    let noUndo = false;
    const R = (a) => getRange(ctx, a, defSheet);
    // Whole columns/rows (A:AO, 1:500) hold millions of cells; trim them to the data area so Excel never has to move that much.
    const RC = async (a) => {
      const rng = R(a); rng.load("cellCount,rowCount,columnCount"); await ctx.sync();
      if (rng.cellCount <= 200000) return rng;
      const used = rng.worksheet.getUsedRangeOrNullObject(true); used.load("address"); await ctx.sync();
      if (used.isNullObject) return rng.getCell(0, 0);
      let anchor = null;
      if (rng.rowCount > 100000 && rng.columnCount <= 1000) anchor = rng.getRow(0);
      else if (rng.columnCount > 1000 && rng.rowCount <= 100000) anchor = rng.getColumn(0);
      const box = anchor ? used.getBoundingRect(anchor) : used;
      const out = rng.getIntersectionOrNullObject(box); out.load("address"); await ctx.sync();
      return out.isNullObject ? rng.getCell(0, 0) : out;
    };
    const snapPush = async (rng) => { const s = await snapshot(ctx, rng); if (s) steps.push({ kind: "restore", snap: s }); else noUndo = true; };

    switch (op.type) {
      case "write": {
        const d = norm2D(op.data);
        const rng = R(op.start).getResizedRange(d.length - 1, d[0].length - 1);
        await snapPush(rng);
        rng.load("formulas"); await ctx.sync();
        rng.formulas = d.map((row, r) => row.map((v, c) => (v === null ? rng.formulas[r][c] : v)));
        if (op.numberFormat) rng.numberFormat = fill2D(d.length, d[0].length, op.numberFormat);
        break;
      }
      case "format": {
        const rng = await RC(op.range); await snapPush(rng);
        rng.load("rowCount,columnCount"); await ctx.sync();
        const f = rng.format;
        if (op.bold != null) f.font.bold = !!op.bold;
        if (op.italic != null) f.font.italic = !!op.italic;
        if (op.underline != null) f.font.underline = op.underline ? "Single" : "None";
        if (op.strikethrough != null) f.font.strikethrough = !!op.strikethrough;
        if (op.fontColor) f.font.color = op.fontColor;
        if (op.fontSize) f.font.size = +op.fontSize;
        if (op.fontName) f.font.name = op.fontName;
        if (op.fill) { if (String(op.fill).toLowerCase() === "none") f.fill.clear(); else f.fill.color = op.fill; }
        if (op.numberFormat && rng.rowCount * rng.columnCount <= 1e6) rng.numberFormat = fill2D(rng.rowCount, rng.columnCount, op.numberFormat);
        if (op.hAlign) f.horizontalAlignment = cap(op.hAlign);
        if (op.vAlign) f.verticalAlignment = cap(op.vAlign);
        if (op.wrap != null) f.wrapText = !!op.wrap;
        if (op.border) {
          let edges = { all: ["EdgeTop", "EdgeBottom", "EdgeLeft", "EdgeRight", "InsideVertical", "InsideHorizontal"], outline: ["EdgeTop", "EdgeBottom", "EdgeLeft", "EdgeRight"], bottom: ["EdgeBottom"], none: ["EdgeTop", "EdgeBottom", "EdgeLeft", "EdgeRight", "InsideVertical", "InsideHorizontal"] }[op.border] || [];
          if (rng.columnCount < 2) edges = edges.filter((e) => e !== "InsideVertical");
          if (rng.rowCount < 2) edges = edges.filter((e) => e !== "InsideHorizontal");
          for (const e of edges) {
            const b = f.borders.getItem(e);
            if (op.border === "none") b.style = "None";
            else { b.style = "Continuous"; b.weight = cap(op.borderWeight || "thin"); b.color = op.borderColor || "#000000"; }
          }
        }
        if (op.columnWidth) f.columnWidth = +op.columnWidth;
        if (op.rowHeight) f.rowHeight = +op.rowHeight;
        if (op.autofit) f.autofitColumns();
        break;
      }
      case "clear": {
        const rng = await RC(op.range); await snapPush(rng);
        rng.clear({ contents: "Contents", formats: "Formats", all: "All" }[op.what] || "Contents");
        break;
      }
      case "sort": {
        const rng = await RC(op.range); await snapPush(rng);
        rng.load("columnIndex"); await ctx.sync();
        const keys = (op.keys && op.keys.length ? op.keys : [{ column: op.column ?? 0, ascending: op.ascending !== false }])
          .map((k) => ({ key: colOffset(k.column, rng.columnIndex), ascending: k.ascending !== false }));
        rng.sort.apply(keys, false, !!op.hasHeaders, "Rows");
        break;
      }
      case "removeDuplicates": {
        const rng = await RC(op.range); await snapPush(rng);
        rng.load("columnIndex,columnCount"); await ctx.sync();
        const cols = op.columns && op.columns.length ? op.columns.map((c) => colOffset(c, rng.columnIndex)) : [...Array(rng.columnCount).keys()];
        rng.removeDuplicates(cols, !!op.hasHeaders);
        break;
      }
      case "findReplace": {
        const rng = await RC(op.range); await snapPush(rng);
        rng.replaceAll(String(op.find ?? ""), String(op.replace ?? ""), { completeMatch: !!op.wholeCell, matchCase: !!op.matchCase });
        break;
      }
      case "merge": {
        const rng = R(op.range); await snapPush(rng); rng.load("address"); await ctx.sync();
        rng.merge(!!op.across); steps.push({ kind: "unmerge", address: rng.address });
        break;
      }
      case "unmerge": { const rng = R(op.range); await snapPush(rng); rng.unmerge(); break; }
      case "conditionalFormat": {
        const rng = R(op.range); rng.load("address"); await ctx.sync();
        const cf = addConditionalFormat(rng, op); cf.load("id"); await ctx.sync();
        steps.push({ kind: "deleteCF", address: rng.address, id: cf.id });
        break;
      }
      case "dataValidation": {
        const rng = R(op.range); rng.load("address"); await ctx.sync();
        rng.dataValidation.clear();
        let rule = null;
        if (op.list) rule = { list: { inCellDropDown: true, source: op.list.map(String).join(",") } };
        else if (op.number) rule = numRule(op.number, op.number.whole ? "wholeNumber" : "decimal");
        else if (op.date) rule = numRule({ min: op.date.min && `=DATEVALUE("${op.date.min}")`, max: op.date.max && `=DATEVALUE("${op.date.max}")` }, "date");
        if (rule) rng.dataValidation.rule = rule;
        rng.dataValidation.errorAlert = { message: op.errorMessage || "That value isn't allowed here.", showAlert: true, style: "Stop", title: "Invalid entry" };
        steps.push({ kind: "clearValidation", address: rng.address });
        break;
      }
      case "chart": {
        const rng = R(op.range); rng.load("address,rowIndex,columnIndex,columnCount"); rng.worksheet.load("name"); await ctx.sync();
        const ws = rng.worksheet;
        const types = { column: "ColumnClustered", stackedcolumn: "ColumnStacked", bar: "BarClustered", line: "Line", linemarkers: "LineMarkers", pie: "Pie", doughnut: "Doughnut", scatter: "XYScatter", area: "Area" };
        const chart = ws.charts.add(types[String(op.chartType || "column").toLowerCase()] || "ColumnClustered", rng, cap(op.seriesBy || "auto"));
        if (op.title) chart.title.text = op.title;
        const isPie = /pie|doughnut/i.test(op.chartType || "");
        if (!isPie && op.xTitle) chart.axes.categoryAxis.title.text = op.xTitle;
        if (!isPie && op.yTitle) chart.axes.valueAxis.title.text = op.yTitle;
        const pos = op.position ? ws.getRange(stripSheet(op.position)) : ws.getCell(rng.rowIndex, rng.columnIndex + rng.columnCount + 1);
        chart.setPosition(pos);
        chart.load("name"); await ctx.sync();
        steps.push({ kind: "deleteChart", sheet: ws.name, name: chart.name });
        break;
      }
      case "table": {
        const rng = await RC(op.range); await snapPush(rng); rng.load("address"); rng.worksheet.load("name"); await ctx.sync();
        const t = rng.worksheet.tables.add(rng, op.hasHeaders !== false);
        if (op.style) t.style = op.style;
        if (op.name) t.name = String(op.name).replace(/[^A-Za-z0-9_]/g, "_");
        t.load("name"); await ctx.sync();
        steps.push({ kind: "tableToRange", sheet: rng.worksheet.name, name: t.name });
        break;
      }
      case "pivot": {
        const src = R(op.source);
        let dest;
        if (!op.destination || op.destination === "newSheet") {
          const sheets = wb.worksheets; sheets.load("items/name"); await ctx.sync();
          let n = 1, name = "Pivot"; const names = sheets.items.map((s) => s.name);
          while (names.includes(name)) name = "Pivot " + (++n);
          const ws = wb.worksheets.add(name); steps.push({ kind: "deleteSheet", name });
          dest = ws.getRange("A3");
        } else dest = R(op.destination);
        const name = "VibePivot" + Date.now().toString().slice(-6);
        const pt = wb.pivotTables.add(name, src, dest);
        await ctx.sync();
        steps.push({ kind: "deletePivot", name });
        const agg = { sum: "Sum", count: "Count", average: "Average", max: "Max", min: "Min" };
        (op.rows || []).forEach((h) => pt.rowHierarchies.add(pt.hierarchies.getItem(h)));
        (op.columns || []).forEach((h) => pt.columnHierarchies.add(pt.hierarchies.getItem(h)));
        (op.filters || []).forEach((h) => pt.filterHierarchies.add(pt.hierarchies.getItem(h)));
        await ctx.sync();
        for (const v of op.values || []) {
          const dh = pt.dataHierarchies.add(pt.hierarchies.getItem(v.field));
          dh.summarizeBy = agg[String(v.summarizeBy || "sum").toLowerCase()] || "Sum";
        }
        if (dest) dest.worksheet.activate();
        break;
      }
      case "freeze": {
        const ws = wb.worksheets.getItem(defSheet); ws.load("name");
        const rows = +op.rows || 0, cols = +op.columns || 0;
        ws.freezePanes.unfreeze();
        if (rows && cols) ws.freezePanes.freezeAt(ws.getRangeByIndexes(0, 0, rows, cols));
        else if (rows) ws.freezePanes.freezeRows(rows);
        else if (cols) ws.freezePanes.freezeColumns(cols);
        await ctx.sync();
        steps.push({ kind: "unfreeze", sheet: ws.name });
        break;
      }
      case "insert": {
        const rng = R(op.range);
        const ins = rng.insert(String(op.shift || "down").toLowerCase() === "right" ? "Right" : "Down");
        ins.load("address"); await ctx.sync();
        steps.push({ kind: "deleteInserted", address: ins.address, shift: String(op.shift || "down").toLowerCase() === "right" ? "Left" : "Up" });
        break;
      }
      case "delete": {
        const rng = R(op.range); rng.load("address"); await ctx.sync();
        const snap = await snapshot(ctx, rng); if (!snap) noUndo = true;
        const left = String(op.shift || "up").toLowerCase() === "left";
        rng.delete(left ? "Left" : "Up");
        steps.push({ kind: "reinsert", address: rng.address, shift: left ? "Right" : "Down", snap });
        break;
      }
      case "name": {
        const nm = cleanName(op.name);
        const rng = R(op.range); rng.load("address"); await ctx.sync();
        const old = wb.names.getItemOrNullObject(nm); old.load("formula"); await ctx.sync();
        const prev = old.isNullObject ? null : old.formula;
        if (!old.isNullObject) old.delete();
        wb.names.add(nm, rng, op.comment || "Named by Vibe Sheet");
        await ctx.sync();
        steps.push({ kind: "unname", name: nm, prev });
        break;
      }
      case "newSheet": {
        const ws = wb.worksheets.add(String(op.name || "Sheet").slice(0, 31)); ws.load("name"); await ctx.sync();
        steps.push({ kind: "deleteSheet", name: ws.name });
        break;
      }
      default: throw new Error("Unknown action " + op.type);
    }
    await ctx.sync();
    return { steps, noUndo };
  });
}
function numRule(o, kind) {
  let operator = "Between", f1 = o.min, f2 = o.max;
  if (o.min == null && o.max != null) { operator = "LessThanOrEqualTo"; f1 = o.max; f2 = undefined; }
  else if (o.min != null && o.max == null) { operator = "GreaterThanOrEqualTo"; }
  const r = { formula1: f1 != null ? String(f1) : "0", operator }; if (f2 != null && operator === "Between") r.formula2 = String(f2);
  return { [kind]: r };
}
function addConditionalFormat(rng, op) {
  const rule = String(op.rule || "").replace(/[\s_-]/g, "").toLowerCase();
  const fmt = (f) => {
    const hasStyle = op.fill || op.fontColor || op.bold;
    if (op.fill || !hasStyle) f.fill.color = op.fill || "#FFC7CE";
    if (op.fontColor || !hasStyle) f.font.color = op.fontColor || "#9C0006";
    if (op.bold) f.font.bold = true;
  };
  const ops = { greaterthan: "GreaterThan", lessthan: "LessThan", between: "Between", equalto: "EqualTo", notequalto: "NotEqualTo", greaterorequal: "GreaterThanOrEqual", lessorequal: "LessThanOrEqual" };
  let cf;
  if (ops[rule]) {
    cf = rng.conditionalFormats.add("CellValue"); fmt(cf.cellValue.format);
    const r = { formula1: toCfFormula(op.value1 ?? 0), operator: ops[rule] };
    if (rule === "between") r.formula2 = toCfFormula(op.value2 ?? 0);
    cf.cellValue.rule = r;
  } else if (rule === "textcontains") {
    cf = rng.conditionalFormats.add("ContainsText"); fmt(cf.textComparison.format);
    cf.textComparison.rule = { operator: "Contains", text: String(op.text ?? op.value1 ?? "") };
  } else if (rule === "duplicates" || rule === "unique") {
    cf = rng.conditionalFormats.add("PresetCriteria"); fmt(cf.preset.format);
    cf.preset.rule = { criterion: rule === "duplicates" ? "DuplicateValues" : "UniqueValues" };
  } else if (rule === "top" || rule === "bottom") {
    cf = rng.conditionalFormats.add("TopBottom"); fmt(cf.topBottom.format);
    cf.topBottom.rule = { rank: +(op.count || op.value1 || 10), type: rule === "top" ? "TopItems" : "BottomItems" };
  } else if (rule === "colorscale") {
    cf = rng.conditionalFormats.add("ColorScale");
    cf.colorScale.criteria = { minimum: { formula: null, type: "LowestValue", color: op.minColor || "#F8696B" },
      midpoint: { formula: "50", type: "Percentile", color: op.midColor || "#FFEB84" }, maximum: { formula: null, type: "HighestValue", color: op.maxColor || "#63BE7B" } };
  } else if (rule === "databar") {
    cf = rng.conditionalFormats.add("DataBar");
    if (op.fill) cf.dataBar.positiveFormat.fillColor = op.fill;
  } else if (rule === "formula" || op.formula) {
    cf = rng.conditionalFormats.add("Custom"); fmt(cf.custom.format);
    cf.custom.rule.formula = String(op.formula || "").startsWith("=") ? op.formula : "=" + op.formula;
  } else throw new Error("Unknown highlight rule: " + op.rule);
  return cf;
}

/* ---------------- undo ---------------- */
async function undoEntry(entry) {
  if (!entry || entry.undone) return;
  setBusy(true, "Undoing…");
  const errors = [];
  for (const s of entry.steps.slice().reverse()) {
    try {
      await Excel.run(async (ctx) => {
        const wb = ctx.workbook;
        switch (s.kind) {
          case "restore": await restore(ctx, s.snap); break;
          case "unmerge": getRange(ctx, s.address).unmerge(); break;
          case "deleteCF": getRange(ctx, s.address).conditionalFormats.getItem(s.id).delete(); break;
          case "clearValidation": getRange(ctx, s.address).dataValidation.clear(); break;
          case "deleteChart": wb.worksheets.getItem(s.sheet).charts.getItem(s.name).delete(); break;
          case "tableToRange": wb.worksheets.getItem(s.sheet).tables.getItem(s.name).convertToRange(); break;
          case "deletePivot": wb.pivotTables.getItem(s.name).delete(); break;
          case "deleteSheet": wb.worksheets.getItem(s.name).delete(); break;
          case "unname": { const n = wb.names.getItemOrNullObject(s.name); await ctx.sync(); if (!n.isNullObject) n.delete(); if (s.prev) wb.names.add(s.name, s.prev); break; }
          case "unfreeze": wb.worksheets.getItem(s.sheet).freezePanes.unfreeze(); break;
          case "deleteInserted": getRange(ctx, s.address).delete(s.shift); break;
          case "reinsert": getRange(ctx, s.address).insert(s.shift); await ctx.sync(); if (s.snap) await restore(ctx, s.snap); break;
        }
        await ctx.sync();
      });
    } catch (e) { errors.push(e.message); }
  }
  entry.undone = true;
  setCellPrompts(getCellPrompts().filter((p) => p.entry !== entry.id));
  setBusy(false);
  setStatus(errors.length ? "Undone, with some issues:<br>" + errors.map(esc).join("<br>") : "↩︎ Undone.", errors.length ? "err" : "ok");
  $("applied").classList.add("hidden");
  renderHistory(); refreshSelection();
}
function undoLast() { undoEntry(state.history.find((h) => !h.undone)); }
function renderHistory() {
  const list = $("historyList");
  if (!state.history.length) { list.innerHTML = `<li class="muted">Nothing yet.</li>`; return; }
  const newest = state.history.find((h) => !h.undone);
  list.innerHTML = state.history.map((h) => `<li class="${h.undone ? "undone" : ""}">
    <div class="h-prompt">${esc(shortVal(h.prompt, 140))}</div>
    <div class="h-meta">${esc(h.summary || "")}</div>
    <div class="h-meta">${esc(h.titles.join(" · "))}</div>
    <div class="h-row"><span class="h-meta">${new Date(h.time).toLocaleTimeString()}${h.undone ? " · undone" : ""}</span>
    ${!h.undone && h === newest ? `<button class="ghost" data-undo="${h.id}">Undo</button>` : ""}</div></li>`).join("");
  list.querySelectorAll("[data-undo]").forEach((b) => b.addEventListener("click", () => undoEntry(state.history.find((h) => h.id === +b.dataset.undo))));
}

/* ---------------- boot ---------------- */
function boot(inExcel) {
  state.inExcel = inExcel;
  initUI();
  if (!inExcel) {
    $("selAddr").textContent = "open inside Excel";
    return;
  }
  if (!Office.context.requirements.isSetSupported("ExcelApi", "1.9")) {
    setStatus("Your Excel version is too old for some features. Use Microsoft 365, Excel 2021+ or Excel on the web.", "err");
  }
  Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, onSelectionChanged);
  refreshSelection();
}
if (typeof Office !== "undefined") {
  Office.onReady((info) => boot(info.host === Office.HostType.Excel));
} else {
  document.addEventListener("DOMContentLoaded", () => boot(false));
}
