/* Probridge Land Bot - chat screen used by both the Excel add-in and the phone app.
 *
 * Inside Excel  (index.html?host=excel): saves rows with Office.js into the open workbook.
 * On the phone  (installed PWA):          saves rows with Microsoft Graph into OneDrive.
 */
"use strict";

const IN_EXCEL = new URLSearchParams(location.search).get("host") === "excel";
const GRAPH_SCOPES = ["Files.ReadWrite", "User.Read"];

const S = {
  cfg: null,
  msal: null,
  account: null,
  sheet: null, // { tableName, headers, keys, rows }
};

const $ = (id) => document.getElementById(id);
const chat = $("chat");

/* ------------------------------------------------------------------ chat */

function say(text, kind = "bot") {
  const div = document.createElement("div");
  div.className = `bubble ${kind === "user" ? "user" : "bot"} ${kind === "error" ? "error" : ""} ${kind === "success" ? "success" : ""}`;
  div.textContent = text;
  chat.appendChild(div);
  chat.scrollTop = chat.scrollHeight;
  return div;
}

function working(text) {
  const div = say("");
  const start = Date.now();
  const draw = () => {
    div.innerHTML = `<span class="spinner"></span>${text} <span class="muted small">${Math.round((Date.now() - start) / 1000)}s</span>`;
  };
  draw();
  const t = setInterval(draw, 1000);
  return { done: () => { clearInterval(t); div.remove(); } };
}

async function api(path, opts = {}) {
  const res = await fetch(path, opts);
  let body = null;
  try { body = await res.json(); } catch { /* not json */ }
  if (!res.ok) throw new Error(body?.detail || `Server error ${res.status}`);
  return body;
}

/* --------------------------------------------------------------- helpers */

const norm = (s) => String(s ?? "").toLowerCase().replace(/[^a-z0-9]/g, "");

function todayText() {
  const d = new Date();
  const p = (n) => String(n).padStart(2, "0");
  return `${p(d.getDate())}.${p(d.getMonth() + 1)}.${d.getFullYear()}`;
}

/** Match each Excel header to a column key from config.json (handles the two "Status" columns). */
function mapHeaders(headers) {
  const used = new Set();
  return headers.map((h) => {
    const nh = norm(h);
    if (!nh) return null;
    const col = S.cfg.columns.find((c) => {
      if (used.has(c.key)) return false;
      const nc = norm(c.header);
      return nh === nc || nh.startsWith(nc) || nc.startsWith(nh);
    });
    if (!col) return null;
    used.add(col.key);
    return col.key;
  });
}

function columnValues(key) {
  if (!S.sheet) return [];
  const i = S.sheet.keys.indexOf(key);
  return i < 0 ? [] : S.sheet.rows.map((r) => r[i]);
}

function nextSerial() {
  const nums = columnValues("sno").map((v) => parseInt(v, 10)).filter((n) => !isNaN(n));
  return nums.length ? Math.max(...nums) + 1 : 1;
}

/** Latest known phone number for each mediator name, from rows already in Excel. */
function mediatorBook() {
  const names = columnValues("mediator_name");
  const nums = columnValues("number");
  const book = new Map();
  names.forEach((n, i) => {
    const name = String(n ?? "").trim();
    const num = String(nums[i] ?? "").replace(/\D/g, "");
    if (name && num.length >= 10) book.set(name.toLowerCase(), num.slice(-10));
    else if (name && !book.has(name.toLowerCase())) book.set(name.toLowerCase(), "");
  });
  return book;
}

function findDuplicate(values) {
  if (!S.sheet) return null;
  const pins = columnValues("google_pin").map(norm);
  const areas = columnValues("exact_area").map(norm);
  const nums = columnValues("number").map((v) => String(v ?? "").replace(/\D/g, "").slice(-10));
  const serials = columnValues("sno");
  const pin = norm(values.google_pin);
  const area = norm(values.exact_area);
  const num = String(values.number ?? "").replace(/\D/g, "").slice(-10);
  for (let i = 0; i < S.sheet.rows.length; i++) {
    if (pin && pins[i] === pin) return `Same Google pin already saved (S.No ${serials[i] ?? "?"}).`;
    if (area && num && areas[i] === area && nums[i] === num)
      return `Same area and mediator number already saved (S.No ${serials[i] ?? "?"}).`;
  }
  return null;
}

/* ----------------------------------------------------------- Excel (laptop) */

function loadOfficeJs() {
  return new Promise((resolve, reject) => {
    const s = document.createElement("script");
    s.src = "https://appsforoffice.microsoft.com/lib/1/hosted/office.js";
    s.onload = () => Office.onReady(resolve);
    s.onerror = () => reject(new Error("Could not load Office.js (check internet)."));
    document.head.appendChild(s);
  });
}

async function excelGetTable(ctx) {
  const name = S.cfg.excel.table_name;
  let table = ctx.workbook.tables.getItemOrNullObject(name);
  await ctx.sync();
  if (table.isNullObject) {
    const all = ctx.workbook.tables.load("items/name");
    await ctx.sync();
    if (all.items.length === 1) return all.items[0];
    throw new Error(
      `Excel table "${name}" not found. Select your data (with headers), press Ctrl+T, ` +
      `then in Table Design set Table Name to "${name}".`);
  }
  return table;
}

async function excelLoadSheet() {
  return Excel.run(async (ctx) => {
    const table = await excelGetTable(ctx);
    table.load("name");
    const head = table.getHeaderRowRange().load("values");
    const body = table.getDataBodyRange().load("values");
    await ctx.sync();
    const headers = head.values[0].map(String);
    const rows = body.values.filter((r) => r.some((v) => v !== "" && v !== null));
    return { tableName: table.name, headers, keys: mapHeaders(headers), rows };
  });
}

async function excelAddRow(row) {
  return Excel.run(async (ctx) => {
    const table = await excelGetTable(ctx);
    table.rows.add(null, [row]);
    await ctx.sync();
    try { // jump to the new row so you can see it
      table.getDataBodyRange().getLastRow().select();
      await ctx.sync();
    } catch { /* selecting is optional */ }
  });
}

/* ------------------------------------------------- Microsoft Graph (phone) */

function graphReady() {
  const id = S.cfg.microsoft_login.client_id;
  return id && !id.startsWith("PASTE-");
}

async function initMsal() {
  if (!graphReady()) {
    say("Phone saving is not set up yet: put your Application (client) ID in config.json " +
        "(see SETUP-GUIDE.md, step 6). You can still test reading messages.", "error");
    return;
  }
  S.msal = new msal.PublicClientApplication({
    auth: {
      clientId: S.cfg.microsoft_login.client_id,
      authority: S.cfg.microsoft_login.authority,
      redirectUri: `${location.origin}/app/`,
    },
    cache: { cacheLocation: "localStorage" },
  });
  await S.msal.initialize();
  const result = await S.msal.handleRedirectPromise().catch((e) => {
    say(`Sign-in failed: ${e.message}`, "error");
    return null;
  });
  S.account = result?.account || S.msal.getAllAccounts()[0] || null;
  if (S.account) S.msal.setActiveAccount(S.account);
  showAccount();
}

function showAccount() {
  $("signin-btn").hidden = !!S.account || IN_EXCEL || !S.msal;
  $("account-btn").hidden = !S.account;
  if (S.account) {
    $("account-btn").textContent = S.account.username.split("@")[0];
    $("account-btn").title = `Signed in as ${S.account.username} - click to sign out`;
  }
}

async function graphToken() {
  if (!S.account) {
    keepDraftForRedirect();
    await S.msal.loginRedirect({ scopes: GRAPH_SCOPES });
    return null;
  }
  try {
    const r = await S.msal.acquireTokenSilent({ scopes: GRAPH_SCOPES, account: S.account });
    return r.accessToken;
  } catch {
    keepDraftForRedirect();
    await S.msal.acquireTokenRedirect({ scopes: GRAPH_SCOPES, account: S.account });
    return null;
  }
}

function graphBase(tableName) {
  const path = S.cfg.excel.onedrive_workbook_path.split("/").map(encodeURIComponent).join("/");
  const base = `https://graph.microsoft.com/v1.0/me/drive/root:/${path}:/workbook`;
  return tableName ? `${base}/tables/${encodeURIComponent(tableName)}` : base;
}

async function graph(url, opts = {}) {
  const token = await graphToken();
  if (!token) throw new Error("Signing in…");
  const res = await fetch(url, {
    ...opts,
    headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json", ...(opts.headers || {}) },
  });
  const body = await res.json().catch(() => null);
  if (!res.ok) {
    const msg = body?.error?.message || `Microsoft error ${res.status}`;
    if (res.status === 404) throw new Error(`${msg} - check "onedrive_workbook_path" and "table_name" in config.json.`);
    throw new Error(msg);
  }
  return body;
}

async function graphLoadSheet() {
  let tableName = S.cfg.excel.table_name;
  let head;
  try {
    head = await graph(`${graphBase(tableName)}/headerRowRange?$select=values`);
  } catch (e) {
    const list = await graph(`${graphBase()}/tables?$select=name`).catch(() => null);
    if (list?.value?.length === 1) {
      tableName = list.value[0].name;
      head = await graph(`${graphBase(tableName)}/headerRowRange?$select=values`);
    } else throw e;
  }
  const body = await graph(`${graphBase(tableName)}/dataBodyRange?$select=values`);
  const headers = head.values[0].map(String);
  const rows = (body.values || []).filter((r) => r.some((v) => v !== "" && v !== null));
  return { tableName, headers, keys: mapHeaders(headers), rows };
}

async function graphAddRow(row) {
  await graph(`${graphBase(S.sheet.tableName)}/rows/add`, {
    method: "POST",
    body: JSON.stringify({ values: [row] }),
  });
}

/* ------------------------------------------------------ sheet (either way) */

async function loadSheet({ quiet = false } = {}) {
  try {
    if (IN_EXCEL) S.sheet = await excelLoadSheet();
    else if (S.account) S.sheet = await graphLoadSheet();
    else return;
    const missing = S.cfg.columns.filter((c) => !S.sheet.keys.includes(c.key)).map((c) => c.header);
    if (!quiet) {
      say(`Connected to table "${S.sheet.tableName}" (${S.sheet.rows.length} rows).` +
          (missing.length ? `\nColumns not found in Excel (will be skipped): ${missing.join(", ")}` : ""));
    }
  } catch (e) {
    if (e.message !== "Signing in…") say(e.message, "error");
  }
}

function buildRow(values) {
  return S.sheet.keys.map((key) => {
    if (key === "sno") return nextSerial();
    if (key === "date") return todayText();
    if (!key) return "";
    return values[key] ?? "";
  });
}

/* --------------------------------------------------------- preview card */

const LONG = new Set(["extent", "exact_area", "negative", "frontage", "previlling"]);

function showPreview(fields, warnings, sourceText, seconds) {
  const card = $("preview-tpl").content.firstElementChild.cloneNode(true);
  const ref = (n) => card.querySelector(`[data-ref="${n}"]`);
  ref("time").textContent = seconds ? `read in ${seconds}s` : "";
  ref("source").textContent = sourceText;

  const book = mediatorBook();
  const values = { ...fields };
  if (!values.status) values.status = S.cfg.default_status;
  if (!values.source) values.source = S.cfg.default_source;
  const w = { ...warnings };
  if (values.mediator_name && !values.number) {
    const known = book.get(values.mediator_name.trim().toLowerCase());
    if (known) { values.number = known; w.number = "Number taken from this mediator's earlier rows."; }
  }

  const form = ref("form");
  const listId = `med-${Date.now()}`;
  const dl = document.createElement("datalist");
  dl.id = listId;
  [...book.keys()].forEach((n) => {
    const o = document.createElement("option");
    o.value = n.replace(/\b\w/g, (c) => c.toUpperCase());
    dl.appendChild(o);
  });
  form.appendChild(dl);

  for (const col of S.cfg.columns) {
    if (col.fill === "auto") continue;
    const wrap = document.createElement("div");
    wrap.className = "field" + (w[col.key] ? " warn" : "");
    if (LONG.has(col.key)) wrap.style.gridColumn = "1 / -1";
    const label = document.createElement("label");
    label.innerHTML = `<span></span><span class="tag ${col.fill === "ai" ? "ai" : "you"}">${col.fill === "ai" ? "AI" : "You"}</span>`;
    label.firstChild.textContent = col.key === "status2" ? `${col.header} (2nd)` : col.header;
    let input;
    if (col.key === "status") {
      input = document.createElement("select");
      for (const opt of S.cfg.status_options) {
        const o = document.createElement("option");
        o.value = o.textContent = opt;
        input.appendChild(o);
      }
    } else if (LONG.has(col.key)) {
      input = document.createElement("textarea");
      input.rows = 1;
    } else {
      input = document.createElement("input");
      if (col.key === "number") input.inputMode = "tel";
      if (col.key === "mediator_name") input.setAttribute("list", listId);
    }
    input.name = col.key;
    input.value = values[col.key] ?? "";
    input.addEventListener("input", () => wrap.classList.remove("warn"));
    if (col.key === "mediator_name") {
      input.addEventListener("change", () => {
        const num = form.elements.number;
        const known = book.get(input.value.trim().toLowerCase());
        if (num && !num.value && known) num.value = known;
      });
    }
    wrap.append(label, input);
    if (w[col.key]) {
      const why = document.createElement("div");
      why.className = "why";
      why.textContent = w[col.key];
      wrap.appendChild(why);
    }
    form.appendChild(wrap);
  }

  const readValues = () => {
    const v = {};
    for (const el of form.elements) if (el.name) v[el.name] = el.value.trim();
    return v;
  };
  const dupBox = ref("dup");
  const checkDup = () => {
    const d = findDuplicate(readValues());
    dupBox.hidden = !d;
    dupBox.textContent = d ? `⚠ Possible duplicate: ${d}` : "";
  };
  checkDup();
  form.addEventListener("change", checkDup);

  ref("discard").onclick = () => card.remove();
  ref("save").onclick = async () => {
    const btn = ref("save");
    btn.disabled = true;
    btn.textContent = "Saving…";
    try {
      if (!IN_EXCEL && !S.msal) throw new Error("Phone saving is not set up yet (client_id missing in config.json).");
      if (!IN_EXCEL && !S.account) { await graphToken(); return; }
      await loadSheet({ quiet: true }); // fresh copy, so S.No and duplicates are correct
      if (!S.sheet) throw new Error("Not connected to Excel yet.");
      const v = readValues();
      const row = buildRow(v);
      if (IN_EXCEL) await excelAddRow(row); else await graphAddRow(row);
      S.sheet.rows.push(row);
      card.classList.add("saved");
      card.querySelectorAll("input,select,textarea").forEach((el) => (el.disabled = true));
      const sno = row[S.sheet.keys.indexOf("sno")];
      say(`✅ Saved to Excel${sno ? ` as S.No ${sno}` : ""}: ${v.exact_area || "property"}`, "success");
    } catch (e) {
      if (e.message !== "Signing in…") say(`Could not save: ${e.message}`, "error");
      btn.disabled = false;
      btn.textContent = "Save to Excel";
    }
  };

  chat.appendChild(card);
  card.scrollIntoView({ behavior: "smooth", block: "start" });
  return card;
}

/* Keep an unsaved preview across a Microsoft sign-in redirect (phone). */
function keepDraftForRedirect() {
  const card = [...chat.querySelectorAll(".preview:not(.saved)")].pop();
  if (!card) return;
  const form = card.querySelector("form");
  const v = {};
  for (const el of form.elements) if (el.name) v[el.name] = el.value;
  try {
    sessionStorage.setItem("draft", JSON.stringify({
      values: v, source: card.querySelector('[data-ref="source"]').textContent,
    }));
  } catch { /* storage blocked */ }
}

function restoreDraft() {
  let d = null;
  try { d = JSON.parse(sessionStorage.getItem("draft") || "null"); sessionStorage.removeItem("draft"); } catch { /* ignore */ }
  if (d) {
    say("Restored your unsaved property after sign-in. Tap Save again.");
    showPreview(d.values, {}, d.source);
  }
}

/* ------------------------------------------------------------ actions */

async function readMessage(text) {
  text = text.trim();
  if (!text) return;
  say(text, "user");
  const w = working("Reading the message with AI on your laptop…");
  try {
    const r = await api("/api/extract", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ text }),
    });
    w.done();
    showPreview(r.fields, r.warnings, text, r.seconds);
  } catch (e) {
    w.done();
    say(e.message, "error");
  }
}

async function transcribe(blob, filename) {
  const w = working("Converting voice to text…");
  try {
    const fd = new FormData();
    fd.append("file", blob, filename);
    const r = await api("/api/transcribe", { method: "POST", body: fd });
    w.done();
    if (!r.text) { say("Could not hear any words in this voice note.", "error"); return; }
    say(`🎤 Voice note (${r.language}) → text`, "bot");
    await readMessage(r.text);
  } catch (e) {
    w.done();
    say(e.message, "error");
  }
}

/* Microphone recording */
const REC = { rec: null, chunks: [], timer: null, cancelled: false };

async function getMic() {
  try {
    return await navigator.mediaDevices.getUserMedia({ audio: true });
  } catch (e) {
    // Excel desktop may need its own permission prompt first
    if (IN_EXCEL && Office.context?.devicePermission) {
      await Office.context.devicePermission.requestPermissionsAsync([Office.DevicePermissionType.microphone]);
      return navigator.mediaDevices.getUserMedia({ audio: true });
    }
    throw e;
  }
}

async function toggleRecording() {
  if (REC.rec) { REC.rec.stop(); return; }
  let stream;
  try {
    stream = await getMic();
  } catch (e) {
    say(`Microphone not available (${e.message}). Use 📎 to attach a voice note instead.`, "error");
    return;
  }
  const type = ["audio/webm;codecs=opus", "audio/webm", "audio/mp4"].find((t) => MediaRecorder.isTypeSupported?.(t)) || "";
  const rec = new MediaRecorder(stream, type ? { mimeType: type } : undefined);
  REC.rec = rec; REC.chunks = []; REC.cancelled = false;
  rec.ondataavailable = (e) => e.data.size && REC.chunks.push(e.data);
  rec.onstop = () => {
    stream.getTracks().forEach((t) => t.stop());
    clearInterval(REC.timer);
    $("rec-bar").hidden = true;
    $("mic-btn").classList.remove("recording");
    $("mic-btn").textContent = "🎤";
    REC.rec = null;
    if (REC.cancelled || !REC.chunks.length) return;
    const ext = (rec.mimeType || "").includes("mp4") ? "m4a" : "webm";
    transcribe(new Blob(REC.chunks, { type: rec.mimeType }), `voice.${ext}`);
  };
  rec.start();
  const start = Date.now();
  $("rec-bar").hidden = false;
  $("mic-btn").classList.add("recording");
  $("mic-btn").textContent = "⏹";
  REC.timer = setInterval(() => {
    const s = Math.floor((Date.now() - start) / 1000);
    $("rec-time").textContent = `${Math.floor(s / 60)}:${String(s % 60).padStart(2, "0")}`;
    if (s >= 300) rec.stop(); // 5 minute limit
  }, 500);
}

/* ------------------------------------------------------------ start */

async function checkHealth() {
  try {
    const h = await api("/api/health");
    const ok = h.ollama && h.model_installed;
    $("ai-dot").className = `dot ${ok ? "ok" : "bad"}`;
    $("ai-dot").title = ok ? `AI ready (${h.model})` : !h.ollama ? "Ollama is not running" : `Model ${h.model} not installed`;
    if (!h.ollama) say("Ollama is not running on the laptop. Open the Ollama app.", "error");
    else if (!h.model_installed) say(`AI model ${h.model} is not installed. On the laptop run: ollama pull ${h.model}`, "error");
  } catch {
    $("ai-dot").className = "dot bad";
    say("Cannot reach the laptop server. Is start.ps1 running (and Tailscale on, for phone)?", "error");
  }
}

async function handleShared() {
  const params = new URLSearchParams(location.search);
  const sid = params.get("shared");
  if (!sid) return;
  history.replaceState(null, "", location.pathname);
  const w = working("Opening shared WhatsApp item…");
  try {
    const r = await api(`/api/shared/${encodeURIComponent(sid)}`);
    w.done();
    if (r.text) await readMessage(r.text);
    else say("Nothing to read in the shared item.", "error");
  } catch (e) {
    w.done();
    say(e.message, "error");
  }
}

async function main() {
  $("where").textContent = IN_EXCEL ? "Excel add-in" : "Phone / browser";
  $("send-btn").onclick = () => { const t = $("msg").value; $("msg").value = ""; readMessage(t); };
  $("msg").addEventListener("keydown", (e) => {
    if (e.key === "Enter" && (e.ctrlKey || e.metaKey)) $("send-btn").click();
  });
  $("mic-btn").onclick = toggleRecording;
  $("rec-cancel").onclick = () => { REC.cancelled = true; REC.rec?.stop(); };
  $("file-input").onchange = (e) => {
    const f = e.target.files[0];
    e.target.value = "";
    if (f) { say(`📎 ${f.name}`, "user"); transcribe(f, f.name); }
  };
  $("signin-btn").onclick = () => S.msal?.loginRedirect({ scopes: GRAPH_SCOPES });
  $("account-btn").onclick = () => {
    if (confirm(`Sign out ${S.account.username}?`)) S.msal.logoutRedirect({ account: S.account });
  };

  try {
    S.cfg = await api("/api/config");
  } catch {
    say("Cannot reach the laptop server. Is start.ps1 running?", "error");
    return;
  }

  if (IN_EXCEL) {
    try { await loadOfficeJs(); } catch (e) { say(e.message, "error"); }
  } else {
    if ("serviceWorker" in navigator) navigator.serviceWorker.register("sw.js").catch(() => {});
    await initMsal();
  }

  say("Paste a mediator's message, record a voice note 🎤, or attach one 📎. " +
      "I'll fill in the columns. You check, then Save.");
  await checkHealth();
  await loadSheet();
  restoreDraft();
  await handleShared();
}

main();
