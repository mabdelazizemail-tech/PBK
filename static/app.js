/* مطابقة المشتريات والمبيعات — واجهة عربية تعمل محلياً */
"use strict";

const $ = (sel) => document.querySelector(sel);
const $$ = (sel) => [...document.querySelectorAll(sel)];

const TITLES = {
  dashboard: "لوحة المتابعة",
  matching: "المطابقة",
  purchases: "المشتريات",
  sales: "المبيعات",
  stockcard: "كارت الصنف",
  eta: "مزامنة ETA",
  settings: "الإعدادات",
};

/* ـــــــــــــــــــــ أدوات ـــــــــــــــــــــ */
function toast(msg, isError = false) {
  const el = document.createElement("div");
  el.className = "toast" + (isError ? " err" : "");
  el.textContent = msg;
  $("#toast-zone").appendChild(el);
  setTimeout(() => el.remove(), 3800);
}

async function api(path, opts = {}) {
  if (opts.json !== undefined) {
    opts.body = JSON.stringify(opts.json);
    opts.headers = { "Content-Type": "application/json", ...(opts.headers || {}) };
    delete opts.json;
  }
  const res = await fetch(path, opts);
  if (!res.ok) {
    let detail = `خطأ ${res.status}`;
    try { detail = (await res.json()).detail || detail; } catch { /* نص غير JSON */ }
    throw new Error(detail);
  }
  return res.json();
}

const nf2 = new Intl.NumberFormat("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
const nf0 = new Intl.NumberFormat("en-US", { maximumFractionDigits: 2 });
const money = (n) => nf2.format(n ?? 0);
const qty = (n) => nf0.format(n ?? 0);
const signedCell = (n, fmt = money) =>
  `<span class="num ${n < 0 ? "neg" : "pos"}">${fmt(n)}</span>`;
const esc = (s) => String(s ?? "").replace(/[&<>"']/g,
  (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
const srcPill = (s) => ({
  eta: '<span class="pill eta">ETA</span>',
  excel: '<span class="pill excel">Excel</span>',
}[s] || '<span class="pill manual">يدوي</span>');

/* ـــــــــــــــــــــ التنقل ـــــــــــــــــــــ */
function showTab(name) {
  $$(".tab").forEach((t) => (t.hidden = t.id !== `tab-${name}`));
  $$(".rail-btn").forEach((b) => b.classList.toggle("active", b.dataset.tab === name));
  $("#page-title").textContent = TITLES[name] || "";
  if (name === "dashboard") loadDashboard();
  if (name === "matching") loadMatching();
  if (name === "purchases") loadLines("purchase");
  if (name === "sales") loadLines("sale");
  if (name === "stockcard") loadStockCard();
  if (name === "eta") initEtaTab();
  if (name === "settings") loadSettings();
}
$$(".rail-btn").forEach((b) => b.addEventListener("click", () => showTab(b.dataset.tab)));
document.addEventListener("click", (e) => {
  const go = e.target.closest("[data-goto]");
  if (go) { e.preventDefault(); showTab(go.dataset.goto); }
});

/* ـــــــــــــــــــــ لوحة المتابعة ـــــــــــــــــــــ */
async function loadDashboard() {
  try {
    const d = await api("/api/dashboard");
    const setNum = (id, v, fmt = money) => {
      const el = $(id);
      el.textContent = fmt(v);
      el.classList.toggle("neg", v < 0);
    };
    $("#kpi-vat-position").textContent = money(d.vat_position) + " ج.م";
    $("#kpi-vat-hint").textContent = d.vat_position < 0
      ? "ضريبة المدخلات أعلى من المخرجات (رصيد دائن لدى المصلحة)"
      : "ضريبة مستحقة السداد عن الفترة";
    $("#kpi-gross-diff").textContent = money(d.gross_diff) + " ج.م";
    $("#kpi-net-margin").textContent = `الهامش الصافي (قبل الضريبة): ${money(d.net_margin)} ج.م`;

    $("#kpi-p-count").textContent = d.purchases.count;
    setNum("#kpi-p-net", d.purchases.net);
    setNum("#kpi-p-vat", d.purchases.vat);
    setNum("#kpi-p-total", d.purchases.total);
    $("#kpi-s-count").textContent = d.sales.count;
    setNum("#kpi-s-net", d.sales.net);
    setNum("#kpi-s-vat", d.sales.vat);
    setNum("#kpi-s-total", d.sales.total);

    $("#kpi-matches").textContent = d.matches_count;
    $("#kpi-unmatched-p").textContent = d.unmatched_purchases;
    $("#kpi-unmatched-s").textContent = d.unmatched_sales;
    $("#kpi-qty-mismatch").textContent = d.qty_mismatch_count;

    const pending = d.unmatched_purchases + d.unmatched_sales;
    const badge = $("#badge-matching");
    badge.hidden = pending === 0;
    badge.textContent = pending;
  } catch (e) { toast(e.message, true); }
}

/* ـــــــــــــــــــــ سطور المشتريات/المبيعات ـــــــــــــــــــــ */
const lineCache = { purchase: [], sale: [] };

async function loadLines(kind) {
  const q = kind === "purchase" ? $("#q-purchases").value : $("#q-sales").value;
  try {
    const rows = await api(`/api/lines?kind=${kind}&q=${encodeURIComponent(q || "")}`);
    lineCache[kind] = rows;
    renderLines(kind, rows);
    refreshDatalists();
  } catch (e) { toast(e.message, true); }
}

function renderLines(kind, rows) {
  const isP = kind === "purchase";
  const tbody = $(isP ? "#tbl-purchases tbody" : "#tbl-sales tbody");
  $(isP ? "#cnt-purchases" : "#cnt-sales").textContent = `(${rows.length})`;
  tbody.innerHTML = rows.map((r) => {
    const net = (r.qty || 0) * (r.unit_price || 0);
    const credit = r.doc_type === "c" || /دائن/.test(r.note || "");
    return `<tr>
      <td class="num">${esc(r.invoice_date || "—")}</td>
      <td>${esc(r.invoice_no)}</td>
      <td>${esc(r.party)}</td>
      <td>${esc(r.item)} ${credit ? '<span class="pill credit">اشعار دائن</span>' : ""}</td>
      <td class="num">${qty(r.qty)}</td>
      <td class="num">${money(r.unit_price)}</td>
      <td class="num">${money(net)}</td>
      <td class="num">${money(r.vat)}</td>
      <td class="num"><b>${money(net + (r.vat || 0))}</b></td>
      ${isP ? "" : `<td>${esc(r.internal_ref || "")}</td>`}
      <td>${srcPill(r.source)}</td>
      <td>${r.match_id ? '<span class="pill linked">مرتبط</span>' : ""}</td>
      <td class="actions">
        <button class="btn tiny ghost" data-edit="${kind}:${r.id}">تعديل</button>
        <button class="btn danger tiny" data-del="${kind}:${r.id}">حذف</button>
      </td>
    </tr>`;
  }).join("") || `<tr><td colspan="13" class="empty">لا توجد سطور — استورد ملف Excel أو زامن من ETA أو أضف يدوياً</td></tr>`;
}

let searchTimer;
["#q-purchases", "#q-sales"].forEach((sel, i) =>
  $(sel).addEventListener("input", () => {
    clearTimeout(searchTimer);
    searchTimer = setTimeout(() => loadLines(i === 0 ? "purchase" : "sale"), 250);
  }));

function refreshDatalists() {
  const parties = new Set(), items = new Set();
  [...lineCache.purchase, ...lineCache.sale].forEach((r) => {
    if (r.party) parties.add(r.party);
    if (r.item) items.add(r.item);
  });
  $("#parties-list").innerHTML = [...parties].map((p) => `<option value="${esc(p)}">`).join("");
  $("#items-list").innerHTML = [...items].map((p) => `<option value="${esc(p)}">`).join("");
}

/* مودال إضافة/تعديل سطر */
const lineModal = $("#line-modal");
let editing = null; // {kind, id} | {kind}

function openLineModal(kind, row = null) {
  editing = { kind, id: row?.id };
  $("#line-modal-title").textContent =
    (row ? "تعديل سطر " : "إضافة سطر ") + (kind === "purchase" ? "مشتريات" : "مبيعات");
  $("#ln-party-label").firstChild.textContent = kind === "purchase" ? "المورد" : "المشتري";
  $("#ln-ref-wrap").hidden = $("#ln-note-wrap").hidden = kind === "purchase";
  $("#ln-date").value = row?.invoice_date || "";
  $("#ln-no").value = row?.invoice_no || "";
  $("#ln-party").value = row?.party || "";
  $("#ln-item").value = row?.item || "";
  $("#ln-qty").value = row?.qty ?? "";
  $("#ln-price").value = row?.unit_price ?? "";
  $("#ln-vat").value = row ? row.vat : "";
  $("#ln-ref").value = row?.internal_ref || "";
  $("#ln-note").value = row?.note || "";
  lineModal.showModal();
}

document.addEventListener("click", (e) => {
  const add = e.target.closest("[data-add]");
  if (add) openLineModal(add.dataset.add);
  const edit = e.target.closest("[data-edit]");
  if (edit) {
    const [kind, id] = edit.dataset.edit.split(":");
    openLineModal(kind, lineCache[kind].find((r) => r.id === +id));
  }
  const del = e.target.closest("[data-del]");
  if (del) {
    const [kind, id] = del.dataset.del.split(":");
    if (confirm("هل تريد حذف هذا السطر؟ سيُحذف أي ربط مرتبط به.")) {
      api(`/api/lines/${kind}/${id}`, { method: "DELETE" })
        .then(() => { toast("تم الحذف"); loadLines(kind); })
        .catch((err) => toast(err.message, true));
    }
  }
});

$("#form-line").addEventListener("submit", (e) => {
  if (e.submitter?.value !== "save") return;
  e.preventDefault();
  const body = {
    invoice_date: $("#ln-date").value || null,
    invoice_no: $("#ln-no").value.trim(),
    party: $("#ln-party").value.trim(),
    item: $("#ln-item").value.trim(),
    qty: parseFloat($("#ln-qty").value),
    unit_price: parseFloat($("#ln-price").value) || 0,
  };
  const vatRaw = $("#ln-vat").value;
  if (vatRaw !== "") body.vat = parseFloat(vatRaw);
  if (editing.kind === "sale") {
    body.internal_ref = $("#ln-ref").value.trim();
    body.note = $("#ln-note").value.trim();
  }
  const req = editing.id
    ? api(`/api/lines/${editing.kind}/${editing.id}`, { method: "PUT", json: body })
    : api("/api/lines", { method: "POST", json: { kind: editing.kind, ...body } });
  req.then(() => {
    lineModal.close();
    toast("تم الحفظ ✓");
    loadLines(editing.kind);
  }).catch((err) => toast(err.message, true));
});

/* ـــــــــــــــــــــ المطابقة ـــــــــــــــــــــ */
let suggestions = [];
let selP = null, selS = null;

async function loadMatching() {
  try {
    const [up, us, matches] = await Promise.all([
      api("/api/lines?kind=purchase&matched=false"),
      api("/api/lines?kind=sale&matched=false"),
      api("/api/matches"),
    ]);
    renderUnmatched("#tbl-unmatched-p", up, "p");
    renderUnmatched("#tbl-unmatched-s", us, "s");
    $("#cnt-unmatched-p").textContent = `(${up.length})`;
    $("#cnt-unmatched-s").textContent = `(${us.length})`;
    renderMatches(matches);
  } catch (e) { toast(e.message, true); }
}

function renderUnmatched(sel, rows, side) {
  const tbody = $(sel + " tbody");
  tbody.innerHTML = rows.map((r) => `
    <tr data-select="${side}:${r.id}">
      <td><input type="radio" name="sel-${side}" ${((side === "p" ? selP : selS) === r.id) ? "checked" : ""}></td>
      <td class="num">${esc(r.invoice_date || "—")}</td>
      <td>${esc(r.invoice_no)}</td>
      <td>${esc(r.party)}</td>
      <td>${esc(r.item)}</td>
      <td class="num">${qty(r.qty)}</td>
    </tr>`).join("") || `<tr><td colspan="6" class="empty">لا يوجد</td></tr>`;
}

document.addEventListener("click", (e) => {
  const tr = e.target.closest("tr[data-select]");
  if (!tr) return;
  const [side, id] = tr.dataset.select.split(":");
  if (side === "p") selP = +id; else selS = +id;
  tr.closest("tbody").querySelectorAll("tr").forEach((row) =>
    row.classList.toggle("selected", row === tr));
  tr.querySelector("input[type=radio]").checked = true;
});

$("#btn-manual-link").addEventListener("click", () => {
  if (!selP || !selS) return toast("اختر سطراً من المشتريات وسطراً من المبيعات أولاً", true);
  api("/api/matches", { method: "POST", json: { purchase_id: selP, sale_id: selS } })
    .then(() => { toast("تم الربط ✓"); selP = selS = null; loadMatching(); loadDashboard(); })
    .catch((err) => toast(err.message, true));
});

function renderMatches(rows) {
  $("#cnt-matches").textContent = `(${rows.length})`;
  $("#tbl-matches tbody").innerHTML = rows.map((m) => `
    <tr>
      <td class="num">${esc(m.p_date || "—")}</td>
      <td>${esc(m.p_no)}</td>
      <td>${esc(m.supplier)}</td>
      <td>${esc(m.p_item)}</td>
      <td class="num">${qty(m.p_qty)}</td>
      <td class="num">${esc(m.s_date || "—")}</td>
      <td>${esc(m.s_no)}</td>
      <td>${esc(m.customer)}</td>
      <td class="num">${qty(m.s_qty)}</td>
      <td>${m.qty_diff ? signedCell(m.qty_diff, qty) : '<span class="num">0</span>'}</td>
      <td class="actions"><button class="btn danger tiny" data-unlink="${m.id}">فك الربط</button></td>
    </tr>`).join("") || `<tr><td colspan="11" class="empty">لا توجد مطابقات بعد</td></tr>`;
}

document.addEventListener("click", (e) => {
  const u = e.target.closest("[data-unlink]");
  if (u) {
    api(`/api/matches/${u.dataset.unlink}`, { method: "DELETE" })
      .then(() => { toast("تم فك الربط"); loadMatching(); loadDashboard(); })
      .catch((err) => toast(err.message, true));
  }
});

$("#btn-auto-match").addEventListener("click", async () => {
  const btn = $("#btn-auto-match");
  btn.disabled = true;
  try {
    suggestions = await api("/api/matches/auto", { method: "POST" });
    renderSuggestions();
    if (!suggestions.length) toast("لا توجد مقترحات — كل السطور المتشابهة مرتبطة بالفعل");
  } catch (e) { toast(e.message, true); }
  btn.disabled = false;
});

function renderSuggestions() {
  const wrap = $("#suggestions-wrap");
  $("#suggestions-empty").hidden = suggestions.length > 0;
  wrap.hidden = suggestions.length === 0;
  $("#btn-accept-all").hidden = suggestions.length === 0;
  $("#tbl-suggestions tbody").innerHTML = suggestions.map((m, i) => `
    <tr>
      <td><span class="pill ${m.score >= 85 ? "score-hi" : "score-md"}">${m.score}%</span></td>
      <td>${esc(m.purchase.invoice_no)} · ${esc(m.purchase.party)}<br><small class="num">${esc(m.purchase.invoice_date || "")}</small></td>
      <td>${esc(m.purchase.item)}</td>
      <td class="num">${qty(m.purchase.qty)}</td>
      <td>${esc(m.sale.invoice_no)} · ${esc(m.sale.party)}<br><small class="num">${esc(m.sale.invoice_date || "")}</small></td>
      <td class="num">${qty(m.sale.qty)}</td>
      <td>${m.qty_diff ? signedCell(m.qty_diff, qty) : "0"}</td>
      <td class="actions">
        <button class="btn tiny primary" data-accept="${i}">اعتماد</button>
        <button class="btn tiny ghost" data-dismiss="${i}">تجاهل</button>
      </td>
    </tr>`).join("");
}

document.addEventListener("click", (e) => {
  const acc = e.target.closest("[data-accept]");
  if (acc) acceptPairs([suggestions[+acc.dataset.accept]]);
  const dis = e.target.closest("[data-dismiss]");
  if (dis) { suggestions.splice(+dis.dataset.dismiss, 1); renderSuggestions(); }
});

$("#btn-accept-all").addEventListener("click", () => acceptPairs(suggestions));

function acceptPairs(list) {
  const pairs = list.map((m) => ({ purchase_id: m.purchase_id, sale_id: m.sale_id }));
  api("/api/matches/accept", { method: "POST", json: { pairs } })
    .then((r) => {
      toast(`تم اعتماد ${r.created} مطابقة ✓`);
      suggestions = suggestions.filter((m) => !list.includes(m));
      renderSuggestions();
      loadMatching();
      loadDashboard();
    })
    .catch((err) => toast(err.message, true));
}

/* ـــــــــــــــــــــ الإعدادات ـــــــــــــــــــــ */
async function loadSettings() {
  try {
    const s = await api("/api/settings");
    $("#set-env").value = s.eta_env;
    $("#set-cid").value = s.eta_client_id;
    $("#set-secret").placeholder = s.has_secret ? "•••••••• (محفوظ)" : "";
    $("#set-vat").value = s.vat_rate;
  } catch (e) { toast(e.message, true); }
}

$("#form-settings").addEventListener("submit", (e) => {
  e.preventDefault();
  api("/api/settings", {
    method: "PUT",
    json: {
      eta_env: $("#set-env").value,
      eta_client_id: $("#set-cid").value.trim(),
      eta_client_secret: $("#set-secret").value,
      vat_rate: parseFloat($("#set-vat").value) || 0.14,
    },
  }).then(() => { toast("تم حفظ الإعدادات ✓"); $("#set-secret").value = ""; loadSettings(); })
    .catch((err) => toast(err.message, true));
});

$("#btn-test-eta").addEventListener("click", async () => {
  const btn = $("#btn-test-eta");
  btn.disabled = true;
  try {
    const r = await api("/api/eta/test", { method: "POST" });
    toast(r.message);
  } catch (e) { toast(e.message, true); }
  btn.disabled = false;
});

/* ـــــــــــــــــــــ مزامنة ETA ـــــــــــــــــــــ */
let pollTimer = null;

async function initEtaTab() {
  const s = await api("/api/settings").catch(() => null);
  $("#eta-cred-warning").hidden = !!(s && s.eta_client_id && s.has_secret);
  if (!$("#sync-from").value) {
    const today = new Date();
    const monthAgo = new Date(today); monthAgo.setDate(today.getDate() - 30);
    $("#sync-from").value = monthAgo.toISOString().slice(0, 10);
    $("#sync-to").value = today.toISOString().slice(0, 10);
  }
  pollSyncStatus();
}

$("#form-sync").addEventListener("submit", async (e) => {
  e.preventDefault();
  const directions = [];
  if ($("#sync-received").checked) directions.push("Received");
  if ($("#sync-sent").checked) directions.push("Sent");
  if (!directions.length) return toast("اختر اتجاهاً واحداً على الأقل", true);
  try {
    const r = await api("/api/eta/sync", {
      method: "POST",
      json: {
        date_from: $("#sync-from").value,
        date_to: $("#sync-to").value,
        directions,
        refresh: $("#sync-refresh").checked,
      },
    });
    toast(r.total ? `بدأت المزامنة… (${r.total} مستند)` : "لا توجد مستندات جديدة");
    pollSyncStatus(true);
  } catch (err) { toast(err.message, true); }
});

function renderSync(st, force = false) {
  const box = $("#sync-progress");
  if (st.started_at || force) {
    box.hidden = false;
    $("#sync-docs").textContent = `${st.stats.documents} مستند`;
    $("#sync-lines").textContent = `${st.stats.lines} سطر`;
    const state = $("#sync-state");
    state.classList.remove("done", "err");
    if (st.running) state.textContent = "جارٍ التنفيذ…";
    else if (st.error) { state.textContent = "فشلت المزامنة"; state.classList.add("err"); }
    else if (st.finished_at) { state.textContent = "اكتملت ✓"; state.classList.add("done"); }
    const log = $("#sync-log");
    log.textContent = (st.log || []).join("\n");
    log.scrollTop = log.scrollHeight;
  }
  $("#btn-sync").disabled = !!st.running;
}

// Read-only status check (tab open / after start). Resumes the drive loop if a
// job is already running — e.g. after reloading the page mid-sync.
async function pollSyncStatus(force = false) {
  clearTimeout(pollTimer);
  try {
    const st = await api("/api/eta/sync/status");
    renderSync(st, force);
    if (st.running) driveSync();
  } catch { /* الخادم غير متاح مؤقتاً */ }
}

// Advances the sync one chunk per call until it finishes. Each /step does a
// bounded, ETA-throttled batch server-side, keeping every request short
// (serverless-friendly) while state persists in the database between calls.
function driveSync() {
  clearTimeout(pollTimer);
  pollTimer = setTimeout(async () => {
    try {
      const st = await api("/api/eta/sync/step", { method: "POST" });
      renderSync(st, true);
      if (st.running) driveSync();
      else if (st.finished_at && !st.error) loadDashboard();
    } catch { /* transient — stop; reopening the tab will resume */ }
  }, 250);
}

/* ـــــــــــــــــــــ استيراد / تصدير ـــــــــــــــــــــ */
$("#btn-export").addEventListener("click", () => {
  window.location.href = "/api/export/excel";
  toast("جارٍ تجهيز ملف Excel…");
});

const importModal = $("#import-modal");
$("#btn-import").addEventListener("click", () => {
  $("#import-result").hidden = true;
  $("#import-file").value = "";
  $("#import-file-name").textContent = "اختر الملف أو أسقطه هنا…";
  importModal.showModal();
});
$("#import-file").addEventListener("change", (e) => {
  $("#import-file-name").textContent = e.target.files[0]?.name || "اختر الملف…";
});

$("#form-import").addEventListener("submit", async (e) => {
  if (e.submitter?.value !== "go") return;
  e.preventDefault();
  const file = $("#import-file").files[0];
  if (!file) return toast("اختر ملفاً أولاً", true);
  const fd = new FormData();
  fd.append("file", file);
  fd.append("mode", document.querySelector("input[name=import-mode]:checked").value);
  const btn = $("#btn-import-go");
  btn.disabled = true;
  try {
    const r = await api("/api/import/excel", { method: "POST", body: fd });
    const res = $("#import-result");
    res.hidden = false;
    res.innerHTML = `تم استيراد <b>${r.purchases}</b> سطر مشتريات و<b>${r.sales}</b> سطر مبيعات
      و<b>${r.matches}</b> مطابقة.` +
      (r.warnings.length
        ? `<br>تحذيرات (${r.warnings.length}):<br>` + r.warnings.map(esc).join("<br>")
        : "");
    toast("اكتمل الاستيراد ✓");
    loadDashboard();
  } catch (err) { toast(err.message, true); }
  btn.disabled = false;
});

/* ـــــــــــــــــــــ كارت الصنف ـــــــــــــــــــــ */
let stockItems = [];
let stockSelKey = null;

async function loadStockCard() {
  try {
    stockItems = await api("/api/stock-card");
    renderStockItems();
  } catch (e) { toast(e.message, true); }
}

function renderStockItems() {
  const q = ($("#q-stock").value || "").trim();
  const rows = q ? stockItems.filter((it) => (it.item_label || "").includes(q)) : stockItems;
  $("#cnt-stock-items").textContent = `(${rows.length})`;
  $("#tbl-stock-items tbody").innerHTML = rows.map((it) => `
    <tr data-stock="${esc(it.item_key)}" class="${it.item_key === stockSelKey ? "selected" : ""}">
      <td>${esc(it.item_label)}</td>
      <td class="num">${qty(it.total_in)}</td>
      <td class="num">${qty(it.total_out)}</td>
      <td class="num"><b>${qty(it.balance)}</b></td>
    </tr>`).join("") ||
    `<tr><td colspan="4" class="empty">لا توجد أصناف — استورد بيانات أو زامن من ETA</td></tr>`;
}

$("#q-stock").addEventListener("input", () => renderStockItems());

document.addEventListener("click", (e) => {
  const tr = e.target.closest("tr[data-stock]");
  if (!tr) return;
  stockSelKey = tr.dataset.stock;
  renderStockItems();
  loadStockCardDetail(stockSelKey);
});

async function loadStockCardDetail(key) {
  try {
    const card = await api(`/api/stock-card/item?key=${encodeURIComponent(key)}`);
    $("#stock-card-title").textContent = `كارت صنف: ${card.item_label}`;
    $("#btn-stock-export-one").hidden = false;
    $("#tbl-stock-card tbody").innerHTML = card.rows.map((r) => `
      <tr>
        <td class="num">${esc(r.date || "—")}</td>
        <td>${esc(r.label)} ${r.doc_type === "c" ? '<span class="pill credit">دائن</span>' : ""}</td>
        <td>${esc(r.invoice_no)}</td>
        <td>${esc(r.party)}</td>
        <td class="num">${r.qty_in ? qty(r.qty_in) : ""}</td>
        <td class="num">${r.qty_out ? qty(r.qty_out) : ""}</td>
        <td class="num"><b>${qty(r.balance)}</b></td>
      </tr>`).join("") || `<tr><td colspan="7" class="empty">لا توجد حركات لهذا الصنف</td></tr>`;
    const tot = $("#stock-totals");
    tot.hidden = false;
    tot.innerHTML = `إجمالي الوارد: <b>${qty(card.total_in)}</b> · ` +
      `إجمالي المنصرف: <b>${qty(card.total_out)}</b> · ` +
      `الرصيد الحالي: <b>${qty(card.balance)}</b>`;
  } catch (e) { toast(e.message, true); }
}

$("#btn-stock-export-one").addEventListener("click", () => {
  if (!stockSelKey) return;
  window.location.href = `/api/stock-card/export?key=${encodeURIComponent(stockSelKey)}`;
  toast("جارٍ تجهيز كارت الصنف…");
});
$("#btn-stock-export-all").addEventListener("click", () => {
  window.location.href = "/api/stock-card/export";
  toast("جارٍ تجهيز ملف الأصناف…");
});

/* ـــــــــــــــــــــ بدء التشغيل ـــــــــــــــــــــ */
loadDashboard();
