import { openStore } from "./store.js";
import { compute, periodIndexFor, todayISO } from "./calc.js";
import { readWorkbook } from "./importer.js";

const SHEETJS = "https://cdn.sheetjs.com/xlsx-0.20.3/package/xlsx.mjs";
const $app = document.getElementById("app");
let store, user = null, authReady = false;
let data = { model: null, state: { overrides: {}, balances: {}, paydays: {} }, entries: [] };
let view = "today", pIdx = null, sheet = null;

// ---------------------------------------------------------------- helpers
const esc = (s) => String(s ?? "").replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
const fmt0 = new Intl.NumberFormat("en-GB", { maximumFractionDigits: 0 });
const fmt2 = new Intl.NumberFormat("en-GB", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
const gbp = (x, dp = 0) => { const v = Math.abs(x) < 0.005 ? 0 : x; return (v < 0 ? "-£" : "£") + (dp ? fmt2 : fmt0).format(Math.abs(v)); };
const pct = (x) => Math.round(x * 100) + "%";
const dlong = (iso) => new Date(iso + "T12:00:00").toLocaleDateString("en-GB", { weekday: "short", day: "numeric", month: "short" });
const dshort = (iso) => new Date(iso + "T12:00:00").toLocaleDateString("en-GB", { day: "numeric", month: "short" });
const uid = () => (crypto.randomUUID ? crypto.randomUUID() : Date.now().toString(36) + Math.random().toString(36).slice(2));
const pillCls = (s) => ({ "over budget": "over", "unplanned": "unplanned", "to come": "tocome", "done": "done", "received": "received", "on track": "ontrack" }[s] || "");
function toast(msg) { const t = document.createElement("div"); t.className = "toast"; t.textContent = msg; document.body.appendChild(t); setTimeout(() => t.remove(), 1800); }

const today = () => todayISO();
const curIdx = () => periodIndexFor(data.model, today());
const calc = () => compute(data.model, { ...data.state, entries: data.entries }, today(), pIdx ?? curIdx());

// ---------------------------------------------------------------- boot
(async () => {
  store = await openStore((d) => { data = d; render(); }, (u) => { user = u; authReady = true; render(); });
  if ("serviceWorker" in navigator) navigator.serviceWorker.register("sw.js").catch(() => {});
})();

// ---------------------------------------------------------------- render
function render() {
  if (!authReady || !store) { $app.innerHTML = `<div class="empty"><div class="muted">Loading…</div></div>`; return; }
  if (store && store.mode === "firebase" && !user) return renderLogin();
  if (!data.model) return renderEmpty();
  const c = calc();
  const isCur = c.pi === curIdx();
  $app.innerHTML = `
    <header>
      <h1>${{ today: "Today", lines: "Line by line", balances: "Balances", journey: "Journey" }[view]}
        <small>${dshort(c.period.start)} to ${dshort(c.period.end)}${isCur ? "" : " · planning view"}</small></h1>
      <div class="pchip"><button data-act="prev">‹</button><span>${esc(c.period.key)}</span><button data-act="next">›</button></div>
      <button class="iconbtn" data-act="settings" aria-label="Settings">⚙︎</button>
    </header>
    <main>${({ today: vToday, lines: vLines, balances: vBalances, journey: vJourney })[view](c)}</main>
    <nav>
      ${navBtn("today", "⌂", "Today")}${navBtn("lines", "☰", "Lines")}
      <button class="add" data-act="add" aria-label="Log money">+</button>
      ${navBtn("balances", "£", "Balances")}${navBtn("journey", "★", "Journey")}
    </nav>
    ${sheet ? `<div class="sheet-bg"><div class="sheet" data-stop>${sheet()}</div></div>` : ""}`;
  bind();
}
const navBtn = (v, i, l) => `<button class="${view === v ? "on" : ""}" data-view="${v}"><i>${i}</i>${l}</button>`;

function renderLogin() {
  $app.innerHTML = `<div class="login">
    <div class="label">Money</div><h2 style="margin:6px 0 22px">Sign in</h2>
    <div class="field"><label>Email</label><input id="em" type="email" autocomplete="username"></div>
    <div class="field"><label>Password</label><input id="pw" type="password" autocomplete="current-password"></div>
    <button class="btn" id="go">Sign in</button><p class="note" id="err"></p></div>`;
  document.getElementById("go").onclick = async () => {
    try { await store.signIn(document.getElementById("em").value.trim(), document.getElementById("pw").value); }
    catch (e) { document.getElementById("err").textContent = "Could not sign in: " + (e.code || e.message); }
  };
}

function renderEmpty() {
  $app.innerHTML = `<header><h1>MONEY<small>${store.mode === "local" ? "Test mode · data stays on this device" : esc(user && user.email)}</small></h1></header>
    <main><div class="card empty"><div class="big">📊</div><h2>Import your cash flow workbook</h2>
    <p class="note">Pick the latest saved AG Cash Flow file. The app reads the plan, this month's position and your debt trajectory from it.</p>
    <label class="btn" style="display:block;text-align:center">Choose .xlsx<input type="file" id="file" accept=".xlsx" class="hidden"></label>
    <p class="note" id="imp"></p></div></main>`;
  document.getElementById("file").onchange = (e) => doImport(e.target.files[0]);
}

// ---------------------------------------------------------------- views
const STATUS = { "over budget": "over plan", "unplanned": "extra", "to come": "to come", "done": "done", "received": "received", "on track": "on track" };

function bar(label, val, fill, right, marker = null) {
  const w = Math.max(0, Math.min(1, val)) * 100;
  return `<div class="bar"><div class="top"><span>${label}</span><b>${right ?? pct(val)}</b></div>
    <div class="track"><div class="fill" style="width:${w}%;background:${fill}"></div>${marker !== null ? `<div class="marker" style="left:${Math.min(99, marker * 100)}%"></div>` : ""}</div></div>`;
}

// ---- the freedom journey: debt free -> safety net -> UK home -> investing -> Zakopane
function journey() {
  const m = data.model, t = m.trajectory || { labels: [], all: [], dmp: [], deposit: [] };
  const ci = curIdx(), key = m.periods[ci].key;
  const ms = m.milestones || [];
  const find = (re) => ms.find((x) => re.test(x.name));
  const monthsTo = (lab) => { const i = m.periods.findIndex((p) => p.key === lab); return i < 0 ? null : i - ci; };
  const ni = t.labels.indexOf(key);
  const owed = ni >= 0 ? (t.all[ni] ?? t.dmp[ni] ?? 0) : 0;
  const fa = t.all.findIndex((v) => v !== null);
  const dmp0 = t.dmp.find((v) => v !== null) ?? 0;
  const cleared = fa >= 0 ? Math.max(0, (dmp0 - (t.dmp[fa] ?? dmp0)) + (t.all[fa] - owed)) : 0;
  const pctDone = cleared + owed > 0 ? cleared / (cleared + owed) : 0;
  const deposit = ni >= 0 ? t.deposit[ni] || 0 : 0;
  const df = find(/debt free/i), ef = find(/emergency/i), mr = find(/mortgage|home/i);
  const subs = ms.filter((x) => /cleared|settled/i.test(x.name));
  const mk = (id, icon, short, title, when, text, extra = {}) => ({ id, icon, short, title, when, text, months: when ? monthsTo(when) : null, ...extra });
  const stages = [
    mk("debt", "🔓", "Debt free", "Debt free", df && df.forecast, `Every debt gone: the DMP, Zopa and both Polish loans. ${gbp(cleared)} cleared so far.`, { frac: pctDone }),
    mk("safety", "🛟", "Safety net", "Emergency fund full", ef && ef.forecast, "A cushion for the unexpected, so nothing knocks the plan off course again."),
    mk("home", "🏡", "Home", "Your UK home", mr && mr.forecast, `Deposit ready with Carley for a home of your own.${deposit > 0 ? " " + gbp(deposit) + " saved so far." : ""}`),
    mk("invest", "📈", "Invest", "Investing for freedom", null, "Low-cost index funds every month through a Stocks & Shares ISA, letting compounding do the heavy lifting."),
    mk("dream", "🏔️", "Zakopane", "Zakopane with Greg", null, "The dream: a holiday let in the mountains, run as a family business. Every step above builds towards it.", { dream: true }),
  ];
  let nowIdx = stages.findIndex((s) => s.months === null || s.months >= 0);
  if (nowIdx < 0) nowIdx = stages.length - 1;
  stages.forEach((s, i) => { s.state = i < nowIdx ? "done" : i === nowIdx ? "now" : "next"; });
  return { stages, nowIdx, cleared, owed, pctDone, df, subs, monthsTo, deposit };
}

function stepper(J) {
  const frac = (J.nowIdx + (J.stages[J.nowIdx].frac || 0)) / (J.stages.length - 1);
  return `<div class="steps"><div class="prog" style="width:calc((100% - 32px) * ${Math.min(1, frac).toFixed(3)})"></div>
    ${J.stages.map((s) => `<div class="step ${s.state}"><div class="dot">${s.state === "done" ? "✓" : s.icon}</div><b>${s.short}</b><span>${esc(s.when || (s.dream ? "the dream" : "next"))}</span></div>`).join("")}</div>`;
}

function greeting() {
  const h = new Date().getHours();
  return h < 12 ? "Good morning" : h < 18 ? "Good afternoon" : "Good evening";
}

function vToday(c) {
  const m = data.model, J = journey();
  const df = J.df, months = df ? J.monthsTo(df.forecast) : null;
  const ahead = c.paceGap < -0.5;
  const watch = c.rows.filter((r) => r.vs < -0.5).sort((a, b) => a.vs - b.vs).slice(0, 5);
  const recent = [...(data.entries || [])].sort((a, b) => (b.ts > a.ts ? 1 : -1)).slice(0, 6);
  const lineName = (id) => id === "__transfer" ? "Sent to mBank" : (m.lines.find((l) => l.id === id) || {}).name || id;
  const age = Math.round((Date.now() - Date.parse(m.importedAt)) / 86400000);
  const gapDay = c.daysLeft > 0 ? -c.spare / c.daysLeft : -c.spare;
  const spareTile = c.spare >= 0
    ? `<div class="label">Spare at payday</div><div class="v green">${gbp(c.spare)}</div><div class="s">Heading straight to your next goal</div>`
    : `<div class="label">Gap to close</div><div class="v orange">${gbp(-c.spare)}</div><div class="s">Just ${gbp(gapDay)} a day less until payday closes it</div>`;
  const sendTile = c.sendNow <= 0.5
    ? `<div class="label">Poland</div><div class="v green">Covered</div><div class="s">${c.sent > 0 ? gbp(c.sent) + " sent already, nice" : "Nothing to send right now"}</div>`
    : `<div class="label">Send to mBank</div><div class="v blue">${gbp(c.sendNow)}</div><div class="s">${c.sent > 0 ? gbp(c.sent) + " sent already, nice" : "Keeps Poland covered to payday"}</div>`;
  const paceTile = ahead
    ? `<div class="label">Spending pace</div><div class="v orange">${gbp(-c.paceGap)} ahead</div><div class="s">${c.spendLeft > 0 ? `Keep to ${gbp(c.perDay)} a day and you land on budget` : "Pause the extras and payday resets it"}</div>`
    : `<div class="label">Spending pace</div><div class="v green">${gbp(c.paceGap)} under</div><div class="s">Ahead of the game, brilliant</div>`;
  return `
    <div class="greet"><b>${greeting()}, Anna 👋</b><p>${months !== null && months > 0 ? `${months} months to debt free. You've cleared ${gbp(J.cleared)} already.` : "Debt free. Look how far you've come."}</p></div>
    <div class="card hero">
      <div class="label">Left to spend until payday</div>
      <div class="big">${gbp(c.spendLeft)}</div>
      <div class="sub">${c.spendLeft > 0 ? `<b>${gbp(c.perDay)} a day</b> for the next ${c.daysLeft} day${c.daysLeft === 1 ? "" : "s"}` : "Budget used for this month, payday brings a fresh start"}</div>
      <div class="chip">💸 Payday ${dlong(c.payday)}</div>
    </div>
    <div class="grid2">
      <div class="tile"><div class="ic ${c.spare >= 0 ? "g" : "o"}">🎯</div>${spareTile}</div>
      <div class="tile"><div class="ic b">🇵🇱</div>${sendTile}</div>
      <div class="tile"><div class="ic ${ahead ? "o" : "g"}">⚡</div>${paceTile}</div>
      <div class="tile" data-view="journey"><div class="ic v">🔓</div><div class="label">Debt free</div><div class="v blue">${esc(df ? df.forecast : "-")}</div>
        <div class="s">${months !== null && months > 0 ? months + " months to go" : "Done"}${df && /on track|ahead/i.test(df.status) ? " · on track ✓" : ""}</div></div>
    </div>
    <div class="card" data-view="journey"><div class="label">Your freedom journey</div>${stepper(J)}</div>
    <div class="card"><div class="label">This month</div>
      ${bar("Month gone", c.elapsed, "#c3cad6")}
      ${bar("Everyday budget used", c.budgetUsed, c.budgetUsed > c.elapsed + 0.02 ? "var(--g-orange)" : "var(--g-blue)", null, c.elapsed)}
      ${bar("Bills paid", c.billsPaid, "linear-gradient(140deg,#5b7fd6,#8aa6e6)")}
      ${bar("Income in", c.incomeIn, "var(--g-green)")}
    </div>
    ${watch.length ? `<div class="card list"><div class="label">Worth a look</div>${watch.map((r) => `
      <div class="row" data-line="${r.id}"><div class="nm"><b>${esc(r.name)}</b><span>plan ${gbp(r.plan)} · heading for ${gbp(r.likely)}</span></div>
      <div class="amt orange">${gbp(-r.vs)}<span class="note">above plan</span></div></div>`).join("")}</div>` : ""}
    <div class="card list"><div class="label">Recent</div>
      ${recent.length ? recent.map((e) => `<div class="row"><div class="nm"><b>${esc(lineName(e.lineId))}</b><span>${dshort(e.date)} · ${esc(e.account)}${e.note ? " · " + esc(e.note) : ""}</span></div>
        <div class="amt ${e.amount > 0 ? "green" : ""}">${gbp(e.amount, 2)}${e.pln ? `<span class="note">${fmt2.format(e.pln)} zł</span>` : ""}</div>
        <button class="del" data-del="${e.id}" aria-label="Delete">×</button></div>`).join("") : `<p class="note">Nothing logged yet. Tap + after you spend and watch the numbers keep up.</p>`}
    </div>
    <p class="note" style="text-align:center">Plan from Excel, imported ${age === 0 ? "today" : age + " day" + (age === 1 ? "" : "s") + " ago"} (balances as at ${dshort(m.snapshot.asAt)})</p>`;
}

function vLines(c) {
  const m = data.model;
  return m.sections.map((s) => {
    const rows = c.rows.filter((r) => r.section === s.name && (r.plan !== 0 || r.soFar !== 0 || r.override !== ""));
    if (!rows.length) return "";
    const tot = rows.reduce((a, r) => a + r.toCome, 0);
    return `<div class="sect label">${esc(s.name)} <span style="float:right">${gbp(tot)} to come</span></div>
      <div class="card list" style="padding:4px 16px">${rows.map((r) => `
        <div class="row" data-line="${r.id}">
          <div class="nm"><b class="${r.override !== "" ? "ovr" : ""}">${esc(r.name)}</b><span>${gbp(r.soFar)} of ${gbp(r.plan)} · ${r.type}${r.from === "PL" ? " · PL" : ""}</span></div>
          <div class="amt">${gbp(r.toCome)}<span><i class="pill ${pillCls(r.status)}">${STATUS[r.status] || "-"}</i></span></div>
        </div>`).join("")}</div>`;
  }).join("") + `<p class="note" style="text-align:center">Orange names have an override. Tap any line to change it.</p>`;
}

function vBalances(c) {
  const b = data.state.balances || {};
  const rate = data.model.plnRate;
  const when = (x) => x && x.ts ? new Date(x.ts).toLocaleString("en-GB", { day: "numeric", month: "short", hour: "2-digit", minute: "2-digit" }) : "never";
  const card = (k, name, est, ic) => `<div class="card"><div style="display:flex;justify-content:space-between;align-items:center"><div class="label">${name}</div><div class="note">updated ${when(b[k])}</div></div>
      <div style="font-size:28px;font-weight:800;margin-top:6px;letter-spacing:-.02em" class="${k === "mbank" ? "orange" : ""}">${gbp(est, 2)}</div>
      ${k === "mbank" ? `<p class="note" style="margin:4px 0 0">Left on the revolving${b.mbank && b.mbank.pln ? ` · ${fmt2.format(b.mbank.pln)} zł last typed` : ""} · £1 = ${rate} zł</p>` : ""}
      <div class="field" style="margin:12px 0 0"><input id="b_${k}" inputmode="decimal" placeholder="${k === "mbank" ? "Owed now in zł, e.g. 10869" : "Balance now in £"}"></div></div>`;
  return `
    <div class="card hero"><div class="label">UK cash now</div><div class="big">${gbp(c.cash)}</div>
      <div class="sub">Barclays and Monzo, kept up to date by what you log. ${gbp(c.buffer)} stays back as your buffer.</div></div>
    ${card("barclays", "Barclays", c.barclays)}${card("monzo", "Monzo", c.monzo)}${card("mbank", "mBank revolving", c.mbank)}
    <button class="btn" data-act="savebal">Save balances</button>
    <p class="note">Only the boxes you fill in are updated. From then on, anything you log on that account moves the figure.</p>`;
}

function chart(t, nowKey) {
  const W = 340, H = 190, L = 34, B = 20, T = 8, R = 6;
  const series = [["all", "#ff4f40", 2.6, ""], ["dmp", "#14233c", 2.2, ""], ["deposit", "#1fa99b", 2.6, ""], ["baseline", "#b3bcc9", 1.4, "4 3"]];
  const n = t.labels.length;
  const max = Math.max(1, ...series.flatMap(([k]) => t[k].filter((v) => v !== null)));
  const top = Math.ceil(max / 10000) * 10000;
  const x = (i) => L + (i / (n - 1)) * (W - L - R), y = (v) => T + (1 - v / top) * (H - T - B);
  const path = (arr) => { let d = "", pen = false; arr.forEach((v, i) => { if (v === null) { pen = false; return; } d += (pen ? "L" : "M") + x(i).toFixed(1) + " " + y(v).toFixed(1); pen = true; }); return d; };
  let g = "";
  for (let v = 0; v <= top; v += top / 4) g += `<line x1="${L}" x2="${W - R}" y1="${y(v)}" y2="${y(v)}" stroke="#e6eaf0"/><text x="${L - 4}" y="${y(v) + 3}" fill="#6e7a8c" font-size="8" font-weight="600" text-anchor="end">£${v / 1000}k</text>`;
  t.labels.forEach((lab, i) => { if (/^Jan-/.test(lab)) g += `<text x="${x(i)}" y="${H - 6}" fill="#6e7a8c" font-size="8" font-weight="600" text-anchor="middle">20${lab.slice(4)}</text>`; });
  const ni = t.labels.indexOf(nowKey);
  if (ni >= 0) g += `<line x1="${x(ni)}" x2="${x(ni)}" y1="${T}" y2="${H - B}" stroke="#14233c" stroke-opacity=".3" stroke-dasharray="2 3"/><text x="${x(ni) + 3}" y="${T + 8}" fill="#14233c" fill-opacity=".6" font-size="8" font-weight="700">you are here</text>`;
  const lines = series.map(([k, col, w, dash]) => `<path d="${path(t[k])}" fill="none" stroke="${col}" stroke-width="${w}" ${dash ? `stroke-dasharray="${dash}"` : ""} stroke-linejoin="round" stroke-linecap="round"/>`).join("");
  return `<svg viewBox="0 0 ${W} ${H}" width="100%" role="img" aria-label="Debt falling and savings rising">${g}${lines}</svg>`;
}

function vJourney(c) {
  const m = data.model, t = m.trajectory, J = journey();
  const ci = curIdx();
  return `
    <div class="card hero hero-o"><div class="label">Debt cleared so far</div><div class="big">${gbp(J.cleared)}</div>
      <div class="track" style="background:rgba(255,255,255,.3);margin:6px 0 10px;position:relative;z-index:1"><div class="fill" style="width:${(J.pctDone * 100).toFixed(1)}%;background:#fff"></div></div>
      <div class="sub"><b>${pct(J.pctDone)} of the way</b> · ${gbp(J.owed)} to go${J.df ? " · free by " + esc(J.df.forecast) : ""}</div></div>
    <div class="card"><div class="label">Your freedom journey</div>${stepper(J)}</div>
    <div class="card">${J.stages.map((s, i) => `
      <div class="stage ${s.state}${s.dream ? " dream" : ""}">
        <div class="num">${s.state === "done" ? "✓" : s.icon}</div>
        <div style="flex:1;min-width:0"><h3>${esc(s.title)}</h3><p>${esc(s.text)}</p>
          ${s.id === "debt" ? `<div class="subm">${J.subs.map((x) => { const mo = J.monthsTo(x.forecast); return `<div><span>${mo !== null && mo < 0 ? "✓ " : ""}${esc(x.name)}</span><span class="${mo !== null && mo < 0 ? "ok" : "muted"}">${esc(x.forecast)}</span></div>`; }).join("")}</div>` : ""}
        </div>
        <div class="when ${s.state === "now" ? "orange" : "blue"}">${esc(s.when || (s.dream ? "✨" : "next"))}${s.months !== null && s.months > 0 ? `<span>${s.months} months</span>` : ""}</div>
      </div>`).join("")}</div>
    ${t.labels.length ? `<div class="card"><div class="label">Debt falling, savings rising</div>${chart(t, m.periods[ci].key)}
      <div class="legend"><span><i style="background:#ff4f40"></i>All debts</span><span><i style="background:#14233c"></i>StepChange DMP</span>
      <span><i style="background:#1fa99b"></i>Deposit saved</span><span><i style="background:#b3bcc9"></i>7 Nov plan</span></div></div>` : ""}
    <div class="card dreamcard"><div class="label" style="color:rgba(255,255,255,.85)">The long game</div>
      <h3 style="margin:6px 0 6px;font-size:19px">🏔️ A holiday let in Zakopane</h3>
      <p class="note" style="color:rgba(255,255,255,.92);margin:0">Debt free, a safety net, your own home, money invested, then a family business with Greg in the mountains. Every pound logged today is part of it.</p></div>
    <p class="note" style="text-align:center">Dates come from the workbook forecast. Re-import after you update Excel to refresh them.</p>`;
}

// ---------------------------------------------------------------- sheets
function openSheet(fn) { sheet = fn; render(); }
function closeSheet() { sheet = null; render(); }

function lineSheet(id) {
  return () => {
    const c = calc(); const r = c.rows.find((x) => x.id === id);
    const o = r.override;
    const mode = o === "" ? "auto" : o === "done" ? "done" : o === "to come" ? "tocome" : "amount";
    const ents = c.entries.filter((e) => e.lineId === id);
    return `<div class="grab"></div><h2>${esc(r.name)}</h2>
      <div class="grid2" style="margin-bottom:14px">
        <div class="tile" style="min-height:0"><div class="label">Plan</div><div class="v">${gbp(r.plan, 2)}</div></div>
        <div class="tile" style="min-height:0"><div class="label">So far</div><div class="v">${gbp(r.soFar, 2)}</div></div>
        <div class="tile" style="min-height:0"><div class="label">Still to come</div><div class="v blue">${gbp(r.toCome, 2)}</div></div>
        <div class="tile" style="min-height:0"><div class="label">Likely vs plan</div><div class="v ${r.vs < -0.5 ? "orange" : "green"}">${gbp(r.vs, 2)}</div></div>
      </div>
      <div class="field"><label>Still to come?</label>
        <div class="seg" id="ovseg">${[["auto", "Auto"], ["done", "Done"], ["tocome", "To come"], ["amount", "Amount"]].map(([k, l]) => `<button data-ov="${k}" class="${mode === k ? "on" : ""}">${l}</button>`).join("")}</div></div>
      <div class="field ${mode === "amount" ? "" : "hidden"}" id="ovamt"><label>Amount still to come (£)</label><input id="ovv" inputmode="decimal" value="${mode === "amount" ? o : ""}"></div>
      <p class="note">Auto works it out from the plan. Done means nothing more this period. To come means the full planned amount is still to come.</p>
      <button class="btn" data-act="saveov" data-id="${id}">Save</button>
      <button class="btn ghost" data-act="addfor" data-id="${id}">Log a ${r.type === "Income" ? "payment in" : "spend"} on this line</button>
      ${ents.length ? `<div class="list" style="margin-top:14px"><div class="label">Logged on the phone this period</div>${ents.map((e) => `<div class="row"><div class="nm"><b>${dshort(e.date)}</b><span>${esc(e.account)}${e.note ? " · " + esc(e.note) : ""}</span></div><div class="amt">${gbp(e.amount, 2)}</div><button class="del" data-del="${e.id}">×</button></div>`).join("")}</div>` : ""}`;
  };
}

let addForm = null;
function addSheet(preset = {}) {
  const m = data.model;
  const first = m.lines.find((l) => l.type === "Variable");
  addForm = { kind: "spend", line: first ? first.id : m.lines[0].id, account: "monzo", cur: "gbp", ...preset };
  const L = m.lines.find((l) => l.id === addForm.line);
  if (L && !preset.account) addForm.account = L.from === "PL" ? "mbank" : "monzo";
  if (L && L.type === "Income" && !preset.kind) addForm.kind = "income";
  return () => {
    const f = addForm;
    const opts = m.sections.map((s) => {
      const ls = m.lines.filter((l) => l.section === s.name && (f.kind === "income" ? l.type === "Income" : l.type !== "Income"));
      return ls.length ? `<optgroup label="${esc(s.name)}">${ls.map((l) => `<option value="${l.id}" ${l.id === f.line ? "selected" : ""}>${esc(l.name)}</option>`).join("")}</optgroup>` : "";
    }).join("");
    return `<div class="grab"></div><h2>Log money ✍️</h2>
      <div class="field"><div class="seg">${[["spend", "Spend"], ["income", "Money in"], ["transfer", "To mBank"]].map(([k, l]) => `<button data-kind="${k}" class="${f.kind === k ? "on" : ""}">${l}</button>`).join("")}</div></div>
      <div class="field"><input id="amt" class="amount" inputmode="decimal" placeholder="${f.cur === "pln" ? "0 zł" : "£0.00"}" value="${esc(f.amt || "")}"></div>
      ${f.kind !== "transfer" ? `<div class="field"><label>Line</label><select id="line">${opts}</select></div>` : ""}
      <div class="field"><label>${f.kind === "transfer" ? "Sent from" : "Account"}</label><div class="seg">${(f.kind === "transfer" ? [["monzo", "Monzo"], ["barclays", "Barclays"]] : [["monzo", "Monzo"], ["barclays", "Barclays"], ["mbank", "mBank"]])
        .map(([k, l]) => `<button data-acc="${k}" class="${f.account === k ? "on" : ""}">${l}</button>`).join("")}</div></div>
      ${f.account === "mbank" && f.kind !== "transfer" ? `<div class="field"><div class="seg">${[["gbp", "£"], ["pln", "zł"]].map(([k, l]) => `<button data-cur="${k}" class="${f.cur === k ? "on" : ""}">${l}</button>`).join("")}</div></div>` : ""}
      <div class="field"><label>Date</label><input id="dt" type="date" value="${f.date || today()}"></div>
      <div class="field"><label>Note (optional)</label><input id="nt" maxlength="80" value="${esc(f.note || "")}"></div>
      <button class="btn" data-act="saveentry">Save</button>`;
  };
}

function settingsSheet() {
  return () => {
    const m = data.model, s = m.snapshot;
    const bts = s.asAt + "T23:59:59.000Z";
    const snapState = { entries: [], overrides: { [s.key]: s.overrides }, paydays: {},
      balances: { barclays: { v: s.balances.barclays || 0, ts: bts }, monzo: { v: s.balances.monzo || 0, ts: bts }, mbank: { v: s.balances.mbank || 0, ts: bts } } };
    const c = compute(m, snapState, s.asAt, m.periods.findIndex((p) => p.key === s.key));
    const ok = (a, b) => Math.abs(a - b) < 0.02;
    const match = s.check && ok(c.spendLeft, s.check.PR_SpendLeft) && ok(c.spare, s.check.PR_Spare) && ok(c.sendNow, s.check.PR_SendMbank);
    const cp = calc();
    return `<div class="grab"></div><h2>Settings</h2>
      <div class="card" style="background:var(--bg)"><div class="label">Workbook</div>
        <p class="note">Last import: <b>${esc(m.fileName)}</b><br>${new Date(m.importedAt).toLocaleString("en-GB")} · ${esc(s.key)} balances as at ${dshort(s.asAt)}<br>
        ${match ? `<span class="ok">✓ Matches PERIOD REVIEW to the penny</span>` : `<span class="orange">Check: the app and PERIOD REVIEW differ. Spending left ${gbp(c.spendLeft, 2)} vs ${gbp(s.check.PR_SpendLeft, 2)}</span>`}</p>
        <label class="btn" style="display:block;text-align:center">Import a newer workbook<input type="file" id="file" accept=".xlsx" class="hidden"></label><p class="note" id="imp"></p></div>
      <div class="field"><label>Payday for ${esc(cp.period.key)}</label><input id="pd" type="date" value="${cp.payday}"></div>
      <button class="btn ghost" data-act="savepd">Save payday</button>
      <p class="note">${store.mode === "local" ? "Test mode: data is kept on this device only. Fill in firebase-config.js to sync to your account." : "Signed in as " + esc(user.email) + ". Data syncs across your devices."}</p>
      ${store.mode === "firebase" ? `<button class="btn danger" data-act="signout">Sign out</button>` : ""}`;
  };
}

// ---------------------------------------------------------------- actions
function bind() {
  $app.querySelectorAll("[data-view]").forEach((b) => b.onclick = () => { view = b.dataset.view; window.scrollTo(0, 0); render(); });
  $app.querySelectorAll("[data-line]").forEach((b) => b.onclick = () => openSheet(lineSheet(b.dataset.line)));
  $app.querySelectorAll("[data-del]").forEach((b) => b.onclick = async (ev) => { ev.stopPropagation(); if (confirm("Delete this entry?")) await store.deleteEntry(b.dataset.del); });
  const bg = $app.querySelector(".sheet-bg");
  if (bg) bg.onclick = (ev) => { if (ev.target === bg) closeSheet(); };
  $app.querySelectorAll("[data-ov]").forEach((b) => b.onclick = () => {
    $app.querySelectorAll("[data-ov]").forEach((x) => x.classList.toggle("on", x === b));
    document.getElementById("ovamt").classList.toggle("hidden", b.dataset.ov !== "amount");
  });
  $app.querySelectorAll("[data-kind]").forEach((b) => b.onclick = () => {
    keepForm(); addForm.kind = b.dataset.kind;
    const m = data.model;
    if (addForm.kind === "income") { const l = m.lines.find((x) => x.type === "Income"); if (l) addForm.line = l.id; addForm.account = "barclays"; }
    else if (addForm.kind === "spend") { const l = m.lines.find((x) => x.id === addForm.line); if (!l || l.type === "Income") addForm.line = (m.lines.find((x) => x.type === "Variable") || m.lines[0]).id; }
    else if (addForm.account === "mbank") addForm.account = "monzo";
    render();
  });
  $app.querySelectorAll("[data-acc]").forEach((b) => b.onclick = () => { keepForm(); addForm.account = b.dataset.acc; if (addForm.account !== "mbank") addForm.cur = "gbp"; render(); });
  $app.querySelectorAll("[data-cur]").forEach((b) => b.onclick = () => { keepForm(); addForm.cur = b.dataset.cur; render(); });
  const ls = document.getElementById("line");
  if (ls) ls.onchange = () => { keepForm(); const l = data.model.lines.find((x) => x.id === ls.value); addForm.line = ls.value; if (l) addForm.account = l.from === "PL" ? "mbank" : addForm.account === "mbank" ? "monzo" : addForm.account; render(); };
  const fi = document.getElementById("file");
  if (fi) fi.onchange = (e) => doImport(e.target.files[0]);
  $app.querySelectorAll("[data-act]").forEach((b) => b.onclick = (ev) => act(b.dataset.act, b, ev));
}

function keepForm() {
  if (!addForm) return;
  const g = (id) => document.getElementById(id);
  if (g("amt")) addForm.amt = g("amt").value;
  if (g("line")) addForm.line = g("line").value;
  if (g("dt")) addForm.date = g("dt").value;
  if (g("nt")) addForm.note = g("nt").value;
}

const parseAmt = (s) => { const v = parseFloat(String(s || "").replace(/[£,\szł]/g, "")); return isFinite(v) ? v : NaN; };

async function act(a, b, ev) {
  const m = data.model;
  if (a === "settings") return openSheet(settingsSheet());
  if (a === "add") return openSheet(addSheet());
  if (a === "prev" || a === "next") { const cur = pIdx ?? curIdx(); pIdx = Math.max(0, Math.min(m.periods.length - 1, cur + (a === "next" ? 1 : -1))); if (pIdx === curIdx()) pIdx = null; return render(); }
  if (a === "addfor") { const L = m.lines.find((l) => l.id === b.dataset.id); return openSheet(addSheet({ line: L.id, kind: L.type === "Income" ? "income" : "spend", account: L.from === "PL" ? "mbank" : (L.type === "Income" ? "barclays" : "monzo") })); }
  if (a === "saveov") {
    const k = ($app.querySelector("[data-ov].on") || {}).dataset?.ov || "auto";
    const per = calc().period.key;
    const st = structuredClone(data.state); st.overrides = st.overrides || {}; st.overrides[per] = st.overrides[per] || {};
    if (k === "auto") delete st.overrides[per][b.dataset.id];
    else if (k === "done") st.overrides[per][b.dataset.id] = "done";
    else if (k === "tocome") st.overrides[per][b.dataset.id] = "to come";
    else { const v = parseAmt(document.getElementById("ovv").value); if (isNaN(v)) return toast("Type an amount"); st.overrides[per][b.dataset.id] = Math.abs(v); }
    await store.setState(st); sheet = null; toast("Saved"); return render();
  }
  if (a === "saveentry") {
    keepForm(); const f = addForm;
    let v = parseAmt(f.amt); if (isNaN(v) || v === 0) return toast("Type an amount");
    v = Math.abs(v); let pln = null;
    if (f.account === "mbank" && f.cur === "pln" && f.kind !== "transfer") { pln = v; v = Math.round((v / m.plnRate) * 100) / 100; }
    const e = { id: uid(), ts: new Date().toISOString(), date: f.date || today(), account: f.account, note: (f.note || "").trim(),
      lineId: f.kind === "transfer" ? "__transfer" : f.line, amount: f.kind === "income" ? v : -v };
    if (pln) e.pln = pln;
    await store.addEntry(e); sheet = null; addForm = null; toast((e.lineId === "__transfer" ? "Sent " + gbp(-e.amount, 2) + " to mBank" : "Logged " + gbp(e.amount, 2)) + " ✓"); return render();
  }
  if (a === "savebal") {
    const st = structuredClone(data.state); st.balances = st.balances || {};
    const ts = new Date().toISOString(); let n = 0;
    for (const k of ["barclays", "monzo", "mbank"]) {
      const raw = document.getElementById("b_" + k).value; if (!raw.trim()) continue;
      const v = parseAmt(raw); if (isNaN(v)) return toast("Check the " + k + " amount");
      st.balances[k] = k === "mbank" ? { v: -Math.abs(v) / m.plnRate, pln: Math.abs(v), ts } : { v, ts }; n++;
    }
    if (!n) return toast("Nothing typed");
    await store.setState(st); toast("Balances saved"); return render();
  }
  if (a === "savepd") { const v = document.getElementById("pd").value; if (!v) return; const st = structuredClone(data.state); st.paydays = st.paydays || {}; st.paydays[calc().period.key] = v; await store.setState(st); sheet = null; toast("Payday saved"); return render(); }
  if (a === "signout") { sheet = null; await store.signOut(); }
}

async function doImport(file) {
  const out = document.getElementById("imp");
  if (!file) return;
  try {
    out.textContent = "Reading " + file.name + "…";
    const XLSX = await import(SHEETJS);
    const model = readWorkbook(XLSX, new Uint8Array(await file.arrayBuffer()), file.name);
    const s = model.snapshot;
    // state: overrides typed in Excel win for that period; balances replaced unless the phone has a newer one
    const st = structuredClone(data.state || {}); st.overrides = st.overrides || {}; st.balances = st.balances || {}; st.paydays = st.paydays || {};
    st.overrides[s.key] = { ...(st.overrides[s.key] || {}), ...s.overrides };
    const bts = s.asAt + "T23:59:59.000Z";
    const keep = (k) => st.balances[k] && st.balances[k].ts && st.balances[k].ts.slice(0, 10) > s.asAt;
    if (!keep("barclays")) st.balances.barclays = { v: s.balances.barclays || 0, ts: bts };
    if (!keep("monzo")) st.balances.monzo = { v: s.balances.monzo || 0, ts: bts };
    if (!keep("mbank")) st.balances.mbank = { v: s.balances.mbank || 0, pln: s.mbankPLN, ts: bts };
    delete st.paydays[s.key];
    await store.setModel(model); await store.setState(st);
    pIdx = null; sheet = null; toast("Imported " + s.key); render();
  } catch (e) {
    console.error(e); if (out) out.textContent = "Import failed: " + e.message;
  }
}
