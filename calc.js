// The same in-month maths as PERIOD REVIEW, plus whatever has been logged on the phone since the last import.
const DAY = 86400000;
const d2n = (iso) => Date.parse(iso + "T00:00:00Z") / DAY;
const r2 = (x) => Math.round(x * 100) / 100;

export function todayISO(now = new Date()) {
  const p = (n) => String(n).padStart(2, "0");
  return `${now.getFullYear()}-${p(now.getMonth() + 1)}-${p(now.getDate())}`;
}

export function periodIndexFor(model, iso) {
  const i = model.periods.findIndex((p) => iso >= p.start && iso <= p.end);
  if (i >= 0) return i;
  return iso < model.periods[0].start ? 0 : model.periods.length - 1;
}

export function defaultPayday(endISO) {
  const d = new Date(Date.parse(endISO + "T00:00:00Z") + DAY);
  const wd = d.getUTCDay(); // 0 Sun, 6 Sat -> previous Friday
  if (wd === 6) d.setTime(d.getTime() - DAY);
  if (wd === 0) d.setTime(d.getTime() - 2 * DAY);
  return d.toISOString().slice(0, 10);
}

export function autoToCome(plan, soFar) {
  if (plan < 0) return Math.min(0, plan - soFar);
  if (plan > 0) return Math.max(0, plan - soFar);
  return 0;
}

export function applyOverride(o, line, plan, soFar) {
  if (o === undefined || o === null || o === "") return autoToCome(plan, soFar);
  if (o === "done") return 0;
  if (o === "to come") return plan;
  if (typeof o === "number" && isFinite(o)) return (line.type === "Income" || plan > 0 ? 1 : -1) * Math.abs(o);
  return autoToCome(plan, soFar);
}

export function statusOf(plan, soFar, toCome, overridden) {
  if (overridden) return Math.abs(toCome) < 0.005 ? (plan < 0 && soFar < plan - 0.5 ? "over budget" : "done") : "to come";
  if (plan === 0 && soFar === 0) return "";
  if (plan > 0) return soFar >= plan - 0.5 ? "received" : "to come";
  if (plan === 0) return "unplanned";
  if (soFar < plan - 0.5) return "over budget";
  return toCome === 0 ? "done" : "on track";
}

// state = { entries:[{id,ts,date,lineId,amount,account,note}], balances:{barclays:{v,ts},monzo:{v,ts},mbank:{v,pln,ts}},
//           overrides:{[periodKey]:{[lineId]:value}}, paydays:{[periodKey]:iso} }
export function compute(model, state, today, periodIdx = null) {
  const pi = periodIdx ?? periodIndexFor(model, today);
  const P = model.periods[pi];
  const snap = model.snapshot;
  const isSnap = P.key === snap.key;
  // entries dated on or before the workbook's "balances as at" day are already in the bank downloads
  const cutoff = isSnap ? snap.asAt : "";
  const entries = (state.entries || []).filter((e) => e.date >= P.start && e.date <= P.end && e.date > cutoff);
  const ov = (state.overrides && state.overrides[P.key]) || {};

  const rows = model.lines.map((L) => {
    const plan = L.plan[pi] || 0;
    const logged = entries.filter((e) => e.lineId === L.id).reduce((a, e) => a + e.amount, 0);
    const soFar = r2((isSnap ? snap.soFar[L.id] || 0 : 0) + logged);
    const o = ov[L.id];
    const overridden = o !== undefined && o !== null && o !== "";
    const toCome = r2(applyOverride(o, L, plan, soFar));
    const likely = r2(soFar + toCome);
    return { ...L, plan, soFar, toCome, likely, vs: r2(likely - plan), override: overridden ? o : "", status: statusOf(plan, soFar, toCome, overridden) };
  });

  const sum = (f, k) => r2(rows.filter(f).reduce((a, r) => a + r[k], 0));

  // UK cash now: last typed balance + anything logged against that account since
  const bal = state.balances || {};
  const acct = (name) => {
    const b = bal[name] || { v: 0, ts: "" };
    const after = (state.entries || []).filter((e) => e.ts > (b.ts || ""));
    let moves = after.filter((e) => e.account === name).reduce((a, e) => a + e.amount, 0);
    if (name === "mbank") moves -= after.filter((e) => e.lineId === "__transfer").reduce((a, e) => a + e.amount, 0); // money sent in
    return r2((b.v || 0) + moves);
  };
  const barclays = acct("barclays"), monzo = acct("monzo");
  const mbank = acct("mbank");
  const cash = r2(barclays + monzo);
  const sent = r2(-entries.filter((e) => e.lineId === "__transfer").reduce((a, e) => a + e.amount, 0)); // amounts out of UK are negative

  const incomeUK = sum((r) => r.type === "Income" && r.from === "UK", "toCome");
  const fixedUK = sum((r) => r.type === "Fixed" && r.from === "UK", "toCome");
  const varUK = sum((r) => r.type === "Variable" && r.from === "UK", "toCome");
  const plToCome = sum((r) => r.from === "PL", "toCome");
  const cover = Math.min(Math.max(0, sent), Math.max(0, -plToCome));
  const ukPayday = r2(cash + cover + incomeUK + fixedUK + varUK);
  const buffer = model.buffer;
  const spare = r2(ukPayday - buffer + plToCome);
  const sendNow = r2(Math.max(0, Math.max(0, spare) - plToCome - cover));

  const payday = (state.paydays && state.paydays[P.key]) || (isSnap ? snap.payday : defaultPayday(P.end));
  const daysLeft = Math.max(0, d2n(payday) - d2n(today));
  const days = d2n(P.end) - d2n(P.start) + 1;
  const daysGone = Math.min(days, Math.max(0, d2n(today) - d2n(P.start) + 1));
  const elapsed = days ? daysGone / days : 0;

  const spendLeft = r2(-sum((r) => r.type === "Variable", "toCome"));
  const perDay = r2(daysLeft > 0 ? spendLeft / daysLeft : spendLeft);
  const varActual = sum((r) => r.type === "Variable", "soFar");
  const varPlan = sum((r) => r.type === "Variable", "plan");
  const paceGap = r2(varActual - varPlan * elapsed); // negative = spending ahead of an even pace
  const ratio = (f) => { const p = sum(f, "plan"); return p === 0 ? 0 : sum(f, "soFar") / p; };

  return {
    pi, period: P, isSnap, rows, payday, daysLeft, days, daysGone, elapsed,
    barclays, monzo, mbank, cash, sent, cover, incomeUK, fixedUK, varUK, plToCome, ukPayday, buffer, spare, sendNow,
    spendLeft, perDay, varActual, varPlan, paceGap,
    budgetUsed: ratio((r) => r.type === "Variable"), billsPaid: ratio((r) => r.type === "Fixed"), incomeIn: ratio((r) => r.type === "Income"),
    entries,
  };
}
