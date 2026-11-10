// Reads the saved AG Cash Flow workbook (values Excel cached on save) and returns the app's plan model.
// Nothing is recalculated here: the workbook stays the engine, the app only reads what Excel worked out.

const DAY = 86400000;
export const serialToISO = (n) => new Date(Date.UTC(1899, 11, 30) + Math.round(n) * DAY).toISOString().slice(0, 10);

function colNum(letters) { let n = 0; for (const ch of letters) n = n * 26 + ch.charCodeAt(0) - 64; return n; }
function colLet(n) { let s = ""; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; }

function parseRef(ref) {
  // "'PERIOD REVIEW'!$N$4" or "PLAN!$M$68:$BL$68"
  const m = ref.match(/^'?(.+?)'?!\$?([A-Z]+)\$?(\d+)(?::\$?([A-Z]+)\$?(\d+))?$/);
  if (!m) return null;
  return { sheet: m[1], c1: colNum(m[2]), r1: +m[3], c2: m[4] ? colNum(m[4]) : colNum(m[2]), r2: m[5] ? +m[5] : +m[3] };
}

export function readWorkbook(XLSX, data, fileName = "") {
  const wb = XLSX.read(data, { type: "array", cellFormula: true, cellDates: false });
  const names = {};
  for (const n of (wb.Workbook && wb.Workbook.Names) || []) {
    if (n.Sheet === undefined || n.Sheet === null || !(n.Name in names)) names[n.Name] = n.Ref;
  }

  const sh = (name) => { const s = wb.Sheets[name]; if (!s) throw new Error(`Sheet "${name}" not found - is this the AG Cash Flow workbook?`); return s; };
  const cell = (sheet, c, r) => sheet[colLet(c) + r];
  const val = (sheet, c, r) => { const x = cell(sheet, c, r); return x ? x.v : undefined; };
  const num = (v) => (typeof v === "number" && isFinite(v) ? v : 0);
  const named = (n) => {
    const ref = names[n]; if (!ref) throw new Error(`Name ${n} missing from the workbook`);
    const p = parseRef(ref); return { p, v: val(sh(p.sheet), p.c1, p.r1) };
  };

  // ---- periods
  const CF = sh("CASH FLOW");
  const periods = [];
  for (let c = colNum("C"); c <= colNum("BB"); c++) {
    const key = val(CF, c, 4), end = val(CF, c, 6);
    if (!key || typeof end !== "number") continue;
    periods.push({ key: String(key), end: serialToISO(end) });
  }
  periods.forEach((p, i) => {
    const prevEnd = i ? periods[i - 1].end : null;
    p.start = prevEnd ? new Date(Date.parse(prevEnd) + DAY).toISOString().slice(0, 10)
                      : new Date(Date.UTC(+p.end.slice(0, 4), +p.end.slice(5, 7) - 2, 25)).toISOString().slice(0, 10);
  });

  // ---- PERIOD REVIEW lines
  const PR = sh("PERIOD REVIEW");
  const linesRef = parseRef(names["PR_Lines"]);
  const hdrRow = linesRef.r1 - 1;
  const cols = {};
  for (let c = 1; c <= 20; c++) {
    const h = val(PR, c, hdrRow); if (typeof h !== "string") continue;
    const k = h.trim().toLowerCase();
    if (k === "line") cols.name = c; else if (k === "type") cols.type = c; else if (k === "from") cols.from = c;
    else if (k === "plan") cols.plan = c; else if (k === "so far") cols.soFar = c; else if (k === "override") cols.override = c;
  }
  for (const k of ["name", "type", "from", "plan", "soFar"]) if (!cols[k]) throw new Error(`Column "${k}" not found on PERIOD REVIEW`);
  const PLAN = sh("PLAN");

  const lines = [], sections = [];
  let section = null;
  const soFar = {}, overrides = {};
  for (let r = linesRef.r1; r <= linesRef.r2; r++) {
    const name = val(PR, cols.name, r);
    const planCell = cell(PR, cols.plan, r);
    const f = planCell && planCell.f ? planCell.f : "";
    const m = f.match(/INDEX\(PLAN!\$?([A-Z]+)\$?(\d+):\$?([A-Z]+)\$?(\d+)/i);
    if (m) {
      const row = +m[2], c0 = colNum(m[1]);
      const id = "L" + row;
      const plan = periods.map((_, i) => Math.round(num(val(PLAN, c0 + i, row)) * 100) / 100);
      const type = String(val(PR, cols.type, r) || "Variable");
      const from = String(val(PR, cols.from, r) || "UK");
      if (!section) { section = { name: "OTHER", lines: [] }; sections.push(section); }
      section.lines.push(id);
      lines.push({ id, name: String(name || id), type, from, section: section.name, plan });
      soFar[id] = Math.round(num(val(PR, cols.soFar, r)) * 100) / 100;
      if (cols.override) {
        const o = val(PR, cols.override, r);
        if (o !== undefined && o !== null && o !== "") overrides[id] = typeof o === "number" ? Math.abs(o) : String(o).trim().toLowerCase();
      }
    } else if (typeof name === "string" && name.trim() && !(planCell && planCell.f)) {
      section = { name: name.trim(), lines: [] }; sections.push(section);
    }
  }

  // ---- current position as Excel saw it
  const idx = num(named("PR_Idx").v);
  const key = String(named("PR_Key").v);
  const asAt = serialToISO(num(named("PR_AsAt").v));
  const payday = serialToISO(num(named("PR_Payday").v));
  const plnRate = num(named("PLN_GBP").v);
  const buffer = num(named("CurBuffer").v);

  // bank balances: section 1 rows - find "Actual now" header above the Total row
  let recHdr = 0;
  for (let r = 1; r < hdrRow; r++) if (val(PR, 2, r) === "Account") { recHdr = r; break; }
  let actualCol = 0;
  for (let c = 2; c <= 12; c++) if (String(val(PR, c, recHdr) || "").toLowerCase() === "actual now") actualCol = c;
  const bal = {};
  let mbankPLN = null;
  for (let r = recHdr + 1; r < recHdr + 6; r++) {
    const label = String(val(PR, 2, r) || "").toLowerCase();
    const v = num(val(PR, actualCol, r));
    if (label.startsWith("barclays")) bal.barclays = v;
    else if (label.startsWith("monzo")) bal.monzo = v;
    else if (label.startsWith("mbank")) {
      bal.mbank = v;
      const fx = cell(PR, actualCol, r);
      const pm = fx && fx.f ? fx.f.match(/^-?\s*(-?[\d.]+)\s*\/\s*PLN_GBP/i) : null;
      mbankPLN = pm ? Math.abs(parseFloat(pm[1])) : Math.round(Math.abs(v) * plnRate * 100) / 100;
    }
  }

  const check = {};
  for (const n of ["PR_SpendLeft", "PR_Spare", "PR_SendMbank", "PR_PerDay", "PR_UKNow2"]) {
    try { check[n] = Math.round(num(named(n).v) * 100) / 100; } catch (e) { /* optional */ }
  }

  // ---- debt trajectory (DASH DATA K40:O..)
  const trajectory = { labels: [], all: [], dmp: [], deposit: [], baseline: [] };
  try {
    const DD = sh("DASH DATA");
    for (let r = 41; r < 200; r++) {
      const lab = val(DD, colNum("K"), r);
      if (lab === undefined || lab === null || lab === "") break;
      const g = (c) => { const v = val(DD, colNum(c), r); return typeof v === "number" ? Math.round(v) : null; };
      trajectory.labels.push(String(lab));
      trajectory.all.push(g("L")); trajectory.dmp.push(g("M")); trajectory.deposit.push(g("N")); trajectory.baseline.push(g("O"));
    }
  } catch (e) { /* dashboard data optional */ }

  const milestones = [];
  try {
    const DB = sh("DASHBOARD");
    for (let r = 40; r < 70; r++) {
      if (val(DB, 2, r) === "Milestone") {
        for (let k = r + 1; k < r + 12; k++) {
          const nm = val(DB, 2, k); if (!nm) break;
          milestones.push({ name: String(nm), forecast: String(val(DB, colNum("I"), k) ?? ""), plan: String(val(DB, colNum("M"), k) ?? ""),
                            status: String(val(DB, colNum("U"), k) ?? "") });
        }
        break;
      }
    }
  } catch (e) { /* optional */ }

  return {
    version: 1, fileName, importedAt: new Date().toISOString(),
    periods, sections, lines, plnRate, buffer, trajectory, milestones,
    snapshot: { key, idx, asAt, payday, soFar, overrides, balances: bal, mbankPLN, check },
  };
}
