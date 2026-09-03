import { useState, useMemo, useEffect, useRef } from "react";

/* ===================================================================
   Portfolio Desk — specialist analytics view for admins.
   Ported from a standalone dashboard artifact; wired here to read
   directly from the CRM's real building/unit data (state.buildings,
   state.units) instead of a hardcoded dataset. All CSS is scoped
   under .pd-root so it can't leak into the rest of the admin UI.
=================================================================== */

const CSS = `
@import url('https://fonts.googleapis.com/css2?family=Fraunces:opsz,wght@9..144,400;9..144,500;9..144,600&family=Inter:wght@400;500;600;700&display=swap');

.pd-root{ --pd-navy:#10233f; --pd-navy-2:#17325a; --pd-cream:#f6f3ec; --pd-cream-2:#efeadd; --pd-gold:#b9924a; --pd-gold-2:#d9b876; --pd-ink:#16233a; --pd-slate:#5c6a80; --pd-line:#e2ddcc; --pd-white:#ffffff; --pd-occ:#2f6d4f; --pd-occ-bg:#e7f1ea; --pd-vac:#a5471f; --pd-vac-bg:#f7ead9; --pd-danger:#b23a2f; --pd-shadow:0 10px 30px -12px rgba(16,35,63,0.25); }
.pd-root, .pd-root *{ box-sizing:border-box; }
.pd-root{ font-family:'Inter',sans-serif; background:var(--pd-cream); color:var(--pd-ink); -webkit-font-smoothing:antialiased; }
.pd-root h1, .pd-root h2, .pd-root h3, .pd-root .serif{ font-family:'Fraunces',serif; }

.pd-root #app{ display:flex; }

.pd-root #sidebar{ width:272px; flex:0 0 272px; background:linear-gradient(180deg,var(--pd-navy) 0%, var(--pd-navy-2) 100%); color:var(--pd-cream); display:flex; flex-direction:column; }
.pd-root #tower-nav{ padding:22px 10px 26px 10px; flex:1; }
.pd-root .nav-group-label{ font-size:10.5px; letter-spacing:1.3px; text-transform:uppercase; color:rgba(246,243,236,0.45); padding:14px 12px 6px 12px; }
.pd-root .tower-item{ display:flex; align-items:center; justify-content:space-between; padding:9px 12px; margin:1px 4px; border-radius:8px; cursor:pointer; font-size:13.5px; color:rgba(246,243,236,0.88); transition:background .15s ease; }
.pd-root .tower-item:hover{ background:rgba(255,255,255,0.06); }
.pd-root .tower-item.active{ background:var(--pd-gold); color:var(--pd-navy); font-weight:600; }
.pd-root .tower-item .cnt{ font-size:11px; color:rgba(246,243,236,0.5); }
.pd-root .tower-item.active .cnt{ color:rgba(16,35,63,0.65); }
.pd-root .all-item{ margin:6px 4px 4px 4px; padding:10px 12px; border-radius:8px; background:rgba(255,255,255,0.06); font-size:13.5px; cursor:pointer; display:flex; justify-content:space-between; align-items:center; }
.pd-root .all-item.active{ background:var(--pd-gold); color:var(--pd-navy); font-weight:600; }

.pd-root #main{ flex:1; min-width:0; padding:26px 34px 60px 34px; }
.pd-root #topbar{ display:flex; align-items:flex-end; justify-content:space-between; gap:20px; margin-bottom:22px; flex-wrap:wrap; }
.pd-root #topbar h1{ font-size:26px; font-weight:500; margin:0 0 4px 0; color:var(--pd-navy); }
.pd-root #topbar .meta{ font-size:13px; color:var(--pd-slate); }
.pd-root #searchbox{ position:relative; width:300px; max-width:100%; }
.pd-root #searchbox input{ width:100%; padding:10px 14px 10px 34px; border-radius:9px; border:1px solid var(--pd-line); background:var(--pd-white); font-size:13.5px; font-family:'Inter',sans-serif; color:var(--pd-ink); }
.pd-root #searchbox input:focus{ outline:2px solid var(--pd-gold-2); outline-offset:-1px; }
.pd-root #searchbox svg{ position:absolute; left:11px; top:50%; transform:translateY(-50%); opacity:.45; }

.pd-root #kpi-row{ display:grid; grid-template-columns:repeat(auto-fit,minmax(170px,1fr)); gap:12px; margin-bottom:24px; }
.pd-root .kpi-card{ background:var(--pd-white); border:1px solid var(--pd-line); border-radius:12px; padding:16px 18px; box-shadow:var(--pd-shadow); }
.pd-root .kpi-card .val{ font-family:'Fraunces',serif; font-size:26px; color:var(--pd-navy); font-weight:500; }
.pd-root .kpi-card .lbl{ font-size:11.5px; color:var(--pd-slate); margin-top:3px; }
.pd-root .kpi-card .sub{ font-size:11px; color:var(--pd-gold); margin-top:6px; font-weight:600; }

.pd-root #tower-grid{ display:grid; grid-template-columns:repeat(auto-fill,minmax(240px,1fr)); gap:14px; margin-bottom:10px; }
.pd-root .tcard{ background:var(--pd-white); border:1px solid var(--pd-line); border-radius:12px; padding:16px 18px; cursor:pointer; transition:transform .12s ease, box-shadow .12s ease; }
.pd-root .tcard:hover{ transform:translateY(-2px); box-shadow:var(--pd-shadow); }
.pd-root .tcard .tname{ font-family:'Fraunces',serif; font-size:17px; color:var(--pd-navy); font-weight:500; }
.pd-root .tcard .tside{ font-size:10.5px; color:var(--pd-gold); text-transform:uppercase; letter-spacing:1px; margin-top:2px; }
.pd-root .tcard .trow{ display:flex; justify-content:space-between; font-size:12.5px; color:var(--pd-slate); margin-top:10px; }
.pd-root .tcard .trow b{ color:var(--pd-ink); font-weight:600; }
.pd-root .occbar{ height:6px; background:var(--pd-cream-2); border-radius:4px; margin-top:10px; overflow:hidden; }
.pd-root .occbar > div{ height:100%; background:var(--pd-occ); }

.pd-root #filters{ display:flex; gap:8px; flex-wrap:wrap; margin-bottom:14px; align-items:center; }
.pd-root .chip-select{ font-size:12.5px; padding:7px 10px; border-radius:8px; border:1px solid var(--pd-line); background:var(--pd-white); color:var(--pd-ink); font-family:'Inter',sans-serif; }
.pd-root .clear-link{ font-size:12px; color:var(--pd-gold); cursor:pointer; text-decoration:underline; margin-left:4px; }

.pd-root #table-wrap{ background:var(--pd-white); border:1px solid var(--pd-line); border-radius:12px; overflow-x:auto; overflow-y:hidden; box-shadow:var(--pd-shadow); -webkit-overflow-scrolling:touch; }
.pd-root table{ width:100%; min-width:880px; border-collapse:collapse; font-size:12.8px; }
.pd-root thead th{ text-align:left; padding:11px 12px; background:var(--pd-cream-2); color:var(--pd-navy); font-weight:600; font-size:11.3px; text-transform:uppercase; letter-spacing:.4px; border-bottom:1px solid var(--pd-line); cursor:pointer; white-space:nowrap; user-select:none; }
.pd-root thead th:hover{ color:var(--pd-gold); }
.pd-root tbody td{ padding:10px 12px; border-bottom:1px solid var(--pd-line); color:var(--pd-ink); white-space:nowrap; }
.pd-root tbody tr{ cursor:pointer; }
.pd-root tbody tr:hover{ background:var(--pd-cream-2); }
.pd-root .status-pill{ display:inline-block; padding:2px 9px; border-radius:20px; font-size:11px; font-weight:600; }
.pd-root .status-pill.occ{ background:var(--pd-occ-bg); color:var(--pd-occ); }
.pd-root .status-pill.vac{ background:var(--pd-vac-bg); color:var(--pd-vac); }
.pd-root .gap-pos{ color:var(--pd-occ); font-weight:600; }
.pd-root .gap-neg{ color:var(--pd-danger); font-weight:600; }
.pd-root .expiry-soon{ color:var(--pd-danger); font-weight:600; }
.pd-root #tablefoot{ padding:10px 14px; font-size:11.5px; color:var(--pd-slate); display:flex; justify-content:space-between; }

.pd-root #overlay{ position:fixed; inset:0; background:rgba(16,23,38,0.42); display:none; z-index:40; }
.pd-root #overlay.show{ display:block; }
.pd-root #drawer{ position:fixed; top:0; right:-540px; width:520px; max-width:94vw; height:100vh; background:var(--pd-cream); z-index:50; box-shadow:-18px 0 40px rgba(0,0,0,0.25); transition:right .25s ease; overflow-y:auto; }
.pd-root #drawer.open{ right:0; }
.pd-root #drawer-head{ background:var(--pd-navy); color:var(--pd-cream); padding:22px 26px 20px 26px; position:sticky; top:0; z-index:2; }
.pd-root #drawer-head .close{ position:absolute; top:18px; right:20px; cursor:pointer; font-size:20px; color:rgba(246,243,236,0.7); }
.pd-root #drawer-head h2{ font-size:22px; margin:0 0 3px 0; font-weight:500; }
.pd-root #drawer-head .sub{ font-size:12.5px; color:var(--pd-gold-2); }
.pd-root #drawer-body{ padding:22px 26px 40px 26px; }
.pd-root .dblock{ background:var(--pd-white); border:1px solid var(--pd-line); border-radius:12px; padding:16px 18px; margin-bottom:14px; }
.pd-root .dblock h3{ font-size:11.5px; text-transform:uppercase; letter-spacing:.8px; color:var(--pd-gold); margin:0 0 12px 0; font-weight:700; }
.pd-root .dgrid{ display:grid; grid-template-columns:1fr 1fr; gap:10px 16px; }
.pd-root .dfield .k{ font-size:10.8px; color:var(--pd-slate); margin-bottom:2px; }
.pd-root .dfield .v{ font-size:14px; color:var(--pd-ink); font-weight:600; }
.pd-root .rentbar-wrap{ margin-top:6px; }
.pd-root .rentbar{ height:10px; background:var(--pd-cream-2); border-radius:6px; position:relative; margin-top:6px; }
.pd-root .rentbar .fill{ position:absolute; top:0; bottom:0; background:var(--pd-gold-2); border-radius:6px; }
.pd-root .rentbar .marker{ position:absolute; top:-4px; width:2px; height:18px; background:var(--pd-navy); }
.pd-root .rentlabels{ display:flex; justify-content:space-between; font-size:10.5px; color:var(--pd-slate); margin-top:5px; }
.pd-root textarea#blurb{ width:100%; min-height:150px; border:1px solid var(--pd-line); border-radius:8px; padding:12px; font-family:'Inter',sans-serif; font-size:12.8px; color:var(--pd-ink); resize:vertical; background:var(--pd-cream); }
.pd-root .btn{ display:inline-flex; align-items:center; gap:6px; padding:9px 15px; border-radius:8px; border:none; font-size:12.8px; font-weight:600; cursor:pointer; font-family:'Inter',sans-serif; }
.pd-root .btn-gold{ background:var(--pd-gold); color:var(--pd-navy); }
.pd-root .btn-gold:hover{ background:var(--pd-gold-2); }
.pd-root .btn-row{ display:flex; gap:8px; flex-wrap:wrap; margin-top:10px; }
.pd-root .msg-tabs{ display:flex; gap:6px; margin-bottom:10px; }
.pd-root .msg-tab{ padding:6px 12px; border-radius:7px; font-size:11.8px; cursor:pointer; background:var(--pd-cream-2); color:var(--pd-slate); }
.pd-root .msg-tab.active{ background:var(--pd-navy); color:var(--pd-cream); }
.pd-root .copied-flag{ font-size:11.5px; color:var(--pd-occ); margin-left:8px; opacity:0; transition:opacity .2s; }
.pd-root .copied-flag.show{ opacity:1; }

.pd-root ::-webkit-scrollbar{ width:9px; height:9px; }
.pd-root ::-webkit-scrollbar-thumb{ background:#cfc7ac; border-radius:5px; }

@media(max-width:880px){
  .pd-root #sidebar{ position:fixed; left:-272px; z-index:60; transition:left .2s ease; }
  .pd-root #sidebar.open{ left:0; }
  .pd-root #main{ padding:18px; }
  .pd-root #menu-btn{ display:inline-flex; }
}
.pd-root #menu-btn{ display:none; align-items:center; gap:6px; background:var(--pd-navy); color:var(--pd-cream); border:none; padding:8px 12px; border-radius:8px; font-size:12.5px; cursor:pointer; margin-bottom:12px; }
`;

/* ---------- helpers ---------- */
const fmt = n => (n === null || n === undefined ? "—" : Number(n).toLocaleString("en-US"));
const fmtAED = n => (n === null || n === undefined ? "—" : "AED " + Number(n).toLocaleString("en-US"));
const round1 = n => Math.round(n * 10) / 10;
const round2 = n => Math.round(n * 100) / 100;

function toNum(v) {
  if (v === null || v === undefined || v === "") return null;
  const n = Number(String(v).replace(/,/g, ""));
  return Number.isFinite(n) ? n : null;
}
function median(sorted) {
  const n = sorted.length;
  if (!n) return null;
  const mid = Math.floor(n / 2);
  return n % 2 ? sorted[mid] : (sorted[mid - 1] + sorted[mid]) / 2;
}
function daysUntil(dateStr) {
  if (!dateStr) return null;
  const d = new Date(dateStr);
  if (isNaN(d.getTime())) return null;
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  d.setHours(0, 0, 0, 0);
  return Math.round((d - today) / 86400000);
}

const COLS = [
  { k: "unit_id", l: "Unit" },
  { k: "tower", l: "Building", hideInTower: true },
  { k: "type", l: "Type" },
  { k: "floor", l: "Floor" },
  { k: "beds", l: "Beds" },
  { k: "sqft", l: "Sqft" },
  { k: "status", l: "Status" },
  { k: "current_rent", l: "Current Rent" },
  { k: "comp_median", l: "Comp Median" },
  { k: "gap_pct", l: "Gap %" },
  { k: "yield_pct", l: "Yield %" },
  { k: "days_to_expiry", l: "Lease Ends" },
];

/* Build display rows from real CRM state (state.buildings / state.units),
   computing yield / comps / gap / lease-expiry client-side. */
function buildRows(state) {
  const buildings = state.buildings || {};
  const units = state.units || {};

  const rows = Object.values(units).map(u => {
    const buildingName = buildings[u.buildingId]?.name || "Unassigned";
    const status = String(u.status || "").toLowerCase() === "occupied" ? "Occupied" : "Vacant";
    const rent = toNum(u.currentRent);
    const purchase = toNum(u.purchasePrice);
    return {
      key: u.key,
      buildingId: u.buildingId,
      tower: buildingName,
      unit_id: u.unitId || "",
      type: u.unitType || "",
      floor: u.floor || "",
      beds: u.bedrooms || "",
      baths: u.bathrooms || "",
      sqft: toNum(u.carpetArea),
      status,
      current_rent: rent,
      lease_start: u.currentLeaseStart || "",
      lease_end: u.currentLeaseEnd || "",
      days_to_expiry: daysUntil(u.currentLeaseEnd),
      purchase_price: purchase,
      ownership_since: u.ownershipStart || "",
      ownerName: u.ownerName || "",
      ownerContact: u.ownerContact || "",
      email: u.email || "",
      yield_pct: rent && purchase ? round2((rent / purchase) * 100) : null,
    };
  });

  // Comps: occupied units of the same type, within the same building.
  const compGroups = {};
  rows.forEach(r => {
    if (r.status === "Occupied" && r.current_rent) {
      const gk = r.buildingId + "::" + r.type;
      (compGroups[gk] = compGroups[gk] || []).push(r.current_rent);
    }
  });
  rows.forEach(r => {
    const gk = r.buildingId + "::" + r.type;
    const rents = (compGroups[gk] || []).slice().sort((a, b) => a - b);
    if (rents.length) {
      r.comp_n = rents.length;
      r.comp_low = rents[0];
      r.comp_high = rents[rents.length - 1];
      r.comp_median = median(rents);
      r.gap_pct = r.current_rent ? round1(((r.current_rent - r.comp_median) / r.comp_median) * 100) : null;
    } else {
      r.comp_n = 0;
      r.comp_low = null;
      r.comp_high = null;
      r.comp_median = null;
      r.gap_pct = null;
    }
  });

  return rows;
}

function buildBuildingSummary(rows, name) {
  const list = rows.filter(r => r.tower === name);
  const total = list.length;
  const occupied = list.filter(r => r.status === "Occupied").length;
  const vacant = total - occupied;
  const occ_pct = total ? round1((occupied / total) * 100) : 0;
  const occRents = list.filter(r => r.status === "Occupied" && r.current_rent).map(r => r.current_rent);
  const avg_rent = occRents.length ? Math.round(occRents.reduce((a, b) => a + b, 0) / occRents.length) : null;
  const yields = list.filter(r => r.yield_pct).map(r => r.yield_pct);
  const avg_yield = yields.length ? round2(yields.reduce((a, b) => a + b, 0) / yields.length) : null;
  const renewals_60d = list.filter(r => r.days_to_expiry !== null && r.days_to_expiry >= 0 && r.days_to_expiry <= 60).length;
  const renewals_90d = list.filter(r => r.days_to_expiry !== null && r.days_to_expiry >= 0 && r.days_to_expiry <= 90).length;
  const types = Array.from(new Set(list.map(r => r.type).filter(Boolean))).sort();
  const beds = Array.from(new Set(list.map(r => r.beds).filter(v => v !== "" && v != null))).sort((a, b) => Number(a) - Number(b));
  return { total, occupied, vacant, occ_pct, avg_rent, avg_yield, renewals_60d, renewals_90d, types, beds };
}

function buildBlurb(mode, u) {
  const rentLine =
    u.status === "Occupied"
      ? `Currently leased at ${fmtAED(u.current_rent)}/yr, ending ${u.lease_end || "TBC"}.`
      : `Currently vacant — ready to let.`;
  const suggestLine = u.comp_median
    ? `Market comps for this type/building are running ${fmtAED(u.comp_low)}–${fmtAED(u.comp_high)}, median ${fmtAED(u.comp_median)}, based on ${u.comp_n} comparable unit${u.comp_n === 1 ? "" : "s"}.`
    : "";
  const yieldLine = u.yield_pct ? `Estimated gross yield: ${u.yield_pct}%.` : "";

  if (mode === "whatsapp") {
    return (
      `*${u.tower} — Unit ${u.unit_id}*\n` +
      `Type ${u.type || "—"} · ${u.beds || "—"} bed · ${fmt(u.sqft)} sqft · Floor ${u.floor || "—"}\n\n` +
      `${rentLine}\n${suggestLine}\n${yieldLine}\n\n` +
      `Happy to arrange a viewing or send more details — let me know a good time.`
    );
  }
  if (mode === "email") {
    return (
      `Subject: ${u.tower} — Unit ${u.unit_id}\n\n` +
      `Hi,\n\nSharing the details for Unit ${u.unit_id} at ${u.tower}:\n\n` +
      `- Type: ${u.type || "—"}\n- Size: ${fmt(u.sqft)} sqft, ${u.beds || "—"} bedroom(s), Floor ${u.floor || "—"}\n` +
      `- Status: ${u.status}\n- ${rentLine}\n- ${suggestLine}\n- ${yieldLine}\n\n` +
      `Happy to arrange a viewing at your convenience.\n\nBest regards,`
    );
  }
  return (
    `Call talking points — ${u.tower} Unit ${u.unit_id}\n\n` +
    `• Type ${u.type || "—"}, ${u.beds || "—"} bed, ${fmt(u.sqft)} sqft, floor ${u.floor || "—"}\n` +
    `• ${rentLine}\n` +
    `• ${suggestLine}\n` +
    `• ${yieldLine}\n` +
    `• Close: offer a viewing slot this week, confirm move-in timeline.`
  );
}

function KpiCard({ val, lbl, sub }) {
  return (
    <div className="kpi-card">
      <div className="val">{val}</div>
      <div className="lbl">{lbl}</div>
      {sub !== undefined && sub !== null && <div className="sub">{sub}</div>}
    </div>
  );
}

function RentBar({ u }) {
  if (!u.comp_low || !u.comp_high || u.comp_low === u.comp_high) {
    return <div style={{ fontSize: 12, color: "var(--pd-slate)" }}>Not enough comparable data to plot a range.</div>;
  }
  const lo = u.comp_low, hi = u.comp_high, span = hi - lo;
  const medPos = (((u.comp_median - lo) / span) * 100).toFixed(1);
  let curPos = null;
  if (u.current_rent) curPos = Math.max(0, Math.min(100, ((u.current_rent - lo) / span) * 100));
  return (
    <div className="rentbar-wrap">
      <div className="rentbar">
        <div className="fill" style={{ left: 0, width: "100%", opacity: 0.35 }}></div>
        <div className="marker" style={{ left: medPos + "%" }} title="Median"></div>
        {curPos !== null && <div className="marker" style={{ left: curPos + "%", background: "var(--pd-navy)" }} title="Current rent"></div>}
      </div>
      <div className="rentlabels">
        <span>{fmtAED(lo)}</span>
        <span>median {fmtAED(u.comp_median)}</span>
        <span>{fmtAED(hi)}</span>
      </div>
    </div>
  );
}

export default function PortfolioDashboard({ state }) {
  const rows = useMemo(() => buildRows(state), [state.buildings, state.units]);
  const buildingNames = useMemo(() => Array.from(new Set(rows.map(r => r.tower))).sort(), [rows]);
  const summaries = useMemo(() => {
    const m = {};
    buildingNames.forEach(name => { m[name] = buildBuildingSummary(rows, name); });
    return m;
  }, [rows, buildingNames]);

  const [tower, setTowerState] = useState(null);
  const [search, setSearch] = useState("");
  const [fStatus, setFStatus] = useState("");
  const [fType, setFType] = useState("");
  const [fBeds, setFBeds] = useState("");
  const [fRenewal, setFRenewal] = useState("");
  const [sortKey, setSortKey] = useState("unit_id");
  const [sortDir, setSortDir] = useState(1);
  const [sidebarOpen, setSidebarOpen] = useState(false);
  const [drawerUnitKey, setDrawerUnitKey] = useState(null);
  const [msgMode, setMsgMode] = useState("whatsapp");
  const [copied, setCopied] = useState(false);
  const [blurb, setBlurb] = useState("");
  const copiedTimeout = useRef(null);

  function setTower(t) {
    setTowerState(t);
    setFStatus(""); setFType(""); setFBeds(""); setFRenewal(""); setSearch("");
    setSidebarOpen(false);
  }

  const overall = useMemo(() => {
    const total = rows.length;
    const occ = rows.filter(r => r.status === "Occupied").length;
    const yields = rows.filter(r => r.yield_pct).map(r => r.yield_pct);
    const avgYield = yields.length ? yields.reduce((a, b) => a + b, 0) / yields.length : 0;
    const ren90 = rows.filter(r => r.days_to_expiry !== null && r.days_to_expiry >= 0 && r.days_to_expiry <= 90).length;
    return { total, occ, avgYield, ren90 };
  }, [rows]);

  const currentBuildingSummary = tower ? summaries[tower] : null;
  const searchActive = tower === null && search.trim() !== "";

  const filteredList = useMemo(() => {
    if (searchActive) {
      const q = search.toLowerCase();
      return rows
        .filter(u => (u.unit_id || "").toLowerCase().includes(q) || (u.tower || "").toLowerCase().includes(q) || (u.type || "").toLowerCase().includes(q))
        .slice(0, 300);
    }
    if (tower === null) return [];
    let list = rows.filter(u => u.tower === tower);
    if (fStatus) list = list.filter(u => u.status === fStatus);
    if (fType) list = list.filter(u => u.type === fType);
    if (fBeds) list = list.filter(u => String(u.beds) === fBeds);
    if (fRenewal === "60") list = list.filter(u => u.days_to_expiry !== null && u.days_to_expiry >= 0 && u.days_to_expiry <= 60);
    if (fRenewal === "90") list = list.filter(u => u.days_to_expiry !== null && u.days_to_expiry >= 0 && u.days_to_expiry <= 90);
    if (fRenewal === "overdue") list = list.filter(u => u.days_to_expiry !== null && u.days_to_expiry < 0 && u.status === "Occupied");
    list = list.slice().sort((a, b) => {
      let av = a[sortKey], bv = b[sortKey];
      if (av === null || av === undefined) av = -Infinity;
      if (bv === null || bv === undefined) bv = -Infinity;
      if (typeof av === "string") return av.localeCompare(bv) * sortDir;
      return (av - bv) * sortDir;
    });
    return list;
  }, [rows, tower, search, searchActive, fStatus, fType, fBeds, fRenewal, sortKey, sortDir]);

  function sortBy(k) {
    if (sortKey === k) setSortDir(d => d * -1);
    else { setSortKey(k); setSortDir(1); }
  }

  const drawerUnit = useMemo(() => (drawerUnitKey ? rows.find(r => r.key === drawerUnitKey) : null), [rows, drawerUnitKey]);

  useEffect(() => {
    if (drawerUnit) setBlurb(buildBlurb(msgMode, drawerUnit));
  }, [drawerUnit, msgMode]);

  function openUnit(key) { setDrawerUnitKey(key); setMsgMode("whatsapp"); }
  function closeDrawer() { setDrawerUnitKey(null); }

  function copyBlurb() {
    navigator.clipboard.writeText(blurb).then(() => {
      setCopied(true);
      if (copiedTimeout.current) clearTimeout(copiedTimeout.current);
      copiedTimeout.current = setTimeout(() => setCopied(false), 1600);
    }).catch(() => {});
  }

  const visibleCols = searchActive ? COLS : COLS.filter(c => !c.hideInTower);
  const showTableView = tower !== null || searchActive;

  return (
    <div className="pd-root">
      <style>{CSS}</style>
      <div id="app">
        <aside id="sidebar" className={sidebarOpen ? "open" : ""}>
          <div id="tower-nav">
            <div className={"all-item" + (tower === null ? " active" : "")} onClick={() => setTower(null)}>
              <span>All buildings</span><span className="cnt">{rows.length}</span>
            </div>
            {buildingNames.length > 0 && <div className="nav-group-label">Buildings</div>}
            {buildingNames.map(name => (
              <div key={name} className={"tower-item" + (tower === name ? " active" : "")} onClick={() => setTower(name)}>
                <span>{name}</span><span className="cnt">{summaries[name]?.total || 0}</span>
              </div>
            ))}
          </div>
        </aside>

        <div id="overlay" className={drawerUnit ? "show" : ""} onClick={closeDrawer}></div>

        <main id="main">
          <button id="menu-btn" onClick={() => setSidebarOpen(o => !o)}>☰ Buildings</button>

          <div id="topbar">
            <div>
              <h1 id="page-title">{tower === null ? "Portfolio Overview" : tower}</h1>
              <div className="meta" id="page-meta">
                {tower === null
                  ? `${buildingNames.length} building${buildingNames.length === 1 ? "" : "s"} · ${rows.length.toLocaleString()} units on file`
                  : `Types ${(currentBuildingSummary?.types || []).join(", ") || "—"}`}
              </div>
            </div>
            <div id="searchbox">
              <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><circle cx="11" cy="11" r="8"></circle><line x1="21" y1="21" x2="16.65" y2="16.65"></line></svg>
              <input id="search-input" type="text" placeholder="Search unit ID, building, type…" value={search} onChange={e => setSearch(e.target.value)} />
            </div>
          </div>

          <div id="kpi-row">
            {tower === null ? (
              <>
                <KpiCard val={overall.total} lbl="Total units tracked" />
                <KpiCard val={`${overall.occ} / ${overall.total}`} lbl="Currently occupied" sub={overall.total ? Math.round((overall.occ / overall.total) * 100) + "% occupancy" : undefined} />
                <KpiCard val={overall.avgYield.toFixed(1) + "%"} lbl="Portfolio avg. gross yield" />
                <KpiCard val={overall.ren90} lbl="Leases expiring in 90 days" />
              </>
            ) : currentBuildingSummary ? (
              <>
                <KpiCard val={currentBuildingSummary.total} lbl="Units on file" />
                <KpiCard val={`${currentBuildingSummary.occupied} / ${currentBuildingSummary.total}`} lbl="Occupied" sub={currentBuildingSummary.occ_pct + "% occupancy"} />
                <KpiCard val={currentBuildingSummary.avg_rent ? fmtAED(currentBuildingSummary.avg_rent) : "—"} lbl="Avg. current rent" />
                <KpiCard val={(currentBuildingSummary.avg_yield || 0) + "%"} lbl="Avg. gross yield" />
                <KpiCard val={currentBuildingSummary.renewals_90d} lbl="Renewals due (90d)" />
              </>
            ) : null}
          </div>

          {tower === null && !searchActive && (
            <div id="tower-grid-wrap">
              <div id="tower-grid">
                {buildingNames.length === 0 && (
                  <div style={{ color: "var(--pd-slate)", fontSize: 13 }}>No buildings on file yet — upload units in the Buildings &amp; Upload tab to populate this dashboard.</div>
                )}
                {buildingNames.map(name => {
                  const s = summaries[name];
                  return (
                    <div key={name} className="tcard" onClick={() => setTower(name)}>
                      <div className="tname">{name}</div>
                      <div className="tside">Types {s.types.join(", ") || "—"}</div>
                      <div className="trow"><span>Units</span><b>{s.total}</b></div>
                      <div className="trow"><span>Avg rent</span><b>{s.avg_rent ? fmtAED(s.avg_rent) : "—"}</b></div>
                      <div className="trow"><span>Avg yield</span><b>{s.avg_yield || "—"}%</b></div>
                      <div className="trow"><span>Renewals (90d)</span><b>{s.renewals_90d}</b></div>
                      <div className="occbar"><div style={{ width: s.occ_pct + "%" }}></div></div>
                    </div>
                  );
                })}
              </div>
            </div>
          )}

          {showTableView && (
            <div id="table-view">
              <div id="filters">
                {searchActive ? (
                  <span className="clear-link" onClick={() => setSearch("")}>Clear search — back to overview</span>
                ) : tower !== null ? (
                  <>
                    <select className="chip-select" value={fStatus} onChange={e => setFStatus(e.target.value)}>
                      <option value="">All statuses</option>
                      <option value="Occupied">Occupied</option>
                      <option value="Vacant">Vacant</option>
                    </select>
                    <select className="chip-select" value={fType} onChange={e => setFType(e.target.value)}>
                      <option value="">All types</option>
                      {(currentBuildingSummary?.types || []).map(t => <option key={t} value={t}>Type {t}</option>)}
                    </select>
                    <select className="chip-select" value={fBeds} onChange={e => setFBeds(e.target.value)}>
                      <option value="">All bed counts</option>
                      {(currentBuildingSummary?.beds || []).map(b => <option key={b} value={String(b)}>{b} bed</option>)}
                    </select>
                    <select className="chip-select" value={fRenewal} onChange={e => setFRenewal(e.target.value)}>
                      <option value="">Any lease timing</option>
                      <option value="60">Renewal due ≤60 days</option>
                      <option value="90">Renewal due ≤90 days</option>
                      <option value="overdue">Lease overdue</option>
                    </select>
                    <span className="clear-link" onClick={() => { setFStatus(""); setFType(""); setFBeds(""); setFRenewal(""); }}>Clear filters</span>
                  </>
                ) : null}
              </div>
              <div id="table-wrap">
                <table>
                  <thead>
                    <tr>
                      {visibleCols.map(c => (
                        <th key={c.k} onClick={() => sortBy(c.k)}>{c.l}{sortKey === c.k ? (sortDir === 1 ? " ▲" : " ▼") : ""}</th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {filteredList.length === 0 ? (
                      <tr><td colSpan={visibleCols.length} style={{ padding: 30, textAlign: "center", color: "var(--pd-slate)" }}>{searchActive ? "No matches." : "No units match these filters."}</td></tr>
                    ) : filteredList.map(u => (
                      <tr key={u.key} onClick={() => openUnit(u.key)}>
                        <td><b>{u.unit_id}</b></td>
                        {searchActive && <td>{u.tower}</td>}
                        <td>{u.type || "—"}</td>
                        <td>{u.floor || "—"}</td>
                        <td>{u.beds || "—"}</td>
                        <td>{fmt(u.sqft)}</td>
                        <td>{u.status === "Occupied" ? <span className="status-pill occ">Occupied</span> : <span className="status-pill vac">Vacant</span>}</td>
                        <td>{fmtAED(u.current_rent)}</td>
                        <td>{fmtAED(u.comp_median)}</td>
                        <td className={u.gap_pct > 0 ? "gap-pos" : u.gap_pct < 0 ? "gap-neg" : ""}>{u.gap_pct !== null && u.gap_pct !== undefined ? u.gap_pct + "%" : "—"}</td>
                        <td>{u.yield_pct ? u.yield_pct + "%" : "—"}</td>
                        <td>{u.days_to_expiry === null ? "—" : u.days_to_expiry < 0 ? <span className="expiry-soon">overdue</span> : u.days_to_expiry <= 90 ? <span className="expiry-soon">{u.days_to_expiry}d</span> : u.days_to_expiry + "d"}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
                <div id="tablefoot">
                  <span>{filteredList.length} {searchActive ? (filteredList.length === 1 ? "match" : "matches") : (filteredList.length === 1 ? "unit" : "units")} {searchActive ? "across all buildings" : "shown"}</span>
                  <span>Click a row for full analysis</span>
                </div>
              </div>
            </div>
          )}
        </main>
      </div>

      <div id="drawer" className={drawerUnit ? "open" : ""}>
        <div id="drawer-head">
          <div className="close" onClick={closeDrawer}>✕</div>
          <h2 id="dh-title">{drawerUnit ? `${drawerUnit.tower} · Unit ${drawerUnit.unit_id}` : "—"}</h2>
          <div className="sub" id="dh-sub">{drawerUnit ? `Type ${drawerUnit.type || "—"} · Floor ${drawerUnit.floor || "—"} · ${fmt(drawerUnit.sqft)} sqft` : "—"}</div>
        </div>
        <div id="drawer-body">
          {drawerUnit && (
            <>
              <div className="dblock">
                <h3>Status &amp; Current Lease</h3>
                <div className="dgrid">
                  <div className="dfield"><div className="k">Status</div><div className="v">{drawerUnit.status === "Occupied" ? <span className="status-pill occ">Occupied</span> : <span className="status-pill vac">Vacant</span>}</div></div>
                  <div className="dfield"><div className="k">Current rent</div><div className="v">{fmtAED(drawerUnit.current_rent)}</div></div>
                  <div className="dfield"><div className="k">Lease start</div><div className="v">{drawerUnit.lease_start || "—"}</div></div>
                  <div className="dfield"><div className="k">Lease end</div><div className="v">{drawerUnit.lease_end || "—"}</div></div>
                </div>
              </div>

              <div className="dblock">
                <h3>Market Position — Comparable Rents</h3>
                <div style={{ fontSize: 12.5, color: "var(--pd-slate)", marginBottom: 2 }}>
                  Based on {drawerUnit.comp_n} comparable unit{drawerUnit.comp_n === 1 ? "" : "s"} — Type {drawerUnit.type || "—"}, {drawerUnit.tower}
                </div>
                <RentBar u={drawerUnit} />
                <div className="dgrid" style={{ marginTop: 14 }}>
                  <div className="dfield"><div className="k">Suggested rent (median)</div><div className="v">{fmtAED(drawerUnit.comp_median)}</div></div>
                  <div className="dfield"><div className="k">Range observed</div><div className="v">{fmtAED(drawerUnit.comp_low)} – {fmtAED(drawerUnit.comp_high)}</div></div>
                  <div className="dfield"><div className="k">Gap vs current rent</div><div className={"v " + (drawerUnit.gap_pct > 0 ? "gap-pos" : drawerUnit.gap_pct < 0 ? "gap-neg" : "")}>{drawerUnit.gap_pct !== null && drawerUnit.gap_pct !== undefined ? (drawerUnit.gap_pct > 0 ? "+" : "") + drawerUnit.gap_pct + "%" : "—"}</div></div>
                  <div className="dfield"><div className="k">Est. gross yield</div><div className="v">{drawerUnit.yield_pct ? drawerUnit.yield_pct + "%" : "—"}</div></div>
                </div>
              </div>

              <div className="dblock">
                <h3>Ownership</h3>
                <div className="dgrid">
                  <div className="dfield"><div className="k">Purchase price</div><div className="v">{fmtAED(drawerUnit.purchase_price)}</div></div>
                  <div className="dfield"><div className="k">Owner since</div><div className="v">{drawerUnit.ownership_since || "—"}</div></div>
                  <div className="dfield"><div className="k">Owner name</div><div className="v">{drawerUnit.ownerName || "—"}</div></div>
                  <div className="dfield"><div className="k">Owner contact</div><div className="v">{drawerUnit.ownerContact || drawerUnit.email || "—"}</div></div>
                </div>
              </div>

              <div className="dblock">
                <h3>Send to Client</h3>
                <div className="msg-tabs">
                  {["whatsapp", "email", "call"].map(m => (
                    <div key={m} className={"msg-tab" + (msgMode === m ? " active" : "")} onClick={() => setMsgMode(m)}>
                      {m === "whatsapp" ? "WhatsApp" : m === "email" ? "Email" : "Call script"}
                    </div>
                  ))}
                </div>
                <textarea id="blurb" value={blurb} onChange={e => setBlurb(e.target.value)} />
                <div className="btn-row">
                  <button className="btn btn-gold" onClick={copyBlurb}>Copy to clipboard</button>
                  <span className={"copied-flag" + (copied ? " show" : "")}>Copied ✓</span>
                </div>
              </div>
            </>
          )}
        </div>
      </div>
    </div>
  );
}
