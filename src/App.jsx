import { useState, useEffect, useMemo, useRef } from "react";
import Papa from "papaparse";
import * as XLSX from "xlsx";
import PortfolioDashboard from "./PortfolioDashboard.jsx";
import {
    Building2, Search, Upload, LogOut, ArrowRightLeft, Phone,
  PhoneMissed, PhoneOff, Check, X, Clock, ChevronRight, RefreshCcw,
  Users, LayoutGrid, ClipboardList, AlertTriangle, Trash2, Settings, Eye, EyeOff, Skull, Calendar,
  MessageCircle, Mail, Briefcase
} from "lucide-react";

/* ---------------------------------------------------------
   THEME
--------------------------------------------------------- */
const THEME = `
@import url('https://fonts.googleapis.com/css2?family=Playfair+Display:wght@700;800;900&family=Inter:wght@400;500;600&family=IBM+Plex+Mono:wght@400;500&display=swap');
:root{
  --bg:#100E0C; --panel:#171412; --panel-2:#1E1A17; --line:#2B2521;
  --text:#F2ECE0; --text-dim:#9A9188; --text-faint:#635B54;
  --accent:#D9552E; --accent-dim:#3A1C11;
  --red:#D9552E; --red-dim:#3A1C11;
  --green:#7E8C6A; --green-dim:#232719;
  --amber:#C08A3E; --amber-dim:#2E2213;
    --violet:#8A7B6E; --violet-dim:#241E1A;
  --gray:#6E655D; --gray-dim:#1F1B18;
  --blue:#5E7A94; --blue-dim:#1A222B;
}
.rcrm{ font-family:'Inter',sans-serif; color:var(--text); background:var(--bg); }
.rcrm .disp{ font-family:'Playfair Display',Georgia,serif; letter-spacing:-0.015em; }
.rcrm .mono{ font-family:'IBM Plex Mono',monospace; }
.rcrm .eyebrow{ font-size:10.5px; text-transform:uppercase; letter-spacing:0.18em; color:var(--text-dim); font-weight:500; }
.rcrm .rule{ height:1px; background:var(--accent); border:none; }
.rcrm ::-webkit-scrollbar{ width:8px; height:8px; }
.rcrm ::-webkit-scrollbar-thumb{ background:var(--line); border-radius:4px; }
.rcrm ::-webkit-scrollbar-thumb:hover{ background:var(--text-faint); }
.rcrm button{ font-family:inherit; cursor:pointer; transition:background 0.12s, border-color 0.12s; }
.rcrm button:disabled{ cursor:not-allowed; opacity:0.4; }
.rcrm input, .rcrm textarea, .rcrm select{ font-family:inherit; box-sizing: border-box;}
.rcrm input:focus, .rcrm textarea:focus, .rcrm select:focus{ border-color:var(--accent); }
.rcrm *:focus-visible{ outline:1px solid var(--accent); outline-offset:2px; }
.tap{ min-height:44px; }
.sweep-btn{ position:relative; overflow:hidden; }
.sweep-btn::before{ content:""; position:absolute; inset:0; background:var(--sweep-color); transform:scaleX(0); transform-origin:left; transition:transform 0.24s ease; z-index:0; }
.sweep-btn:hover::before{ transform:scaleX(1); }
.sweep-btn:hover{ color:#000000 !important; }
.pulse-btn{ transition: transform 0.28s cubic-bezier(0.22, 1, 0.36, 1), background-color 0.28s ease, color 0.28s ease, font-weight 0.28s ease, border-color 0.28s ease; }
.pulse-btn:hover{ transform: scale(1.15); font-weight: 1000; }
.login-pulse:hover{ background: var(--bg) !important; color: var(--accent) !important; border-color: var(--text-dim) !important; }
.logout-pulse:hover{ background: var(--accent) !important; color: #000000 !important; }
.wipe-pulse:hover{ background: var(--red) !important; color: #000000 !important; }
.explorer-window{ height:13px; flex:1; border-radius:1.5px; background:var(--panel-2); }
.explorer-window.has-data{ background:var(--text-faint); }
.explorer-window.match{ background:var(--accent); box-shadow:0 0 5px 0 var(--accent); }
.explorer-unit-card{ border:1px solid var(--line); border-radius:8px; padding:14px 10px; background:var(--panel-2); cursor:pointer; text-align:center; transition:all .12s ease; position:relative; }
.explorer-unit-card:hover{ transform:translateY(-2px); border-color:var(--accent); }
.explorer-unit-card.match{ border-color:var(--accent); border-width:2px; box-shadow:0 0 0 2px var(--accent), 0 0 16px 1px var(--accent-dim); background:var(--accent-dim); }
.explorer-unit-card.dim{ opacity:.32; }
.explorer-dot{ width:7px; height:7px; border-radius:50%; position:absolute; top:10px; right:10px; }
.explorer-dot.full{ background:var(--green); }
.explorer-dot.partial{ background:var(--amber); }
.explorer-dot.none{ background:var(--text-faint); }
.explorer-rented-x{ position:absolute; top:8px; left:9px; font-size:12px; font-weight:700; color:var(--text-faint); line-height:1; }
@keyframes explorerJumpPulse{ 0%{ box-shadow:0 0 0 0 var(--accent); } 70%{ box-shadow:0 0 0 14px transparent; } 100%{ box-shadow:0 0 0 0 transparent; } }
.explorer-unit-card.jumped{ border-color:var(--accent); animation:explorerJumpPulse 1s ease-out 2; }
`;/* ---------------------------------------------------------
   HELPERS: dates, stats, storage keys
--------------------------------------------------------- */
const dayMs = 86400000;
const addDays = (d, n) => new Date(new Date(d).getTime() + n * dayMs);
const addMonths = (d, n) => { const dt = new Date(d); dt.setMonth(dt.getMonth() + n); return dt; };
const startOfNextDay = (d) => { const dt = addDays(d, 1); dt.setHours(8, 0, 0, 0); return dt; };
const isDue = (iso) => !iso || new Date(iso).getTime() <= Date.now();
const fmtDate = (iso) => iso ? new Date(iso).toLocaleDateString(undefined, { month: "short", day: "numeric", year: "numeric" }) : "—";
const daysUntil = (iso) => Math.ceil((new Date(iso).getTime() - Date.now()) / dayMs);

const mean = (arr) => arr.length ? arr.reduce((a, b) => a + b, 0) / arr.length : 0;
const stdev = (arr) => {
  if (arr.length < 2) return 0;
  const m = mean(arr);
  return Math.sqrt(mean(arr.map(v => (v - m) ** 2)));
};
const clamp = (v, lo, hi) => Math.max(lo, Math.min(hi, v));

const STORAGE_KEY = "rental-crm-state-v1";

/* ---------------------------------------------------------
   FIELD MATCHING for CSV import (lenient header matching)
--------------------------------------------------------- */
const CANDIDATES = {
  unitId: ["unit id", "unitid", "unit", "unit number", "unit no", "unit no.", "property no", "property number"],
  unitType: ["unit type", "type", "unit type code", "apartment type"],
  floor: ["floor", "floor level"],
  carpetArea: ["carpet area", "area", "sqft", "sq ft", "size", "built-up area (sqft)", "built-up area"],
  bedrooms: ["bedrooms", "bed", "beds", "bhk"],
  bathrooms: ["bathrooms", "bath", "baths"],
  ownerName: ["owner name", "owner", "name"],
  ownerContact: ["owner contact", "contact", "phone", "owner phone", "mobile"],
  secondaryContact: ["secondary contact", "secondary number", "secondary phone", "alternate number", "alternate contact", "second phone"],
  email: ["email", "email address", "owner email", "e-mail"],
  ownershipStart: ["ownership start", "ownership start date", "purchase date", "acquired"],
  purchasePrice: ["purchase price", "price", "purchase amount"],
  currentRent: ["current rent", "rent", "monthly rent", "asking rent"],
  currentLeaseStart: ["current lease start", "lease start", "current lease start date", "start date"],
  currentLeaseEnd: ["current lease end", "lease end", "current lease end date", "end date"],
  status: ["status", "occupancy", "occupied", "vacant"],
};
function getField(row, key) {
  const cands = CANDIDATES[key] || [];
  const rowKeys = Object.keys(row || {});
  for (const c of cands) {
    const hit = rowKeys.find(k => k.trim().toLowerCase() === c);
    if (hit && String(row[hit]).trim() !== "") return String(row[hit]).trim();
  }
  return "";
}
function normStatus(raw) {
  const s = (raw || "").toLowerCase();
  if (s.includes("vac")) return "vacant";
  if (s.includes("occ") || s.includes("rent")) return "occupied";
  return "";
}
// Strips everything but letters/numbers and lowercases, so "Unit 101",
// "unit-101", " 101 " and "UNIT101" all join to the same key across sheets
// that format IDs differently. Display always uses the original string from
// the unit sheet — this is only the join key.
function normalizeId(raw) {
  return (raw || "").toLowerCase().replace(/[^a-z0-9]/g, "");
}
// Same normalization, used for matching column HEADER NAMES (not cell
// values) against the candidate phrase list. Comparing on the normalized
// form instead of exact lowercase equality means "Unit No.", "Unit-No",
// "UnitNumber" and "unit no" all match the same candidate.
function normText(s) {
  return (s || "").toLowerCase().replace(/[^a-z0-9]/g, "");
}
function guessColumn(headers, field) {
  const cands = (CANDIDATES[field] || []).map(normText).filter(Boolean);
  let best = "", bestScore = 0;
  headers.forEach(h => {
    const nh = normText(h);
    if (!nh) return;
    cands.forEach(nc => {
      let score = 0;
      if (nh === nc) score = 100;
      else if (nh.includes(nc) || nc.includes(nh)) score = 50 + Math.min(nc.length, nh.length);
      if (score > bestScore) { bestScore = score; best = h; }
    });
  });
  return best;
}

/* ---------------------------------------------------------
   SCORING — deliberately built to work off the Unit sheet alone
   (no separate per-tenancy history file). Volatility is a proxy
   based on how soon the current lease turns over or whether the
   unit has no lease on file at all (vacant/unknown = most "in play").
   Demand compares current rent against same-bedroom-count units in
   the building, since Layout isn't reliably present in the data.
--------------------------------------------------------- */
function computeVolatility(unit) {
  if (!unit.currentLeaseStart && !unit.currentLeaseEnd) {
    // Nothing on file — likely vacant, or simply unknown. Either way
    // there's no lease locking the owner in, so treat as high urgency.
    return 90;
  }
  if (!unit.currentLeaseEnd) {
    // Lease started but no end date on file — some uncertainty, but
    // not as strong a signal as a known, approaching expiry.
    return 55;
  }
  const days = daysUntil(unit.currentLeaseEnd);
  if (days <= 0) return 95;   // already expired
  if (days <= 90) return 80;  // expiring soon — good time to call
  if (days <= 180) return 55;
  if (days <= 365) return 30;
  return 10; // long runway left on the lease, low urgency
}
function computeDemand(unit, buildingUnits) {
  const rent = Number(unit.currentRent);
  if (!rent) return 50;
  const comps = buildingUnits
    .filter(u => u.key !== unit.key && u.bedrooms && u.bedrooms === unit.bedrooms && Number(u.currentRent) > 0)
    .map(u => Number(u.currentRent));
  if (!comps.length) return 50;
  const avg = mean(comps);
  const diff = (rent - avg) / avg;
  return Math.round(clamp(50 + diff * 100, 0, 100));
}

/* ---------------------------------------------------------
   CRM bucket logic
--------------------------------------------------------- */
const BUCKET_LABEL = {
  toBeCalled: "To Be Called",
  callMeBack: "Call Me Back",
  unreachable: "Unreachable",
  declined: "Declined",
  cold: "Cold",
  yes: "Active Client",
  whatsapp: "WhatsApp",
  linkedin: "LinkedIn",
  email: "Email",
};
const OUTCOME_LABEL = { yes: "Yes", callMeBack: "Call Me Back", no: "No / Declined", noAnswer: "No Answer" };
const BUCKET_COLOR = {
  toBeCalled: "var(--accent)",
  callMeBack: "var(--amber)",
  unreachable: "var(--gray)",
  declined: "var(--red)",
  cold: "var(--violet)",
  yes: "var(--green)",
  whatsapp: "var(--green)",
  linkedin: "var(--blue)",
  email: "var(--gray)",
};

function defaultCrm() {
  return { bucket: "toBeCalled", notes: "", noAnswerCount: 0, nextActionDate: null, tag: null, lastOutcome: null, lastCallDate: null };
}

// Where does this unit actually show up right now, given dates?
function effectiveBucket(crm) {
  if (!crm) return { display: "toBeCalled", resurfaced: false };
  if (crm.bucket === "toBeCalled") return { display: "toBeCalled", resurfaced: false };
  if (crm.bucket === "yes") return { display: "yes", resurfaced: false };
    if (["callMeBack", "unreachable", "declined", "cold"].includes(crm.bucket) && isDue(crm.nextActionDate)) {
    return { display: "toBeCalled", resurfaced: true, origin: crm.bucket };
  }
  return { display: crm.bucket, resurfaced: false };
}

function applyOutcome(crm, outcome) {
  const c = { ...crm };
  const now = new Date();
  c.lastOutcome = outcome;
  c.lastCallDate = now.toISOString();
  if (outcome === "yes") {
    c.bucket = "yes"; c.nextActionDate = null; c.tag = null;
  } else if (outcome === "callMeBack") {
    c.bucket = "callMeBack"; c.nextActionDate = startOfNextDay(now).toISOString();
  } else if (outcome === "no") {
    c.bucket = "declined"; c.nextActionDate = addMonths(now, 6).toISOString(); c.tag = "declined";
    } else if (outcome === "noAnswer") {
    c.noAnswerCount = (c.noAnswerCount || 0) + 1;
    if (c.noAnswerCount >= 3) {
      c.bucket = "cold"; c.nextActionDate = addMonths(now, 6).toISOString(); c.tag = "cold";
    } else {
      c.bucket = "unreachable"; c.nextActionDate = addDays(now, 3).toISOString();
    }
  } else if (["whatsapp", "linkedin", "email"].includes(outcome)) {
    c.bucket = outcome;
  }
  return c;
}

/* ---------------------------------------------------------
   STORAGE
--------------------------------------------------------- */
const DEFAULT_USERS = {
  huzeif: { username: "huzeif", passkey: "huzeif1234", role: "admin", displayName: "Huzeif" },
  imad: { username: "imad", passkey: "imad252", role: "agent", displayName: "Imad" },
  abdullah: { username: "abdullah", passkey: "abdullah1234", role: "agent", displayName: "Abdullah" },
};

import { createClient } from "@supabase/supabase-js";

const SUPABASE_URL = "https://pekmujkfrgvvhzpnrtqi.supabase.co";
const SUPABASE_ANON_KEY = "sb_publishable_THUzycPHN9xqABnE6wRJ3w_5jYGwpTV";
const supabase = createClient(SUPABASE_URL, SUPABASE_ANON_KEY, {
  auth: { persistSession: false },
  realtime: { params: { eventsPerSecond: 5 } },
});
const STATE_ROW_ID = "singleton";

// The whole app state lives in a single JSONB column of a single row. This is
// the simplest possible schema that still gives cross-device sync: every save
// overwrites the blob, every load reads it, and last write wins. That's fine
// for a 3-person internal tool where only one person is typing at a time; it
// would be wrong for anything with real concurrency.
async function loadState() {
  try {
    const { data, error } = await supabase
      .from("app_state")
      .select("data")
      .eq("id", STATE_ROW_ID)
      .maybeSingle();
    if (error) throw error;
    if (data && data.data && Object.keys(data.data).length > 0) {
      const parsed = data.data;
      if (!parsed.users || Object.keys(parsed.users).length === 0) parsed.users = DEFAULT_USERS;
      // fill any missing top-level keys so older data doesn't crash newer code
      return { buildings: {}, units: {}, assignments: {}, crm: {}, callLog: [], users: DEFAULT_USERS, ...parsed };
    }
  } catch (e) {
    console.error("loadState failed, using empty state:", e);
  }
  return { buildings: {}, units: {}, assignments: {}, crm: {}, callLog: [], users: DEFAULT_USERS };
}

async function saveState(state) {
  try {
    const { error } = await supabase
      .from("app_state")
      .update({ data: state, updated_at: new Date().toISOString() })
      .eq("id", STATE_ROW_ID);
    if (error) throw error;
  } catch (e) {
    console.error("Save failed", e);
  }
}

// Every device subscribes to the singleton row. When any other device writes,
// this fires and lets the app refresh itself so the second device sees the
// change without a manual reload. The unsubscribe returned is called on
// unmount to clean up the socket.
function subscribeToState(onChange) {
  const channel = supabase
    .channel("app_state_singleton")
    .on(
      "postgres_changes",
      { event: "UPDATE", schema: "public", table: "app_state", filter: `id=eq.${STATE_ROW_ID}` },
      (payload) => {
        if (payload.new && payload.new.data) onChange(payload.new.data);
      }
    )
    .subscribe();
  return () => supabase.removeChannel(channel);
}

/* ---------------------------------------------------------
   SAMPLE DATA (for demoing before real CSVs exist)
--------------------------------------------------------- */
function buildSample() {
  const buildingId = "shoreline-a";
  const bedroomOptions = ["0", "1", "1", "2"];
  const units = {};
  for (let i = 1; i <= 12; i++) {
    const unitId = `${100 + i}`;
    const key = `${buildingId}::${unitId}`;
    const bedrooms = bedroomOptions[i % bedroomOptions.length];
    const baseRent = bedrooms === "0" ? 55000 : bedrooms === "1" ? 78000 : 105000;
    const rentJitter = Math.round((Math.random() - 0.5) * 20000);
    const currentRent = baseRent + rentJitter;
    const hasLease = i % 6 !== 0; // some units vacant/unknown on purpose, to exercise that state
    const leaseStart = hasLease ? addMonths(new Date(), -(Math.floor(Math.random() * 10) + 1)) : null;
    const leaseEnd = hasLease ? addMonths(leaseStart, 12) : null;
    units[key] = {
      key, unitId, buildingId, floor: String(Math.ceil(i / 3)),
      bedrooms, carpetArea: bedrooms === "0" ? "480" : bedrooms === "1" ? "740" : "1120",
      bathrooms: bedrooms === "0" ? "1" : bedrooms === "1" ? "1" : "2",
      ownerName: `Owner ${i}`, ownerContact: i % 4 === 0 ? "" : `+971 5${i}0 0000${i}`,
      ownershipStart: addMonths(new Date(), -Math.floor(Math.random() * 60) - 12).toISOString(),
      purchasePrice: String(1000000 + i * 45000),
      currentRent: String(currentRent),
      currentLeaseStart: leaseStart ? leaseStart.toISOString() : "",
      currentLeaseEnd: leaseEnd ? leaseEnd.toISOString() : "",
      status: i % 5 === 0 ? "vacant" : "occupied",
      statusUpdated: new Date().toISOString(),
      assignedTo: null,
    };
  }
  return {
    buildings: { [buildingId]: { id: buildingId, name: "Shoreline Building A" } },
    units,
    assignments: { [buildingId]: ["imad", "abdullah"] },
    crm: {},
  };
}

/* ---------------------------------------------------------
   ROOT APP
--------------------------------------------------------- */
export default function App() {
  const [state, setState] = useState(null);
  const [loading, setLoading] = useState(true);
  const [loggedInUsername, setLoggedInUsername] = useState(() => localStorage.getItem("der-crm-user") || null);
  const [saving, setSaving] = useState(false);

  const saveTimerRef = useRef(null);
  const pendingStateRef = useRef(null);
  // Every write we make bumps a token; when the realtime channel echoes our
  // own write back, the token on the incoming row matches ours and we skip
  // applying it. This stops the "save then receive my own save then re-save"
  // feedback loop that otherwise happens on every edit.
  const localWriteTokenRef = useRef(0);

  useEffect(() => {
    loadState().then(s => { setState(s); setLoading(false); });
    const unsub = subscribeToState((remote) => {
      // Ignore an echo of our own most recent write.
      if (remote && remote.__writeToken === localWriteTokenRef.current) return;
      setState(remote);
    });
    return unsub;
  }, []);

  // Every user action calls persist(nextState). We update local state
  // immediately for snappy UI, then debounce the actual DB write by 400ms so
  // that a burst of rapid edits (typing in a notes field, for example)
  // collapses to a single write instead of one per keystroke.
  const persist = (next) => {
    localWriteTokenRef.current += 1;
    const stamped = { ...next, __writeToken: localWriteTokenRef.current };
    setState(stamped);
    pendingStateRef.current = stamped;
    setSaving(true);
    if (saveTimerRef.current) clearTimeout(saveTimerRef.current);
    saveTimerRef.current = setTimeout(async () => {
      const toSave = pendingStateRef.current;
      pendingStateRef.current = null;
      saveTimerRef.current = null;
      await saveState(toSave);
      setSaving(false);
    }, 400);
  };

  if (loading) {
    return (
      <div className="rcrm" style={{ minHeight: 500, display: "flex", alignItems: "center", justifyContent: "center", background: "var(--bg)" }}>
        <style>{THEME}</style>
        <div style={{ color: "var(--text-dim)" }}>Loading…</div>
      </div>
    );
  }

  const currentUser = loggedInUsername ? state.users[loggedInUsername] : null;

  if (!currentUser) {
    return <LoginScreen users={state.users} onLogin={(u) => { localStorage.setItem("der-crm-user", u); setLoggedInUsername(u); }} />;
  }

  return (
    <div className="rcrm" style={{ height: "100vh", overflow: "hidden", background: "var(--bg)", display: "flex", flexDirection: "column" }}>
      <style>{THEME}</style>
      {currentUser.role === "admin"
        ? <AdminApp state={state} setState={persist} onLogout={() => { localStorage.removeItem("der-crm-user"); setLoggedInUsername(null); }} saving={saving} />
        : <UserApp state={state} setState={persist} user={currentUser.username} onLogout={() => { localStorage.removeItem("der-crm-user"); setLoggedInUsername(null); }} saving={saving} />}
    </div>
  );
}

/* ---------------------------------------------------------
   LOGIN
--------------------------------------------------------- */
function LoginScreen({ users, onLogin }) {
  const [username, setUsername] = useState("");
  const [passkey, setPasskey] = useState("");
  const [showPasskey, setShowPasskey] = useState(false);
  const [error, setError] = useState("");

  const submit = () => {
    const u = (users || {})[username.trim().toLowerCase()];
    if (u && u.passkey === passkey) {
      setError("");
      onLogin(u.username);
    } else {
      setError("Incorrect username or passkey.");
    }
  };
  const onKeyDown = (e) => { if (e.key === "Enter") submit(); };

  return (
    <div className="rcrm" style={{ minHeight: 640, background: "var(--bg)", display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center", padding: "40px 20px" }}>
      <style>{THEME}</style>
      <div style={{ width: 320 }}>
        <hr className="rule" style={{ margin: "0 0 14px" }} />
        <div style={{ textAlign: "center", marginBottom: 40 }}>
          <span className="eyebrow">Dubai Estate Radar</span>
        </div>

        <div className="disp" style={{ fontSize: 42, fontWeight: 800, lineHeight: 1.02, marginBottom: 10, textAlign: "center" }}>
          Sign in.
        </div>
        <div style={{ width: 56, height: 2, background: "var(--accent)", margin: "0 auto 18px" }} />
        <div style={{ color: "var(--text-dim)", fontSize: 13.5, lineHeight: 1.5, marginBottom: 28, textAlign: "center" }}>
          Prospecting and call tracking.
        </div>

        <div style={{ display: "flex", flexDirection: "column", gap: 14 }}>
          <div>
            <div className="eyebrow" style={{ marginBottom: 6 }}>Username</div>
            <input value={username} onChange={e => { setUsername(e.target.value); setError(""); }} onKeyDown={onKeyDown} autoFocus
              className="tap" style={{ width: "100%", background: "transparent", border: "none", borderBottom: "1px solid var(--line)", color: "var(--text)", padding: "8px 0", fontSize: 15, outline: "none" }} />
          </div>
         <div>
            <div className="eyebrow" style={{ marginBottom: 6 }}>Passkey</div>
            <div style={{ position: "relative" }}>
              <input type={showPasskey ? "text" : "password"} value={passkey} onChange={e => { setPasskey(e.target.value); setError(""); }} onKeyDown={onKeyDown}
                className="tap" style={{ width: "100%", background: "transparent", border: "none", borderBottom: "1px solid var(--line)", color: "var(--text)", padding: "8px 24px 8px 0", fontSize: 15, outline: "none" }} />
              <button type="button" onClick={() => setShowPasskey(v => !v)} tabIndex={-1}
                style={{ position: "absolute", right: 0, top: "50%", transform: "translateY(-50%)", background: "transparent", border: "none", color: "var(--text-dim)", cursor: "pointer", display: "flex", alignItems: "center", padding: 4 }}>
                {showPasskey ? <EyeOff size={16} /> : <Eye size={16} />}
              </button>
            </div>
          </div>
          {error && (
            <div style={{ borderLeft: "2px solid var(--accent)", background: "var(--accent-dim)", color: "var(--text)", padding: "10px 12px", fontSize: 12.5 }}>
              {error}
            </div>
          )}
          <button type="button" onClick={submit} className="tap pulse-btn login-pulse"

            style={{ marginTop: 8, background: "var(--accent)", color: "#fff", border: "1px solid transparent", borderRadius: 5, padding: "13px", fontSize: 12, fontWeight: 600, textTransform: "uppercase", letterSpacing: "0.14em" }}>
            Log in
          </button>
        </div>
      </div>
    </div>
  );
}

/* ---------------------------------------------------------
   TOP BAR (shared)
--------------------------------------------------------- */
function TopBar({ title, subtitle, right, onLogout, saving, extraMenuItems }) {
  const [settingsOpen, setSettingsOpen] = useState(false);
  const wrapperRef = useRef(null);

  // close when clicking outside the settings cluster
  useEffect(() => {
    if (!settingsOpen) return;
    const onDoc = (e) => { if (wrapperRef.current && !wrapperRef.current.contains(e.target)) setSettingsOpen(false); };
    document.addEventListener("mousedown", onDoc);
    return () => document.removeEventListener("mousedown", onDoc);
  }, [settingsOpen]);

  const items = [
    ...(extraMenuItems || []),
    { label: "Log out", icon: LogOut, onClick: onLogout },
  ];

  return (
    <div style={{ background: "var(--panel)", borderBottom: "1px solid var(--line)" }}>
      <hr className="rule" style={{ margin: 0 }} />
      <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", padding: "13px 20px", gap: 12, flexWrap: "wrap" }}>
        <div>
          <div className="disp" style={{ fontSize: 20, fontWeight: 800, lineHeight: 1.1 }}>{title}</div>
          {subtitle && <div className="eyebrow" style={{ marginTop: 3 }}>{subtitle}</div>}
        </div>
        <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
          {saving && <span className="mono" style={{ fontSize: 10.5, color: "var(--text-faint)" }}>saving…</span>}
          {right}
          <div ref={wrapperRef} style={{ display: "flex", alignItems: "center", gap: 6, position: "relative" }}>
            {/* items slide out to the LEFT of the gear when open */}
            <div style={{
              display: "flex", alignItems: "center", gap: 6, overflow: "hidden",
              maxWidth: settingsOpen ? 500 : 0, opacity: settingsOpen ? 1 : 0,
              transform: settingsOpen ? "translateX(0)" : "translateX(12px)",
              transition: "max-width 0.25s ease, opacity 0.2s ease, transform 0.25s ease",
              pointerEvents: settingsOpen ? "auto" : "none",
            }}>
              {items.map((it, i) => (
                <button key={i} onClick={() => { it.onClick(); setSettingsOpen(false); }} className={`tap${it.label === "Log out" ? " pulse-btn logout-pulse" : ""}`}
                  style={{
                    display: "flex", alignItems: "center", gap: 6, whiteSpace: "nowrap",
                    background: "transparent", border: "1px solid var(--line)", color: "var(--text-dim)",
                    padding: "8px 12px", borderRadius: 6, fontSize: 11, textTransform: "uppercase", letterSpacing: "0.1em"
                  }}>
                  <it.icon size={13} /> {it.label}
                </button>
              ))}
            </div>
            <button onClick={() => setSettingsOpen(v => !v)} className="tap" aria-label="Settings"
              style={{
                display: "flex", alignItems: "center", justifyContent: "center",
                width: 34, height: 34, background: "transparent",
                border: "none", color: "var(--text-dim)",
                transition: "transform 0.25s ease, color 0.15s",
                transform: settingsOpen ? "rotate(90deg)" : "rotate(0)",
              }}>
              <Settings size={14} />
            </button>
          </div>
        </div>
      </div>
    </div>
  );
}

/* ---------------------------------------------------------
   ADMIN APP
--------------------------------------------------------- */
function AdminApp({ state, setState, onLogout, saving }) {
  const [tab, setTab] = useState("buildings");
  const [confirmingWipe, setConfirmingWipe] = useState(false);
  const buildings = Object.values(state.buildings);

  const wipeAll = () => setState({ buildings: {}, units: {}, assignments: {}, crm: {}, callLog: [], users: state.users });

  return (
    <>
      <TopBar title="Admin Portal" subtitle="Upload data, assign buildings and units to agents" onLogout={onLogout} saving={saving}
        extraMenuItems={[{ label: "Wipe data", icon: Trash2, onClick: () => setConfirmingWipe(true) }]} />
      {confirmingWipe && (
        <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.7)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 1000 }}>
          <div style={{ background: "var(--panel)", border: "1px solid var(--line)", borderRadius: 10, padding: 32, maxWidth: 360, textAlign: "center" }}>
            <Skull size={40} style={{ color: "var(--red)", marginBottom: 14 }} />
            <div className="disp" style={{ fontSize: 18, fontWeight: 800, marginBottom: 22 }}>SURE YOU WANT TO WIPE ALL DATA?</div>
            <div style={{ display: "flex", gap: 10, justifyContent: "center" }}>
              <button onClick={() => setConfirmingWipe(false)} className="tap pulse-btn logout-pulse"
                style={{ background: "transparent", border: "1px solid var(--line)", color: "var(--text-dim)", padding: "11px 22px", borderRadius: 6, fontSize: 11.5, fontWeight: 600, textTransform: "uppercase", letterSpacing: "0.1em" }}>
                No
              </button>
              <button onClick={() => { wipeAll(); setConfirmingWipe(false); }} className="tap pulse-btn wipe-pulse"
                style={{ background: "transparent", border: "1px solid var(--line)", color: "var(--text-dim)", padding: "11px 22px", borderRadius: 6, fontSize: 11.5, fontWeight: 600, textTransform: "uppercase", letterSpacing: "0.1em" }}>
                Wipe
              </button>
            </div>
          </div>
        </div>
      )}
      <div style={{ display: "flex", borderBottom: "1px solid var(--line)", background: "var(--panel)" }}>
{[["buildings", "Buildings & Upload"], ["assign", "Assignments"], ["users", "Users"], ["tracking", "Tracking"], ["explorer", "Explorer"], ["portfolio", "Portfolio Desk"]].map(([id, label]) => (          <button key={id} onClick={() => setTab(id)} className="tap"
            style={{
              padding: "13px 18px", background: "transparent", border: "none",
              borderBottom: tab === id ? "1px solid var(--accent)" : "1px solid transparent",
              color: tab === id ? "var(--text)" : "var(--text-dim)", fontSize: 10.5, fontWeight: 600,
              textTransform: "uppercase", letterSpacing: "0.12em"
            }}>{label}</button>
        ))}
      </div>
      <div style={{ flex: 1, overflow: "auto", padding: 20 }}>
        {tab === "buildings" && <BuildingsUpload state={state} setState={setState} />}
        {tab === "assign" && <AssignmentPanel state={state} setState={setState} />}
        {tab === "users" && <UsersPanel state={state} setState={setState} />}
        {tab === "tracking" && <TrackingPanel state={state} />}
        {tab === "explorer" && <BuildingExplorer state={state} role="admin" />}
        {tab === "portfolio" && <PortfolioDashboard state={state} />}
      </div>
    </>
  );
}

function TrackingPanel({ state }) {
  const [mode, setMode] = useState("date");
  const [singleDate, setSingleDate] = useState(() => new Date().toISOString().slice(0, 10));
  const [period, setPeriod] = useState("week");
  const [rangeStart, setRangeStart] = useState("");
  const [rangeEnd, setRangeEnd] = useState("");
  const [rangeApplied, setRangeApplied] = useState(null);

  const log = state.callLog || [];

  const singleDateCalls = useMemo(() => {
    if (!singleDate) return [];
    return log.filter(c => new Date(c.timestamp).toDateString() === new Date(singleDate + "T00:00:00").toDateString());
  }, [log, singleDate]);

  const periodCalls = useMemo(() => {
    const now = new Date();
    const startOfToday = new Date(now.getFullYear(), now.getMonth(), now.getDate());
    let start = null, end = null;
    if (period === "week") {
      start = new Date(startOfToday); start.setDate(start.getDate() - start.getDay());
      end = new Date(start); end.setDate(end.getDate() + 7);
    } else if (period === "month") {
      start = new Date(now.getFullYear(), now.getMonth(), 1);
      end = new Date(now.getFullYear(), now.getMonth() + 1, 1);
    } else if (period === "lastMonth") {
      start = new Date(now.getFullYear(), now.getMonth() - 1, 1);
      end = new Date(now.getFullYear(), now.getMonth(), 1);
    }
    return log.filter(c => {
      const t = new Date(c.timestamp);
      return (!start || t >= start) && (!end || t < end);
    });
  }, [log, period]);

  const rangeCalls = useMemo(() => {
    if (!rangeApplied) return [];
    const start = new Date(rangeApplied.start + "T00:00:00");
    const end = new Date(rangeApplied.end + "T23:59:59");
    return log.filter(c => { const t = new Date(c.timestamp); return t >= start && t <= end; });
  }, [log, rangeApplied]);

  const activeCalls = mode === "date" ? singleDateCalls : mode === "period" ? periodCalls : rangeCalls;
  const periodLabels = { week: "This week", month: "This month", lastMonth: "Last month", allTime: "Since the beginning" };
  const activeLabel = mode === "date"
    ? new Date(singleDate + "T00:00:00").toLocaleDateString(undefined, { month: "long", day: "numeric", year: "numeric" })
    : mode === "period"
    ? periodLabels[period]
    : rangeApplied ? `${rangeApplied.start} to ${rangeApplied.end}` : "Custom range";

  return (
    <div style={{ display: "flex", gap: 40, alignItems: "flex-start" }}>
      <div style={{ flex: 1, minWidth: 0 }}>
        <div style={{ fontSize: 12, fontWeight: 600, color: "var(--accent)", textTransform: "uppercase", letterSpacing: "0.1em", marginBottom: 16 }}>{activeLabel}</div>
       <StatsBlock calls={activeCalls} users={state.users} state={state} />
      </div>

      <div style={{ width: 260, flexShrink: 0, display: "flex", flexDirection: "column", gap: 14 }}>
        <div onClick={() => setMode("date")} style={{ padding: 14, borderRadius: 8, border: mode === "date" ? "1px solid var(--accent)" : "1px solid var(--line)", background: "var(--panel)", cursor: "pointer" }}>
          <div style={{ fontSize: 11, fontWeight: 600, color: "var(--text-dim)", textTransform: "uppercase", letterSpacing: "0.1em", marginBottom: 8 }}>Specific date</div>
          <CustomDatePicker value={singleDate} onChange={(v) => { setSingleDate(v); setMode("date"); }} />
        </div>

        <div onClick={() => setMode("period")} style={{ padding: 14, borderRadius: 8, border: mode === "period" ? "1px solid var(--accent)" : "1px solid var(--line)", background: "var(--panel)", cursor: "pointer" }}>
          <div style={{ fontSize: 11, fontWeight: 600, color: "var(--text-dim)", textTransform: "uppercase", letterSpacing: "0.1em", marginBottom: 8 }}>Period</div>
          <select value={period} onChange={e => { setPeriod(e.target.value); setMode("period"); }} className="tap"
            style={{ width: "100%", background: "var(--bg)", border: "1px solid var(--line)", borderRadius: 6, padding: "8px 10px", color: "var(--text)", fontSize: 13 }}>
            <option value="week">This week</option>
            <option value="month">This month</option>
            <option value="lastMonth">Last month</option>
            <option value="allTime">Since the beginning</option>
          </select>
        </div>

        <div style={{ padding: 14, borderRadius: 8, border: mode === "range" ? "1px solid var(--accent)" : "1px solid var(--line)", background: "var(--panel)" }}>
          <div style={{ fontSize: 11, fontWeight: 600, color: "var(--text-dim)", textTransform: "uppercase", letterSpacing: "0.1em", marginBottom: 8 }}>Custom range</div>
          <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
            <CustomDatePicker value={rangeStart} onChange={setRangeStart} />
            <CustomDatePicker value={rangeEnd} onChange={setRangeEnd} />
            <button onClick={() => { if (rangeStart && rangeEnd) { setRangeApplied({ start: rangeStart, end: rangeEnd }); setMode("range"); } }} className="tap"
              style={{ background: "var(--accent)", border: "none", color: "#fff", padding: "8px 10px", borderRadius: 6, fontSize: 11.5, fontWeight: 600, textTransform: "uppercase", letterSpacing: "0.1em" }}>
              View
            </button>
          </div>
        </div>
      </div>
    </div>
  );
}

function StatsBlock({ calls, users, state }) {
  const [expandedAgent, setExpandedAgent] = useState(null);
  const byAgent = {};
  calls.forEach(c => { byAgent[c.agent] = byAgent[c.agent] || []; byAgent[c.agent].push(c); });
  const agentNames = Object.keys(byAgent).sort();

  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 24 }}>
      <div>
        <div style={{ fontSize: 12, color: "var(--text-dim)", textTransform: "uppercase", letterSpacing: "0.1em", marginBottom: 6 }}>Total calls</div>
        <div style={{ fontSize: 64, fontWeight: 800, lineHeight: 1 }}>{calls.length}</div>
      </div>
      {agentNames.length === 0 ? (
        <div style={{ color: "var(--text-dim)", fontSize: 13 }}>No calls logged in this range.</div>
      ) : (
        <div style={{ display: "flex", flexDirection: "column", gap: 6 }}>
          {agentNames.map(agent => {
            const isOpen = expandedAgent === agent;
            const assignedUnits = Object.values(state.units).filter(u => effectiveAssignedUsers(u, state.assignments).includes(agent));
            const completedUnits = assignedUnits.filter(u => state.crm[u.key] && state.crm[u.key].lastOutcome);
            const pct = assignedUnits.length ? Math.round((completedUnits.length / assignedUnits.length) * 100) : 0;
            return (
              <div key={agent} style={{ background: "var(--panel)", border: "1px solid var(--line)", borderRadius: 6, overflow: "hidden" }}>
               <button onClick={() => setExpandedAgent(isOpen ? null : agent)} className="tap"
                  style={{ display: "grid", gridTemplateColumns: "100px 1fr 70px", alignItems: "center", width: "100%", background: "transparent", border: "none", padding: "12px 16px", color: "var(--text)", fontSize: 14, textAlign: "left", cursor: "pointer", gap: 12 }}>
                  <span style={{ whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis" }}>{users[agent]?.displayName || agent}</span>
                  <div style={{ justifySelf: "center", display: "flex", alignItems: "center", gap: 10 }}>
                    <div style={{ width: 190, height: 9, background: "var(--line)", borderRadius: 5, overflow: "hidden", border: "1px solid #000000" }} title={`${completedUnits.length} of ${assignedUnits.length} assigned units completed overall`}>
                      <div style={{ width: `${pct}%`, height: "100%", background: "var(--accent)" }} />
                    </div>
                    <span style={{ fontSize: 11, color: "var(--text-faint)", whiteSpace: "nowrap" }}>{pct}%</span>
                  </div>
                  <span style={{ fontWeight: 700, whiteSpace: "nowrap", textAlign: "right" }}>{byAgent[agent].length} calls</span>
                </button>
                {isOpen && (
                  <div style={{ display: "flex", flexDirection: "column", gap: 6, padding: "0 16px 14px", borderTop: "1px solid var(--line)" }}>
                    {["yes", "callMeBack", "no", "noAnswer"].map(outcome => (
                      <div key={outcome} style={{ display: "flex", justifyContent: "space-between", padding: "8px 4px", fontSize: 13, borderBottom: "1px solid var(--line)" }}>
                        <span style={{ color: "var(--text-dim)" }}>{OUTCOME_LABEL[outcome]}</span>
                        <span>{byAgent[agent].filter(c => c.outcome === outcome).length}</span>
                      </div>
                    ))}
                  </div>
                )}
              </div>
            );
          })}
        </div>
      )}
    </div>
  );
}

const MONTH_NAMES = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
const WHEEL_YEARS = [2026, 2027, 2028, 2029, 2030];

function CustomDatePicker({ value, onChange }) {
  const [open, setOpen] = useState(false);
  const [pickerMode, setPickerMode] = useState("calendar");
  const initial = value ? new Date(value + "T00:00:00") : new Date();
  const [viewMonth, setViewMonth] = useState(initial.getMonth());
  const [viewYear, setViewYear] = useState(initial.getFullYear());
  const [wheelMonth, setWheelMonth] = useState(initial.getMonth());
  const [wheelYear, setWheelYear] = useState(Math.min(Math.max(initial.getFullYear(), 2026), 2030));
  const ref = useRef(null);

  useEffect(() => {
    const onDocClick = (e) => { if (ref.current && !ref.current.contains(e.target)) { setOpen(false); setPickerMode("calendar"); } };
    document.addEventListener("mousedown", onDocClick);
    return () => document.removeEventListener("mousedown", onDocClick);
  }, []);

  const daysInMonth = (m, y) => new Date(y, m + 1, 0).getDate();
  const firstWeekday = (m, y) => new Date(y, m, 1).getDay();

  const selectDay = (day) => {
    const d = new Date(viewYear, viewMonth, day);
    onChange(d.toISOString().slice(0, 10));
    setOpen(false);
  };

  const applyWheel = () => { setViewMonth(wheelMonth); setViewYear(wheelYear); setPickerMode("calendar"); };

  const cells = [];
  const totalDays = daysInMonth(viewMonth, viewYear);
  const startOffset = firstWeekday(viewMonth, viewYear);
  for (let i = 0; i < startOffset; i++) cells.push(null);
  for (let d = 1; d <= totalDays; d++) cells.push(d);

  const displayValue = value ? new Date(value + "T00:00:00").toLocaleDateString(undefined, { month: "short", day: "numeric", year: "numeric" }) : "Select date";

  return (
    <div style={{ position: "relative" }} ref={ref}>
      <button type="button" onClick={() => setOpen(o => !o)} className="tap"
        style={{ display: "flex", alignItems: "center", justifyContent: "space-between", gap: 8, width: "100%", background: "var(--bg)", border: "1px solid var(--line)", borderRadius: 6, padding: "8px 10px", color: value ? "var(--text)" : "var(--text-dim)", fontSize: 13 }}>
        {displayValue}
        <Calendar size={14} style={{ color: "var(--text-dim)" }} />
      </button>

      {open && (
        <div style={{ position: "absolute", top: "calc(100% + 6px)", left: 0, zIndex: 50, width: 240, background: "var(--panel)", border: "1px solid var(--line)", borderRadius: 8, padding: 12, boxShadow: "0 8px 24px rgba(0,0,0,0.4)" }}>
          {pickerMode === "calendar" ? (
            <>
              <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 10 }}>
                <button type="button" onClick={() => { const d = new Date(viewYear, viewMonth - 1, 1); setViewMonth(d.getMonth()); setViewYear(d.getFullYear()); }}
                  style={{ background: "transparent", border: "none", color: "var(--text-dim)", cursor: "pointer", padding: 4, display: "flex" }}>
                  <ChevronRight size={14} style={{ transform: "rotate(180deg)" }} />
                </button>
                <button type="button" onClick={() => { setWheelMonth(viewMonth); setWheelYear(Math.min(Math.max(viewYear, 2026), 2030)); setPickerMode("wheel"); }}
                  style={{ background: "transparent", border: "none", color: "var(--text)", fontSize: 12.5, fontWeight: 700, cursor: "pointer" }}>
                  {MONTH_NAMES[viewMonth]} {viewYear}
                </button>
                <button type="button" onClick={() => { const d = new Date(viewYear, viewMonth + 1, 1); setViewMonth(d.getMonth()); setViewYear(d.getFullYear()); }}
                  style={{ background: "transparent", border: "none", color: "var(--text-dim)", cursor: "pointer", padding: 4, display: "flex" }}>
                  <ChevronRight size={14} />
                </button>
              </div>
              <div style={{ display: "grid", gridTemplateColumns: "repeat(7, 1fr)", gap: 2, marginBottom: 4 }}>
                {["S", "M", "T", "W", "T", "F", "S"].map((d, i) => (
                  <div key={i} style={{ fontSize: 9.5, color: "var(--text-faint)", textAlign: "center" }}>{d}</div>
                ))}
              </div>
              <div style={{ display: "grid", gridTemplateColumns: "repeat(7, 1fr)", gap: 2 }}>
                {cells.map((day, i) => {
                  const isSelected = value && day && new Date(value + "T00:00:00").toDateString() === new Date(viewYear, viewMonth, day).toDateString();
                  return (
                    <button key={i} type="button" disabled={!day} onClick={() => day && selectDay(day)}
                      style={{ aspectRatio: "1", background: isSelected ? "var(--accent)" : "transparent", border: "none", borderRadius: 4, color: !day ? "transparent" : isSelected ? "#000" : "var(--text)", fontSize: 11.5, cursor: day ? "pointer" : "default" }}>
                      {day || ""}
                    </button>
                  );
                })}
              </div>
            </>
          ) : (
            <>
              <div style={{ display: "flex", gap: 8, marginBottom: 14 }}>
                <div style={{ flex: 1, maxHeight: 140, overflow: "auto", border: "1px solid var(--line)", borderRadius: 6 }}>
                  {MONTH_NAMES.map((m, i) => (
                    <div key={m} onClick={() => setWheelMonth(i)}
                      style={{ padding: "7px 6px", textAlign: "center", fontSize: 11.5, cursor: "pointer", background: wheelMonth === i ? "var(--panel-2)" : "transparent", color: wheelMonth === i ? "var(--text)" : "var(--text-dim)", fontWeight: wheelMonth === i ? 700 : 400 }}>
                      {m}
                    </div>
                  ))}
                </div>
                <div style={{ flex: 1, maxHeight: 140, overflow: "auto", border: "1px solid var(--line)", borderRadius: 6 }}>
                  {WHEEL_YEARS.map(y => (
                    <div key={y} onClick={() => setWheelYear(y)}
                      style={{ padding: "7px 6px", textAlign: "center", fontSize: 11.5, cursor: "pointer", background: wheelYear === y ? "var(--panel-2)" : "transparent", color: wheelYear === y ? "var(--text)" : "var(--text-dim)", fontWeight: wheelYear === y ? 700 : 400 }}>
                      {y}
                    </div>
                  ))}
                </div>
              </div>
              <div style={{ display: "flex", justifyContent: "flex-end" }}>
                <button type="button" onClick={applyWheel} className="tap pulse-btn login-pulse"
                  style={{ display: "flex", alignItems: "center", justifyContent: "center", width: 32, height: 32, background: "transparent", border: "1px solid var(--line)", borderRadius: 6, color: "var(--text-dim)", cursor: "pointer" }}>
                  <Check size={15} />
                </button>
              </div>
            </>
          )}
        </div>
      )}
    </div>
  );
}

function BuildingsUpload({ state, setState }) {  const [newBuildingName, setNewBuildingName] = useState("");
  const [targetBuilding, setTargetBuilding] = useState("");
  const [confirmWipeBuilding, setConfirmWipeBuilding] = useState(null);
  const [unitFile, setUnitFile] = useState(null);
  const [ownerFile, setOwnerFile] = useState(null);
  const [unitHeaders, setUnitHeaders] = useState([]);
  const [ownerHeaders, setOwnerHeaders] = useState([]);
  const [unitMap, setUnitMap] = useState({});
  const [ownerMap, setOwnerMap] = useState({});
  const [log, setLog] = useState("");

  const buildings = Object.values(state.buildings);

      const wipeBuilding = (buildingId) => {
    const newUnits = {};
    Object.entries(state.units).forEach(([k, u]) => { if (u.buildingId !== buildingId) newUnits[k] = u; });
    const newCrm = {};
    Object.entries(state.crm).forEach(([k, c]) => { if (!k.startsWith(`${buildingId}::`)) newCrm[k] = c; });
    const newAssignments = { ...state.assignments };
    delete newAssignments[buildingId];
    const newBuildings = { ...state.buildings };
    delete newBuildings[buildingId];
    setState({ ...state, buildings: newBuildings, units: newUnits, crm: newCrm, assignments: newAssignments });
    if (targetBuilding === buildingId) setTargetBuilding("");
  };


  const createBuilding = () => {
    if (!newBuildingName.trim()) return;
    const id = newBuildingName.trim().toLowerCase().replace(/\s+/g, "-") + "-" + Math.random().toString(36).slice(2, 6);
    const next = { ...state, buildings: { ...state.buildings, [id]: { id, name: newBuildingName.trim() } } };
    setState(next);
    setNewBuildingName("");
    setTargetBuilding(id);
  };

  // Strips a UTF-8 BOM if present (common on CSVs exported from Excel) and
  // trims whitespace — a BOM on the first header is the single most common
  // reason a "Unit ID" column silently fails to match.
  const cleanHeader = (h) => (h || "").replace(/^\uFEFF/, "").trim();

  // Some spreadsheets don't have column titles in row 1 — a title/banner
  // row, a notes line, or a blank row above the real headers is common in
  // hand-built sheets. Blindly treating row 1 as headers in that case
  // shifts every column by one row and produces exactly the kind of
  // garbage values ("rent" reading as a unit number) that showed up here.
  // Instead: scan the first several rows of each page for the one that
  // actually looks like a header row — the row containing something that
  // matches the Unit ID candidate list — and treat that as row 1.
  const ALL_CANDIDATES_FLAT = Object.values(CANDIDATES).flat().map(normText);
  const UNITID_CANDS = (CANDIDATES.unitId || []).map(normText);
  const scoreHeaderRow = (row) => {
    let unitIdHits = 0, anyHits = 0;
    (row || []).forEach(cell => {
      const nc = normText(cell);
      if (!nc) return;
      if (UNITID_CANDS.some(c => nc === c || nc.includes(c) || c.includes(nc))) unitIdHits++;
      if (ALL_CANDIDATES_FLAT.some(c => nc === c)) anyHits++;
    });
    return unitIdHits * 100 + anyHits;
  };
  const detectHeaderRowIndex = (aoa) => {
    const scan = Math.min(aoa.length, 10);
    let best = 0, bestScore = 0;
    for (let i = 0; i < scan; i++) {
      const score = scoreHeaderRow(aoa[i]);
      if (score > bestScore) { bestScore = score; best = i; }
    }
    return best; // falls back to row 0 if nothing scores — same as before
  };
  const rowsFromAoa = (aoa) => {
    if (!aoa || !aoa.length) return { headers: [], rows: [] };
    const hIdx = detectHeaderRowIndex(aoa);
    const headerRow = (aoa[hIdx] || []).map(h => cleanHeader(String(h ?? "")));
    const rows = [];
    for (let i = hIdx + 1; i < aoa.length; i++) {
      const r = aoa[i] || [];
      if (r.every(c => c === "" || c == null)) continue;
      const obj = {};
      headerRow.forEach((h, ci) => { if (h) obj[h] = r[ci] ?? ""; });
      if (Object.keys(obj).length) rows.push(obj);
    }
    return { headers: headerRow.filter(Boolean), rows };
  };

  // Works for CSV, TSV, XLSX, and XLS. For a real spreadsheet workbook, every
  // sheet/page is read and its rows concatenated together — a workbook with
  // Building A on one tab and Building B on another (or the same data split
  // across pages some other way) all comes in, not just the first tab.
  // Each row keeps the header names from whichever sheet it came from, and
  // the mapping dropdown shows the union of every header seen anywhere in
  // the file, so mapping still works even if a name only appears on page 2.
  const parseAnyFile = (file) => new Promise((resolve) => {
    const name = file.name.toLowerCase();
    const isCsv = name.endsWith(".csv") || name.endsWith(".tsv") || name.endsWith(".txt");
    if (isCsv) {
      Papa.parse(file, {
        header: false, skipEmptyLines: true,
        complete: (res) => resolve(rowsFromAoa(res.data || [])),
      });
      return;
    }
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const wb = XLSX.read(e.target.result, { type: "array", cellDates: true });
        let rows = [];
        const headerSet = [];
        const seen = new Set();
        wb.SheetNames.forEach(sheetName => {
          const ws = wb.Sheets[sheetName];
          const aoa = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "", raw: false });
          const { headers, rows: sheetRows } = rowsFromAoa(aoa);
          headers.forEach(h => { if (!seen.has(h)) { seen.add(h); headerSet.push(h); } });
          rows.push(...sheetRows);
        });
        resolve({ rows, headers: headerSet, sheetCount: wb.SheetNames.length });
      } catch (err) {
        resolve({ rows: [], headers: [], error: String(err) });
      }
    };
    reader.readAsArrayBuffer(file);
  });

  // unitId is guessed first and "claims" its column; every other field is
  // then guessed only from the remaining, unclaimed columns. Without this,
  // two fields with weak/fuzzy matches could both resolve to the same
  // column (e.g. rent and lease date both reading the Unit ID column) —
  // this makes that structurally impossible.
  const onSelectFile = async (file, setFile, setHeaders, setMap, fieldList) => {
    setFile(file);
    if (!file) { setHeaders([]); setMap({}); return; }
    const res = await parseAnyFile(file);
    setHeaders(res.headers);
    const guessed = {};
    const claimed = new Set();
    const ordered = ["unitId", ...fieldList.filter(f => f !== "unitId")];
    ordered.forEach(f => {
      if (!fieldList.includes(f)) return;
      const available = res.headers.filter(h => !claimed.has(h));
      const hit = guessColumn(available, f);
      guessed[f] = hit;
      if (hit) claimed.add(hit);
    });
    setMap(guessed);
  };

  const getMapped = (row, map, field) => {
    const col = map[field];
    if (!col) return "";
    const v = row[col];
    if (v == null) return "";
    if (v instanceof Date) return v.toISOString().slice(0, 10);
    return String(v).trim();
  };

  const runImport = async () => {
    if (!targetBuilding) { setLog("Pick or create a building first."); return; }
    if (!unitFile) { setLog("Unit data file is required at minimum (owner data is optional and can be added later)."); return; }
    if (!unitMap.unitId) { setLog(`Couldn't find anything in the unit file that looks like a Unit ID / Unit Number / Property No column. Rename that column to one of those, or open "Unit data" below and pick it manually.`); return; }

    const unitRes = await parseAnyFile(unitFile);
    const ownerRes = ownerFile ? await parseAnyFile(ownerFile) : null;
    const unitRows = unitRes.rows || [];
    const ownerRows = ownerRes?.rows || [];

    const unitIdSet = new Set(
      unitRows.map(r => normalizeId(getMapped(r, unitMap, "unitId"))).filter(Boolean)
    );

    const ownerByUnit = {};
    let ownerMatched = 0; const ownerUnmatchedSample = [];
    if (ownerFile && !ownerMap.unitId) {
      setLog(`Owner file uploaded, but couldn't find a Unit ID column in it to join against. Owner data won't attach to any unit. Open "Owner data" below and pick it manually, or remove the file.`);
      return;
    }
    ownerRows.forEach(r => {
      const uidRaw = getMapped(r, ownerMap, "unitId");
      if (!uidRaw) return;
      const nid = normalizeId(uidRaw);
      ownerByUnit[nid] = {
        ownerName: getMapped(r, ownerMap, "ownerName"),
        ownerContact: getMapped(r, ownerMap, "ownerContact"),
        secondaryContact: getMapped(r, ownerMap, "secondaryContact"),
        email: getMapped(r, ownerMap, "email"),
      };
      if (unitIdSet.has(nid)) ownerMatched++;
      else if (ownerUnmatchedSample.length < 5) ownerUnmatchedSample.push(uidRaw);
    });

    const units = { ...state.units };
    let count = 0, skipped = 0;
    unitRows.forEach(r => {
      const uid = getMapped(r, unitMap, "unitId");
      if (!uid) { skipped++; return; }
      const nid = normalizeId(uid);
      const key = `${targetBuilding}::${uid}`;
      const existing = units[key] || {};
      const owner = ownerByUnit[nid] || {};
      // If the same Unit ID appears more than once (duplicate rows, or the
      // same unit split across two sheet pages), fill in whatever's blank
      // instead of a later row silently overwriting fields the earlier one
      // already had.
      const pick = (val, prevField) => val || existing[prevField] || "";
      units[key] = {
        ...existing,
        key, unitId: uid, unitType: pick(getMapped(r, unitMap, "unitType"), "unitType"), buildingId: targetBuilding,
        floor: pick(getMapped(r, unitMap, "floor"), "floor"),
        carpetArea: pick(getMapped(r, unitMap, "carpetArea"), "carpetArea"),
        bedrooms: pick(getMapped(r, unitMap, "bedrooms"), "bedrooms"),
        bathrooms: pick(getMapped(r, unitMap, "bathrooms"), "bathrooms"),
        ownerName: owner.ownerName || existing.ownerName || "",
        ownerContact: owner.ownerContact || existing.ownerContact || "",
        secondaryContact: owner.secondaryContact || existing.secondaryContact || "",
        email: owner.email || existing.email || "",
        ownershipStart: pick(getMapped(r, unitMap, "ownershipStart"), "ownershipStart") || existing.ownershipStart || "",
        purchasePrice: pick(getMapped(r, unitMap, "purchasePrice"), "purchasePrice") || existing.purchasePrice || "",
        currentRent: pick(getMapped(r, unitMap, "currentRent"), "currentRent"),
        currentLeaseStart: pick(getMapped(r, unitMap, "currentLeaseStart"), "currentLeaseStart"),
        currentLeaseEnd: pick(getMapped(r, unitMap, "currentLeaseEnd"), "currentLeaseEnd"),
        status: normStatus(getMapped(r, unitMap, "status")) || existing.status || "",
        statusUpdated: new Date().toISOString(),
        assignedTo: existing.assignedTo ?? null,
      };
      count++;
    });

    if (count === 0) {
      setLog(`0 units imported. ${skipped} row(s) had nothing in the detected Unit ID column ("${unitMap.unitId}"). Check "Unit data" below and confirm that's really the right column.`);
      return;
    }

    let diag = "";
    if (unitRes.sheetCount > 1) diag += ` Read across ${unitRes.sheetCount} pages in the unit file.`;
    if (ownerRes?.sheetCount > 1) diag += ` Read across ${ownerRes.sheetCount} pages in the owner file.`;
    if (ownerFile) {
      diag += ` Owner rows: ${ownerMatched}/${ownerRows.length} matched a unit.`;
      if (ownerMatched < ownerRows.length && ownerUnmatchedSample.length) diag += ` Unmatched owner-sheet IDs looked like: ${ownerUnmatchedSample.join(", ")}. Compare the formatting against your unit sheet's Unit ID column.`;
    }
    const mappedSummary = Object.entries(unitMap).filter(([, v]) => v).map(([f, v]) => `${FIELD_LABEL[f].replace(" *", "")} ← "${v}"`).join(", ");

    setState({ ...state, units });
    setLog(`Imported ${count} unit${count === 1 ? "" : "s"}.${skipped ? ` Skipped ${skipped} row(s) with no value in the Unit ID column.` : ""}${diag}${!ownerFile ? " No owner file uploaded, so owner name and contact will show as unavailable until added." : ""} Detected: ${mappedSummary}.`);
    setUnitFile(null); setOwnerFile(null);
    setUnitHeaders([]); setOwnerHeaders([]);
    setUnitMap({}); setOwnerMap({});
  };

  const FIELD_LABEL = {
    unitId: "Unit ID / Unit Number *", unitType: "Unit Type", floor: "Floor", carpetArea: "Carpet Area",
    bedrooms: "Bedrooms", bathrooms: "Bathrooms", currentRent: "Current Rent",
    currentLeaseStart: "Current Lease Start", currentLeaseEnd: "Current Lease End",
    status: "Status (occupied/vacant)",
    ownerName: "Owner Name", ownerContact: "Owner Contact", secondaryContact: "Secondary Number", email: "Email", ownershipStart: "Ownership Start", purchasePrice: "Purchase Price",
  };

  // Fully automatic by default — the file is read, columns are detected,
  // done. The mapping grid only appears if you open "Adjust detected
  // columns", and even then it's pre-filled with the best guess, not blank.
  const UploadBlock = ({ title, hint, file, setter, headers, setHeaders, fieldList, map, setMap }) => {
    const [showOverride, setShowOverride] = useState(false);
    const [inputKey, setInputKey] = useState(0);
    const foundCount = fieldList.filter(f => map[f]).length;
    return (
      <div style={{ marginBottom: 16 }}>
        <div style={{ fontSize: 12.5, fontWeight: 600, marginBottom: 3 }}>{title}</div>
        <div style={{ fontSize: 11, color: "var(--text-faint)", marginBottom: 6 }}>{hint}</div>
        <div style={{ display: "flex", alignItems: "center", gap: 8, flexWrap: "wrap" }}>
          <label className="tap" style={{ display: "inline-flex", alignItems: "center", gap: 8, background: "var(--panel-2)", border: "1px dashed var(--line)", borderRadius: 5, padding: "9px 14px", fontSize: 12.5, color: file ? "var(--text)" : "var(--text-dim)" }}>
            <Upload size={14} />
            {file ? file.name : "Choose data file"}
            <input key={inputKey} type="file" accept=".csv,.tsv,.txt,.xlsx,.xls,.xlsm" style={{ display: "none" }} onChange={e => onSelectFile(e.target.files[0] || null, setter, setHeaders, setMap, fieldList)} />
          </label>
          {file && (
            <button onClick={() => { setter(null); setHeaders([]); setMap({}); setShowOverride(false); setInputKey(k => k + 1); }} className="tap"
              style={{ display: "flex", alignItems: "center", gap: 5, background: "transparent", border: "1px solid var(--line)", color: "var(--text-dim)", padding: "8px 12px", borderRadius: 5, fontSize: 12.5 }}>
              <Trash2 size={13} /> Clear
            </button>
          )}
        </div>
        {file && headers.length === 0 && (
          <div style={{ fontSize: 11.5, color: "var(--red)", marginTop: 6 }}>No header row detected. Does the file have column titles in the first row of every page?</div>
        )}
        {file && headers.length > 0 && (
          <div style={{ marginTop: 8, display: "flex", alignItems: "center", gap: 10, flexWrap: "wrap" }}>
            <span style={{ fontSize: 11.5, color: map.unitId ? "var(--text-dim)" : "var(--red)" }}>
              {map.unitId ? `Detected ${foundCount} of ${fieldList.length} fields automatically.` : "Couldn't confidently detect a Unit ID column."}
            </span>
            <button onClick={() => setShowOverride(!showOverride)} className="tap"
              style={{ fontSize: 11.5, background: "transparent", border: "none", color: "var(--accent)", padding: "2px 0" }}>
              {showOverride ? "Hide" : "Adjust detected columns"}
            </button>
          </div>
        )}
        {file && headers.length > 0 && showOverride && (
          <div style={{ marginTop: 10, background: "var(--panel-2)", borderRadius: 5, padding: 12 }}>
            <div style={{ fontSize: 10.5, color: "var(--text-faint)", marginBottom: 8 }}>
              Columns found across this file: {headers.join(", ")}
            </div>
            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 8 }}>
              {fieldList.map(f => (
                <div key={f} style={{ display: "flex", flexDirection: "column", gap: 3 }}>
                  <label style={{ fontSize: 10.5, color: "var(--text-dim)" }}>{FIELD_LABEL[f]}</label>
                  <select value={map[f] || ""} onChange={e => setMap({ ...map, [f]: e.target.value })}
                    className="tap" style={{ background: "var(--panel)", border: "1px solid var(--line)", color: "var(--text)", borderRadius: 5, padding: "6px 8px", fontSize: 12 }}>
                    <option value="">(not in this file)</option>
                    {headers.map(h => <option key={h} value={h}>{h}</option>)}
                  </select>
                </div>
              ))}
            </div>
          </div>
        )}
      </div>
    );
  };

  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 24, maxWidth: 780 , margin: "0 auto" }}>
      <div style={{ background: "var(--panel)", border: "1px solid var(--line)", borderRadius: 6, padding: 18 }}>
        <div style={{ fontWeight: 600, fontSize: 14, marginBottom: 10 }}>1. Create or pick a building</div>
        <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
          <input value={newBuildingName} onChange={e => setNewBuildingName(e.target.value)} placeholder="New building name"
            style={{ flex: 1, minWidth: 180, background: "var(--panel-2)", border: "1px solid var(--line)", borderRadius: 5, padding: "10px 12px", color: "var(--text)", fontSize: 13.5 }} />
          <button onClick={createBuilding} className="tap" style={{ background: "var(--accent)", color: "#fff", border: "none", borderRadius: 5, padding: "10px 16px", fontSize: 13.5, fontWeight: 600 }}>Create</button>
        </div>
        {buildings.length > 0 && (
          <div style={{ marginTop: 14 }}>
           <div style={{ fontSize: 12, color: "var(--text-dim)", marginBottom: 6 }}>Existing buildings</div>
                                                            <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
                      {buildings.map(b => (
                <button key={b.id} onClick={() => setTargetBuilding(b.id)} className="tap"
                  style={{
                    padding: "8px 14px", borderRadius: 5, fontSize: 13,
                    background: targetBuilding === b.id ? "var(--accent-dim)" : "var(--panel-2)",
                    border: targetBuilding === b.id ? "1px solid var(--accent)" : "1px solid var(--line)",
                    color: "var(--text)"
                  }}>{b.name}</button>
              ))}
            </div>
          </div>
        )}
      </div>

      <div style={{ background: "var(--panel)", border: "1px solid var(--line)", borderRadius: 6, padding: 18 }}>
        <div style={{ fontWeight: 600, fontSize: 14, marginBottom: 4 }}>2. Upload data for: <span style={{ color: "var(--accent)" }}>{state.buildings[targetBuilding]?.name || "no building selected"}</span></div>
        <div style={{ fontSize: 12, color: "var(--text-dim)", marginBottom: 14, lineHeight: 1.5 }}>
          Upload any spreadsheet based database.
        </div>

        <UploadBlock title="Unit data (required)" hint="Needs at minimum a Unit ID / Unit Number / Property No column."
          file={unitFile} setter={setUnitFile} headers={unitHeaders} setHeaders={setUnitHeaders}
          fieldList={["unitId", "unitType", "floor", "carpetArea", "bedrooms", "bathrooms", "currentRent", "currentLeaseStart", "currentLeaseEnd", "status", "ownershipStart", "purchasePrice"]}
          map={unitMap} setMap={setUnitMap} />

        <UploadBlock title="Owner data (optional)" hint="Must also include a Unit ID column so it can join to the unit file."
          file={ownerFile} setter={setOwnerFile} headers={ownerHeaders} setHeaders={setOwnerHeaders}
          fieldList={["unitId", "ownerName", "ownerContact", "secondaryContact", "email"]}
          map={ownerMap} setMap={setOwnerMap} />

        <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginTop: 6, flexWrap: "wrap", gap: 10 }}>
          <button onClick={runImport} disabled={!targetBuilding} className="tap" style={{ background: "var(--accent)", color: "#fff", border: "none", borderRadius: 5, padding: "10px 18px", fontSize: 13.5, fontWeight: 600 }}>
            Import into this building
          </button>
          {confirmWipeBuilding === targetBuilding ? (
            <div style={{ display: "flex", gap: 6 }}>
              <button onClick={() => { wipeBuilding(targetBuilding); setConfirmWipeBuilding(null); }} className="tap"
                style={{ background: "var(--red)", border: "none", color: "#fff", padding: "10px 14px", borderRadius: 5, fontSize: 12, fontWeight: 600 }}>Confirm wipe</button>
              <button onClick={() => setConfirmWipeBuilding(null)} className="tap"
                style={{ background: "transparent", border: "1px solid var(--line)", color: "var(--text-dim)", padding: "10px 14px", borderRadius: 5, fontSize: 12 }}>Cancel</button>
            </div>
          ) : (
            <button onClick={() => targetBuilding && setConfirmWipeBuilding(targetBuilding)} disabled={!targetBuilding} title="Wipe data for this building"
              style={{ background: "transparent", border: "1px solid var(--line)", color: targetBuilding ? "var(--text-dim)" : "var(--text-faint)", padding: "10px 12px", borderRadius: 5, cursor: targetBuilding ? "pointer" : "not-allowed", opacity: targetBuilding ? 1 : 0.5 }}>
              <Trash2 size={14} />
            </button>
          )}
        </div>
        {log && <div style={{ marginTop: 10, fontSize: 12.5, color: "var(--text-dim)", background: "var(--panel-2)", borderRadius: 5, padding: "10px 12px" }}>{log}</div>}
      </div>

      {targetBuilding && state.units && (() => {
        const buildingUnits = Object.values(state.units).filter(u => u.buildingId === targetBuilding);
        const withLease = buildingUnits.filter(u => u.currentLeaseStart || u.currentLeaseEnd).length;
        return (
          <div style={{ fontSize: 12.5, color: "var(--text-dim)" }}>
            {buildingUnits.length} unit(s) currently on file for this building.
            {buildingUnits.length > 0 && (
              <span> {" "}{withLease} of them have current lease dates on file (drives the volatility signal).</span>
            )}
          </div>
        );
      })()}
    </div>
  );
}

function UnitOverrides({ bUnits, agents, toggleUnitAgent }) {
  const [expanded, setExpanded] = useState(false);
  const [searching, setSearching] = useState(false);
  const [q, setQ] = useState("");
  const filtered = q.trim() ? bUnits.filter(u => u.unitId.toLowerCase().includes(q.trim().toLowerCase())) : bUnits;

  return (
    <div>
      <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", gap: 8 }}>
        <div onClick={() => setExpanded(e => !e)} style={{ display: "flex", alignItems: "center", gap: 6, cursor: "pointer", fontSize: 11.5, color: "var(--text-dim)", textTransform: "uppercase", letterSpacing: "0.1em" }}>
          <ChevronRight size={12} style={{ transform: expanded ? "rotate(90deg)" : "none", transition: "transform 0.15s ease" }} />
          Unit overrides
        </div>
        {expanded && (
          <button onClick={() => setSearching(s => !s)} className="tap"
            style={{ background: "transparent", border: "none", color: "var(--text-dim)", cursor: "pointer", padding: 4, display: "flex" }}>
            <Search size={13} />
          </button>
        )}
      </div>
      {expanded && searching && (
        <input value={q} onChange={e => setQ(e.target.value)} placeholder="Search unit…" autoFocus
          style={{ width: "100%", marginTop: 8, background: "var(--panel-2)", border: "1px solid var(--line)", borderRadius: 5, padding: "6px 10px", color: "var(--text)", fontSize: 12 }} />
      )}
      {expanded && (
        <div style={{ maxHeight: 190, overflow: "auto", border: "1px solid var(--line)", borderRadius: 6, marginTop: 8 }}>
          {filtered.map(u => (
            <div key={u.key} style={{ display: "flex", alignItems: "center", justifyContent: "space-between", gap: 8, padding: "8px 10px", borderBottom: "1px solid var(--line)" }}>
              <div style={{ fontSize: 12.5 }}>Unit {u.unitId}</div>
              <div style={{ display: "flex", gap: 5, flexWrap: "wrap" }}>
                {agents.map(a => {
                  const isAssigned = (u.assignedTo || []).includes(a.username);
                  return (
                    <button key={a.username} onClick={() => toggleUnitAgent(u.key, a.username)} className="tap"
                      style={{ padding: "4px 8px", borderRadius: 4, fontSize: 10.5, background: isAssigned ? "var(--accent-dim)" : "transparent", border: isAssigned ? "1px solid var(--accent)" : "1px solid var(--line)", color: "var(--text-dim)" }}>
                      {a.displayName}
                    </button>
                  );
                })}
              </div>
            </div>
          ))}
          {filtered.length === 0 && <div style={{ padding: 10, fontSize: 12, color: "var(--text-faint)" }}>No matching units.</div>}
        </div>
      )}
    </div>
  );
}

function AssignmentPanel({ state, setState }) {  const buildings = Object.values(state.buildings);
  const units = Object.values(state.units);
  const agents = Object.values(state.users || {}).filter(u => u.role === "agent");

  const toggleBuildingAgent = (buildingId, username) => {
    const current = state.assignments[buildingId] || [];
    const next = current.includes(username) ? current.filter(u => u !== username) : [...current, username];
    setState({ ...state, assignments: { ...state.assignments, [buildingId]: next } });
  };

  const toggleUnitAgent = (unitKey, username) => {
    const unit = state.units[unitKey];
    const current = unit.assignedTo || [];
    const next = current.includes(username) ? current.filter(u => u !== username) : [...current, username];
    setState({ ...state, units: { ...state.units, [unitKey]: { ...unit, assignedTo: next } } });
  };

  if (!buildings.length) return <div style={{ color: "var(--text-dim)", fontSize: 13.5 }}>Create a building and import data first.</div>;
  if (!agents.length) return <div style={{ color: "var(--text-dim)", fontSize: 13.5 }}>No agent logins yet. Create one in the Users tab first.</div>;

  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 20, maxWidth: 900 , margin: "0 auto" }}>
      {buildings.map(b => {
        const bUnits = units.filter(u => u.buildingId === b.id);
        const buildingAssigned = state.assignments[b.id] || [];
        return (
          <div key={b.id} style={{ background: "var(--panel)", border: "1px solid var(--line)", borderRadius: 6, padding: 16 }}>
            <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", flexWrap: "wrap", gap: 10 }}>
              <div>
                <div style={{ fontWeight: 600, fontSize: 14.5 }}>{b.name}</div>
                <div style={{ fontSize: 12, color: "var(--text-dim)" }}>{bUnits.length} unit(s)</div>
              </div>
              <div style={{ display: "flex", gap: 6, flexWrap: "wrap" }}>
                {agents.map(a => (
                  <button key={a.username} onClick={() => toggleBuildingAgent(b.id, a.username)} className="tap"
                    style={{
                      padding: "7px 12px", borderRadius: 5, fontSize: 12.5,
                      background: buildingAssigned.includes(a.username) ? "var(--accent-dim)" : "var(--panel-2)",
                      border: buildingAssigned.includes(a.username) ? "1px solid var(--accent)" : "1px solid var(--line)",
                      color: "var(--text)"
                    }}>{a.displayName}</button>
                ))}
              {buildingAssigned.length === 0 && <span style={{ fontSize: 11.5, color: "var(--text-faint)", alignSelf: "center" }}>Unassigned</span>}
              </div>
            </div>
            <div style={{ marginTop: 14, paddingTop: 14, borderTop: "1px solid var(--line)" }}>
              <UnitOverrides bUnits={bUnits} agents={agents} toggleUnitAgent={toggleUnitAgent} />
            </div>
          </div>
        );
      })}
    </div>
  );
}

/* ---------------------------------------------------------
   USERS — admin creates/manages logins
--------------------------------------------------------- */
function UsersPanel({ state, setState }) {
  const [newUsername, setNewUsername] = useState("");
  const [newDisplayName, setNewDisplayName] = useState("");
  const [newPasskey, setNewPasskey] = useState("");
  const [newRole, setNewRole] = useState("agent");
  const [error, setError] = useState("");
  const [confirmingDelete, setConfirmingDelete] = useState(null);
  const [editingPasskeyFor, setEditingPasskeyFor] = useState(null);
  const [passkeyDraft, setPasskeyDraft] = useState("");

  const users = Object.values(state.users || {});

  const createUser = () => {
    const uname = newUsername.trim().toLowerCase();
    if (!uname || !newDisplayName.trim() || !newPasskey.trim()) { setError("Username, display name, and passkey are all required."); return; }
    if (state.users[uname]) { setError(`"${uname}" is already taken.`); return; }
    setState({ ...state, users: { ...state.users, [uname]: { username: uname, displayName: newDisplayName.trim(), passkey: newPasskey.trim(), role: newRole } } });
    setNewUsername(""); setNewDisplayName(""); setNewPasskey(""); setNewRole("agent"); setError("");
  };

  const savePasskey = (username) => {
    if (!passkeyDraft.trim()) return;
    setState({ ...state, users: { ...state.users, [username]: { ...state.users[username], passkey: passkeyDraft.trim() } } });
    setEditingPasskeyFor(null); setPasskeyDraft("");
  };

  const deleteUser = (username) => {
    const next = { ...state.users };
    delete next[username];
    // clear this user out of any building/unit assignments so nothing points at a login that no longer exists
    const assignments = {};
    Object.entries(state.assignments).forEach(([bid, list]) => { assignments[bid] = (list || []).filter(u => u !== username); });
    const units = {};
    Object.entries(state.units).forEach(([key, u]) => { units[key] = { ...u, assignedTo: u.assignedTo ? u.assignedTo.filter(a => a !== username) : u.assignedTo }; });
    setState({ ...state, users: next, assignments, units });
    setConfirmingDelete(null);
  };

  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 20, maxWidth: 620 , margin: "0 auto" }}>
      <div style={{ background: "var(--panel)", border: "1px solid var(--line)", borderRadius: 6, padding: 18 }}>
        <div style={{ fontWeight: 600, fontSize: 14, marginBottom: 12 }}>Create a login</div>
        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10 }}>
          <div>
            <div style={{ fontSize: 11, color: "var(--text-dim)", marginBottom: 3 }}>Username</div>
            <input value={newUsername} onChange={e => setNewUsername(e.target.value)}
              className="tap" style={{ width: "100%", background: "var(--panel-2)", border: "1px solid var(--line)", color: "var(--text)", borderRadius: 5, padding: "9px 11px", fontSize: 13.5 }} />
          </div>
          <div>
            <div style={{ fontSize: 11, color: "var(--text-dim)", marginBottom: 3 }}>Display name</div>
            <input value={newDisplayName} onChange={e => setNewDisplayName(e.target.value)}
              className="tap" style={{ width: "100%", background: "var(--panel-2)", border: "1px solid var(--line)", color: "var(--text)", borderRadius: 5, padding: "9px 11px", fontSize: 13.5 }} />
          </div>
          <div>
            <div style={{ fontSize: 11, color: "var(--text-dim)", marginBottom: 3 }}>Passkey</div>
            <input value={newPasskey} onChange={e => setNewPasskey(e.target.value)}
              className="tap" style={{ width: "100%", background: "var(--panel-2)", border: "1px solid var(--line)", color: "var(--text)", borderRadius: 5, padding: "9px 11px", fontSize: 13.5 }} />
          </div>
          <div>
            <div style={{ fontSize: 11, color: "var(--text-dim)", marginBottom: 3 }}>Role</div>
            <select value={newRole} onChange={e => setNewRole(e.target.value)}
              className="tap" style={{ width: "100%", background: "var(--panel-2)", border: "1px solid var(--line)", color: "var(--text)", borderRadius: 5, padding: "9px 11px", fontSize: 13.5 }}>
              <option value="agent">Agent</option>
              <option value="admin">Admin</option>
            </select>
          </div>
        </div>
        {error && <div style={{ color: "var(--red)", fontSize: 12, marginTop: 8 }}>{error}</div>}
        <button onClick={createUser} className="tap" style={{ marginTop: 12, background: "var(--accent)", color: "#fff", border: "none", borderRadius: 5, padding: "9px 16px", fontSize: 13.5, fontWeight: 600 }}>Create login</button>
      </div>

      <div style={{ background: "var(--panel)", border: "1px solid var(--line)", borderRadius: 6, padding: 18 }}>
        <div style={{ fontWeight: 600, fontSize: 14, marginBottom: 12 }}>Existing logins</div>
        <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
          {users.map(u => (
            <div key={u.username} style={{ padding: "10px 12px", background: "var(--panel-2)", borderRadius: 8 }}>
              <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", gap: 10, flexWrap: "wrap" }}>
                <div>
                  <div style={{ fontSize: 13.5, fontWeight: 600 }}>{u.displayName} <span style={{ color: "var(--text-faint)", fontWeight: 400 }}>· {u.username}</span></div>
                  <div style={{ fontSize: 11.5, color: "var(--text-dim)" }}>{u.role === "admin" ? "Admin" : "Agent"} · passkey: {u.passkey}</div>
                </div>
                {confirmingDelete === u.username ? (
                  <div style={{ display: "flex", gap: 6 }}>
                    <button onClick={() => deleteUser(u.username)} className="tap" style={{ background: "var(--red)", border: "none", color: "#fff", padding: "6px 10px", borderRadius: 5, fontSize: 11.5, fontWeight: 600 }}>Confirm</button>
                    <button onClick={() => setConfirmingDelete(null)} className="tap" style={{ background: "transparent", border: "1px solid var(--line)", color: "var(--text-dim)", padding: "6px 10px", borderRadius: 5, fontSize: 11.5 }}>Cancel</button>
                  </div>
                ) : editingPasskeyFor === u.username ? null : (
                  <div style={{ display: "flex", gap: 6 }}>
                    <button onClick={() => { setEditingPasskeyFor(u.username); setPasskeyDraft(""); }} className="tap"
                      style={{ background: "transparent", border: "1px solid var(--line)", color: "var(--text-dim)", padding: "6px 10px", borderRadius: 5, fontSize: 11.5 }}>Change passkey</button>
                    <button onClick={() => setConfirmingDelete(u.username)} className="tap" style={{ background: "transparent", border: "1px solid var(--line)", color: "var(--text-dim)", padding: "6px 10px", borderRadius: 5, fontSize: 11.5 }}>Remove</button>
                  </div>
                )}
              </div>
              {editingPasskeyFor === u.username && (
                <div style={{ display: "flex", gap: 6, marginTop: 8 }}>
                  <input value={passkeyDraft} onChange={e => setPasskeyDraft(e.target.value)} placeholder="New passkey" autoFocus
                    onKeyDown={e => { if (e.key === "Enter") savePasskey(u.username); if (e.key === "Escape") setEditingPasskeyFor(null); }}
                    className="tap" style={{ flex: 1, background: "var(--panel)", border: "1px solid var(--line)", color: "var(--text)", borderRadius: 5, padding: "7px 10px", fontSize: 12.5 }} />
                  <button onClick={() => savePasskey(u.username)} className="tap" style={{ background: "var(--accent)", border: "none", color: "#fff", padding: "6px 12px", borderRadius: 5, fontSize: 11.5, fontWeight: 600 }}>Save</button>
                  <button onClick={() => setEditingPasskeyFor(null)} className="tap" style={{ background: "transparent", border: "1px solid var(--line)", color: "var(--text-dim)", padding: "6px 10px", borderRadius: 5, fontSize: 11.5 }}>Cancel</button>
                </div>
              )}
            </div>
          ))}
        </div>
      </div>
    </div>
  );
}

/* ---------------------------------------------------------
   USER APP (agent calling portal)
--------------------------------------------------------- */
function effectiveAssignedUsers(unit, assignments) {
  if (unit.assignedTo && unit.assignedTo.length) return unit.assignedTo;
  return assignments[unit.buildingId] || [];
}

function UserApp({ state, setState, user, onLogout, saving }) {
  const [rightPanel, setRightPanel] = useState("detail"); // 'detail' | 'lookup'
  const [buildingToggle, setBuildingToggle] = useState(null);
  const [bucketTab, setBucketTab] = useState("toBeCalled");
  const [activeUnitKey, setActiveUnitKey] = useState(null);
  const displayName = state.users[user]?.displayName || user;

  const myUnits = useMemo(() => {
    return Object.values(state.units).filter(u => effectiveAssignedUsers(u, state.assignments).includes(user));
  }, [state.units, state.assignments, user]);

  // Assignment is determined by the buildings/assignments table itself, NOT by
  // whether units happen to exist yet. Assigning a building before its CSV is
  // uploaded (or an import that silently produced 0 rows) must still show up
  // here — otherwise "assigned but empty" looks identical to "not assigned".
  const myBuildingIds = useMemo(() => {
    const fromBuildingLevel = Object.keys(state.buildings).filter(bid => (state.assignments[bid] || []).includes(user));
    const fromUnitOverride = [...new Set(
      Object.values(state.units).filter(u => u.assignedTo && u.assignedTo.includes(user)).map(u => u.buildingId)
    )];
    return [...new Set([...fromBuildingLevel, ...fromUnitOverride])];
  }, [state.buildings, state.assignments, state.units, user]);

  useEffect(() => {
    if ((!buildingToggle || !myBuildingIds.includes(buildingToggle)) && myBuildingIds.length) {
      setBuildingToggle(myBuildingIds[0]);
    }
  }, [myBuildingIds]); // eslint-disable-line

  // reassigned-to-me units outside the currently toggled building (cross-building call scenario)
  const reassignedElsewhere = useMemo(() =>
    myUnits.filter(u => u.assignedTo && u.assignedTo.includes(user) && u.buildingId !== buildingToggle),
    [myUnits, buildingToggle, user]);

  const buildingUnits = useMemo(() => myUnits.filter(u => u.buildingId === buildingToggle), [myUnits, buildingToggle]);

  const activeUnit = state.units[activeUnitKey] || null;
  const activeUnitBuildingUnits = activeUnit ? Object.values(state.units).filter(u => u.buildingId === activeUnit.buildingId) : [];

  const updateCrm = (key, nextCrmPartial) => {
    const crm = { ...(state.crm[key] || defaultCrm()), ...nextCrmPartial };
    setState({ ...state, crm: { ...state.crm, [key]: crm } });
  };
  const logOutcome = (key, outcome) => {
    const crm = applyOutcome(state.crm[key] || defaultCrm(), outcome);
    const entry = { unit: key, agent: user, outcome, timestamp: new Date().toISOString() };
    setState({ ...state, crm: { ...state.crm, [key]: crm }, callLog: [...(state.callLog || []), entry] });
  };
  // Deliberate manual override for a misclick — pulls a unit out of
  // whichever bucket it landed in (Call Me Back, No Answer, Reject, Active)
  // and back into To Be Called. Notes and call history stay intact; only
  // the bucket/date/tag reset. If the last thing logged was a No Answer,
  // that single count is rolled back too, since a misclick shouldn't count
  // as a real missed call.
  const moveToQueue = (key) => {
    const prev = state.crm[key] || defaultCrm();
    const crm = {
      ...prev,
      bucket: "toBeCalled",
      nextActionDate: null,
      tag: null,
      noAnswerCount: prev.lastOutcome === "noAnswer" ? Math.max(0, (prev.noAnswerCount || 0) - 1) : (prev.noAnswerCount || 0),
    };
    setState({ ...state, crm: { ...state.crm, [key]: crm } });
  };

  return (
    <>
      <TopBar
        title={displayName}
        subtitle={state.buildings[buildingToggle]?.name || "No building assigned yet"}
        onLogout={onLogout} saving={saving}
        right={
          <div style={{ display: "flex", gap: 8, alignItems: "center" }}>
            {myBuildingIds.length > 0 && (
              <select value={buildingToggle || ""} onChange={e => { setBuildingToggle(e.target.value); setActiveUnitKey(null); }}
                className="tap" style={{ background: "var(--panel-2)", border: "1px solid var(--line)", color: "var(--text)", borderRadius: 5, padding: "8px 10px", fontSize: 13 }}>
                {myBuildingIds.map(bid => {
                  const count = Object.values(state.units).filter(u => u.buildingId === bid && effectiveAssignedUsers(u, state.assignments).includes(user)).length;
                  return <option key={bid} value={bid}>{state.buildings[bid]?.name}{count === 0 ? " (no units yet)" : ""}</option>;
                })}
              </select>
            )}
          </div>
        } />

      {reassignedElsewhere.length > 0 && (
        <div style={{ background: "var(--amber-dim)", borderBottom: "1px solid var(--line)", padding: "8px 20px", fontSize: 12.5, color: "var(--amber)", display: "flex", alignItems: "center", gap: 8 }}>
          <ArrowRightLeft size={13} />
          {reassignedElsewhere.length} unit(s) reassigned to you from another building: {reassignedElsewhere.map(u => `${state.buildings[u.buildingId]?.name} #${u.unitId}`).join(", ")}
        </div>
      )}

      {myBuildingIds.length === 0 ? (
        <div style={{ padding: 40, color: "var(--text-dim)", textAlign: "center" }}>No buildings assigned to you yet. Ask the admin to assign one in the Assignments tab.</div>
      ) : buildingUnits.length === 0 ? (
        <div style={{ padding: 40, color: "var(--text-dim)", textAlign: "center" }}>
          <div style={{ marginBottom: 6 }}>{state.buildings[buildingToggle]?.name} is assigned to you, but no units have been imported for it yet.</div>
          <div style={{ fontSize: 12.5, color: "var(--text-faint)" }}>Ask the admin to upload the unit CSV for this building.</div>
        </div>
      ) : (
        // Split screen is permanent: CRM queue always occupies the left pane.
        // The right pane switches between the active call detail and the
        // standalone lookup tool — lookup never replaces the queue anymore.
        <div style={{ display: "flex", flex: 1, overflow: "hidden" }}>
          <QueuePane
            bucketTab={bucketTab} setBucketTab={setBucketTab}
            units={buildingUnits} crmMap={state.crm}
            activeUnitKey={activeUnitKey}
            setActiveUnitKey={(key) => { setActiveUnitKey(key); setRightPanel("detail"); }}
            onMoveToQueue={moveToQueue}
          />
          <div style={{ flex: 1, display: "flex", flexDirection: "column", overflow: "hidden" }}>
            <div style={{ display: "flex", borderBottom: "1px solid var(--line)", background: "var(--panel)" }}>
{[["detail", "Call Detail"], ["lookup", "Lookup"], ["explorer", "Explorer"]].map(([id, label]) => (                <button key={id} onClick={() => setRightPanel(id)} className="tap"
                  style={{
                    padding: "11px 16px", background: "transparent", border: "none",
                    borderBottom: rightPanel === id ? "1px solid var(--accent)" : "1px solid transparent",
                    color: rightPanel === id ? "var(--text)" : "var(--text-dim)", fontSize: 10, fontWeight: 600,
                    textTransform: "uppercase", letterSpacing: "0.12em",
                    display: "flex", alignItems: "center", gap: 6
                  }}>
{id === "lookup" && <Search size={13} />} {id === "explorer" && <LayoutGrid size={13} />} {label}                </button>
              ))}
            </div>
            <div style={{ flex: 1, overflow: "hidden", display: "flex" }}>
              {rightPanel === "detail" ? (
                <DetailPane
                  unit={activeUnit} crm={activeUnit ? (state.crm[activeUnit.key] || defaultCrm()) : null}
                  buildingUnits={activeUnitBuildingUnits}
                  onNotes={(notes) => activeUnit && updateCrm(activeUnit.key, { notes })}
                  onOutcome={(outcome) => activeUnit && logOutcome(activeUnit.key, outcome)}
                  onMoveToQueue={() => activeUnit && moveToQueue(activeUnit.key)}
                />
              ) : rightPanel === "lookup" ? (
                <LookupTool state={state} myUnits={myUnits} />
              ) : (
                <BuildingExplorer state={state} user={user} role="agent" />
              )}
            </div>
          </div>
        </div>
      )}
    </>
  );
}

/* ---------------------------------------------------------
   QUEUE PANE (left)
--------------------------------------------------------- */
const TABS = [
  ["toBeCalled", "To Be Called"],
  ["callMeBack", "Call Me Back"],
  ["unreachable", "Unreachable"],
  ["reject", "Reject"],
  ["yes", "Active"],
];

const SORT_OPTIONS = [
  ["unitAsc", "Unit # ↑"],
  ["unitDesc", "Unit # ↓"],
  ["easy", "Easy to Rent"],
  ["hard", "Hard to Rent"],
];
const OCCUPANCY_OPTIONS = [
  ["all", "All"],
  ["vacant", "Vacant"],
  ["occupied", "Occupied"],
];

function compareUnitId(a, b) {
  const na = parseFloat(a), nb = parseFloat(b);
  if (!isNaN(na) && !isNaN(nb) && na !== nb) return na - nb;
  return a.localeCompare(b, undefined, { numeric: true });
}

function QueuePane({ bucketTab, setBucketTab, units, crmMap, activeUnitKey, setActiveUnitKey, onMoveToQueue }) {
  const [sortMode, setSortMode] = useState("easy");
  const [occupancyFilter, setOccupancyFilter] = useState("all");
  const [filterOpen, setFilterOpen] = useState(false);
  const [searchOpen, setSearchOpen] = useState(false);
  const [searchQuery, setSearchQuery] = useState("");
  const [channelSubTab, setChannelSubTab] = useState(null);
  const [rejectSubTab, setRejectSubTab] = useState(null);

  const rows = useMemo(() => {
    return units.map(u => {
      const crm = crmMap[u.key] || defaultCrm();
      const eff = effectiveBucket(crm);
      const volatility = computeVolatility(u);
      const demand = computeDemand(u, units);
      return { unit: u, crm, eff, volatility, demand, rank: volatility - demand };
    });
  }, [units, crmMap]);

  const searchActive = searchOpen && searchQuery.trim().length > 0;
  const q = searchQuery.trim().toLowerCase();

    const bucketFiltered = rows.filter(r => {
    if (searchActive) {
      // When search is active, scope is the whole CRM (every bucket), matched by unit id or owner name — not affected by any filter, occupancy included
      return r.unit.unitId.toLowerCase().includes(q) || (r.unit.ownerName || "").toLowerCase().includes(q);
    }
    if (occupancyFilter !== "all" && r.unit.status !== occupancyFilter) return false;
        if (bucketTab === "reject") {
      const inGroup = ["declined", "cold"].includes(r.eff.display);
      if (!inGroup) return false;
      return rejectSubTab ? r.eff.display === rejectSubTab : true;
    }
    if (bucketTab === "unreachable") {
      const inGroup = ["unreachable", "whatsapp", "linkedin", "email"].includes(r.eff.display);
      if (!inGroup) return false;
      return channelSubTab ? r.eff.display === channelSubTab : true;
    }
    return r.eff.display === bucketTab;
  });

  const sorted = [...bucketFiltered].sort((a, b) => {
    if (sortMode === "unitAsc") return compareUnitId(a.unit.unitId, b.unit.unitId);
    if (sortMode === "unitDesc") return compareUnitId(b.unit.unitId, a.unit.unitId);
    if (sortMode === "hard") return a.rank - b.rank; // hard to rent: low volatility, high demand — lowest rank first
    return b.rank - a.rank; // easy to rent: high volatility, low demand — highest rank first
  });

    const counts = Object.fromEntries(TABS.map(([id]) => [
    id,
       id === "reject"
      ? rows.filter(r => ["declined", "cold"].includes(r.eff.display)).length
      : id === "unreachable"
      ? rows.filter(r => ["unreachable", "whatsapp", "linkedin", "email"].includes(r.eff.display)).length
      : rows.filter(r => r.eff.display === id).length
  ]));

  const activeSortLabel = SORT_OPTIONS.find(([id]) => id === sortMode)?.[1];

  return (
    <div style={{ width: "38%", minWidth: 300, borderRight: "1px solid var(--line)", display: "flex", flexDirection: "column", background: "var(--panel)" }}>
      <div style={{ display: "flex", flexWrap: "wrap", borderBottom: "1px solid var(--line)" }}>
               {TABS.map(([id, label]) => (
                   <button key={id} onClick={() => { setBucketTab(id); setChannelSubTab(null); setRejectSubTab(null); }} className="tap"
            style={{
              flex: "1 0 auto", padding: "11px 8px", background: "transparent", border: "none",
              borderBottom: bucketTab === id ? "1px solid var(--accent)" : "1px solid transparent",
              color: bucketTab === id ? "var(--text)" : "var(--text-dim)", fontSize: 9.5, fontWeight: 600,
              textTransform: "uppercase", letterSpacing: "0.1em"
            }}>{label} <span style={{ color: bucketTab === id ? "var(--accent)" : "var(--text-faint)" }}>{counts[id]}</span></button>
        ))}
      </div>

           {bucketTab === "unreachable" && (
        <div style={{ display: "flex", gap: 6, padding: "8px 12px", borderBottom: "1px solid var(--line)", flexWrap: "wrap" }}>
          {[[null, "All"], ["whatsapp", "WhatsApp"], ["email", "Email"], ["linkedin", "LinkedIn"]].map(([id, label]) => (
            <button key={label} onClick={() => setChannelSubTab(id)} className="tap"
              style={{
                padding: "5px 11px", borderRadius: 14, fontSize: 10.5,
                background: channelSubTab === id ? "var(--accent)" : "transparent",
                border: channelSubTab === id ? "1px solid var(--accent)" : "1px solid var(--line)",
                color: channelSubTab === id ? "#fff" : "var(--text-dim)", fontWeight: channelSubTab === id ? 600 : 400
              }}>{label}</button>
          ))}
        </div>
      )}

      {bucketTab === "reject" && (
        <div style={{ display: "flex", gap: 6, padding: "8px 12px", borderBottom: "1px solid var(--line)", flexWrap: "wrap" }}>
          {[[null, "All"], ["declined", "Declined"], ["cold", "Cold"]].map(([id, label]) => (
            <button key={label} onClick={() => setRejectSubTab(id)} className="tap"
              style={{
                padding: "5px 11px", borderRadius: 14, fontSize: 10.5,
                background: rejectSubTab === id ? "var(--accent)" : "transparent",
                border: rejectSubTab === id ? "1px solid var(--accent)" : "1px solid var(--line)",
                color: rejectSubTab === id ? "#fff" : "var(--text-dim)", fontWeight: rejectSubTab === id ? 600 : 400
              }}>{label}</button>
          ))}
        </div>
      )}

      <div style={{ borderBottom: "1px solid var(--line)" }}>
        <div style={{ display: "flex", alignItems: "stretch" }}>
          <button onClick={() => setFilterOpen(!filterOpen)} className="tap"
            style={{
              flex: 1, display: "flex", alignItems: "center", justifyContent: "space-between",
              padding: "11px 12px", background: filterOpen ? "var(--panel-2)" : "transparent", border: "none",
              color: "var(--text)", fontSize: 12
            }}>
            <span style={{ display: "flex", alignItems: "center", gap: 7 }}>
              <span className="eyebrow">Filter</span>
              <span style={{ color: "var(--accent)", fontSize: 12 }}>{activeSortLabel}{occupancyFilter !== "all" ? ` · ${OCCUPANCY_OPTIONS.find(([id]) => id === occupancyFilter)[1]}` : ""}</span>
            </span>
            <ChevronRight size={14} style={{ transform: filterOpen ? "rotate(90deg)" : "none", transition: "transform 0.15s", color: "var(--text-faint)" }} />
          </button>
          <button onClick={() => { setSearchOpen(v => !v); if (searchOpen) setSearchQuery(""); }} className="tap"
            aria-label="Search all units"
            style={{
              width: 42, display: "flex", alignItems: "center", justifyContent: "center",
              background: searchOpen ? "var(--panel-2)" : "transparent",
              border: "none", borderLeft: "1px solid var(--line)",
              color: searchOpen ? "var(--accent)" : "var(--text-dim)"
            }}>
            <Search size={14} />
          </button>
        </div>
        {searchOpen && (
          <div style={{ padding: "8px 12px", borderTop: "1px solid var(--line)", background: "var(--panel-2)" }}>
            <input value={searchQuery} onChange={e => setSearchQuery(e.target.value)} placeholder="Find unit"
              autoFocus className="tap"
              style={{ width: "100%", background: "var(--panel)", border: "1px solid var(--line)", color: "var(--text)", borderRadius: 6, padding: "8px 10px", fontSize: 12.5, outline: "none" }} />
            {searchActive && (
              <div className="eyebrow" style={{ marginTop: 6, color: "var(--text-faint)" }}>
                Showing matches across every bucket
              </div>
            )}
          </div>
        )}
        {filterOpen && (
          <div style={{ padding: "0 12px 12px", display: "flex", flexDirection: "column", gap: 4 }}>
            {SORT_OPTIONS.map(([id, label]) => (
              <button key={id} onClick={() => { setSortMode(id); setFilterOpen(false); }} className="tap"
                style={{
                  display: "flex", alignItems: "center", justifyContent: "space-between",
                  padding: "9px 11px", borderRadius: 6, fontSize: 12.5, textAlign: "left",
                  background: sortMode === id ? "var(--accent-dim)" : "var(--panel-2)",
                  border: sortMode === id ? "1px solid var(--accent)" : "1px solid var(--line)",
                  color: "var(--text)"
                }}>
                {label}
                {sortMode === id && <Check size={13} color="var(--accent)" />}
              </button>
            ))}
            <div className="eyebrow" style={{ marginTop: 8, marginBottom: 2, color: "var(--text-faint)" }}>Occupancy</div>
            {OCCUPANCY_OPTIONS.map(([id, label]) => (
              <button key={id} onClick={() => setOccupancyFilter(id)} className="tap"
                style={{
                  display: "flex", alignItems: "center", justifyContent: "space-between",
                  padding: "9px 11px", borderRadius: 6, fontSize: 12.5, textAlign: "left",
                  background: occupancyFilter === id ? "var(--accent-dim)" : "var(--panel-2)",
                  border: occupancyFilter === id ? "1px solid var(--accent)" : "1px solid var(--line)",
                  color: "var(--text)"
                }}>
                {label}
                {occupancyFilter === id && <Check size={13} color="var(--accent)" />}
              </button>
            ))}
          </div>
        )}
      </div>

      <div style={{ overflow: "auto", flex: 1 }}>
        {sorted.length === 0 && (
          <div style={{ padding: 24, color: "var(--text-faint)", fontSize: 13, textAlign: "center" }}>
            {searchActive ? "No unit or owner matches that search." : "Nothing in this bucket right now."}
          </div>
        )}
        {sorted.map(({ unit, crm, eff, volatility, demand }) => (
          <UnitRow key={unit.key} unit={unit} crm={crm} eff={eff} volatility={volatility} demand={demand}
            active={unit.key === activeUnitKey}
            onToggle={() => setActiveUnitKey(unit.key === activeUnitKey ? null : unit.key)}
            onMoveToQueue={bucketTab !== "toBeCalled" ? () => onMoveToQueue(unit.key) : null} />
        ))}
      </div>
    </div>
  );
}

function MiniBar({ label, value, colorVar }) {
  return (
    <div style={{ display: "flex", alignItems: "center", gap: 6, minWidth: 0 }}>
      <span className="mono" style={{ fontSize: 9.5, color: "var(--text-faint)", width: 12 }}>{label}</span>
      <div style={{ flex: 1, height: 2, background: "var(--line)", overflow: "hidden" }}>
        <div style={{ width: `${value}%`, height: "100%", background: colorVar }} />
      </div>
    </div>
  );
}

function UnitRow({ unit, crm, eff, volatility, demand, active, onToggle, onMoveToQueue }) {
  const [confirmingMove, setConfirmingMove] = useState(false);
  return (
    <div onClick={onToggle} className="tap"
      style={{
        padding: "13px 14px", borderBottom: "1px solid var(--line)", cursor: "pointer",
        background: active ? "var(--panel-2)" : "transparent",
        borderLeft: active ? "2px solid var(--accent)" : "2px solid transparent",
        display: "flex", flexDirection: "column", gap: 7
      }}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
        <div className="mono" style={{ fontSize: 14, fontWeight: 500, color: active ? "var(--accent)" : "var(--text)" }}>{unit.unitId}</div>
        <div style={{ display: "flex", alignItems: "center", gap: 6 }}>
          {eff.resurfaced && (
            <span style={{ fontSize: 10, padding: "2px 6px", borderRadius: 5, background: BUCKET_COLOR[eff.origin] + "22", color: BUCKET_COLOR[eff.origin] }}>
                            back from {BUCKET_LABEL[eff.origin]}{eff.origin === "unreachable" ? ` ×${crm.noAnswerCount}` : ""}
            </span>
          )}
          {crm.bucket !== "toBeCalled" && !eff.resurfaced && (
            <span style={{ fontSize: 10, padding: "2px 6px", borderRadius: 5, background: BUCKET_COLOR[crm.bucket] + "22", color: BUCKET_COLOR[crm.bucket] }}>
                            {crm.bucket === "callMeBack" || crm.bucket === "unreachable" ? `in ${daysUntil(crm.nextActionDate)}d` : BUCKET_LABEL[crm.bucket]}
            </span>
          )}
          {onMoveToQueue && (confirmingMove ? (
            <>
              <button onClick={(e) => { e.stopPropagation(); onMoveToQueue(); setConfirmingMove(false); }}
                className="tap" title="Confirm move"
                style={{ padding: "3px 7px", background: "var(--red)", border: "none", borderRadius: 5, color: "#fff", fontSize: 10.5, fontWeight: 600 }}>
                Confirm
              </button>
              <button onClick={(e) => { e.stopPropagation(); setConfirmingMove(false); }}
                className="tap" title="Cancel"
                style={{ padding: "3px 7px", background: "transparent", border: "1px solid var(--line)", borderRadius: 5, color: "var(--text-dim)", fontSize: 10.5 }}>
                ✕
              </button>
            </>
          ) : (
            <button
              onClick={(e) => { e.stopPropagation(); setConfirmingMove(true); }}
              className="tap" title="Move back to To Be Called"
              style={{ display: "flex", alignItems: "center", padding: 4, background: "transparent", border: "1px solid var(--line)", borderRadius: 5, color: "var(--text-dim)" }}>
              <RefreshCcw size={12} />
            </button>
          ))}
        </div>
      </div>
      <div style={{ fontSize: 11.5, color: "var(--text-dim)" }}>{unit.bedrooms ? `${unit.bedrooms} bed` : "—"} · Floor {unit.floor || "—"} · {unit.ownerName || <em style={{ color: "var(--text-faint)" }}>owner not available</em>}</div>
      <div style={{ display: "flex", gap: 10 }}>
        <MiniBar label="V" value={volatility} colorVar="var(--red)" />
        <MiniBar label="D" value={demand} colorVar="var(--green)" />
      </div>
    </div>
  );
}

/* ---------------------------------------------------------
   DETAIL PANE (right) — the split-screen "everything at a touch" view
--------------------------------------------------------- */

function Field({ label, value, emptyText }) {
  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 2 }}>
      <div className="eyebrow" style={{ color: "var(--text-faint)" }}>{label}</div>
      <div style={{ fontSize: 14, color: value ? "var(--text)" : "var(--text-faint)" }}>
        {value || <em>{emptyText || "not available"}</em>}
      </div>
    </div>
  );
}

function DetailPane({ unit, crm, buildingUnits, onNotes, onOutcome, onMoveToQueue }) {
  const [confirmingMove, setConfirmingMove] = useState(false);
  if (!unit) {
    return (
      <div style={{ flex: 1, display: "flex", alignItems: "center", justifyContent: "center", color: "var(--text-faint)", fontSize: 13.5 }}>
        Tap a unit on the left to load it here.
      </div>
    );
  }
  const volatility = computeVolatility(unit);
  const demand = computeDemand(unit, buildingUnits);

  return (
    <div style={{ flex: 1, overflow: "auto", padding: 20, display: "flex", flexDirection: "column", gap: 18 }}>
      {/* Confirmation banner — the "headline moment" for the call in progress */}
      <div>
        <hr className="rule" style={{ margin: "0 0 12px" }} />
        <div style={{ display: "flex", alignItems: "flex-start", justifyContent: "space-between", flexWrap: "wrap", gap: 10 }}>
          <div style={{ minWidth: 0 }}>
            <div className="eyebrow" style={{ marginBottom: 6 }}>
              Unit {unit.unitId} · {unit.status === "vacant" ? "Vacant" : unit.status === "occupied" ? "Occupied" : "Status unknown"}
            </div>
            <div className="disp" style={{ fontSize: 30, fontWeight: 800, lineHeight: 1.05 }}>
              {unit.ownerName || "Owner not available"}
            </div>
            <div className="mono" style={{ fontSize: 13, color: unit.ownerContact ? "var(--accent)" : "var(--text-faint)", marginTop: 7 }}>
              {unit.ownerContact || "No contact on file"}
            </div>
            <div className="mono" style={{ fontSize: 12, color: unit.secondaryContact ? "var(--text-dim)" : "var(--text-faint)", marginTop: 3 }}>
              {unit.secondaryContact || "No secondary number"}
            </div>
          </div>
          <div className="eyebrow" style={{ color: "var(--accent)", whiteSpace: "nowrap", paddingTop: 2 }}>Calling now</div>
        </div>
      </div>

      {/* Scores */}
      <div style={{ background: "var(--panel)", border: "1px solid var(--line)", borderRadius: 6, padding: 16 }}>
        <div className="eyebrow" style={{ marginBottom: 12 }}>Signal</div>
        <div style={{ display: "flex", flexDirection: "column", gap: 12 }}>
          <ScoreBar label="Volatility" value={volatility} colorVar="var(--accent)" />
          <ScoreBar label="Demand" value={demand} colorVar="var(--green)" />
        </div>
        <div style={{ fontSize: 12, color: "var(--text-dim)", marginTop: 14, lineHeight: 1.6, borderLeft: "2px solid var(--line)", paddingLeft: 12 }}>
          {volatility > 55 && demand < 45 && "Easy to rent: lease has turned over or is turning over soon, and rent sits below comps. Good candidate to pitch."}
          {volatility < 45 && demand > 55 && "Hard to rent: lease is locked in for a while and rent is already at or above comps. Owner likely has no pressing pain point."}
          {!(volatility > 55 && demand < 45) && !(volatility < 45 && demand > 55) && "Mixed profile. Check the lease dates and rent below before deciding your angle."}
        </div>
      </div>

      {/* Core fields */}
      <div style={{ background: "var(--panel)", border: "1px solid var(--line)", borderRadius: 6, padding: 16, display: "grid", gridTemplateColumns: "1fr 1fr", gap: 14 }}>
        <Field label="Floor" value={unit.floor} />
        <Field label="Unit Type" value={unit.unitType} />
        <Field label="Carpet Area" value={unit.carpetArea ? `${unit.carpetArea} sqft` : ""} />
        <Field label="Bed / Bath" value={unit.bedrooms || unit.bathrooms ? `${unit.bedrooms || "—"} bed / ${unit.bathrooms || "—"} bath` : ""} />
        <Field label="Current Rent" value={unit.currentRent ? `AED ${Number(unit.currentRent).toLocaleString()}` : ""} />
        <Field label="Current Lease Start" value={unit.currentLeaseStart ? fmtDate(unit.currentLeaseStart) : ""} />
        <Field label="Current Lease End" value={unit.currentLeaseEnd ? fmtDate(unit.currentLeaseEnd) : ""} />
        <Field label="Purchase Price" value={unit.purchasePrice ? `AED ${Number(unit.purchasePrice).toLocaleString()}` : ""} />
        <Field label="Ownership Since" value={unit.ownershipStart ? fmtDate(unit.ownershipStart) : ""} />
        <Field label="Current Market Price" value={unit.marketPrice ? `AED ${Number(unit.marketPrice).toLocaleString()}` : ""} />
        <Field label="Email" value={unit.email} emptyText="No email" />
      </div>

      {/* Notes */}
      <div>
        <div style={{ fontSize: 12, fontWeight: 600, color: "var(--text-dim)", marginBottom: 8 }}>NOTES</div>
        <div style={{ position: "relative" }}>
          <textarea value={crm.notes} onChange={e => onNotes(e.target.value)} placeholder="What does this owner need? What did they say last call?"
            style={{ width: "100%", minHeight: 90, background: "var(--panel)", border: "1px solid var(--line)", borderRadius: 6, padding: 12, paddingBottom: 34, color: "var(--text)", fontSize: 13, resize: "vertical" }} />
          <button onClick={() => onNotes("")} className="sweep-btn" title="Clear notes"
            style={{ position: "absolute", right: 8, bottom: 8, display: "flex", alignItems: "center", justifyContent: "center", width: 22, height: 22, background: "transparent", border: "1px solid var(--line)", borderRadius: 4, color: "var(--text-dim)", cursor: "pointer", "--sweep-color": "var(--red)" }}>
            <span style={{ position: "relative", zIndex: 1, display: "flex" }}><X size={13} /></span>
          </button>
        </div>
      </div>

      {/* Outcome logging */}
      <div>
        <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 8 }}>
          <div style={{ fontSize: 12, fontWeight: 600, color: "var(--text-dim)" }}>LOG OUTCOME</div>
          {crm.bucket !== "toBeCalled" && (confirmingMove ? (
            <div style={{ display: "flex", alignItems: "center", gap: 6 }}>
              <span style={{ fontSize: 11.5, color: "var(--red)" }}>Move back?</span>
              <button onClick={() => { onMoveToQueue(); setConfirmingMove(false); }} className="tap"
                style={{ background: "var(--red)", border: "none", color: "#fff", padding: "5px 10px", borderRadius: 5, fontSize: 11.5, fontWeight: 600 }}>Confirm</button>
              <button onClick={() => setConfirmingMove(false)} className="tap"
                style={{ background: "transparent", border: "1px solid var(--line)", color: "var(--text-dim)", padding: "5px 10px", borderRadius: 5, fontSize: 11.5 }}>Cancel</button>
            </div>
          ) : (
            <button onClick={() => setConfirmingMove(true)}
              className="tap" style={{ display: "flex", alignItems: "center", gap: 5, background: "transparent", border: "1px solid var(--line)", color: "var(--text-dim)", padding: "5px 10px", borderRadius: 5, fontSize: 11.5 }}>
              <RefreshCcw size={12} /> Move back to To Be Called
            </button>
          ))}
                </div>
        <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
          <OutcomeBtn icon={Check} label="Yes" color="var(--green)" onClick={() => onOutcome("yes")} />
          <OutcomeBtn icon={Clock} label="Call Me Back" color="var(--amber)" onClick={() => onOutcome("callMeBack")} />
          <OutcomeBtn icon={X} label="No" color="var(--red)" onClick={() => onOutcome("no")} />
          {!["whatsapp", "linkedin", "email"].includes(crm.bucket) && (
            <OutcomeBtn icon={PhoneMissed} label="No Answer" color="var(--gray)" onClick={() => onOutcome("noAnswer")} />
          )}
        </div>
        {crm.bucket === "unreachable" && (
          <div style={{ marginTop: 18 }}>
            <div style={{ fontSize: 12, fontWeight: 600, color: "var(--text-dim)", marginBottom: 8 }}>FOLLOW UP THROUGH</div>
            <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
              <OutcomeBtn icon={MessageCircle} label="WhatsApp" color="var(--green)" onClick={() => onOutcome("whatsapp")} />
              <OutcomeBtn icon={Mail} label="Email" color="var(--gray)" onClick={() => onOutcome("email")} />
              <OutcomeBtn icon={Briefcase} label="LinkedIn" color="var(--blue)" onClick={() => onOutcome("linkedin")} />
            </div>
          </div>
        )}
        {crm.lastOutcome && (
          <div style={{ fontSize: 11.5, color: "var(--text-faint)", marginTop: 8 }}>
            Last logged: {BUCKET_LABEL[crm.bucket] || crm.lastOutcome} on {fmtDate(crm.lastCallDate)}
            {crm.nextActionDate && ` · next action ${fmtDate(crm.nextActionDate)}`}
          </div>
        )}
      </div>
    </div>
  );
}

function ScoreBar({ label, value, colorVar }) {
  return (
    <div>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "baseline", marginBottom: 6 }}>
        <span className="eyebrow">{label}</span>
        <span className="mono" style={{ color: colorVar, fontSize: 13 }}>{value}</span>
      </div>
      <div style={{ height: 3, background: "var(--line)", overflow: "hidden" }}>
        <div style={{ width: `${value}%`, height: "100%", background: colorVar, transition: "width 0.3s" }} />
      </div>
    </div>
  );
}

function OutcomeBtn({ icon: Icon, label, color, onClick }) {
  return (
    <button onClick={onClick} className="tap sweep-btn" style={{
      display: "flex", alignItems: "center", gap: 7, background: "transparent", border: "1px solid var(--line)",
      color, padding: "11px 14px", borderRadius: 5, fontSize: 10.5, fontWeight: 600,
      textTransform: "uppercase", letterSpacing: "0.1em", "--sweep-color": color
    }}>
      <span style={{ position: "relative", zIndex: 1, display: "flex", alignItems: "center", gap: 7 }}>
        <Icon size={14} /> {label}
      </span>
    </button>
  );
}

/* ---------------------------------------------------------
   LOOKUP TOOL (standalone, informational only)
--------------------------------------------------------- */
function LookupTool({ myUnits }) {
  const [q, setQ] = useState("");
  const results = useMemo(() => {
    if (!q.trim()) return [];
    const query = q.trim().toLowerCase();
    return myUnits.filter(u =>
      u.unitId.toLowerCase().includes(query) ||
      (u.ownerName || "").toLowerCase().includes(query)
    ).slice(0, 20);
  }, [q, myUnits]);

  return (
    <div style={{ flex: 1, padding: 24, overflow: "auto" }}>
      <div style={{ maxWidth: 560, margin: "0 auto" }}>
        <div style={{ position: "relative", marginBottom: 20 }}>
          <Search size={16} style={{ position: "absolute", left: 14, top: 13, color: "var(--text-faint)" }} />
          <input value={q} onChange={e => setQ(e.target.value)} placeholder="Search by unit number or owner name…" autoFocus
            className="tap" style={{ width: "100%", background: "var(--panel)", border: "1px solid var(--line)", borderRadius: 6, padding: "12px 14px 12px 40px", color: "var(--text)", fontSize: 14 }} />
        </div>

        {q.trim() && results.length === 0 && (
          <div style={{ color: "var(--text-faint)", fontSize: 13.5, textAlign: "center", marginTop: 20 }}>No match in your assigned units.</div>
        )}

        <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
          {results.map(u => {
            const isRented = u.status === "occupied";
            const isVacant = u.status === "vacant";
            return (
              <div key={u.key} style={{
                background: "var(--panel)", border: `1px solid ${isRented ? "var(--red)" : "var(--line)"}`,
                borderRadius: 6, padding: 16
              }}>
                <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                  <div className="mono" style={{ fontSize: 15, fontWeight: 700 }}>Unit {u.unitId}</div>
                  <span style={{
                    fontSize: 11, fontWeight: 700, padding: "3px 9px", borderRadius: 5,
                    background: isRented ? "var(--red-dim)" : isVacant ? "var(--green-dim)" : "var(--gray-dim)",
                    color: isRented ? "var(--red)" : isVacant ? "var(--green)" : "var(--gray)"
                  }}>
                    {isRented ? "RENTED OUT" : isVacant ? "VACANT" : "STATUS UNKNOWN"}
                  </span>
                </div>
                <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10, marginTop: 12 }}>
                  <Field label="Owner" value={u.ownerName} />
                  <Field label="Contact" value={u.ownerContact} />
                  <Field label="Secondary Number" value={u.secondaryContact} emptyText="No secondary number" />
                  <Field label="Email" value={u.email} emptyText="No email" />
                  <Field label="Unit Type" value={u.unitType} />
                  <Field label="Bedrooms" value={u.bedrooms} />
                  <Field label="Built-up Area" value={u.carpetArea ? `${u.carpetArea} sqft` : ""} />
                 <Field label="Current Rent" value={u.currentRent ? `AED ${Number(u.currentRent).toLocaleString()}` : ""} />
                  <Field label="Current Market Price" value={u.marketPrice ? `AED ${Number(u.marketPrice).toLocaleString()}` : ""} />
                </div>
                <div style={{ fontSize: 10.5, color: "var(--text-faint)", marginTop: 10 }}>Reference only. Outcomes can't be logged from lookup. Open the CRM queue to call this unit.</div>
              </div>
            );
          })}
        </div>
      </div>
    </div>
  );
}

/* ---------------------------------------------------------
   BUILDING EXPLORER — visual floor-by-floor browser.
   Admin sees every building. Agents see only buildings they're
   assigned to at the building level (full access), plus any
   individual units assigned to them via override in buildings
   they otherwise have no access to (restricted to just those units).
--------------------------------------------------------- */
function buildingExplorerAccess(state, user, role) {
  const buildings = Object.values(state.buildings);
  if (role === "admin") {
    return buildings.map(b => ({ id: b.id, name: b.name, restricted: false, allowedKeys: null }));
  }
  const out = [];
  buildings.forEach(b => {
    const buildingLevel = (state.assignments[b.id] || []).includes(user);
    if (buildingLevel) { out.push({ id: b.id, name: b.name, restricted: false, allowedKeys: null }); return; }
    const overrideUnits = Object.values(state.units).filter(u => u.buildingId === b.id && u.assignedTo && u.assignedTo.includes(user));
    if (overrideUnits.length) {
      out.push({ id: b.id, name: b.name, restricted: true, allowedKeys: new Set(overrideUnits.map(u => u.key)) });
    }
  });
  return out;
}

function sortFloors(floors) {
  return floors.slice().sort((a, b) => {
    const na = parseFloat(a), nb = parseFloat(b);
    const aNum = !isNaN(na), bNum = !isNaN(nb);
    if (aNum && bNum) return nb - na;
    if (aNum) return -1;
    if (bNum) return 1;
    return a.localeCompare(b);
  });
}
function explorerDataLevel(u) {
  if (u.ownerName && u.ownerContact) return "full";
  if (u.ownerName || u.ownerContact) return "partial";
  return "none";
}
function explorerBedMatch(u, filter) {
  if (filter === "all") return true;
  const beds = Number(u.bedrooms);
  if (filter === "4plus") return beds >= 4;
  return beds === Number(filter);
}

function BuildingExplorer({ state, user, role }) {
  const access = useMemo(() => buildingExplorerAccess(state, user, role), [state.buildings, state.units, state.assignments, user, role]);
  const [buildingId, setBuildingId] = useState(access[0]?.id || null);
  const [floor, setFloor] = useState(null);
  const [buildingFilter, setBuildingFilter] = useState("all");
  const [floorFilter, setFloorFilter] = useState("all");
  const [searchOpen, setSearchOpen] = useState(false);
  const [searchQ, setSearchQ] = useState("");
  const [jumpedKey, setJumpedKey] = useState(null);
  const [modalUnit, setModalUnit] = useState(null);
  const searchRef = useRef(null);

  useEffect(() => {
    if (!access.length) return;
    if (!access.find(a => a.id === buildingId)) setBuildingId(access[0].id);
  }, [access]);

  useEffect(() => {
    const onDoc = (e) => { if (searchRef.current && !searchRef.current.contains(e.target)) setSearchOpen(false); };
    document.addEventListener("mousedown", onDoc);
    return () => document.removeEventListener("mousedown", onDoc);
  }, []);

  const currentAccess = access.find(a => a.id === buildingId) || access[0];

  const unitsInBuilding = useMemo(() => {
    if (!currentAccess) return [];
    return Object.values(state.units).filter(u => u.buildingId === currentAccess.id && (!currentAccess.restricted || currentAccess.allowedKeys.has(u.key)));
  }, [state.units, currentAccess]);

  const byFloor = {};
  unitsInBuilding.forEach(u => {
    const f = u.floor && String(u.floor).trim() ? String(u.floor).trim() : "—";
    (byFloor[f] = byFloor[f] || []).push(u);
  });
  const numericFloors = sortFloors(Object.keys(byFloor).filter(f => f !== "—"));
  const floorOrder = byFloor["—"] ? [...numericFloors, "—"] : numericFloors;

  useEffect(() => {
    if (!floorOrder.length) { setFloor(null); return; }
    if (!floorOrder.includes(floor)) setFloor(floorOrder[0]);
  }, [buildingId, floorOrder.join(",")]);

  if (!access.length) {
    return (
      <div style={{ flex: 1, display: "flex", alignItems: "center", justifyContent: "center", color: "var(--text-faint)", fontSize: 13.5, textAlign: "center", padding: 24 }}>
        No buildings assigned to you yet — you'll see this once the admin assigns you a building or specific units.
      </div>
    );
  }

  const runSearch = (q) => {
    const query = q.trim().toLowerCase();
    if (!query) return [];
    const out = [];
    access.forEach(a => {
      Object.values(state.units)
        .filter(u => u.buildingId === a.id && (!a.restricted || a.allowedKeys.has(u.key)))
        .forEach(u => {
          const hit = (u.ownerName || "").toLowerCase().includes(query) ||
            (u.ownerContact || "").toLowerCase().includes(query) ||
            (u.unitId || "").toLowerCase().includes(query);
          if (hit) out.push({ ...u, _buildingId: a.id, _buildingName: a.name });
        });
    });
    return out.slice(0, 8);
  };
  const searchResults = runSearch(searchQ);

  const goToUnit = (u, highlight) => {
    if (u.buildingId !== buildingId) setBuildingId(u.buildingId);
    const f = u.floor && String(u.floor).trim() ? String(u.floor).trim() : "—";
    setFloor(f);
    setFloorFilter("all");
    setSearchOpen(false);
    setSearchQ("");
    if (highlight) {
      setJumpedKey(u.key);
      setTimeout(() => setJumpedKey(null), 2200);
    }
  };

  const floorUnits = byFloor[floor] || [];
  const floorMatchCount = floorFilter !== "all" ? floorUnits.filter(u => explorerBedMatch(u, floorFilter)).length : null;

  return (
    <div style={{ flex: 1, overflow: "auto", padding: 24 }}>
      <div style={{ maxWidth: 1000, margin: "0 auto" }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", gap: 12, marginBottom: 16, flexWrap: "wrap" }}>
          <div style={{ display: "flex", gap: 6, flexWrap: "wrap" }}>
            {access.map(a => (
              <button key={a.id} onClick={() => { setBuildingId(a.id); setBuildingFilter("all"); setFloorFilter("all"); }} className="tap"
                style={{ padding: "8px 14px", borderRadius: 6, fontSize: 12.5, background: buildingId === a.id ? "var(--accent-dim)" : "var(--panel-2)", border: buildingId === a.id ? "1px solid var(--accent)" : "1px solid var(--line)", color: "var(--text)", display: "flex", alignItems: "center", gap: 6 }}>
                {a.name}{a.restricted && <span style={{ fontSize: 9.5, color: "var(--text-faint)", textTransform: "uppercase", letterSpacing: "0.06em" }}>· partial</span>}
              </button>
            ))}
          </div>
          <div style={{ position: "relative" }} ref={searchRef}>
            <button onClick={() => setSearchOpen(o => !o)} className="tap" title="Search owner / contact / unit"
              style={{ width: 36, height: 36, display: "flex", alignItems: "center", justifyContent: "center", background: "var(--panel)", border: searchOpen ? "1px solid var(--accent)" : "1px solid var(--line)", borderRadius: 8, color: searchOpen ? "var(--accent)" : "var(--text-dim)", cursor: "pointer" }}>
              <Search size={16} />
            </button>
            {searchOpen && (
              <div style={{ position: "absolute", top: "calc(100% + 6px)", right: 0, width: 300, background: "var(--panel)", border: "1px solid var(--line)", borderRadius: 10, boxShadow: "0 10px 30px rgba(0,0,0,0.45)", zIndex: 40, padding: 10 }}>
                <input value={searchQ} onChange={e => setSearchQ(e.target.value)} placeholder="Owner, contact, or unit…" autoFocus
                  className="tap" style={{ width: "100%", background: "var(--panel-2)", border: "1px solid var(--line)", borderRadius: 6, padding: "9px 11px", color: "var(--text)", fontSize: 13 }} />
                <div style={{ marginTop: 8, maxHeight: 260, overflowY: "auto" }}>
                  {searchQ.trim() && searchResults.length === 0 && (
                    <div style={{ padding: "14px 8px", fontSize: 12, color: "var(--text-faint)", textAlign: "center" }}>No match in your accessible units.</div>
                  )}
                  {searchResults.map(u => (
                    <div key={u.key} style={{ padding: "10px 8px", borderBottom: "1px solid var(--line)", display: "flex", justifyContent: "space-between", alignItems: "center", gap: 8 }}>
                      <div style={{ minWidth: 0, cursor: "pointer" }} onClick={() => goToUnit(u, false)}>
                        <div style={{ fontSize: 13, fontWeight: 600 }}>{u.ownerName || "Owner not on file"}</div>
                        <div style={{ fontSize: 10.5, color: "var(--text-dim)", marginTop: 2 }}>{u._buildingName} · Unit {u.unitId} · Floor {u.floor || "—"}</div>
                      </div>
                      <button onClick={() => goToUnit(u, true)} className="tap" style={{ fontFamily: "monospace", fontSize: 10, background: "transparent", border: "1px solid var(--line)", color: "var(--accent)", padding: "6px 9px", borderRadius: 5, cursor: "pointer", flexShrink: 0 }}>
                        Jump →
                      </button>
                    </div>
                  ))}
                </div>
              </div>
            )}
          </div>
        </div>

        {currentAccess?.restricted && (
          <div style={{ background: "var(--panel-2)", border: "1px solid var(--line)", borderRadius: 6, padding: "9px 12px", fontSize: 12, color: "var(--text-dim)", marginBottom: 16 }}>
            You're only assigned specific units in {currentAccess.name}, not the whole building — showing {unitsInBuilding.length} unit(s) assigned to you here.
          </div>
        )}

        <div className="eyebrow" style={{ marginBottom: 8 }}>Highlight by size — across the whole building</div>
        <div style={{ display: "flex", gap: 8, marginBottom: 20, flexWrap: "wrap" }}>
          {[["all", "All Units"], ["1", "1 Bed"], ["2", "2 Bed"], ["3", "3 Bed"], ["4plus", "4+ Bed"]].map(([id, label]) => (
            <button key={id} onClick={() => setBuildingFilter(id)} className="tap"
              style={{ fontFamily: "monospace", fontSize: 11, padding: "7px 14px", borderRadius: 20, border: buildingFilter === id ? "1px solid var(--accent)" : "1px solid var(--line)", background: buildingFilter === id ? "var(--accent)" : "var(--panel)", color: buildingFilter === id ? "#fff" : "var(--text-dim)", fontWeight: buildingFilter === id ? 600 : 400 }}>
              {label}
            </button>
          ))}
        </div>

        <div style={{ display: "grid", gridTemplateColumns: "200px 1fr", gap: 20, alignItems: "start" }}>
          <div style={{ background: "var(--panel)", border: "1px solid var(--line)", borderRadius: 10, padding: "12px 10px" }}>
            <div className="eyebrow" style={{ marginBottom: 10, paddingLeft: 2 }}>Floors</div>
            <div style={{ maxHeight: 560, overflowY: "auto" }}>
              {floorOrder.map(f => (
                <div key={f} onClick={() => { setFloor(f); setFloorFilter("all"); }}
                  style={{ display: "flex", alignItems: "center", gap: 8, padding: "6px 6px", borderRadius: 6, cursor: "pointer", background: f === floor ? "var(--panel-2)" : "transparent", border: f === floor ? "1px solid var(--accent)" : "1px solid transparent" }}>
                  <div style={{ width: 26, fontFamily: "monospace", fontSize: 10.5, color: f === floor ? "var(--accent)" : "var(--text-dim)", fontWeight: f === floor ? 600 : 400, textAlign: "right", flexShrink: 0 }}>{f}</div>
                  <div style={{ display: "flex", gap: 2, flex: 1 }}>
                    {byFloor[f].map(u => {
                      let cls = "";
                      if (explorerDataLevel(u) !== "none") cls = "has-data";
                      if (buildingFilter !== "all" && explorerBedMatch(u, buildingFilter)) cls = "match";
                      return <div key={u.key} className={`explorer-window ${cls}`} />;
                    })}
                  </div>
                </div>
              ))}
            </div>
          </div>

          <div style={{ background: "var(--panel)", border: "1px solid var(--line)", borderRadius: 10, padding: "20px 22px 24px", minHeight: 420 }}>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "baseline", flexWrap: "wrap", gap: 8, marginBottom: 4 }}>
              <div className="disp" style={{ fontSize: 20, fontWeight: 800 }}>Floor {floor}</div>
              <span className="mono" style={{ color: "var(--text-dim)", fontSize: 11.5 }}>{floorUnits.length} unit(s)</span>
            </div>
            <div style={{ fontSize: 12, color: "var(--text-dim)", marginBottom: 16 }}>Tap a unit to view owner, contact, and rental status on file.</div>

            <div style={{ display: "flex", alignItems: "center", gap: 8, flexWrap: "wrap", paddingBottom: 16, marginBottom: 16, borderBottom: "1px solid var(--line)" }}>
              <span className="eyebrow" style={{ marginRight: 2 }}>Highlight on this floor</span>
              {[["all", "All"], ["1", "1 Bed"], ["2", "2 Bed"], ["3", "3 Bed"], ["4plus", "4+ Bed"]].map(([id, label]) => (
                <button key={id} onClick={() => setFloorFilter(id)} className="tap"
                  style={{ fontFamily: "monospace", fontSize: 11, padding: "7px 14px", borderRadius: 20, border: floorFilter === id ? "1px solid var(--accent)" : "1px solid var(--line)", background: floorFilter === id ? "var(--accent)" : "var(--panel)", color: floorFilter === id ? "#fff" : "var(--text-dim)", fontWeight: floorFilter === id ? 600 : 400 }}>
                  {label}
                </button>
              ))}
            </div>
            {floorMatchCount === 0 && (
              <div style={{ fontSize: 12.5, color: "var(--text-faint)", fontStyle: "italic", marginBottom: 12 }}>No units on Floor {floor} match that size.</div>
            )}

            <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(120px, 1fr))", gap: 12 }}>
              {floorUnits.map(u => {
                const buildingActive = buildingFilter !== "all";
                const floorActive = floorFilter !== "all";
                const matchesBuilding = explorerBedMatch(u, buildingFilter);
                const matchesFloor = explorerBedMatch(u, floorFilter);
                let cls = "explorer-unit-card";
                if ((buildingActive && matchesBuilding) || (floorActive && matchesFloor)) cls += " match";
                else if ((buildingActive && !matchesBuilding) || (floorActive && !matchesFloor)) cls += " dim";
                if (jumpedKey === u.key) cls += " jumped";
                const level = explorerDataLevel(u);
                return (
                  <div key={u.key} className={cls} onClick={() => setModalUnit(u)}>
                    {u.status === "occupied" && <div className="explorer-rented-x" title="Currently rented">&times;</div>}
                    <div className={`explorer-dot ${level}`} />
                    <div className="mono" style={{ fontSize: 12, color: "var(--text-dim)" }}>{u.unitId}</div>
                    <div style={{ fontSize: 12.5, fontWeight: 600, margin: "6px 0 2px" }}>{u.bedrooms ? `${u.bedrooms} Bed` : "—"}</div>
                    <div className="mono" style={{ fontSize: 10.5, color: "var(--text-dim)" }}>{u.carpetArea ? `${u.carpetArea} sqft` : "no data"}</div>
                  </div>
                );
              })}
            </div>
          </div>
        </div>

        {modalUnit && (
          <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.72)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 100, padding: 20 }} onClick={() => setModalUnit(null)}>
            <div style={{ background: "var(--panel)", border: "1px solid var(--line)", borderRadius: 12, maxWidth: 420, width: "100%", padding: "26px 26px 22px", position: "relative" }} onClick={e => e.stopPropagation()}>
              <button onClick={() => setModalUnit(null)} className="tap" style={{ position: "absolute", top: 14, right: 16, background: "none", border: "none", color: "var(--text-dim)", fontSize: 20, cursor: "pointer" }}>&times;</button>
              <div className="mono" style={{ color: "var(--accent)", fontSize: 11.5, letterSpacing: "0.05em" }}>UNIT {modalUnit.unitId} · FLOOR {modalUnit.floor || "—"}</div>
              <div className="disp" style={{ fontSize: 22, fontWeight: 800, margin: "6px 0 8px" }}>
                {modalUnit.bedrooms ? `${modalUnit.bedrooms} Bedroom` : "Unit"}{modalUnit.carpetArea ? ` · ${Number(modalUnit.carpetArea).toLocaleString()} sqft` : ""}
              </div>
              {modalUnit.status === "occupied" && (
                <div style={{ display: "inline-block", fontSize: 10.5, color: "var(--text-faint)", border: "1px solid var(--line)", borderRadius: 12, padding: "3px 10px", marginBottom: 14 }}>&times; Currently rented</div>
              )}
              <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 14, marginBottom: 16 }}>
                <Field label="Unit Type" value={modalUnit.unitType} />
                <Field label="Owner" value={modalUnit.ownerName} />
                <Field label="Contact" value={modalUnit.ownerContact} />
                <Field label="Current Rent" value={modalUnit.currentRent ? `AED ${Number(modalUnit.currentRent).toLocaleString()}` : ""} />
                <Field label="Current Market Price" value={modalUnit.marketPrice ? `AED ${Number(modalUnit.marketPrice).toLocaleString()}` : ""} />
              </div>
            </div>
          </div>
        )}
      </div>
    </div>
  );
}

