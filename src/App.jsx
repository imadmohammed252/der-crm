import { useState, useEffect, useMemo, useRef } from "react";
import Papa from "papaparse";
import * as XLSX from "xlsx";
import {
  Building2, Search, Upload, LogOut, ArrowRightLeft, Phone,
  PhoneMissed, PhoneOff, Check, X, Clock, ChevronRight, RefreshCcw,
  Users, LayoutGrid, ClipboardList, AlertTriangle, Trash2, Settings
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
.rcrm input, .rcrm textarea, .rcrm select{ font-family:inherit; box-sizzing: border-box;}
.rcrm input:focus, .rcrm textarea:focus, .rcrm select:focus{ border-color:var(--accent); }
.rcrm *:focus-visible{ outline:1px solid var(--accent); outline-offset:2px; }
.tap{ min-height:44px; }
`;

/* ---------------------------------------------------------
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
  floor: ["floor", "floor level"],
  carpetArea: ["carpet area", "area", "sqft", "sq ft", "size", "built-up area (sqft)", "built-up area"],
  bedrooms: ["bedrooms", "bed", "beds", "bhk"],
  bathrooms: ["bathrooms", "bath", "baths"],
  ownerName: ["owner name", "owner", "name"],
  ownerContact: ["owner contact", "contact", "phone", "owner phone", "mobile"],
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
  noAnswer: "No Answer",
  declined: "Declined",
  unreachable: "Unreachable",
  yes: "Active Client",
};
const BUCKET_COLOR = {
  toBeCalled: "var(--accent)",
  callMeBack: "var(--amber)",
  noAnswer: "var(--gray)",
  declined: "var(--red)",
  unreachable: "var(--violet)",
  yes: "var(--green)",
};

function defaultCrm() {
  return { bucket: "toBeCalled", notes: "", noAnswerCount: 0, nextActionDate: null, tag: null, lastOutcome: null, lastCallDate: null };
}

// Where does this unit actually show up right now, given dates?
function effectiveBucket(crm) {
  if (!crm) return { display: "toBeCalled", resurfaced: false };
  if (crm.bucket === "toBeCalled") return { display: "toBeCalled", resurfaced: false };
  if (crm.bucket === "yes") return { display: "yes", resurfaced: false };
  if (["callMeBack", "noAnswer", "declined", "unreachable"].includes(crm.bucket) && isDue(crm.nextActionDate)) {
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
      c.bucket = "unreachable"; c.nextActionDate = addMonths(now, 6).toISOString(); c.tag = "unreachable";
    } else {
      c.bucket = "noAnswer"; c.nextActionDate = addDays(now, 3).toISOString();
    }
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
      return { buildings: {}, units: {}, assignments: {}, crm: {}, users: DEFAULT_USERS, ...parsed };
    }
  } catch (e) {
    console.error("loadState failed, using empty state:", e);
  }
  return { buildings: {}, units: {}, assignments: {}, crm: {}, users: DEFAULT_USERS };
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
  const [loggedInUsername, setLoggedInUsername] = useState(null);
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
    return <LoginScreen users={state.users} onLogin={setLoggedInUsername} />;
  }

  return (
    <div className="rcrm" style={{ minHeight: 640, background: "var(--bg)", display: "flex", flexDirection: "column" }}>
      <style>{THEME}</style>
      {currentUser.role === "admin"
        ? <AdminApp state={state} setState={persist} onLogout={() => setLoggedInUsername(null)} saving={saving} />
        : <UserApp state={state} setState={persist} user={currentUser.username} onLogout={() => setLoggedInUsername(null)} saving={saving} />}
    </div>
  );
}

/* ---------------------------------------------------------
   LOGIN
--------------------------------------------------------- */
function LoginScreen({ users, onLogin }) {
  const [username, setUsername] = useState("");
  const [passkey, setPasskey] = useState("");
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
            <input type="password" value={passkey} onChange={e => { setPasskey(e.target.value); setError(""); }} onKeyDown={onKeyDown}
              className="tap" style={{ width: "100%", background: "transparent", border: "none", borderBottom: "1px solid var(--line)", color: "var(--text)", padding: "8px 0", fontSize: 15, outline: "none" }} />
          </div>
          {error && (
            <div style={{ borderLeft: "2px solid var(--accent)", background: "var(--accent-dim)", color: "var(--text)", padding: "10px 12px", fontSize: 12.5 }}>
              {error}
            </div>
          )}
          <button type="button" onClick={submit} className="tap"
            style={{ marginTop: 8, background: "var(--accent)", color: "#fff", border: "none", borderRadius: 5, padding: "13px", fontSize: 12, fontWeight: 600, textTransform: "uppercase", letterSpacing: "0.14em" }}>
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
                <button key={i} onClick={() => { it.onClick(); setSettingsOpen(false); }} className="tap"
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

  const wipeAll = () => setState({ buildings: {}, units: {}, assignments: {}, crm: {}, users: state.users });

  return (
    <>
      <TopBar title="Admin Portal" subtitle="Upload data, assign buildings and units to agents" onLogout={onLogout} saving={saving}
        extraMenuItems={[{ label: "Wipe data", icon: Trash2, onClick: () => setConfirmingWipe(true) }]} />
      {confirmingWipe && (
        <div style={{ background: "var(--accent-dim)", borderBottom: "1px solid var(--accent)", padding: "10px 20px", display: "flex", alignItems: "center", justifyContent: "space-between", gap: 12, flexWrap: "wrap" }}>
          <span style={{ fontSize: 12.5, color: "var(--text)" }}>Wipe every building, unit, and CRM record? User logins will stay.</span>
          <div style={{ display: "flex", gap: 6 }}>
            <button onClick={() => { wipeAll(); setConfirmingWipe(false); }} className="tap"
              style={{ background: "var(--accent)", border: "none", color: "#fff", padding: "7px 12px", borderRadius: 6, fontSize: 11, fontWeight: 600, textTransform: "uppercase", letterSpacing: "0.1em" }}>Yes, wipe</button>
            <button onClick={() => setConfirmingWipe(false)} className="tap"
              style={{ background: "transparent", border: "1px solid var(--line)", color: "var(--text-dim)", padding: "7px 12px", borderRadius: 6, fontSize: 11, textTransform: "uppercase", letterSpacing: "0.1em" }}>Cancel</button>
          </div>
        </div>
      )}
      <div style={{ display: "flex", borderBottom: "1px solid var(--line)", background: "var(--panel)" }}>
        {[["buildings", "Buildings & Upload"], ["assign", "Assignments"], ["users", "Users"]].map(([id, label]) => (
          <button key={id} onClick={() => setTab(id)} className="tap"
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
      </div>
    </>
  );
}

function BuildingsUpload({ state, setState }) {
  const [newBuildingName, setNewBuildingName] = useState("");
  const [targetBuilding, setTargetBuilding] = useState("");
  const [unitFile, setUnitFile] = useState(null);
  const [ownerFile, setOwnerFile] = useState(null);
  const [unitHeaders, setUnitHeaders] = useState([]);
  const [ownerHeaders, setOwnerHeaders] = useState([]);
  const [unitMap, setUnitMap] = useState({});
  const [ownerMap, setOwnerMap] = useState({});
  const [log, setLog] = useState("");

  const buildings = Object.values(state.buildings);

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
        ownershipStart: getMapped(r, ownerMap, "ownershipStart"),
        purchasePrice: getMapped(r, ownerMap, "purchasePrice"),
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
        key, unitId: uid, buildingId: targetBuilding,
        floor: pick(getMapped(r, unitMap, "floor"), "floor"),
        carpetArea: pick(getMapped(r, unitMap, "carpetArea"), "carpetArea"),
        bedrooms: pick(getMapped(r, unitMap, "bedrooms"), "bedrooms"),
        bathrooms: pick(getMapped(r, unitMap, "bathrooms"), "bathrooms"),
        ownerName: owner.ownerName || existing.ownerName || "",
        ownerContact: owner.ownerContact || existing.ownerContact || "",
        ownershipStart: owner.ownershipStart || existing.ownershipStart || "",
        purchasePrice: owner.purchasePrice || existing.purchasePrice || "",
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
    unitId: "Unit ID / Unit Number *", floor: "Floor", carpetArea: "Carpet Area",
    bedrooms: "Bedrooms", bathrooms: "Bathrooms", currentRent: "Current Rent",
    currentLeaseStart: "Current Lease Start", currentLeaseEnd: "Current Lease End",
    status: "Status (occupied/vacant)",
    ownerName: "Owner Name", ownerContact: "Owner Contact", ownershipStart: "Ownership Start", purchasePrice: "Purchase Price",
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
              <X size={13} /> Clear
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
          fieldList={["unitId", "floor", "carpetArea", "bedrooms", "bathrooms", "currentRent", "currentLeaseStart", "currentLeaseEnd", "status"]}
          map={unitMap} setMap={setUnitMap} />

        <UploadBlock title="Owner data (optional)" hint="Must also include a Unit ID column so it can join to the unit file."
          file={ownerFile} setter={setOwnerFile} headers={ownerHeaders} setHeaders={setOwnerHeaders}
          fieldList={["unitId", "ownerName", "ownerContact", "ownershipStart", "purchasePrice"]}
          map={ownerMap} setMap={setOwnerMap} />

        <button onClick={runImport} disabled={!targetBuilding} className="tap" style={{ marginTop: 6, background: "var(--accent)", color: "#fff", border: "none", borderRadius: 5, padding: "10px 18px", fontSize: 13.5, fontWeight: 600 }}>
          Import into this building
        </button>
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

function AssignmentPanel({ state, setState }) {
  const buildings = Object.values(state.buildings);
  const units = Object.values(state.units);
  const agents = Object.values(state.users || {}).filter(u => u.role === "agent");

  const toggleBuildingAgent = (buildingId, username) => {
    const current = state.assignments[buildingId] || [];
    const next = current.includes(username) ? current.filter(u => u !== username) : [...current, username];
    setState({ ...state, assignments: { ...state.assignments, [buildingId]: next } });
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
    setState({ ...state, crm: { ...state.crm, [key]: crm } });
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
              {[["detail", "Call Detail"], ["lookup", "Lookup"]].map(([id, label]) => (
                <button key={id} onClick={() => setRightPanel(id)} className="tap"
                  style={{
                    padding: "11px 16px", background: "transparent", border: "none",
                    borderBottom: rightPanel === id ? "1px solid var(--accent)" : "1px solid transparent",
                    color: rightPanel === id ? "var(--text)" : "var(--text-dim)", fontSize: 10, fontWeight: 600,
                    textTransform: "uppercase", letterSpacing: "0.12em",
                    display: "flex", alignItems: "center", gap: 6
                  }}>
                  {id === "lookup" && <Search size={13} />} {label}
                </button>
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
              ) : (
                <LookupTool state={state} myUnits={myUnits} />
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
  ["noAnswer", "No Answer"],
  ["reject", "Reject"],
  ["yes", "Active"],
];

const SORT_OPTIONS = [
  ["unitAsc", "Unit # ↑"],
  ["unitDesc", "Unit # ↓"],
  ["easy", "Easy to Rent"],
  ["hard", "Hard to Rent"],
];

function compareUnitId(a, b) {
  const na = parseFloat(a), nb = parseFloat(b);
  if (!isNaN(na) && !isNaN(nb) && na !== nb) return na - nb;
  return a.localeCompare(b, undefined, { numeric: true });
}

function QueuePane({ bucketTab, setBucketTab, units, crmMap, activeUnitKey, setActiveUnitKey, onMoveToQueue }) {
  const [sortMode, setSortMode] = useState("easy");
  const [filterOpen, setFilterOpen] = useState(false);
  const [searchOpen, setSearchOpen] = useState(false);
  const [searchQuery, setSearchQuery] = useState("");

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
      // When search is active, scope is the whole CRM (every bucket), matched by unit id or owner name
      return r.unit.unitId.toLowerCase().includes(q) || (r.unit.ownerName || "").toLowerCase().includes(q);
    }
    if (bucketTab === "reject") return ["declined", "unreachable"].includes(r.eff.display);
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
      ? rows.filter(r => ["declined", "unreachable"].includes(r.eff.display)).length
      : rows.filter(r => r.eff.display === id).length
  ]));

  const activeSortLabel = SORT_OPTIONS.find(([id]) => id === sortMode)?.[1];

  return (
    <div style={{ width: "38%", minWidth: 300, borderRight: "1px solid var(--line)", display: "flex", flexDirection: "column", background: "var(--panel)" }}>
      <div style={{ display: "flex", flexWrap: "wrap", borderBottom: "1px solid var(--line)" }}>
        {TABS.map(([id, label]) => (
          <button key={id} onClick={() => setBucketTab(id)} className="tap"
            style={{
              flex: "1 0 auto", padding: "11px 8px", background: "transparent", border: "none",
              borderBottom: bucketTab === id ? "1px solid var(--accent)" : "1px solid transparent",
              color: bucketTab === id ? "var(--text)" : "var(--text-dim)", fontSize: 9.5, fontWeight: 600,
              textTransform: "uppercase", letterSpacing: "0.1em"
            }}>{label} <span style={{ color: bucketTab === id ? "var(--accent)" : "var(--text-faint)" }}>{counts[id]}</span></button>
        ))}
      </div>

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
              <span style={{ color: "var(--accent)", fontSize: 12 }}>{activeSortLabel}</span>
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
              back from {BUCKET_LABEL[eff.origin]}{eff.origin === "noAnswer" ? ` ×${crm.noAnswerCount}` : ""}
            </span>
          )}
          {crm.bucket !== "toBeCalled" && !eff.resurfaced && (
            <span style={{ fontSize: 10, padding: "2px 6px", borderRadius: 5, background: BUCKET_COLOR[crm.bucket] + "22", color: BUCKET_COLOR[crm.bucket] }}>
              {crm.bucket === "callMeBack" || crm.bucket === "noAnswer" ? `in ${daysUntil(crm.nextActionDate)}d` : BUCKET_LABEL[crm.bucket]}
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
function Field({ label, value }) {
  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 2 }}>
      <div className="eyebrow" style={{ color: "var(--text-faint)" }}>{label}</div>
      <div style={{ fontSize: 14, color: value ? "var(--text)" : "var(--text-faint)" }}>
        {value || <em>not available</em>}
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
        <Field label="Carpet Area" value={unit.carpetArea ? `${unit.carpetArea} sqft` : ""} />
        <Field label="Bed / Bath" value={unit.bedrooms || unit.bathrooms ? `${unit.bedrooms || "—"} bed / ${unit.bathrooms || "—"} bath` : ""} />
        <Field label="Current Rent" value={unit.currentRent ? `AED ${Number(unit.currentRent).toLocaleString()}` : ""} />
        <Field label="Current Lease Start" value={unit.currentLeaseStart ? fmtDate(unit.currentLeaseStart) : ""} />
        <Field label="Current Lease End" value={unit.currentLeaseEnd ? fmtDate(unit.currentLeaseEnd) : ""} />
        <Field label="Purchase Price" value={unit.purchasePrice ? `AED ${Number(unit.purchasePrice).toLocaleString()}` : ""} />
        <Field label="Ownership Since" value={unit.ownershipStart ? fmtDate(unit.ownershipStart) : ""} />
      </div>

      {/* Notes */}
      <div>
        <div style={{ fontSize: 12, fontWeight: 600, color: "var(--text-dim)", marginBottom: 8 }}>NOTES (persists across every stage)</div>
        <textarea value={crm.notes} onChange={e => onNotes(e.target.value)} placeholder="What does this owner need? What did they say last call?"
          style={{ width: "100%", minHeight: 90, background: "var(--panel)", border: "1px solid var(--line)", borderRadius: 6, padding: 12, color: "var(--text)", fontSize: 13, resize: "vertical" }} />
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
          <OutcomeBtn icon={PhoneMissed} label="No Answer" color="var(--gray)" onClick={() => onOutcome("noAnswer")} />
        </div>
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
    <button onClick={onClick} className="tap" style={{
      display: "flex", alignItems: "center", gap: 7, background: "transparent", border: `1px solid ${color}66`,
      color, padding: "11px 14px", borderRadius: 5, fontSize: 10.5, fontWeight: 600,
      textTransform: "uppercase", letterSpacing: "0.1em"
    }}>
      <Icon size={14} /> {label}
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
                  <Field label="Bedrooms" value={u.bedrooms} />
                  <Field label="Built-up Area" value={u.carpetArea ? `${u.carpetArea} sqft` : ""} />
                  <Field label="Current Rent" value={u.currentRent ? `AED ${Number(u.currentRent).toLocaleString()}` : ""} />
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
