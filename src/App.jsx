import { useState, useEffect, useMemo, useRef } from "react";
import { subscribeCollection, subscribeTemplate, subscribeConfigDoc, addDocs, removeDoc, removeDocs, saveTemplateDoc, saveConfigDoc, subscribeAuth, login, logout, setUserDoc, saveDocIn, changePassword, resetPassword, IS_TEST } from "./firebase.js";
import { fillOfficialSheet, fillForeignTrip, readCountryRates } from "./fillTemplate.js";

// ExcelJS se načítá až při práci se šablonou (lazy-load — zmenší úvodní bundle)
let _exceljs = null;
async function getExcelJS() {
  if (!_exceljs) {
    const mod = await import("exceljs/dist/exceljs.min.js");
    _exceljs = mod.default || mod;
  }
  return _exceljs;
}

// Šablony jsou uložené v databázi (za přihlášením), ne na veřejné adrese.
async function getTemplateBuffer(tpl) {
  if (tpl && tpl.b64) return b64ToBuf(tpl.b64);
  throw new Error("Šablona výkazu není nahraná — nahraje ji správce na záložce Export.");
}

// ---------- Pomocné funkce ----------
const today = () => new Date().toISOString().slice(0, 10);
const fmtDate = (iso) => {
  if (!iso) return "";
  const [y, m, d] = iso.split("-");
  return `${parseInt(d)}. ${parseInt(m)}. ${y}`;
};
const fmtDateShort = (iso) => {
  if (!iso) return "";
  const [, m, d] = iso.split("-");
  return `${parseInt(d)}. ${parseInt(m)}.`;
};
const fmtHours = (h) => {
  const n = Number(h) || 0;
  return n.toLocaleString("cs-CZ", { minimumFractionDigits: 0, maximumFractionDigits: 2 });
};
const monthKey = (iso) => (iso || "").slice(0, 7);

const MESICE = ["Leden", "Únor", "Březen", "Duben", "Květen", "Červen", "Červenec", "Srpen", "Září", "Říjen", "Listopad", "Prosinec"];

// Seznam činností přesně podle šablony (Tabulky!F2:F15)
const CINNOSTI = [
  "Práce",
  "Pracovní cesta ČR",
  "Práce + prac. cesta ČR",
  "Pracovní cesta zahraniční",
  "Práce + prac. cesta zahraniční",
  "Dovolená",
  "Dovolená 1/2",
  "Home Office",
  "Home Office 1/2",
  "Nemoc",
  "OČR",
  "Náhradní volno",
  "Neplacené volno",
  "Jiné",
];
const CESTA_TYPY = ["Pracovní cesta ČR", "Práce + prac. cesta ČR", "Pracovní cesta zahraniční", "Práce + prac. cesta zahraniční"];
const ABSENCE_TYPY = ["Dovolená", "Dovolená 1/2", "Nemoc", "OČR", "Náhradní volno", "Neplacené volno"];
const jeCesta = (c) => CESTA_TYPY.includes(c);

// ---- Oprávnění ----
// Moduly, ke kterým se dá účtu povolit nebo zakázat přístup
const MODULY = [
  { id: "vykazy", label: "Výkazy (docházka)", icon: "🕐" },
  { id: "cas", label: "Čas na projektech", icon: "⏱️" },
  { id: "export", label: "Export výkazů", icon: "📤" },
  { id: "zahr", label: "Zahraniční cesty", icon: "✈️" },
];
// Nové účty mají ve výchozím stavu docházku a čas; starší účty bez nastavení
// si ponechají přístup ke všemu, aby se jim nic neztratilo.
const VYCHOZI_PERMS = { vykazy: true, cas: true, export: false, zahr: false };
const permsUctu = (profile) =>
  profile.role === "admin"
    ? { vykazy: true, cas: true, export: true, zahr: true }
    : profile.perms || { vykazy: true, cas: true, export: true, zahr: true };

// ---- Zahraniční pracovní cesta ----
const DOPRAVA = ["Služební auto", "Letadlo", "Vlak", "Služební auto + letadlo"];
// Co se k cestě zapíše do měsíčního výkazu (formulář cesty je jeho příloha)
const ZAHR_DO_VYKAZU = [
  { v: "Pracovní cesta zahraniční", label: "Pracovní cesta zahraniční" },
  { v: "Práce + prac. cesta zahraniční", label: "Práce + prac. cesta zahraniční" },
  { v: "Neplacené volno", label: "Neplacené volno (platí zahraniční entita)" },
  { v: "", label: "Nezapisovat — doplním ručně" },
];

// Výpočet stravného přesně podle vzorců ve formuláři (sloupce H, I, J, K)
function vypocetDne(od, doC, jidla, sazba, kapesneDen) {
  const f1 = parseTimeToFraction(od);
  const f2 = parseTimeToFraction(doC);
  if (f1 == null || f2 == null) return { min: 0, diety: 0, kapesne: 0, celkem: 0 };
  const min = Math.round((f2 - f1) * 1440);
  const j = Number(jidla) || 0;
  const s = Number(sazba) || 0;
  const k = Number(kapesneDen) || 0;
  let h = 0, kap = 0;
  if (min > 1080) { h = s - 0.25 * j * s; kap = k; }
  else if (min > 720) { h = s / 1.5 - 0.35 * j * (s / 1.5); kap = k / 1.5; }
  else if (min > 299) { h = s / 3 - 0.7 * j * (s / 3); kap = k / 3; }
  const diety = Math.max(h, 0);
  return { min, diety, kapesne: kap, celkem: Math.round((diety + kap) * 100) / 100 };
}

// Kurz ČNB k datu (pokud je začátek o víkendu, použije se páteční kurz)
function kurzDatumProCestu(iso) {
  if (!iso) return iso;
  const d = new Date(iso + "T00:00:00Z");
  const dow = d.getUTCDay();
  if (dow === 6) d.setUTCDate(d.getUTCDate() - 1);
  else if (dow === 0) d.setUTCDate(d.getUTCDate() - 2);
  return d.toISOString().slice(0, 10);
}

// Kurzy stahuje jednou denně GitHub Action do souboru kurzy.json (ČNB nepovoluje
// volání přímo z prohlížeče). Když datum v souboru není, zkusí se ještě přímé volání.
let _kurzyCache = null;
let _kurzyChybi = false;
async function nactiKurzCNB(mena, datum) {
  if (!_kurzyCache && !_kurzyChybi) {
    try {
      const r = await fetch(import.meta.env.BASE_URL + "kurzy.json", { cache: "no-cache" });
      if (r.ok) _kurzyCache = await r.json();
      else _kurzyChybi = true;
    } catch (e) {
      _kurzyChybi = true;
      console.warn("kurzy.json se nepodařilo načíst:", e);
    }
  }
  // ČNB vyhlašuje kurzy jen v pracovní dny — když datum chybí, jdeme zpět až 7 dní
  if (_kurzyCache) {
    const d = new Date(datum + "T00:00:00Z");
    for (let i = 0; i < 8; i++) {
      const key = d.toISOString().slice(0, 10);
      const den = _kurzyCache[key];
      if (den && den[mena]) return { kurz: den[mena], datum: key };
      d.setUTCDate(d.getUTCDate() - 1);
    }
  }
  // Záložní pokus o přímé volání (funguje jen tam, kde to prohlížeč povolí)
  try {
    const res = await fetch(`https://api.cnb.cz/cnbapi/exrates/daily?date=${datum}&lang=CZ`);
    if (res.ok) {
      const data = await res.json();
      const row = (data.rates || []).find((r) => r.currencyCode === mena);
      if (row) return { kurz: Math.round((row.rate / (row.amount || 1)) * 1e8) / 1e8, datum };
    }
  } catch (e) {
    /* CORS nebo offline — spadneme do chyby níže */
  }
  if (_kurzyChybi) {
    throw new Error("Tabulka kurzů zatím není v aplikaci — spusťte v GitHubu workflow Kurzy ČNB. Zatím zadejte kurz ručně.");
  }
  throw new Error(`Kurz ${mena} k ${fmtDate(datum)} v tabulce není — zadejte ho ručně.`);
}
const jeAbsence = (c) => ABSENCE_TYPY.includes(c);

const P = {
  bg: "#F3F5F7",
  panel: "#FFFFFF",
  ink: "#1C2530",
  muted: "#5C6B7A",
  line: "#DDE3E9",
  accent: "#0E6E6B",
  accentSoft: "#E3F2F1",
  warn: "#B3541E",
};

// ---------- Responsivita ----------
function useIsMobile() {
  const [mobile, setMobile] = useState(() =>
    typeof window !== "undefined" ? window.matchMedia("(max-width: 767px)").matches : true
  );
  useEffect(() => {
    const mq = window.matchMedia("(max-width: 767px)");
    const fn = (e) => setMobile(e.matches);
    mq.addEventListener("change", fn);
    return () => mq.removeEventListener("change", fn);
  }, []);
  return mobile;
}

// ---------- Styly (mobile-first) ----------
const S = {
  app: {
    minHeight: "100vh",
    background: P.bg,
    color: P.ink,
    fontFamily: "'Segoe UI', -apple-system, 'Helvetica Neue', Arial, sans-serif",
    fontSize: 15,
    WebkitTapHighlightColor: "transparent",
  },
  mono: { fontFamily: "'SF Mono', 'Consolas', 'Menlo', monospace", fontVariantNumeric: "tabular-nums" },
  panel: {
    background: P.panel,
    border: `1px solid ${P.line}`,
    borderRadius: 12,
    padding: 16,
  },
  input: {
    padding: "12px 12px",
    border: `1px solid ${P.line}`,
    borderRadius: 8,
    fontSize: 16,
    background: "#fff",
    color: P.ink,
    outline: "none",
    minWidth: 0,
    width: "100%",
    boxSizing: "border-box",
    WebkitAppearance: "none",
  },
  label: { fontSize: 11, textTransform: "uppercase", letterSpacing: "0.06em", color: P.muted, fontWeight: 600, marginBottom: 5, display: "block" },
  btn: {
    padding: "13px 18px",
    borderRadius: 10,
    border: "none",
    background: P.accent,
    color: "#fff",
    fontWeight: 600,
    fontSize: 15,
    cursor: "pointer",
    minHeight: 46,
  },
  btnGhost: {
    padding: "11px 16px",
    borderRadius: 10,
    border: `1px solid ${P.line}`,
    background: "#fff",
    color: P.ink,
    fontWeight: 500,
    fontSize: 14,
    cursor: "pointer",
    minHeight: 44,
  },
  btnDanger: {
    padding: "8px 10px",
    borderRadius: 8,
    border: "none",
    background: "transparent",
    color: P.warn,
    cursor: "pointer",
    fontSize: 15,
    minWidth: 40,
    minHeight: 40,
  },
  chip: (on) => ({
    padding: "10px 16px",
    borderRadius: 999,
    border: `1.5px solid ${on ? P.accent : P.line}`,
    background: on ? P.accentSoft : "#fff",
    color: on ? P.accent : P.ink,
    fontWeight: on ? 700 : 500,
    cursor: "pointer",
    fontSize: 14,
    minHeight: 42,
  }),
  h2: { margin: "0 0 12px", fontSize: 15, fontWeight: 700 },
};

const uid = () => Math.random().toString(36).slice(2, 10) + Date.now().toString(36);

function bufToB64(buf) {
  const bytes = new Uint8Array(buf);
  let bin = "";
  const chunk = 0x8000;
  for (let i = 0; i < bytes.length; i += chunk) {
    bin += String.fromCharCode.apply(null, bytes.subarray(i, i + chunk));
  }
  return btoa(bin);
}
function b64ToBuf(b64) {
  const bin = atob(b64);
  const bytes = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
  return bytes.buffer;
}

function parseTimeToFraction(t) {
  if (!t) return null;
  const m = String(t).match(/^(\d{1,2}):(\d{2})$/);
  if (!m) return null;
  const mins = parseInt(m[1]) * 60 + parseInt(m[2]);
  if (mins < 0 || mins >= 1440) return null;
  return mins / 1440;
}

// České státní svátky pro daný rok (včetně pohyblivých velikonočních)
function czHolidays(year) {
  // Velikonoce — Meeusův algoritmus
  const a = year % 19, b = Math.floor(year / 100), c = year % 100;
  const d = Math.floor(b / 4), e = b % 4, f = Math.floor((b + 8) / 25);
  const g = Math.floor((b - f + 1) / 3), h = (19 * a + b - d - g + 15) % 30;
  const i = Math.floor(c / 4), k = c % 4, l = (32 + 2 * e + 2 * i - h - k) % 7;
  const m = Math.floor((a + 11 * h + 22 * l) / 451);
  const month = Math.floor((h + l - 7 * m + 114) / 31);
  const day = ((h + l - 7 * m + 114) % 31) + 1;
  const easter = new Date(Date.UTC(year, month - 1, day));
  const shift = (dt, days) => {
    const d2 = new Date(dt);
    d2.setUTCDate(d2.getUTCDate() + days);
    return d2.toISOString().slice(0, 10);
  };
  const fix = (mo, da) => `${year}-${String(mo).padStart(2, "0")}-${String(da).padStart(2, "0")}`;
  return new Set([
    fix(1, 1), fix(5, 1), fix(5, 8), fix(7, 5), fix(7, 6), fix(9, 28),
    fix(10, 28), fix(11, 17), fix(12, 24), fix(12, 25), fix(12, 26),
    shift(easter, -2), // Velký pátek
    shift(easter, 1),  // Velikonoční pondělí
  ]);
}

// Seznam ISO dat v rozsahu; volitelně jen pracovní dny (bez So/Ne a svátků)
function listDays(fromIso, toIso, workdaysOnly) {
  if (!toIso || toIso < fromIso) return [fromIso];
  const out = [];
  const holidays = workdaysOnly ? czHolidays(parseInt(fromIso.slice(0, 4))) : null;
  if (workdaysOnly && toIso.slice(0, 4) !== fromIso.slice(0, 4)) {
    czHolidays(parseInt(toIso.slice(0, 4))).forEach((h) => holidays.add(h));
  }
  const d = new Date(fromIso + "T00:00:00Z");
  const end = new Date(toIso + "T00:00:00Z");
  while (d <= end && out.length < 62) {
    const iso = d.toISOString().slice(0, 10);
    const dow = d.getUTCDay();
    if (!workdaysOnly || (dow !== 0 && dow !== 6 && !holidays.has(iso))) out.push(iso);
    d.setUTCDate(d.getUTCDate() + 1);
  }
  return out;
}

// ---------- CSV export ----------
function buildCsv(rows) {
  const esc = (v) => {
    const s = String(v ?? "");
    return /[;"\n]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s;
  };
  return "\uFEFF" + rows.map((r) => r.map(esc).join(";")).join("\r\n");
}

function downloadBlob(name, blob) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = name;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  setTimeout(() => URL.revokeObjectURL(url), 2000);
}
function downloadFile(name, content) {
  downloadBlob(name, new Blob([content], { type: "text/csv;charset=utf-8" }));
}

// ---------- Přihlašovací obrazovka ----------
function LoginScreen() {
  const [email, setEmail] = useState("");
  const [pass, setPass] = useState("");
  const [busy, setBusy] = useState(false);
  const [err, setErr] = useState(null);
  const [info, setInfo] = useState(null);

  const forgot = async () => {
    const e = email.trim();
    if (!e || !e.includes("@")) return setErr("Vyplňte nahoře svůj e-mail a klikněte znovu.");
    setBusy(true);
    setErr(null);
    setInfo(null);
    try {
      await resetPassword(e);
      setInfo(`Na ${e} jsme poslali odkaz pro nastavení nového hesla. Zkontrolujte i spam.`);
    } catch (ex) {
      setErr(ex.code === "auth/invalid-email" ? "Neplatný formát e-mailu." : "Odeslání se nepodařilo: " + ex.message);
    }
    setBusy(false);
  };

  const submit = async () => {
    if (!email.trim() || !pass) return setErr("Vyplňte e-mail a heslo.");
    setBusy(true);
    setErr(null);
    try {
      await login(email.trim(), pass);
    } catch (e) {
      const map = {
        "auth/invalid-credential": "Nesprávný e-mail nebo heslo.",
        "auth/invalid-email": "Neplatný formát e-mailu.",
        "auth/user-disabled": "Účet je zablokovaný.",
        "auth/too-many-requests": "Příliš mnoho pokusů — chvíli počkejte.",
        "auth/network-request-failed": "Chyba připojení k internetu.",
      };
      setErr(map[e.code] || "Přihlášení selhalo: " + e.message);
    }
    setBusy(false);
  };

  return (
    <div style={{ ...S.app, display: "flex", alignItems: "center", justifyContent: "center", padding: 16 }}>
      <div style={{ ...S.panel, width: "100%", maxWidth: 380, padding: 24 }}>
        <div style={{ fontSize: 20, fontWeight: 700, marginBottom: 4 }}>
          Výkaz práce <span style={{ color: P.accent, fontWeight: 400 }}>· tým</span>
          {IS_TEST && (
            <span style={{ marginLeft: 8, background: "#B3541E", color: "#fff", fontSize: 11, fontWeight: 700, padding: "3px 8px", borderRadius: 6, letterSpacing: "0.05em", verticalAlign: "middle" }}>
              TEST
            </span>
          )}
        </div>
        <p style={{ margin: "0 0 18px", color: P.muted, fontSize: 13 }}>Přihlaste se firemním účtem.</p>
        <div style={{ display: "grid", gap: 12 }}>
          <div>
            <label style={S.label}>E-mail</label>
            <input type="email" autoComplete="username" value={email} onChange={(e) => setEmail(e.target.value)} style={S.input} onKeyDown={(e) => e.key === "Enter" && submit()} />
          </div>
          <div>
            <label style={S.label}>Heslo</label>
            <input type="password" autoComplete="current-password" value={pass} onChange={(e) => setPass(e.target.value)} style={S.input} onKeyDown={(e) => e.key === "Enter" && submit()} />
          </div>
          {err && <div style={{ color: P.warn, fontSize: 13 }}>{err}</div>}
          {info && <div style={{ color: P.accent, fontSize: 13 }}>{info}</div>}
          <button onClick={submit} style={S.btn} disabled={busy}>
            {busy ? "Přihlašuji…" : "Přihlásit se"}
          </button>
          <button onClick={forgot} style={{ border: "none", background: "transparent", color: P.muted, fontSize: 13, cursor: "pointer", textDecoration: "underline", padding: 6 }} disabled={busy}>
            Zapomenuté heslo?
          </button>
        </div>
      </div>
    </div>
  );
}

// ---------- Vstupní bod: řeší přihlášení a profil ----------
export default function App() {
  const [authUser, setAuthUser] = useState(undefined);

  useEffect(() => subscribeAuth(setAuthUser), []);

  if (authUser === undefined) {
    return (
      <div style={{ ...S.app, display: "flex", alignItems: "center", justifyContent: "center" }}>
        <div style={{ color: P.muted }}>Načítám…</div>
      </div>
    );
  }
  if (authUser === null) return <LoginScreen />;
  return <AuthedGate authUser={authUser} />;
}

function AuthedGate({ authUser }) {
  const [users, setUsers] = useState(null);
  const [gateError, setGateError] = useState(false);
  const email = (authUser.email || "").toLowerCase();

  useEffect(() => {
    return subscribeCollection("users", setUsers, (err) => {
      console.error("users:", err);
      setGateError(true);
    });
  }, []);

  // První přihlášený účet se automaticky stane správcem
  useEffect(() => {
    if (users && users.length === 0 && email) {
      setUserDoc(email, { role: "admin", personIds: [] }).catch(console.error);
    }
  }, [users, email]);

  if (gateError) {
    return (
      <CenterMessage
        title="Nedaří se načíst uživatelské účty"
        text="Zkontrolujte připojení a Firestore pravidla (musí povolovat přístup přihlášeným)."
        onLogout={logout}
      />
    );
  }
  if (users === null) {
    return (
      <div style={{ ...S.app, display: "flex", alignItems: "center", justifyContent: "center" }}>
        <div style={{ color: P.muted }}>Načítám účet…</div>
      </div>
    );
  }

  const profile = users.find((u) => u.id === email);
  if (!profile) {
    if (users.length === 0) {
      return (
        <div style={{ ...S.app, display: "flex", alignItems: "center", justifyContent: "center" }}>
          <div style={{ color: P.muted }}>Zakládám účet správce…</div>
        </div>
      );
    }
    return (
      <CenterMessage
        title="Účet není přiřazen"
        text={`Účet ${email} zatím nemá v aplikaci roli. Požádejte správce, ať vás přidá v sekci Správa.`}
        onLogout={logout}
      />
    );
  }

  return <Workspace authUser={authUser} profile={profile} users={users} />;
}

function CenterMessage({ title, text, onLogout }) {
  return (
    <div style={{ ...S.app, display: "flex", alignItems: "center", justifyContent: "center", padding: 16 }}>
      <div style={{ ...S.panel, maxWidth: 400, textAlign: "center" }}>
        <div style={{ fontWeight: 700, marginBottom: 8 }}>{title}</div>
        <p style={{ color: P.muted, fontSize: 14, margin: "0 0 14px" }}>{text}</p>
        <button onClick={onLogout} style={S.btnGhost}>Odhlásit se</button>
      </div>
    </div>
  );
}

// ---------- Pracovní plocha (po přihlášení) ----------
function Workspace({ authUser, profile, users }) {
  const [tab, setTab] = useState("vykazy");
  const [pwOpen, setPwOpen] = useState(false);
  const [people, setPeople] = useState(null);
  const [projects, setProjects] = useState(null);
  const [entries, setEntries] = useState(null);
  const [trips, setTrips] = useState(null);
  const [timesheet, setTimesheet] = useState(null);
  const [template, setTemplate] = useState(undefined);
  const [zahrTemplate, setZahrTemplate] = useState(undefined); // undefined = načítám, null = žádná
  const [toast, setToast] = useState(null);
  const [connError, setConnError] = useState(false);
  const isMobile = useIsMobile();

  useEffect(() => {
    const onErr = (err) => {
      console.error("Firestore:", err);
      setConnError(true);
    };
    const u1 = subscribeCollection("people", (d) => setPeople(d.sort((a, b) => a.name.localeCompare(b.name, "cs"))), onErr);
    const u2 = subscribeCollection("projects", (d) => setProjects(d.sort((a, b) => (a.code + a.name).localeCompare(b.code + b.name, "cs"))), onErr);
    const u3 = subscribeCollection("entries", setEntries, onErr);
    const u5 = subscribeCollection("trips", setTrips, onErr);
    const u6 = subscribeCollection("timesheet", setTimesheet, onErr);
    const u4 = subscribeTemplate(setTemplate, onErr);
    const u7 = subscribeConfigDoc("zahrTemplate", setZahrTemplate, onErr);
    return () => { u1(); u2(); u3(); u4(); u5(); u6(); u7(); };
  }, []);

  const showToast = (msg) => {
    setToast(msg);
    setTimeout(() => setToast(null), 3500);
  };

  const guard = async (fn, okMsg) => {
    try {
      const res = await fn();
      if (okMsg) showToast(okMsg);
      return res === undefined ? true : res;
    } catch (err) {
      console.error(err);
      showToast("Uložení selhalo: " + (err.code === "permission-denied" ? "chybí oprávnění (zkontrolujte Firestore pravidla)." : err.message));
      return false;
    }
  };

  const actions = {
    addEntries: (items, msg) => guard(() => addDocs("entries", items), msg),
    removeEntry: (id) => guard(() => removeDoc("entries", id)),
    addPeople: (items, msg) => guard(() => addDocs("people", items), msg),
    removePerson: (id) => guard(() => removeDoc("people", id)),
    addProjects: (items, msg) => guard(() => addDocs("projects", items), msg),
    removeProject: (id) => guard(() => removeDoc("projects", id)),
    removePeopleBulk: (ids, msg) => guard(() => removeDocs("people", ids), msg),
    removeProjectsBulk: (ids, msg) => guard(() => removeDocs("projects", ids), msg),
    removeEntriesBulk: (ids, msg) => guard(() => removeDocs("entries", ids), msg),
    saveTemplate: async (meta) => guard(() => saveTemplateDoc(meta), `Šablona uložena (rok ${meta.year}). Sdílí ji celý tým.`),
    saveZahrTemplate: async (meta) => guard(() => saveConfigDoc("zahrTemplate", meta), "Šablona zahraniční cesty uložena."),
    setUser: (email, data, msg) => guard(() => setUserDoc(email, data), msg),
    saveTrip: (id, data, msg) => guard(() => saveDocIn("trips", id, data), msg),
    addTimesheet: (items, msg) => guard(() => addDocs("timesheet", items), msg),
    removeTimesheetBulk: (ids, msg) => guard(() => removeDocs("timesheet", ids), msg),
    saveProject: (id, data, msg) => guard(() => saveDocIn("projects", id, data), msg),
    savePerson: (id, data, msg) => guard(() => saveDocIn("people", id, data), msg),
    removeTrip: (id, msg) => guard(() => removeDoc("trips", id), msg),
    removeUser: (email, msg) => guard(() => removeDoc("users", email), msg),
  };

  const loading = people === null || projects === null || entries === null || trips === null || timesheet === null || template === undefined || zahrTemplate === undefined;

  if (loading) {
    return (
      <div style={{ ...S.app, display: "flex", alignItems: "center", justifyContent: "center", flexDirection: "column", gap: 10 }}>
        <div style={{ color: P.muted }}>Načítám data…</div>
        {connError && (
          <div style={{ color: P.warn, fontSize: 13, maxWidth: 320, textAlign: "center" }}>
            Nedaří se připojit k databázi. Zkontrolujte připojení k internetu a Firestore pravidla (viz README).
          </div>
        )}
      </div>
    );
  }

  const isAdmin = profile.role === "admin";
  const myIds = profile.personIds || [];
  // Vedoucí vidí jen svůj tým a jeho záznamy; správce vše
  const visPeople = isAdmin ? people : people.filter((p) => myIds.includes(p.id));
  const visEntries = isAdmin ? entries : entries.filter((e) => myIds.includes(e.personId));
  const visTrips = isAdmin ? trips : trips.filter((t) => myIds.includes(t.personId));
  const visTime = isAdmin ? timesheet : timesheet.filter((t) => myIds.includes(t.personId));
  const data = { people: visPeople, projects, entries: visEntries, trips: visTrips, timesheet: visTime };
  // Šablona z Firestore má přednost, jinak zabudovaná
  const effectiveTemplate = template;
  const perms = permsUctu(profile);
  const zakladni = [
    { id: "vykazy", label: "Výkazy", icon: "🕐" },
    { id: "cas", label: "Čas", icon: "⏱️" },
    { id: "export", label: "Export", icon: "📤" },
    { id: "zahr", label: "Cesty", icon: "✈️" },
  ].filter((t) => perms[t.id]);
  const tabs = isAdmin
    ? [
        ...zakladni,
        { id: "projekty", label: "Projekty", icon: "📁" },
        { id: "lide", label: "Lidé", icon: "👥" },
        { id: "sprava", label: "Správa", icon: "⚙️" },
      ]
    : zakladni;

  // Pokud přihlášený nemá povolený žádný modul
  if (tabs.length === 0) {
    return (
      <CenterMessage
        title="Účet nemá povolený žádný modul"
        text="Požádejte správce, ať vám v sekci Správa nastaví přístup."
        onLogout={logout}
      />
    );
  }
  // Kdyby aktivní záložka nebyla povolená (např. po změně práv), přepneme na první
  const aktivni = tabs.some((t) => t.id === tab) ? tab : tabs[0].id;

  return (
    <div style={S.app}>
      {/* Horní lišta */}
      <header
        style={{
          background: P.ink,
          color: "#fff",
          padding: isMobile ? "12px 16px" : "14px 20px",
          paddingTop: isMobile ? "max(12px, env(safe-area-inset-top))" : 14,
          display: "flex",
          alignItems: "center",
          gap: 16,
          position: "sticky",
          top: 0,
          zIndex: 40,
        }}
      >
        <div style={{ fontSize: isMobile ? 16 : 17, fontWeight: 700, letterSpacing: "0.02em", flex: isMobile ? 1 : "0 0 auto" }}>
          Výkaz práce <span style={{ color: "#7FC7C4", fontWeight: 400 }}>· tým</span>
          {IS_TEST && (
            <span style={{ marginLeft: 8, background: "#B3541E", color: "#fff", fontSize: 11, fontWeight: 700, padding: "3px 8px", borderRadius: 6, letterSpacing: "0.05em", verticalAlign: "middle" }}>
              TEST
            </span>
          )}
        </div>
        {!isMobile && (
          <nav style={{ display: "flex", gap: 4, flexWrap: "wrap" }}>
            {tabs.map((t) => (
              <button
                key={t.id}
                onClick={() => setTab(t.id)}
                style={{
                  padding: "8px 16px",
                  borderRadius: 8,
                  border: "none",
                  cursor: "pointer",
                  fontSize: 14,
                  fontWeight: 600,
                  background: aktivni === t.id ? P.accent : "transparent",
                  color: aktivni === t.id ? "#fff" : "#B9C4CE",
                }}
              >
                {t.label}
              </button>
            ))}
          </nav>
        )}
        <div style={{ marginLeft: "auto", display: "flex", alignItems: "center", gap: 10, whiteSpace: "nowrap" }}>
          <span style={{ fontSize: 12, color: connError ? "#E8A87C" : "#8A97A3" }}>
            {connError ? "offline" : isMobile ? "" : (authUser.email || "")}
          </span>
          <button
            onClick={() => setPwOpen(true)}
            title="Změnit heslo"
            style={{ padding: "6px 10px", borderRadius: 8, border: "1px solid #3A4655", background: "transparent", color: "#B9C4CE", fontSize: 12, cursor: "pointer" }}
          >
            Heslo
          </button>
          <button
            onClick={() => logout()}
            title="Odhlásit se"
            style={{ padding: "6px 10px", borderRadius: 8, border: "1px solid #3A4655", background: "transparent", color: "#B9C4CE", fontSize: 12, cursor: "pointer" }}
          >
            Odhlásit
          </button>
        </div>
      </header>

      {/* Obsah */}
      <main
        style={{
          maxWidth: 1080,
          margin: "0 auto",
          padding: isMobile ? "14px 12px" : "20px 16px",
          paddingBottom: isMobile ? "calc(84px + env(safe-area-inset-bottom))" : 60,
        }}
      >
        {aktivni === "vykazy" && <TabVykazy data={data} actions={actions} showToast={showToast} isMobile={isMobile} />}
        {aktivni === "export" && <TabExport data={data} actions={actions} template={effectiveTemplate} showToast={showToast} isMobile={isMobile} isAdmin={isAdmin} />}
        {aktivni === "cas" && <TabCas data={data} actions={actions} showToast={showToast} isMobile={isMobile} isAdmin={isAdmin} />}
        {aktivni === "zahr" && <TabZahranicni data={data} actions={actions} showToast={showToast} isMobile={isMobile} isAdmin={isAdmin} zahrTemplate={zahrTemplate} />}
        {aktivni === "projekty" && isAdmin && <TabProjekty data={data} actions={actions} showToast={showToast} />}
        {aktivni === "lide" && isAdmin && <TabLide data={data} actions={actions} showToast={showToast} />}
        {aktivni === "sprava" && isAdmin && <TabSprava users={users} people={people} actions={actions} showToast={showToast} myEmail={(authUser.email || "").toLowerCase()} />}
      </main>

      {/* Spodní navigace (mobil) */}
      {isMobile && (
        <nav
          style={{
            position: "fixed",
            bottom: 0,
            left: 0,
            right: 0,
            background: "#fff",
            borderTop: `1px solid ${P.line}`,
            display: "flex",
            zIndex: 40,
            paddingBottom: "env(safe-area-inset-bottom)",
            boxShadow: "0 -2px 12px rgba(0,0,0,0.06)",
          }}
        >
          {tabs.map((t) => (
            <button
              key={t.id}
              onClick={() => setTab(t.id)}
              style={{
                flex: 1,
                padding: "8px 4px 10px",
                border: "none",
                background: "transparent",
                cursor: "pointer",
                display: "flex",
                flexDirection: "column",
                alignItems: "center",
                gap: 2,
                color: aktivni === t.id ? P.accent : P.muted,
              }}
            >
              <span style={{ fontSize: 20, filter: aktivni === t.id ? "none" : "grayscale(1) opacity(0.7)" }}>{t.icon}</span>
              <span style={{ fontSize: 11, fontWeight: aktivni === t.id ? 700 : 500 }}>{t.label}</span>
            </button>
          ))}
        </nav>
      )}

      {pwOpen && <PasswordDialog onClose={() => setPwOpen(false)} showToast={showToast} isMobile={isMobile} />}

      {toast && (
        <div
          style={{
            position: "fixed",
            bottom: isMobile ? "calc(76px + env(safe-area-inset-bottom))" : 20,
            left: "50%",
            transform: "translateX(-50%)",
            background: P.ink,
            color: "#fff",
            padding: "12px 18px",
            borderRadius: 10,
            fontSize: 13,
            boxShadow: "0 4px 16px rgba(0,0,0,0.25)",
            zIndex: 60,
            maxWidth: "92vw",
            textAlign: "center",
          }}
        >
          {toast}
        </div>
      )}
    </div>
  );
}

// ---------- Záložka: Výkazy ----------
function TabVykazy({ data, actions, showToast, isMobile }) {
  const { people, projects, entries } = data;
  const [date, setDate] = useState(today());
  const [dateTo, setDateTo] = useState("");
  const [workdaysOnly, setWorkdaysOnly] = useState(true);
  const [cinnost, setCinnost] = useState("Práce");
  const [note, setNote] = useState("");
  const [odkudKam, setOdkudKam] = useState("");
  const [ucel, setUcel] = useState("");
  const [casOd, setCasOd] = useState("");
  const [casDo, setCasDo] = useState("");
  const [jidla, setJidla] = useState("");
  const [selected, setSelected] = useState([]);
  const [saving, setSaving] = useState(false);
  const [filterPerson, setFilterPerson] = useState("");
  const [filterMonth, setFilterMonth] = useState(monthKey(today()));
  const [uprava, setUprava] = useState(false);
  const [vybrane, setVybrane] = useState([]);

  const toggleVyber = (id) => setVybrane((s) => (s.includes(id) ? s.filter((x) => x !== id) : [...s, id]));
  const smazatVybrane = async () => {
    if (vybrane.length === 0) return showToast("Nic není vybráno.");
    if (!window.confirm(`Smazat ${vybrane.length} ${vybrane.length === 1 ? "záznam" : vybrane.length < 5 ? "záznamy" : "záznamů"} z výkazu?`)) return;
    const ok = await actions.removeEntriesBulk(vybrane, `Smazáno ${vybrane.length} záznamů.`);
    if (ok) { setVybrane([]); setUprava(false); }
  };

  const absence = jeAbsence(cinnost);
  const cesta = jeCesta(cinnost);

  const togglePerson = (id) =>
    setSelected((s) => (s.includes(id) ? s.filter((x) => x !== id) : [...s, id]));

  const addEntries = async () => {
    if (!date) return showToast("Vyberte datum.");
    if (dateTo && dateTo < date) return showToast("Datum 'do' nesmí být před datem 'od'.");
    if (selected.length === 0) return showToast("Vyberte alespoň jednu osobu.");
    if (cesta) {
      if (casOd && !parseTimeToFraction(casOd)) return showToast("Čas 'od' zadejte ve formátu HH:MM.");
      if (casDo && !parseTimeToFraction(casDo)) return showToast("Čas 'do' zadejte ve formátu HH:MM.");
    }

    if (saving) return; // ochrana proti dvojkliku
    const days = listDays(date, dateTo, workdaysOnly && !!dateTo);
    if (days.length === 0) return showToast("Ve zvoleném rozsahu nejsou žádné pracovní dny.");
    if (days.length * selected.length > 400) return showToast("Příliš velký rozsah — zmenšete počet dní nebo osob.");

    const newOnes = days.flatMap((day) => selected.map((pid) => ({
      date: day,
      personId: pid,
      cinnost,
      note: note.trim(),
      odkudKam: cesta ? odkudKam.trim() : "",
      ucel: cesta ? ucel.trim() : "",
      casOd: cesta ? casOd.trim() : "",
      casDo: cesta ? casDo.trim() : "",
      jidla: cesta && jidla !== "" ? parseInt(jidla) || 0 : null,
      createdAt: new Date().toISOString(),
    })));

    // Kontrola duplicit: stejná osoba + den + činnost už ve výkazu je
    const isDupe = (ne) =>
      entries.some(
        (e) =>
          e.personId === ne.personId &&
          e.date === ne.date &&
          (e.cinnost || "Práce") === (ne.cinnost || "Práce")
      );
    const fresh = newOnes.filter((ne) => !isDupe(ne));
    const dupCount = newOnes.length - fresh.length;
    let toSave = newOnes;
    if (dupCount > 0) {
      if (fresh.length === 0) {
        if (!window.confirm(dupCount === 1
          ? "Tento den už má u této osoby stejnou činnost zapsanou. Uložit přesto ještě jednou?"
          : `Všech ${dupCount} záznamů už ve výkazu je (stejná osoba, den, činnost). Uložit je přesto ještě jednou?`)) return;
      } else {
        if (!window.confirm(`${dupCount} z ${newOnes.length} záznamů už existuje — uloží se jen ${fresh.length} nových. Pokračovat?`)) return;
        toSave = fresh;
      }
    }

    const nDays = new Set(toSave.map((e) => e.date)).size;
    const nPeople = new Set(toSave.map((e) => e.personId)).size;
    const msg = toSave.length === 1
      ? "Výkaz uložen."
      : nDays === 1
      ? `Uloženo pro ${nPeople} osob.`
      : `Uloženo ${toSave.length} záznamů (${nDays} dní × ${nPeople} os.).`;
    setSaving(true);
    const ok = await actions.addEntries(toSave, msg);
    setSaving(false);
    if (ok) {
      setNote("");
      setOdkudKam("");
      setUcel("");
      setCasOd("");
      setCasDo("");
      setJidla("");
    }
  };

  const personName = (id) => people.find((p) => p.id === id)?.name || "(smazaná osoba)";
  const projectName = (id) => {
    if (!id) return "—";
    const pr = projects.find((p) => p.id === id);
    return pr ? (pr.code ? `${pr.code} – ${pr.name}` : pr.name) : "(smazaný projekt)";
  };

  const filtered = useMemo(() => {
    return entries
      .filter((e) => (!filterPerson || e.personId === filterPerson) && (!filterMonth || monthKey(e.date) === filterMonth))
      .sort((a, b) => (a.date < b.date ? 1 : a.date > b.date ? -1 : 0));
  }, [entries, filterPerson, filterMonth]);

  const noSetup = people.length === 0;

  return (
    <div style={{ display: "grid", gap: 14 }}>
      {noSetup && (
        <div style={{ ...S.panel, borderLeft: `4px solid ${P.warn}`, fontSize: 14 }}>
          Nejdřív přidejte lidi (záložka Lidé), pak můžete vykazovat.
        </div>
      )}

      {/* Formulář zadání */}
      <section style={S.panel}>
        <h2 style={S.h2}>Nový záznam</h2>
        <div style={{ display: "grid", gridTemplateColumns: isMobile ? "1fr 1fr" : "repeat(auto-fit, minmax(160px, 1fr))", gap: 12, marginBottom: 12, alignItems: "end" }}>
          <div>
            <label style={{ ...S.label, whiteSpace: "nowrap" }}>Datum od</label>
            <input type="date" value={date} onChange={(e) => setDate(e.target.value)} style={S.input} />
          </div>
          <div>
            <label style={{ ...S.label, whiteSpace: "nowrap" }} title="Vyplňte pro zadání více dní najednou">Do — nepovinné</label>
            <input type="date" value={dateTo} min={date} onChange={(e) => setDateTo(e.target.value)} style={S.input} />
          </div>
          <div style={{ gridColumn: isMobile ? "1 / -1" : "auto" }}>
            <label style={S.label}>Činnost</label>
            <select value={cinnost} onChange={(e) => setCinnost(e.target.value)} style={S.input}>
              {CINNOSTI.map((c) => (
                <option key={c} value={c}>{c}</option>
              ))}
            </select>
          </div>
        </div>

        {dateTo && (
          <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 12, flexWrap: "wrap" }}>
            <label style={{ display: "flex", alignItems: "center", gap: 8, fontSize: 14, cursor: "pointer" }}>
              <input type="checkbox" checked={workdaysOnly} onChange={(e) => setWorkdaysOnly(e.target.checked)} style={{ accentColor: P.accent, width: 18, height: 18 }} />
              Jen pracovní dny (bez víkendů a svátků)
            </label>
            <span style={{ fontSize: 13, color: P.accent, fontWeight: 600 }}>
              → {listDays(date, dateTo, workdaysOnly).length} dní
            </span>
            <button onClick={() => setDateTo("")} style={{ ...S.btnDanger, color: P.muted, fontSize: 13, padding: "4px 8px", minHeight: 32 }}>
              zrušit rozsah
            </button>
          </div>
        )}

        {cesta && (
          <div style={{ display: "grid", gridTemplateColumns: isMobile ? "1fr 1fr" : "repeat(auto-fit, minmax(140px, 1fr))", gap: 12, marginBottom: 12, padding: 12, background: P.accentSoft, borderRadius: 10 }}>
            <div style={{ gridColumn: "1 / -1" }}>
              <label style={S.label}>Pracovní cesta odkud – kam</label>
              <input value={odkudKam} onChange={(e) => setOdkudKam(e.target.value)} placeholder="Praha – Opava" style={S.input} />
            </div>
            <div style={{ gridColumn: "1 / -1" }}>
              <label style={S.label}>Účel cesty</label>
              <input value={ucel} onChange={(e) => setUcel(e.target.value)} placeholder="Servis u klienta" style={S.input} />
            </div>
            <div>
              <label style={S.label}>Od (HH:MM)</label>
              <input value={casOd} onChange={(e) => setCasOd(e.target.value)} placeholder="7:00" inputMode="numeric" style={{ ...S.input, ...S.mono }} />
            </div>
            <div>
              <label style={S.label}>Do (HH:MM)</label>
              <input value={casDo} onChange={(e) => setCasDo(e.target.value)} placeholder="16:30" inputMode="numeric" style={{ ...S.input, ...S.mono }} />
            </div>
            <div>
              <label style={S.label}>Bezplatná jídla</label>
              <input type="number" min="0" max="3" value={jidla} onChange={(e) => setJidla(e.target.value)} placeholder="0" style={{ ...S.input, ...S.mono }} />
            </div>
          </div>
        )}

        <div style={{ marginBottom: 12 }}>
          <label style={S.label}>Poznámka (nepovinné)</label>
          <input type="text" value={note} onChange={(e) => setNote(e.target.value)} placeholder="Co se dělalo…" style={S.input} />
        </div>

        <label style={S.label}>Vykázat pro — jednu nebo více osob</label>
        <div style={{ display: "flex", flexWrap: "wrap", gap: 8, marginBottom: 14 }}>
          {people.map((p) => {
            const on = selected.includes(p.id);
            return (
              <button key={p.id} onClick={() => togglePerson(p.id)} style={S.chip(on)}>
                {on ? "✓ " : ""}{p.name}
              </button>
            );
          })}
          {people.length > 1 && (
            <button
              onClick={() => setSelected(selected.length === people.length ? [] : people.map((p) => p.id))}
              style={{ ...S.chip(false), borderStyle: "dashed" }}
            >
              {selected.length === people.length ? "Zrušit výběr" : "Celá směna"}
            </button>
          )}
        </div>

        <button onClick={addEntries} style={{ ...S.btn, width: isMobile ? "100%" : "auto", opacity: saving ? 0.6 : 1 }} disabled={saving || noSetup}>
          {(() => {
            if (saving) return "Ukládám…";
            const nd = dateTo ? listDays(date, dateTo, workdaysOnly).length : 1;
            if (nd > 1 && selected.length > 1) return `Uložit (${selected.length} osob × ${nd} dní)`;
            if (nd > 1) return `Uložit záznam (${nd} dní)`;
            if (selected.length > 1) return `Uložit záznam (${selected.length} osob)`;
            return "Uložit záznam";
          })()}
        </button>
      </section>

      {/* Přehled */}
      <section style={S.panel}>
        <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 12 }}>
          <h2 style={{ ...S.h2, marginBottom: 0 }}>Přehled záznamů</h2>
          {filtered.length > 0 && (
            <button
              onClick={() => { setUprava(!uprava); setVybrane([]); }}
              style={{ ...S.btnDanger, color: P.accent, fontWeight: 600, fontSize: 13 }}
            >
              {uprava ? "Hotovo" : "Upravit"}
            </button>
          )}
        </div>
        {uprava && (
          <div style={{ display: "flex", gap: 8, alignItems: "center", marginBottom: 10, flexWrap: "wrap" }}>
            <button onClick={() => setVybrane(vybrane.length === filtered.length ? [] : filtered.map((e) => e.id))} style={S.btnGhost}>
              {vybrane.length === filtered.length ? "Zrušit výběr" : "Vybrat vše"}
            </button>
            <button onClick={smazatVybrane} style={{ ...S.btnGhost, color: P.warn, borderColor: P.warn, opacity: vybrane.length ? 1 : 0.5 }} disabled={vybrane.length === 0}>
              Smazat vybrané ({vybrane.length})
            </button>
          </div>
        )}
        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10, marginBottom: 12 }}>
          <div>
            <label style={S.label}>Osoba</label>
            <select value={filterPerson} onChange={(e) => setFilterPerson(e.target.value)} style={S.input}>
              <option value="">Všichni</option>
              {people.map((p) => (
                <option key={p.id} value={p.id}>{p.name}</option>
              ))}
            </select>
          </div>
          <div>
            <label style={S.label}>Měsíc</label>
            <input type="month" value={filterMonth} onChange={(e) => setFilterMonth(e.target.value)} style={S.input} />
          </div>
        </div>

        {filtered.length === 0 ? (
          <div style={{ color: P.muted, padding: "12px 0" }}>Pro zvolený filtr nejsou žádné záznamy.</div>
        ) : isMobile ? (
          /* Mobil: karty */
          <div style={{ display: "grid", gap: 8 }}>
            {filtered.map((e) => (
              <div
                key={e.id}
                onClick={uprava ? () => toggleVyber(e.id) : undefined}
                style={{
                  border: `1px solid ${uprava && vybrane.includes(e.id) ? P.warn : P.line}`,
                  background: uprava && vybrane.includes(e.id) ? "#FBEFE7" : "#fff",
                  borderRadius: 10, padding: "10px 12px", cursor: uprava ? "pointer" : "default",
                }}
              >
                <div style={{ display: "flex", alignItems: "baseline", gap: 8 }}>
                  {uprava && (
                    <input type="checkbox" readOnly checked={vybrane.includes(e.id)} style={{ accentColor: P.warn, width: 16, height: 16, flexShrink: 0, pointerEvents: "none" }} />
                  )}
                  <span style={{ ...S.mono, fontWeight: 700, fontSize: 14 }}>{fmtDateShort(e.date)}</span>
                  <span style={{ fontWeight: 600, fontSize: 14, flex: 1, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{personName(e.personId)}</span>
                  {uprava && (
                    <button style={{ ...S.btnDanger, padding: "2px 6px", minHeight: 32, minWidth: 32 }} onClick={(ev) => { ev.stopPropagation(); actions.removeEntry(e.id); }} title="Smazat">✕</button>
                  )}
                </div>
                <div style={{ fontSize: 13, color: P.muted, marginTop: 2 }}>
                  <span style={{ color: (e.cinnost || "Práce") === "Práce" ? P.muted : P.warn, fontWeight: 600 }}>{e.cinnost || "Práce"}</span>
                  {[e.odkudKam, e.note].filter(Boolean).length > 0 && (
                    <div style={{ marginTop: 2 }}>{[e.odkudKam, e.note].filter(Boolean).join(" · ")}</div>
                  )}
                </div>
              </div>
            ))}
          </div>
        ) : (
          /* Desktop: tabulka */
          <div style={{ overflowX: "auto" }}>
            <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
              <thead>
                <tr>
                  {["Datum", "Osoba", "Činnost", "Poznámka"].concat(uprava ? [""] : []).map((h, i) => (
                    <th key={i} style={{ textAlign: "left", padding: "8px 10px", borderBottom: `2px solid ${P.line}`, fontSize: 11, textTransform: "uppercase", letterSpacing: "0.05em", color: P.muted }}>
                      {h}
                    </th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {filtered.map((e) => (
                  <tr
                    key={e.id}
                    onClick={uprava ? () => toggleVyber(e.id) : undefined}
                    style={{ cursor: uprava ? "pointer" : "default", background: uprava && vybrane.includes(e.id) ? "#FBEFE7" : "transparent" }}
                  >
                    <td style={{ padding: "8px 10px", borderBottom: `1px solid ${P.line}`, whiteSpace: "nowrap", ...S.mono }}>
                      {uprava && (
                        <input type="checkbox" readOnly checked={vybrane.includes(e.id)} style={{ accentColor: P.warn, width: 15, height: 15, marginRight: 8, verticalAlign: "middle", pointerEvents: "none" }} />
                      )}
                      {fmtDate(e.date)}
                    </td>
                    <td style={{ padding: "8px 10px", borderBottom: `1px solid ${P.line}` }}>{personName(e.personId)}</td>
                    <td style={{ padding: "8px 10px", borderBottom: `1px solid ${P.line}`, color: (e.cinnost || "Práce") === "Práce" ? P.muted : P.warn }}>{e.cinnost || "Práce"}</td>

                    <td style={{ padding: "8px 10px", borderBottom: `1px solid ${P.line}`, color: P.muted }}>{[e.odkudKam, e.note].filter(Boolean).join(" · ")}</td>
                    {uprava && (
                      <td style={{ padding: "4px 6px", borderBottom: `1px solid ${P.line}`, textAlign: "right" }}>
                        <button style={S.btnDanger} onClick={(ev) => { ev.stopPropagation(); actions.removeEntry(e.id); }} title="Smazat záznam">✕</button>
                      </td>
                    )}
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )}

        {filtered.length > 0 && (
          <div style={{ marginTop: 14, paddingTop: 12, borderTop: `2px solid ${P.line}`, fontSize: 13, color: P.muted }}>
            Zapsaných dní: <b style={{ ...S.mono, color: P.ink }}>{new Set(filtered.map((e) => e.date + e.personId)).size}</b>
          </div>
        )}
      </section>
    </div>
  );
}

// ---------- Záložka: Export ----------
function TabExport({ data, actions, template, showToast, isMobile, isAdmin }) {
  const { people, projects, entries } = data;
  const [exportMonth, setExportMonth] = useState(new Date().getMonth() + 1);
  const [selected, setSelected] = useState(people.map((p) => p.id));
  const [busy, setBusy] = useState(false);
  const [dialogOpen, setDialogOpen] = useState(false);
  const [dialogSel, setDialogSel] = useState([]);
  // CSV část
  const [from, setFrom] = useState(monthKey(today()) + "-01");
  const [to, setTo] = useState(today());
  const [csvMode, setCsvMode] = useState("combined");
  const fileRef = useRef(null);

  const togglePerson = (id) =>
    setSelected((s) => (s.includes(id) ? s.filter((x) => x !== id) : [...s, id]));

  const personName = (id) => people.find((p) => p.id === id)?.name || "neznamy";
  const projectLabel = (id) => {
    if (!id) return "";
    const pr = projects.find((p) => p.id === id);
    return pr ? (pr.code ? `${pr.code} – ${pr.name}` : pr.name) : "(smazaný projekt)";
  };
  const safeName = (s) => s.normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/[^a-zA-Z0-9]+/g, "-").replace(/^-|-$/g, "");

  // ---- Nahrání šablony ----
  const onTemplateFile = async (ev) => {
    const f = ev.target.files?.[0];
    ev.target.value = "";
    if (!f) return;
    setBusy(true);
    try {
      const buf = await f.arrayBuffer();
      const ExcelJS = await getExcelJS();
      const wb = new ExcelJS.Workbook();
      await wb.xlsx.load(buf);
      const leden = wb.getWorksheet("Leden");
      if (!leden) {
        showToast("Soubor nevypadá jako výkazová šablona (chybí list Leden).");
        setBusy(false);
        return;
      }
      const yearVal = leden.getCell("S2").value;
      const year = typeof yearVal === "number" ? yearVal : new Date().getFullYear();
      const b64 = bufToB64(buf);
      if (b64.length > 900000) {
        showToast("Šablona je příliš velká pro uložení do databáze (limit ~0,9 MB). Zmenšete ji, nebo se ozvěte.");
        setBusy(false);
        return;
      }
      await actions.saveTemplate({ name: f.name, year, uploadedAt: new Date().toISOString(), b64 });
    } catch (err) {
      console.error(err);
      showToast("Chyba při čtení šablony: " + err.message);
    }
    setBusy(false);
  };

  // ---- Načtení lidí ze šablony (Tabulky!A8 dolů) ----
  const importPeopleFromTemplate = async () => {
    if (!template) return;
    setBusy(true);
    try {
      const ExcelJS = await getExcelJS();
      const templateBuf = await getTemplateBuffer(template);
      const wb = new ExcelJS.Workbook();
      await wb.xlsx.load(templateBuf);
      const tab = wb.getWorksheet("Tabulky");
      if (!tab) throw new Error("List Tabulky nenalezen.");
      const names = [];
      for (let r = 8; r <= 200; r++) {
        const v = tab.getCell(`A${r}`).value;
        if (v == null || String(v).trim() === "") break;
        names.push(String(v).trim());
      }
      const existing = new Set(people.map((p) => p.name.toLowerCase()));
      const newOnes = names.filter((n) => !existing.has(n.toLowerCase())).map((n) => ({ name: n }));
      if (newOnes.length === 0) {
        showToast("Všichni lidé ze šablony už v aplikaci jsou.");
      } else {
        await actions.addPeople(newOnes, `Načteno ${newOnes.length} lidí ze šablony.`);
      }
    } catch (err) {
      console.error(err);
      showToast("Chyba: " + err.message);
    }
    setBusy(false);
  };

  // ---- Sestavení denních buněk pro oficiální výkaz ----
  // Struktura listu: E=Popis (činnost, rozbalovací seznam), F=Odkud–Kam / popis,
  // G=Účel cesty, H=čas od, I=čas do, J=bezplatná jídla. Hodiny výkaz needviduje.
  const buildDayCells = (dayEntries) => {
    // Rozhodující je činnost dne — vezmeme "nejsilnější" (cesta > absence > práce)
    const priority = (c) =>
      jeCesta(c) ? 3 : jeAbsence(c) ? 2 : c === "Home Office" || c === "Home Office 1/2" ? 2 : 1;
    let main = dayEntries[0];
    dayEntries.forEach((e) => {
      if (priority(e.cinnost || "Práce") > priority(main.cinnost || "Práce")) main = e;
    });
    const cinnost = main.cinnost || "Práce";
    const cesta = jeCesta(cinnost);

    // F = popis. U cesty: odkud–kam. Jinak: projekty + poznámky ze všech záznamů dne.
    const fParts = [];
    if (cesta && main.odkudKam) fParts.push(main.odkudKam);
    dayEntries.forEach((e) => {
      const bits = [];
      const pl = projectLabel(e.projectId);
      if (pl) bits.push(pl);
      if (e.note) bits.push(e.note);
      if (bits.length) fParts.push(bits.join(" – "));
    });
    // deduplikace
    const fText = [...new Set(fParts)].join("; ");

    // Časy a jídla jen u cest
    let odMin = null, doMax = null, jidlaSum = null;
    if (cesta) {
      dayEntries.forEach((e) => {
        const od = parseTimeToFraction(e.casOd);
        const doV = parseTimeToFraction(e.casDo);
        if (od != null) odMin = odMin == null ? od : Math.min(odMin, od);
        if (doV != null) doMax = doMax == null ? doV : Math.max(doMax, doV);
        if (e.jidla != null) jidlaSum = (jidlaSum || 0) + e.jidla;
      });
    }

    return {
      E: cinnost,
      F: fText || null,
      G: cesta && main.ucel ? main.ucel : null,
      H: odMin,
      I: doMax,
      J: jidlaSum,
    };
  };

  // ---- Dialog exportu ----
  const entriesCountFor = (pid, month) => {
    if (!template) return 0;
    const mm = String(month).padStart(2, "0");
    return entries.filter((e) => e.personId === pid && monthKey(e.date) === `${template.year}-${mm}`).length;
  };
  const peopleWithEntries = (month) => people.filter((p) => entriesCountFor(p.id, month) > 0).map((p) => p.id);

  const openExportDialog = () => {
    setDialogSel(peopleWithEntries(exportMonth));
    setDialogOpen(true);
  };
  const onDialogMonthChange = (m) => {
    setExportMonth(m);
    setDialogSel(peopleWithEntries(m));
  };
  const toggleDialogPerson = (id) =>
    setDialogSel((s) => (s.includes(id) ? s.filter((x) => x !== id) : [...s, id]));

  const dialogFileCount = dialogSel.filter((pid) => entriesCountFor(pid, exportMonth) > 0).length;

  // ---- Export oficiálních výkazů ----
  const exportOfficial = async () => {
    if (!template) return showToast("Nejdřív nahrajte šablonu výkazu.");
    if (dialogSel.length === 0) return showToast("Vyberte alespoň jednoho zaměstnance.");
    setBusy(true);
    try {
      const templateBuf = await getTemplateBuffer(template);
      const sheetName = MESICE[exportMonth - 1];
      const year = template.year;
      const mm = String(exportMonth).padStart(2, "0");
      let files = 0;

      for (const pid of dialogSel) {
        const name = personName(pid);
        const personEntries = entries.filter(
          (e) => e.personId === pid && monthKey(e.date) === `${year}-${mm}`
        );
        if (personEntries.length === 0) continue;

        // Sestavit denní buňky
        const byDay = {};
        personEntries.forEach((e) => {
          const day = parseInt(e.date.slice(8, 10));
          (byDay[day] = byDay[day] || []).push(e);
        });
        const dayCells = {};
        Object.entries(byDay).forEach(([day, list]) => {
          dayCells[day] = buildDayCells(list);
        });

        // Vyplnění přímou úpravou XML — šablona zůstává nedotčená
        const out = await fillOfficialSheet(templateBuf.slice(0), sheetName, name, dayCells);

        const fname = `ALT_Mesicni_vykaz_${year}_${mm}_${safeName(name)}.xlsx`;
        downloadBlob(fname, new Blob([out], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }));
        files++;
        await new Promise((res) => setTimeout(res, 400));
      }

      if (files === 0) {
        showToast(`Za ${MESICE[exportMonth - 1].toLowerCase()} ${year} nejsou u vybraných zaměstnanců žádné záznamy.`);
      } else {
        showToast(`Vygenerováno ${files} souborů.`);
        setDialogOpen(false);
      }
    } catch (err) {
      console.error(err);
      showToast("Chyba při exportu: " + err.message);
    }
    setBusy(false);
  };

  // ---- CSV export ----
  const inRange = (e) => (!from || e.date >= from) && (!to || e.date <= to);
  const rowsFor = (pids) => {
    const rows = [["Datum", "Osoba", "Činnost", "Projekt", "Hodiny", "Poznámka"]];
    const list = entries
      .filter((e) => pids.includes(e.personId) && inRange(e))
      .sort((a, b) => (a.personId + a.date > b.personId + b.date ? 1 : -1));
    list.forEach((e) =>
      rows.push([fmtDate(e.date), personName(e.personId), e.cinnost || "Práce", projectLabel(e.projectId), e.hours ? String(e.hours).replace(".", ",") : "", [e.odkudKam, e.note].filter(Boolean).join(" · ")])
    );
    const total = list.reduce((s, e) => s + (e.hours || 0), 0);
    rows.push([]);
    rows.push(["", "", "", "Celkem hodin", String(total).replace(".", ","), ""]);
    return { rows, count: list.length };
  };

  const doCsvExport = () => {
    if (selected.length === 0) return showToast("Vyberte alespoň jednu osobu.");
    const range = `${from}_${to}`;
    if (csvMode === "combined") {
      const { rows, count } = rowsFor(selected);
      if (count === 0) return showToast("V daném období nejsou žádné záznamy.");
      downloadFile(`vykaz_${range}.csv`, buildCsv(rows));
      showToast(`Exportováno ${count} záznamů.`);
    } else {
      let files = 0;
      selected.forEach((pid, i) => {
        const { rows, count } = rowsFor([pid]);
        if (count > 0) {
          setTimeout(() => downloadFile(`vykaz_${safeName(personName(pid))}_${range}.csv`, buildCsv(rows)), i * 350);
          files++;
        }
      });
      showToast(files === 0 ? "V daném období nejsou žádné záznamy." : `Stahuje se ${files} souborů…`);
    }
  };

  return (
    <div style={{ display: "grid", gap: 14 }}>
      {/* Oficiální výkaz */}
      <section style={{ ...S.panel, borderLeft: `4px solid ${P.accent}` }}>
        <h2 style={{ ...S.h2, marginBottom: 6 }}>Oficiální měsíční výkaz (Excel)</h2>
        <p style={{ margin: "0 0 12px", color: P.muted, fontSize: 13 }}>
          Vyplní firemní šablonu ALT — jméno, činnost, projekty s hodinami, pracovní cesty a časy.
          Vzorce a formátování zůstávají zachované. Každý zaměstnanec = samostatný soubor.
        </p>

        {!template ? (
          <div style={{ display: "grid", gap: 10 }}>
            <div style={{ fontSize: 14, color: P.warn }}>
              Šablona měsíčního výkazu zatím není nahraná — bez ní nelze exportovat.
            </div>
            {isAdmin && (
              <button onClick={() => fileRef.current?.click()} style={S.btn} disabled={busy}>
                Nahrát šablonu výkazu (.xlsx)
              </button>
            )}
          </div>
        ) : (
          <>
            <div style={{ display: "flex", gap: 8, alignItems: "center", flexWrap: "wrap", marginBottom: 12, fontSize: 13 }}>
              <span style={{ background: P.accentSoft, color: P.accent, padding: "6px 12px", borderRadius: 8, fontWeight: 600, overflow: "hidden", textOverflow: "ellipsis", maxWidth: "100%" }}>
                Šablona: {template.name} (rok {template.year})
              </span>
              {isAdmin && (
                <>
                  <button onClick={() => fileRef.current?.click()} style={S.btnGhost} disabled={busy}>Nahrát novější šablonu…</button>
                  <button onClick={importPeopleFromTemplate} style={S.btnGhost} disabled={busy}>Načíst lidi ze šablony</button>
                </>
              )}
            </div>
            <button onClick={openExportDialog} style={{ ...S.btn, width: isMobile ? "100%" : "auto" }} disabled={busy}>
              Exportovat výkazy…
            </button>
            <p style={{ margin: "10px 0 0", fontSize: 12, color: P.muted }}>
              Tip: jména osob musí přesně odpovídat listu Tabulky, aby se dotáhlo pracoviště — použijte „Načíst lidi ze šablony".
            </p>
          </>
        )}
        <input ref={fileRef} type="file" accept=".xlsx" onChange={onTemplateFile} style={{ display: "none" }} />
      </section>

      {/* CSV export */}
      <section style={S.panel}>
        <h2 style={S.h2}>Rychlý export dat (CSV)</h2>
        <div style={{ display: "grid", gridTemplateColumns: isMobile ? "1fr 1fr" : "repeat(3, minmax(140px, 220px))", gap: 10, marginBottom: 12 }}>
          <div>
            <label style={S.label}>Od</label>
            <input type="date" value={from} onChange={(e) => setFrom(e.target.value)} style={S.input} />
          </div>
          <div>
            <label style={S.label}>Do</label>
            <input type="date" value={to} onChange={(e) => setTo(e.target.value)} style={S.input} />
          </div>
          <div style={{ gridColumn: isMobile ? "1 / -1" : "auto" }}>
            <label style={S.label}>Formát</label>
            <select value={csvMode} onChange={(e) => setCsvMode(e.target.value)} style={S.input}>
              <option value="combined">Jeden soubor za všechny vybrané</option>
              <option value="separate">Soubor pro každou osobu</option>
            </select>
          </div>
        </div>
        <label style={S.label}>Osoby pro CSV</label>
        <div style={{ display: "flex", flexWrap: "wrap", gap: 8, marginBottom: 12 }}>
          {people.map((p) => (
            <button key={p.id} onClick={() => togglePerson(p.id)} style={S.chip(selected.includes(p.id))}>
              {selected.includes(p.id) ? "✓ " : ""}{p.name}
            </button>
          ))}
          {people.length > 1 && (
            <button onClick={() => setSelected(selected.length === people.length ? [] : people.map((p) => p.id))} style={{ ...S.chip(false), borderStyle: "dashed" }}>
              {selected.length === people.length ? "Zrušit výběr" : "Vybrat všechny"}
            </button>
          )}
        </div>
        <button onClick={doCsvExport} style={{ ...S.btnGhost, width: isMobile ? "100%" : "auto" }}>Stáhnout CSV</button>
      </section>

      {/* Dialog exportu oficiálních výkazů */}
      {dialogOpen && template && (
        <div
          style={{
            position: "fixed",
            inset: 0,
            background: "rgba(28,37,48,0.55)",
            display: "flex",
            alignItems: isMobile ? "stretch" : "center",
            justifyContent: "center",
            zIndex: 100,
            padding: isMobile ? 0 : 16,
          }}
          onClick={(e) => { if (e.target === e.currentTarget && !busy) setDialogOpen(false); }}
        >
          <div
            style={{
              background: "#fff",
              borderRadius: isMobile ? 0 : 14,
              width: "100%",
              maxWidth: isMobile ? "100%" : 520,
              maxHeight: isMobile ? "100%" : "85vh",
              height: isMobile ? "100%" : "auto",
              display: "flex",
              flexDirection: "column",
              boxShadow: "0 12px 40px rgba(0,0,0,0.3)",
            }}
          >
            <div style={{ padding: "16px 18px", paddingTop: isMobile ? "max(16px, env(safe-area-inset-top))" : 16, borderBottom: `1px solid ${P.line}` }}>
              <div style={{ display: "flex", alignItems: "center", marginBottom: 12 }}>
                <h3 style={{ margin: 0, fontSize: 17, flex: 1 }}>Export měsíčních výkazů</h3>
                <button onClick={() => !busy && setDialogOpen(false)} style={{ ...S.btnDanger, color: P.muted, fontSize: 20 }} title="Zavřít">✕</button>
              </div>
              <label style={S.label}>Měsíc</label>
              <select value={exportMonth} onChange={(e) => onDialogMonthChange(parseInt(e.target.value))} style={S.input} disabled={busy}>
                {MESICE.map((m, i) => (
                  <option key={m} value={i + 1}>{m} {template.year}</option>
                ))}
              </select>
            </div>

            <div style={{ padding: "12px 18px", overflowY: "auto", flex: 1, WebkitOverflowScrolling: "touch" }}>
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 8 }}>
                <label style={{ ...S.label, marginBottom: 0 }}>Zaměstnanci</label>
                <button
                  onClick={() => setDialogSel(dialogSel.length === people.length ? [] : people.map((p) => p.id))}
                  style={{ ...S.btnDanger, color: P.accent, fontWeight: 600, fontSize: 13 }}
                  disabled={busy}
                >
                  {dialogSel.length === people.length ? "Zrušit výběr" : "Vybrat všechny"}
                </button>
              </div>
              {people.length === 0 ? (
                <div style={{ color: P.muted, fontSize: 14 }}>V aplikaci zatím nejsou žádní zaměstnanci.</div>
              ) : (
                <div style={{ display: "grid", gap: 6 }}>
                  {people.map((p) => {
                    const count = entriesCountFor(p.id, exportMonth);
                    const on = dialogSel.includes(p.id);
                    return (
                      <label
                        key={p.id}
                        style={{
                          display: "flex", alignItems: "center", gap: 12, padding: "12px 12px",
                          border: `1px solid ${on ? P.accent : P.line}`,
                          background: on ? P.accentSoft : "#fff",
                          borderRadius: 10, cursor: "pointer", fontSize: 15,
                          opacity: count === 0 ? 0.55 : 1,
                          minHeight: 48,
                          boxSizing: "border-box",
                        }}
                      >
                        <input type="checkbox" checked={on} onChange={() => toggleDialogPerson(p.id)} disabled={busy} style={{ accentColor: P.accent, width: 18, height: 18, flexShrink: 0 }} />
                        <span style={{ flex: 1, fontWeight: on ? 600 : 400, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{p.name}</span>
                        <span style={{ ...S.mono, fontSize: 12, color: count === 0 ? P.warn : P.muted, flexShrink: 0 }}>
                          {count === 0 ? "bez záznamů" : `${count} zázn.`}
                        </span>
                      </label>
                    );
                  })}
                </div>
              )}
            </div>

            <div
              style={{
                padding: "14px 18px",
                paddingBottom: isMobile ? "max(14px, env(safe-area-inset-bottom))" : 14,
                borderTop: `1px solid ${P.line}`,
                display: "grid",
                gap: 10,
              }}
            >
              <span style={{ fontSize: 13, color: P.muted }}>
                {dialogFileCount === 0
                  ? "Nevygeneruje se žádný soubor"
                  : dialogFileCount === 1
                  ? "Vygeneruje se 1 soubor"
                  : dialogFileCount < 5
                  ? `Vygenerují se ${dialogFileCount} soubory`
                  : `Vygeneruje se ${dialogFileCount} souborů`}
                {dialogSel.length > dialogFileCount && ` (${dialogSel.length - dialogFileCount} vybraných bez záznamů se přeskočí)`}
              </span>
              <div style={{ display: "flex", gap: 10 }}>
                <button onClick={() => setDialogOpen(false)} style={{ ...S.btnGhost, flex: 1 }} disabled={busy}>Zrušit</button>
                <button onClick={exportOfficial} style={{ ...S.btn, flex: 2, opacity: dialogFileCount === 0 ? 0.5 : 1 }} disabled={busy || dialogFileCount === 0}>
                  {busy ? "Generuji…" : `Stáhnout (${dialogFileCount})`}
                </button>
              </div>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

// ---------- Záložka: Projekty ----------
function TabProjekty({ data, actions, showToast }) {
  const { projects, entries } = data;
  const [name, setName] = useState("");
  const [code, setCode] = useState("");
  const [importText, setImportText] = useState("");
  const [bulk, setBulk] = useState(false);
  const [bulkSel, setBulkSel] = useState([]);
  const fileRef = useRef(null);

  const toggleBulk = (id) => setBulkSel((s) => (s.includes(id) ? s.filter((x) => x !== id) : [...s, id]));
  const projectUsed = (id) => entries.some((e) => e.projectId === id);

  const deleteBulk = async () => {
    if (bulkSel.length === 0) return showToast("Nic není vybráno.");
    const deletable = bulkSel.filter((id) => !projectUsed(id));
    const skipped = bulkSel.length - deletable.length;
    if (deletable.length === 0) return showToast("Všechny vybrané projekty mají výkazy — nesmazáno. Smažte je jednotlivě, pokud opravdu chcete.");
    if (!window.confirm(`Smazat ${deletable.length} projektů?${skipped ? ` (${skipped} s výkazy se přeskočí)` : ""}`)) return;
    const ok = await actions.removeProjectsBulk(deletable, `Smazáno ${deletable.length} projektů.${skipped ? ` ${skipped} s výkazy přeskočeno.` : ""}`);
    if (ok) {
      setBulkSel([]);
      setBulk(false);
    }
  };

  const addProject = async () => {
    if (!name.trim()) return showToast("Zadejte název projektu.");
    const ok = await actions.addProjects([{ name: name.trim(), code: code.trim() }], "Projekt přidán.");
    if (ok) {
      setName("");
      setCode("");
    }
  };

  const parseAndImport = async (text) => {
    const lines = text.split(/\r?\n/).map((l) => l.trim()).filter(Boolean);
    if (lines.length === 0) return showToast("Nic k importu.");
    const existing = new Set(projects.map((p) => (p.code + "|" + p.name).toLowerCase()));
    const newOnes = [];
    lines.forEach((line) => {
      let c = "", n = line;
      const m = line.split(/[;\t]/);
      if (m.length >= 2) {
        // Explicitní oddělení středníkem nebo tabulátorem
        c = m[0].trim();
        n = m.slice(1).join(" ").trim();
      } else if (line.includes(",") && !line.split(",")[0].trim().includes(" ")) {
        // Formát "KÓD,Název" — jen když první část vypadá jako kód (bez mezer)
        const cm = line.split(",");
        c = cm[0].trim();
        n = cm.slice(1).join(",").trim();
      } else {
        // Rozpoznání kódu na začátku řádku: "P250176 - Název" / "P250176 – Název"
        const dm = line.match(/^([A-Za-z]{1,4}\d{3,}[\w./-]*)\s*[-–—:]\s+(.+)$/);
        if (dm) {
          c = dm[1].trim();
          n = dm[2].trim();
        }
      }
      if (!n) return;
      const key = (c + "|" + n).toLowerCase();
      if (!existing.has(key)) {
        existing.add(key);
        newOnes.push({ name: n, code: c });
      }
    });
    if (newOnes.length === 0) return showToast("Žádné nové projekty (duplicity přeskočeny).");
    const ok = await actions.addProjects(newOnes, `Importováno ${newOnes.length} projektů.`);
    if (ok) setImportText("");
  };

  const onFile = (ev) => {
    const f = ev.target.files?.[0];
    if (!f) return;
    const reader = new FileReader();
    reader.onload = () => parseAndImport(String(reader.result || ""));
    reader.readAsText(f, "utf-8");
    ev.target.value = "";
  };

  const removeProject = (id) => {
    const used = entries.some((e) => e.projectId === id);
    if (used && !window.confirm("Na tento projekt existují výkazy. Opravdu smazat? Záznamy zůstanou, ale bez názvu projektu.")) return;
    actions.removeProject(id);
  };

  return (
    <div style={{ display: "grid", gap: 14 }}>
      <section style={S.panel}>
        <h2 style={S.h2}>Přidat projekt ručně</h2>
        <div style={{ display: "grid", gap: 10 }}>
          <div style={{ display: "grid", gridTemplateColumns: "110px 1fr", gap: 10 }}>
            <input placeholder="Kód" value={code} onChange={(e) => setCode(e.target.value)} style={S.input} />
            <input placeholder="Název projektu" value={name} onChange={(e) => setName(e.target.value)} style={S.input} onKeyDown={(e) => e.key === "Enter" && addProject()} />
          </div>
          <button onClick={addProject} style={S.btn}>Přidat</button>
        </div>
      </section>

      <section style={S.panel}>
        <h2 style={{ ...S.h2, marginBottom: 6 }}>Import seznamu projektů</h2>
        <p style={{ margin: "0 0 10px", color: P.muted, fontSize: 13 }}>
          Vložte seznam (jeden projekt na řádek) nebo nahrajte soubor CSV/TXT. Formáty řádku:
          {" "}<code>Název</code>, <code>KÓD;Název</code>, <code>KÓD,Název</code>. Duplicity se přeskočí.
        </p>
        <textarea
          value={importText}
          onChange={(e) => setImportText(e.target.value)}
          placeholder={"P001;Rekonstrukce hala A\nP002;Servis klient XY\nInterní režie"}
          rows={5}
          style={{ ...S.input, fontFamily: "monospace", fontSize: 14, resize: "vertical" }}
        />
        <div style={{ display: "grid", gap: 10, marginTop: 10 }}>
          <button onClick={() => parseAndImport(importText)} style={S.btn}>Importovat vložený text</button>
          <button onClick={() => fileRef.current?.click()} style={S.btnGhost}>Nahrát soubor CSV/TXT…</button>
          <input ref={fileRef} type="file" accept=".csv,.txt" onChange={onFile} style={{ display: "none" }} />
        </div>
      </section>

      <section style={S.panel}>
        <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 12 }}>
          <h2 style={{ ...S.h2, marginBottom: 0 }}>Seznam projektů ({projects.length})</h2>
          {projects.length > 0 && (
            <button
              onClick={() => { setBulk(!bulk); setBulkSel([]); }}
              style={{ ...S.btnDanger, color: P.accent, fontWeight: 600, fontSize: 13 }}
            >
              {bulk ? "Hotovo" : "Upravit"}
            </button>
          )}
        </div>
        {bulk && (
          <div style={{ display: "flex", gap: 8, alignItems: "center", marginBottom: 10, flexWrap: "wrap" }}>
            <button onClick={() => setBulkSel(bulkSel.length === projects.length ? [] : projects.map((p) => p.id))} style={S.btnGhost}>
              {bulkSel.length === projects.length ? "Zrušit výběr" : "Vybrat vše"}
            </button>
            <button onClick={deleteBulk} style={{ ...S.btnGhost, color: P.warn, borderColor: P.warn, opacity: bulkSel.length ? 1 : 0.5 }} disabled={bulkSel.length === 0}>
              Smazat vybrané ({bulkSel.length})
            </button>
          </div>
        )}
        {projects.length === 0 ? (
          <div style={{ color: P.muted }}>Zatím žádné projekty. Přidejte je ručně nebo importem výše.</div>
        ) : (
          <div style={{ display: "grid", gap: 6 }}>
            {projects.map((p) => (
              <div
                key={p.id}
                onClick={bulk ? () => toggleBulk(p.id) : undefined}
                style={{
                  display: "flex", alignItems: "center", gap: 10, padding: "10px 12px",
                  border: `1px solid ${bulk && bulkSel.includes(p.id) ? P.warn : P.line}`,
                  background: bulk && bulkSel.includes(p.id) ? "#FBEFE7" : "#fff",
                  borderRadius: 10, minHeight: 44, boxSizing: "border-box",
                  cursor: bulk ? "pointer" : "default",
                }}
              >
                {bulk && (
                  <input type="checkbox" readOnly checked={bulkSel.includes(p.id)} style={{ accentColor: P.warn, width: 18, height: 18, flexShrink: 0, pointerEvents: "none" }} />
                )}
                {p.code && <span style={{ ...S.mono, fontSize: 12, background: P.accentSoft, color: P.accent, padding: "3px 8px", borderRadius: 6, fontWeight: 700, flexShrink: 0 }}>{p.code}</span>}
                <span style={{ flex: 1, overflow: "hidden", textOverflow: "ellipsis" }}>{p.name}</span>
                {bulk && projectUsed(p.id) && <span style={{ fontSize: 11, color: P.muted, flexShrink: 0 }}>má výkazy</span>}
                {!bulk && <button style={S.btnDanger} onClick={() => removeProject(p.id)} title="Smazat projekt">✕</button>}
              </div>
            ))}
          </div>
        )}
      </section>
    </div>
  );
}

// ---------- Záložka: Lidé ----------
function TabLide({ data, actions, showToast }) {
  const { people, entries } = data;
  const [name, setName] = useState("");
  const [importText, setImportText] = useState("");
  const [bulk, setBulk] = useState(false);
  const [bulkSel, setBulkSel] = useState([]);
  const fileRef = useRef(null);

  const toggleBulk = (id) => setBulkSel((s) => (s.includes(id) ? s.filter((x) => x !== id) : [...s, id]));
  const personUsed = (id) => entries.some((e) => e.personId === id);

  const deleteBulk = async () => {
    if (bulkSel.length === 0) return showToast("Nikdo není vybrán.");
    const deletable = bulkSel.filter((id) => !personUsed(id));
    const skipped = bulkSel.length - deletable.length;
    if (deletable.length === 0) return showToast("Všichni vybraní mají výkazy — nesmazáno. Smažte je jednotlivě, pokud opravdu chcete.");
    if (!window.confirm(`Smazat ${deletable.length} lidí?${skipped ? ` (${skipped} s výkazy se přeskočí)` : ""}`)) return;
    const ok = await actions.removePeopleBulk(deletable, `Smazáno ${deletable.length} lidí.${skipped ? ` ${skipped} s výkazy přeskočeno.` : ""}`);
    if (ok) {
      setBulkSel([]);
      setBulk(false);
    }
  };

  const addPerson = async () => {
    if (!name.trim()) return showToast("Zadejte jméno.");
    if (people.some((p) => p.name.toLowerCase() === name.trim().toLowerCase())) return showToast("Tato osoba už existuje.");
    const ok = await actions.addPeople([{ name: name.trim() }], "Osoba přidána.");
    if (ok) setName("");
  };

  const parseAndImportPeople = async (text) => {
    const lines = text.split(/\r?\n/).map((l) => l.trim()).filter(Boolean);
    if (lines.length === 0) return showToast("Nic k importu.");
    const existing = new Set(people.map((p) => p.name.toLowerCase()));
    const newOnes = [];
    lines.forEach((line) => {
      const n = line.split(/[;\t,]/)[0].trim();
      if (!n) return;
      const key = n.toLowerCase();
      if (!existing.has(key)) {
        existing.add(key);
        newOnes.push({ name: n });
      }
    });
    if (newOnes.length === 0) return showToast("Žádní noví lidé (duplicity přeskočeny).");
    const ok = await actions.addPeople(newOnes, `Importováno ${newOnes.length} lidí.`);
    if (ok) setImportText("");
  };

  const onFile = (ev) => {
    const f = ev.target.files?.[0];
    if (!f) return;
    const reader = new FileReader();
    reader.onload = () => parseAndImportPeople(String(reader.result || ""));
    reader.readAsText(f, "utf-8");
    ev.target.value = "";
  };

  const removePerson = (id) => {
    const used = entries.some((e) => e.personId === id);
    if (used && !window.confirm("Tato osoba má výkazy. Opravdu smazat? Záznamy zůstanou, ale bez jména.")) return;
    actions.removePerson(id);
  };

  const hoursFor = (id) => entries.filter((e) => e.personId === id).reduce((s, e) => s + (e.hours || 0), 0);

  return (
    <div style={{ display: "grid", gap: 14 }}>
      <section style={S.panel}>
        <h2 style={S.h2}>Přidat osobu</h2>
        <p style={{ margin: "0 0 10px", color: P.muted, fontSize: 13 }}>
          Jméno zadejte přesně jako ve firemním výkazu (Příjmení Jméno), aby export dotáhl pracoviště.
          Seznam lze načíst i z nahrané šablony na záložce Export.
        </p>
        <div style={{ display: "grid", gap: 10 }}>
          <input placeholder="Příjmení Jméno" value={name} onChange={(e) => setName(e.target.value)} style={S.input} onKeyDown={(e) => e.key === "Enter" && addPerson()} />
          <button onClick={addPerson} style={S.btn}>Přidat</button>
        </div>
      </section>

      <section style={S.panel}>
        <h2 style={{ ...S.h2, marginBottom: 6 }}>Import seznamu lidí</h2>
        <p style={{ margin: "0 0 10px", color: P.muted, fontSize: 13 }}>
          Vložte více jmen najednou (jedno na řádek) nebo nahrajte CSV/TXT.
          Z Excelu se bere první sloupec, další sloupce se ignorují. Duplicity se přeskočí.
        </p>
        <textarea
          value={importText}
          onChange={(e) => setImportText(e.target.value)}
          placeholder={"Novák Jan\nSvobodová Petra\nDvořák Martin"}
          rows={5}
          style={{ ...S.input, fontFamily: "monospace", fontSize: 14, resize: "vertical" }}
        />
        <div style={{ display: "grid", gap: 10, marginTop: 10 }}>
          <button onClick={() => parseAndImportPeople(importText)} style={S.btn}>Importovat vložený text</button>
          <button onClick={() => fileRef.current?.click()} style={S.btnGhost}>Nahrát soubor CSV/TXT…</button>
          <input ref={fileRef} type="file" accept=".csv,.txt" onChange={onFile} style={{ display: "none" }} />
        </div>
      </section>

      <section style={S.panel}>
        <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 12 }}>
          <h2 style={{ ...S.h2, marginBottom: 0 }}>Členové týmu ({people.length})</h2>
          {people.length > 0 && (
            <button
              onClick={() => { setBulk(!bulk); setBulkSel([]); }}
              style={{ ...S.btnDanger, color: P.accent, fontWeight: 600, fontSize: 13 }}
            >
              {bulk ? "Hotovo" : "Upravit"}
            </button>
          )}
        </div>
        {bulk && (
          <div style={{ display: "flex", gap: 8, alignItems: "center", marginBottom: 10, flexWrap: "wrap" }}>
            <button onClick={() => setBulkSel(bulkSel.length === people.length ? [] : people.map((p) => p.id))} style={S.btnGhost}>
              {bulkSel.length === people.length ? "Zrušit výběr" : "Vybrat vše"}
            </button>
            <button onClick={deleteBulk} style={{ ...S.btnGhost, color: P.warn, borderColor: P.warn, opacity: bulkSel.length ? 1 : 0.5 }} disabled={bulkSel.length === 0}>
              Smazat vybrané ({bulkSel.length})
            </button>
          </div>
        )}
        {people.length === 0 ? (
          <div style={{ color: P.muted }}>Zatím nikdo. Přidejte členy týmu výše.</div>
        ) : (
          <div style={{ display: "grid", gap: 6 }}>
            {people.map((p) => (
              <div
                key={p.id}
                onClick={bulk ? () => toggleBulk(p.id) : undefined}
                style={{
                  display: "flex", alignItems: "center", gap: 10, padding: "10px 12px",
                  border: `1px solid ${bulk && bulkSel.includes(p.id) ? P.warn : P.line}`,
                  background: bulk && bulkSel.includes(p.id) ? "#FBEFE7" : "#fff",
                  borderRadius: 10, minHeight: 44, boxSizing: "border-box",
                  cursor: bulk ? "pointer" : "default",
                }}
              >
                {bulk && (
                  <input type="checkbox" readOnly checked={bulkSel.includes(p.id)} style={{ accentColor: P.warn, width: 18, height: 18, flexShrink: 0, pointerEvents: "none" }} />
                )}
                <span style={{ flex: 1, fontWeight: 600, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{p.name}</span>
                {bulk && personUsed(p.id) ? (
                  <span style={{ fontSize: 11, color: P.muted, flexShrink: 0 }}>má výkazy</span>
                ) : (
                  <span style={{ ...S.mono, fontSize: 12, color: P.muted, flexShrink: 0 }}>{fmtHours(hoursFor(p.id))} h</span>
                )}
                {!bulk && <button style={S.btnDanger} onClick={() => removePerson(p.id)} title="Smazat osobu">✕</button>}
              </div>
            ))}
          </div>
        )}
      </section>
    </div>
  );
}


// ---------- Záložka: Správa (jen správce) ----------
function TabSprava({ users, people, actions, showToast, myEmail }) {
  const [email, setEmail] = useState("");
  const [role, setRole] = useState("vedouci");
  const [uprava, setUprava] = useState(false);

  const addUser = async () => {
    const e = email.trim().toLowerCase();
    if (!e || !e.includes("@")) return showToast("Zadejte platný e-mail.");
    if (users.some((u) => u.id === e)) return showToast("Tento účet už v aplikaci je.");
    const ok = await actions.setUser(
      e,
      { role, personIds: [], perms: role === "admin" ? { vykazy: true, cas: true, export: true, zahr: true } : { ...VYCHOZI_PERMS } },
      "Účet přidán. Nezapomeňte mu založit přihlášení ve Firebase konzoli."
    );
    if (ok) setEmail("");
  };

  // Zachová stávající nastavení a změní jen zadanou část
  const uprav = (u, patch, msg) =>
    actions.setUser(u.id, {
      role: u.role,
      personIds: u.personIds || [],
      perms: u.perms || { ...VYCHOZI_PERMS },
      ...patch,
    }, msg);

  const prepniModul = (u, modul) => {
    const p = { ...(u.perms || VYCHOZI_PERMS) };
    p[modul] = !p[modul];
    uprav(u, { perms: p });
  };

  const vybratVsechny = (u) =>
    uprav(u, { personIds: (u.personIds || []).length === people.length ? [] : people.map((p) => p.id) });

  const changeRole = (u, newRole) => {
    if (u.id === myEmail && newRole !== "admin") {
      const admins = users.filter((x) => x.role === "admin");
      if (admins.length <= 1) return showToast("Nemůžete odebrat roli poslednímu správci.");
    }
    uprav(u, { role: newRole }, "Role změněna.");
  };

  const toggleTeamMember = (u, personId) => {
    const ids = u.personIds || [];
    const next = ids.includes(personId) ? ids.filter((x) => x !== personId) : [...ids, personId];
    uprav(u, { personIds: next });
  };

  const removeUser = (u) => {
    if (u.id === myEmail) return showToast("Vlastní účet smazat nemůžete.");
    if (!window.confirm(`Odebrat účet ${u.id} z aplikace? (Přihlašovací účet ve Firebase zůstane — smažte ho tam, pokud chcete.)`)) return;
    actions.removeUser(u.id, "Účet odebrán.");
  };

  const sortedUsers = users.slice().sort((a, b) => (a.role === b.role ? a.id.localeCompare(b.id) : a.role === "admin" ? -1 : 1));

  return (
    <div style={{ display: "grid", gap: 14 }}>
      <section style={S.panel}>
        <h2 style={{ ...S.h2, marginBottom: 6 }}>Přidat účet</h2>
        <p style={{ margin: "0 0 10px", color: P.muted, fontSize: 13 }}>
          Zadejte e-mail a roli. Samotné přihlášení (e-mail + heslo) se zakládá ve Firebase konzoli:
          Authentication → Users → Add user. E-mail tady musí být stejný.
        </p>
        <div style={{ display: "grid", gap: 10 }}>
          <input type="email" placeholder="jmeno@altepro.cz" value={email} onChange={(e) => setEmail(e.target.value)} style={S.input} onKeyDown={(e) => e.key === "Enter" && addUser()} />
          <select value={role} onChange={(e) => setRole(e.target.value)} style={S.input}>
            <option value="vedouci">Běžný účet — přístup a lidi nastavíte níže</option>
            <option value="admin">Správce — vidí a spravuje vše</option>
          </select>
          <button onClick={addUser} style={S.btn}>Přidat účet</button>
        </div>
      </section>

      <section style={S.panel}>
        <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 12 }}>
          <h2 style={{ ...S.h2, marginBottom: 0 }}>Účty a týmy ({users.length})</h2>
          <button onClick={() => setUprava(!uprava)} style={{ ...S.btnDanger, color: P.accent, fontWeight: 600, fontSize: 13 }}>
            {uprava ? "Hotovo" : "Upravit"}
          </button>
        </div>
        <p style={{ margin: "0 0 12px", color: P.muted, fontSize: 13 }}>
          U každého účtu nastavte, do kterých modulů se dostane a za koho smí vykazovat.
          Kdo není vybraný, toho účet vůbec neuvidí — ani v přehledech, ani při zadávání.
          Pro vykazování jen za sebe vyberte jedinou osobu.
        </p>
        <div style={{ display: "grid", gap: 10 }}>
          {sortedUsers.map((u) => (
            <div key={u.id} style={{ border: `1px solid ${P.line}`, borderRadius: 10, padding: 12 }}>
              <div style={{ display: "flex", alignItems: "center", gap: 10, flexWrap: "wrap", marginBottom: 10 }}>
                <span style={{ fontWeight: 700, flex: 1, minWidth: 160, overflow: "hidden", textOverflow: "ellipsis" }}>
                  {u.id}{u.id === myEmail ? " (vy)" : ""}
                </span>
                <select value={u.role} onChange={(e) => changeRole(u, e.target.value)} style={{ ...S.input, width: "auto", padding: "8px 10px", fontSize: 14 }}>
                  <option value="vedouci">Běžný účet</option>
                  <option value="admin">Správce</option>
                </select>
                {uprava && <button style={S.btnDanger} onClick={() => removeUser(u)} title="Odebrat účet">✕</button>}
              </div>
              {u.role === "admin" ? (
                <div style={{ fontSize: 13, color: P.muted, marginTop: 4 }}>
                  Správce má přístup ke všem modulům a může vykazovat za kohokoli.
                </div>
              ) : (
                <>
                  {/* Přístup k modulům */}
                  <label style={S.label}>Přístup k modulům</label>
                  <div style={{ display: "flex", flexWrap: "wrap", gap: 6, marginBottom: 12 }}>
                    {MODULY.map((m) => {
                      const on = !!(u.perms || VYCHOZI_PERMS)[m.id];
                      return (
                        <button
                          key={m.id}
                          onClick={() => prepniModul(u, m.id)}
                          style={{ ...S.chip(on), padding: "7px 12px", fontSize: 13, minHeight: 36 }}
                          title={on ? "Kliknutím zakážete" : "Kliknutím povolíte"}
                        >
                          {m.icon} {m.label} {on ? "✓" : "✕"}
                        </button>
                      );
                    })}
                  </div>

                  {/* Za koho smí vykazovat */}
                  <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", gap: 8, flexWrap: "wrap" }}>
                    <label style={{ ...S.label, marginBottom: 0 }}>
                      Smí vykazovat za ({(u.personIds || []).length})
                    </label>
                    {people.length > 0 && (
                      <button onClick={() => vybratVsechny(u)} style={{ ...S.btnDanger, color: P.accent, fontWeight: 600, fontSize: 12 }}>
                        {(u.personIds || []).length === people.length ? "Zrušit vše" : "Vybrat vše"}
                      </button>
                    )}
                  </div>
                  {people.length === 0 ? (
                    <div style={{ color: P.muted, fontSize: 13 }}>Nejdřív přidejte lidi na záložce Lidé.</div>
                  ) : (
                    <div style={{ display: "flex", flexWrap: "wrap", gap: 6, marginTop: 6 }}>
                      {people.map((p) => {
                        const on = (u.personIds || []).includes(p.id);
                        return (
                          <button key={p.id} onClick={() => toggleTeamMember(u, p.id)} style={{ ...S.chip(on), padding: "7px 12px", fontSize: 13, minHeight: 36 }}>
                            {on ? "✓ " : ""}{p.name}
                          </button>
                        );
                      })}
                    </div>
                  )}
                  {(u.personIds || []).length === 0 && (
                    <div style={{ color: P.warn, fontSize: 12, marginTop: 6 }}>
                      Nikdo nevybrán — tento účet zatím nemůže nic vykázat.
                    </div>
                  )}
                </>
              )}
            </div>
          ))}
        </div>
      </section>
    </div>
  );
}


// ---------- Dialog: změna hesla ----------
function PasswordDialog({ onClose, showToast, isMobile }) {
  const [current, setCurrent] = useState("");
  const [pw1, setPw1] = useState("");
  const [pw2, setPw2] = useState("");
  const [busy, setBusy] = useState(false);
  const [err, setErr] = useState(null);

  const submit = async () => {
    if (!current) return setErr("Zadejte současné heslo.");
    if (pw1.length < 6) return setErr("Nové heslo musí mít alespoň 6 znaků.");
    if (pw1 !== pw2) return setErr("Nová hesla se neshodují.");
    if (pw1 === current) return setErr("Nové heslo je stejné jako současné.");
    setBusy(true);
    setErr(null);
    try {
      await changePassword(current, pw1);
      showToast("Heslo změněno.");
      onClose();
    } catch (e) {
      const map = {
        "auth/invalid-credential": "Současné heslo není správné.",
        "auth/wrong-password": "Současné heslo není správné.",
        "auth/weak-password": "Nové heslo je příliš slabé (min. 6 znaků).",
        "auth/too-many-requests": "Příliš mnoho pokusů — chvíli počkejte.",
        "auth/network-request-failed": "Chyba připojení k internetu.",
      };
      setErr(map[e.code] || "Změna se nepodařila: " + e.message);
    }
    setBusy(false);
  };

  return (
    <div
      style={{ position: "fixed", inset: 0, background: "rgba(28,37,48,0.55)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 100, padding: 16 }}
      onClick={(e) => { if (e.target === e.currentTarget && !busy) onClose(); }}
    >
      <div style={{ background: "#fff", borderRadius: 14, width: "100%", maxWidth: 380, padding: 20, boxShadow: "0 12px 40px rgba(0,0,0,0.3)" }}>
        <div style={{ display: "flex", alignItems: "center", marginBottom: 14 }}>
          <h3 style={{ margin: 0, fontSize: 17, flex: 1 }}>Změna hesla</h3>
          <button onClick={() => !busy && onClose()} style={{ ...S.btnDanger, color: P.muted, fontSize: 20 }} title="Zavřít">✕</button>
        </div>
        <div style={{ display: "grid", gap: 12 }}>
          <div>
            <label style={S.label}>Současné heslo</label>
            <input type="password" autoComplete="current-password" value={current} onChange={(e) => setCurrent(e.target.value)} style={S.input} />
          </div>
          <div>
            <label style={S.label}>Nové heslo (min. 6 znaků)</label>
            <input type="password" autoComplete="new-password" value={pw1} onChange={(e) => setPw1(e.target.value)} style={S.input} />
          </div>
          <div>
            <label style={S.label}>Nové heslo znovu</label>
            <input type="password" autoComplete="new-password" value={pw2} onChange={(e) => setPw2(e.target.value)} style={S.input} onKeyDown={(e) => e.key === "Enter" && submit()} />
          </div>
          {err && <div style={{ color: P.warn, fontSize: 13 }}>{err}</div>}
          <button onClick={submit} style={S.btn} disabled={busy}>
            {busy ? "Ukládám…" : "Změnit heslo"}
          </button>
        </div>
      </div>
    </div>
  );
}

// ---------- Záložka: Zahraniční pracovní cesty ----------
function TabZahranicni({ data, actions, showToast, isMobile, isAdmin, zahrTemplate }) {
  const { people, projects, trips } = data;
  const [zeme, setZeme] = useState([]);          // tabulka zemí ze šablony
  const [form, setForm] = useState(null);        // rozpracovaná cesta (null = zavřeno)
  const [busy, setBusy] = useState(false);
  const [kurzBusy, setKurzBusy] = useState(false);
  const autoNastaveno = useRef(false);

  const sablonaRef = useRef(null);

  // Načtení tabulky zemí ze šablony uložené v databázi
  useEffect(() => {
    if (!zahrTemplate || !zahrTemplate.b64) { setZeme([]); return; }
    (async () => {
      try {
        const list = await readCountryRates(b64ToBuf(zahrTemplate.b64));
        setZeme(list);
      } catch (e) {
        console.error("Sazby:", e);
        showToast("Sazby zemí se nepodařilo načíst ze šablony.");
      }
    })();
  }, [zahrTemplate]);

  // Nahrání šablony zahraniční cesty (jen správce)
  const nahratSablonu = async (ev) => {
    const f = ev.target.files?.[0];
    ev.target.value = "";
    if (!f) return;
    setBusy(true);
    try {
      const buf = await f.arrayBuffer();
      const list = await readCountryRates(buf);
      if (list.length === 0) throw new Error("V souboru chybí list Sazby se zeměmi.");
      const b64 = bufToB64(buf);
      if (b64.length > 900000) throw new Error("Soubor je příliš velký (limit ~0,9 MB).");
      await actions.saveZahrTemplate({ name: f.name, uploadedAt: new Date().toISOString(), b64 });
    } catch (e) {
      console.error(e);
      showToast("Chyba: " + e.message);
    }
    setBusy(false);
  };

  const personName = (id) => people.find((p) => p.id === id)?.name || "(neznámý)";
  const zemeInfo = (nazev) => zeme.find((z) => z.zeme === nazev) || null;
  const safeName = (s) => s.normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/[^a-zA-Z0-9]+/g, "-").replace(/^-|-$/g, "");

  // ---- Nová / editace cesty ----
  const novaCesta = () => (autoNastaveno.current = false, autoKurzKlic.current = "", setForm({
    id: null, personId: people[0]?.id || "", ucel: "", cil: "", zeme: "",
    doprava: DOPRAVA[0], zacatek: today(), konec: today(),
    kurz: "", kapesne: "", zaloha: "", days: [],
    vykazCinnost: "Pracovní cesta zahraniční",
  }));

  const editCesta = (t) => setForm({
    ...t, kurz: t.kurz ?? "", kapesne: t.kapesne ?? "", zaloha: t.zaloha ?? "",
    vykazCinnost: t.vykazCinnost ?? "Pracovní cesta zahraniční",
  });

  // Přegenerování dnů podle rozsahu (zachová už zadané hodnoty)
  const syncDays = (f) => {
    const list = listDays(f.zacatek, f.konec, false).slice(0, 14);
    const old = {};
    (f.days || []).forEach((d) => { old[d.date] = d; });
    return list.map((date, i) => old[date] || {
      date,
      misto: f.cil || "",
      od: i === 0 ? "8:00" : "0:01",
      do: i === list.length - 1 ? "18:00" : "23:59",
      jidla: "",
    });
  };

  useEffect(() => {
    if (!form) return;
    const next = syncDays(form);
    const same = JSON.stringify(next.map((d) => d.date)) === JSON.stringify((form.days || []).map((d) => d.date));
    if (!same) setForm((f) => ({ ...f, days: next }));
  }, [form?.zacatek, form?.konec]);

  const setF = (patch) => setForm((f) => ({ ...f, ...patch }));
  const setDay = (i, patch) => setForm((f) => ({ ...f, days: f.days.map((d, ix) => (ix === i ? { ...d, ...patch } : d)) }));

  // ---- Kurz ČNB ----
  const stahniKurz = async () => {
    const info = zemeInfo(form.zeme);
    if (!info) return showToast("Nejdřív vyberte zemi.");
    if (info.mena === "CZK") return showToast("Země používá koruny — kurz není potřeba.");
    setKurzBusy(true);
    try {
      const datum = kurzDatumProCestu(form.zacatek);
      const { kurz, datum: platnyOd } = await nactiKurzCNB(info.mena, datum);
      setF({ kurz: String(kurz) });
      showToast(`Kurz ČNB k ${fmtDate(platnyOd)}: 1 ${info.mena} = ${kurz} Kč`);
    } catch (e) {
      console.error(e);
      showToast(e.message);
    }
    setKurzBusy(false);
  };

  // Automatické doplnění kurzu po výběru země nebo změně data začátku.
  // Ruční hodnotu nepřepisujeme — jen doplňujeme prázdné pole.
  const autoKurzKlic = useRef("");
  useEffect(() => {
    if (!form || !form.zeme || zeme.length === 0) return;
    const info = zemeInfo(form.zeme);
    if (!info || info.mena === "CZK") return;
    const klic = info.mena + "|" + form.zacatek;
    if (autoKurzKlic.current === klic) return;
    const rucne = String(form.kurz || "").trim() !== "" && !autoNastaveno.current;
    if (rucne) return;
    autoKurzKlic.current = klic;
    let zruseno = false;
    (async () => {
      try {
        const { kurz } = await nactiKurzCNB(info.mena, kurzDatumProCestu(form.zacatek));
        if (!zruseno) {
          autoNastaveno.current = true;
          setF({ kurz: String(kurz) });
        }
      } catch (e) {
        // Tichý neúspěch — uživatel může použít tlačítko nebo zadat ručně
        console.warn("Automatický kurz:", e.message);
      }
    })();
    return () => { zruseno = true; };
  }, [form?.zeme, form?.zacatek, zeme]);

  // ---- Souhrn výpočtu ----
  const souhrn = useMemo(() => {
    if (!form) return null;
    const info = zemeInfo(form.zeme);
    if (!info) return null;
    let diety = 0, kap = 0;
    (form.days || []).forEach((d) => {
      const v = vypocetDne(d.od, d.do, d.jidla, info.sazba, form.kapesne);
      diety += v.diety;
      kap += v.kapesne;
    });
    const celkem = Math.round((diety + kap) * 100) / 100;
    const zaloha = Number(form.zaloha) || 0;
    const kurz = Number(String(form.kurz).replace(",", ".")) || 0;
    return {
      mena: info.mena, sazba: info.sazba, diety, kap, celkem,
      kProplaceni: Math.round((celkem - zaloha) * 100) / 100,
      kc: kurz ? Math.ceil((celkem - zaloha) * kurz) : null,
    };
  }, [form, zeme]);

  // ---- Propsání cesty do měsíčního výkazu (sdílené pro uložení i tlačítko) ----
  const syncDoVykazu = async (rec, id) => {
    const stare = (data.entries || []).filter((e) => e.tripId === id).map((e) => e.id);
    if (stare.length) await actions.removeEntriesBulk(stare);
    if (!rec.vykazCinnost) return 0;
    const projekt = projects.find((p) => p.code && rec.ucel && p.code.toLowerCase() === rec.ucel.toLowerCase());
    const dny = listDays(rec.zacatek, rec.konec, true); // jen pracovní dny (bez víkendů a svátků)
    if (dny.length === 0) return 0;
    const nove = dny.map((date) => ({
      date,
      personId: rec.personId,
      cinnost: rec.vykazCinnost,
      projectId: projekt ? projekt.id : "",
      hours: jeAbsence(rec.vykazCinnost) ? 0 : 8,
      note: `Zahraniční cesta – ${rec.cil}`,
      odkudKam: "",
      ucel: rec.ucel || "",
      casOd: "",
      casDo: "",
      jidla: null,
      tripId: id,
      createdAt: new Date().toISOString(),
    }));
    const ok = await actions.addEntries(nove);
    return ok ? nove.length : 0;
  };

  // Ruční propsání z tlačítka u cesty
  const propsat = async (t) => {
    if (!t.vykazCinnost) {
      return showToast("U této cesty není zvolena činnost — otevřete Upravit a vyberte, co se má do výkazu zapsat.");
    }
    const stare = (data.entries || []).filter((e) => e.tripId === t.id).length;
    const dny = listDays(t.zacatek, t.konec, true).length;
    const txt = `Zapsat do měsíčního výkazu ${dny} pracovních dní jako „${t.vykazCinnost}"?` +
      (stare ? `\n\nStávajících ${stare} navázaných záznamů se nahradí.` : "");
    if (!window.confirm(txt)) return;
    setBusy(true);
    const n = await syncDoVykazu(t, t.id);
    setBusy(false);
    showToast(n ? `Do výkazu zapsáno ${n} dní.` : "Nebylo co zapsat (v rozsahu nejsou pracovní dny).");
  };

  // ---- Uložení ----
  const ulozit = async () => {
    if (!form.personId) return showToast("Vyberte osobu.");
    if (!form.zeme) return showToast("Vyberte zemi.");
    if (!form.cil.trim()) return showToast("Vyplňte cíl cesty.");
    if (form.konec < form.zacatek) return showToast("Konec cesty nesmí být před začátkem.");
    if ((form.days || []).length > 14) return showToast("Formulář pojme maximálně 14 dní.");
    const info = zemeInfo(form.zeme);
    const rec = {
      personId: form.personId, ucel: form.ucel.trim(), cil: form.cil.trim(), zeme: form.zeme,
      doprava: form.doprava, zacatek: form.zacatek, konec: form.konec,
      mena: info?.mena || "", sazba: info?.sazba || 0,
      kurz: Number(String(form.kurz).replace(",", ".")) || 0,
      kapesne: Number(form.kapesne) || 0,
      zaloha: Number(form.zaloha) || 0,
      days: (form.days || []).map((d) => ({ date: d.date, misto: d.misto || "", od: d.od || "", do: d.do || "", jidla: Number(d.jidla) || 0 })),
      updatedAt: new Date().toISOString(),
    };
    rec.vykazCinnost = form.vykazCinnost || "";
    setBusy(true);
    const tripId = await actions.saveTrip(form.id, rec);
    if (!tripId) { setBusy(false); return; }
    const id = form.id || tripId;

    // Synchronizace se měsíčním výkazem
    const zapsano = await syncDoVykazu(rec, id);
    setBusy(false);
    showToast(
      (form.id ? "Cesta upravena." : "Cesta uložena.") +
        (zapsano ? ` Do měsíčního výkazu zapsáno ${zapsano} dní jako „${rec.vykazCinnost}".` : "")
    );
    setForm(null);
  };

  // ---- Export jedné cesty ----
  const exportCesta = async (t) => {
    setBusy(true);
    try {
      if (!zahrTemplate || !zahrTemplate.b64) throw new Error("Šablona zahraniční cesty není nahraná.");
      const out = await fillForeignTrip(b64ToBuf(zahrTemplate.b64), { ...t, employeeName: personName(t.personId) });
      const fname = `Zahr_cesta_${t.zacatek}_${safeName(personName(t.personId))}_${safeName(t.cil)}.xlsx`;
      downloadBlob(fname, new Blob([out], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }));
      showToast("Formulář vygenerován.");
    } catch (e) {
      console.error(e);
      showToast("Chyba při exportu: " + e.message);
    }
    setBusy(false);
  };

  const smazat = async (t) => {
    const navazane = (data.entries || []).filter((e) => e.tripId === t.id);
    const txt = `Smazat cestu ${personName(t.personId)} — ${t.cil} (${fmtDate(t.zacatek)})?` +
      (navazane.length ? `\n\nSmažou se i navázané záznamy v měsíčním výkazu (${navazane.length} dní).` : "");
    if (!window.confirm(txt)) return;
    if (navazane.length) await actions.removeEntriesBulk(navazane.map((e) => e.id));
    actions.removeTrip(t.id, "Cesta smazána.");
  };

  const sorted = (trips || []).slice().sort((a, b) => (a.zacatek < b.zacatek ? 1 : -1));

  // ================= FORMULÁŘ =================
  if (form) {
    const info = zemeInfo(form.zeme);
    const maxKapesne = info ? Math.round(info.sazba * 0.4 * 100) / 100 : null;
    return (
      <div style={{ display: "grid", gap: 14 }}>
        <section style={S.panel}>
          <div style={{ display: "flex", alignItems: "center", marginBottom: 14 }}>
            <h2 style={{ ...S.h2, marginBottom: 0, flex: 1 }}>{form.id ? "Úprava cesty" : "Nová zahraniční cesta"}</h2>
            <button onClick={() => setForm(null)} style={{ ...S.btnDanger, color: P.muted, fontSize: 20 }} title="Zavřít">✕</button>
          </div>

          <div style={{ display: "grid", gridTemplateColumns: isMobile ? "1fr" : "repeat(auto-fit, minmax(200px, 1fr))", gap: 12 }}>
            <div>
              <label style={S.label}>Zaměstnanec</label>
              <select value={form.personId} onChange={(e) => setF({ personId: e.target.value })} style={S.input}>
                {people.map((p) => <option key={p.id} value={p.id}>{p.name}</option>)}
              </select>
            </div>
            <div>
              <label style={S.label}>Účel cesty / projekt</label>
              <input list="projekty-list" value={form.ucel} onChange={(e) => setF({ ucel: e.target.value })} placeholder="P260114" style={S.input} />
              <datalist id="projekty-list">
                {projects.map((p) => <option key={p.id} value={p.code || p.name}>{p.name}</option>)}
              </datalist>
            </div>
            <div>
              <label style={S.label}>Cíl cesty (město)</label>
              <input value={form.cil} onChange={(e) => setF({ cil: e.target.value })} placeholder="Schwabenheim" style={S.input} />
            </div>
            <div>
              <label style={S.label}>Země</label>
              <select value={form.zeme} onChange={(e) => setF({ zeme: e.target.value })} style={S.input}>
                <option value="">— vyberte zemi —</option>
                {zeme.map((z) => <option key={z.zeme} value={z.zeme}>{z.zeme} ({z.sazba} {z.mena})</option>)}
              </select>
            </div>
            <div>
              <label style={S.label}>Způsob dopravy</label>
              <select value={form.doprava} onChange={(e) => setF({ doprava: e.target.value })} style={S.input}>
                {DOPRAVA.map((d) => <option key={d} value={d}>{d}</option>)}
              </select>
            </div>
            <div>
              <label style={S.label}>Začátek cesty</label>
              <input type="date" value={form.zacatek} onChange={(e) => setF({ zacatek: e.target.value })} style={S.input} />
            </div>
            <div style={{ gridColumn: isMobile ? "auto" : "1 / -1" }}>
              <label style={S.label}>Zapsat do měsíčního výkazu jako</label>
              <select value={form.vykazCinnost} onChange={(e) => setF({ vykazCinnost: e.target.value })} style={S.input}>
                {ZAHR_DO_VYKAZU.map((o) => <option key={o.v} value={o.v}>{o.label}</option>)}
              </select>
              <div style={{ fontSize: 11, color: P.muted, marginTop: 4 }}>
                Tento formulář je příloha měsíčního výkazu — do něj se za pracovní dny cesty
                (bez víkendů a svátků) zapíše zvolená činnost. Při úpravě nebo smazání cesty se záznamy upraví taky.
              </div>
            </div>
            <div>
              <label style={S.label}>Konec cesty</label>
              <input type="date" value={form.konec} min={form.zacatek} onChange={(e) => setF({ konec: e.target.value })} style={S.input} />
            </div>
          </div>

          {/* Finanční část */}
          <div style={{ display: "grid", gridTemplateColumns: isMobile ? "1fr" : "repeat(auto-fit, minmax(180px, 1fr))", gap: 12, marginTop: 12, padding: 12, background: P.accentSoft, borderRadius: 10 }}>
            <div>
              <label style={S.label}>Kurz ČNB {info && info.mena !== "CZK" ? `(1 ${info.mena} = ? Kč)` : ""}</label>
              <div style={{ display: "flex", gap: 8 }}>
                <input inputMode="decimal" value={form.kurz} onChange={(e) => { autoNastaveno.current = false; setF({ kurz: e.target.value }); }} placeholder="24,17" style={{ ...S.input, ...S.mono }} />
                <button onClick={stahniKurz} style={{ ...S.btnGhost, whiteSpace: "nowrap" }} disabled={kurzBusy || !info}>
                  {kurzBusy ? "…" : "Stáhnout"}
                </button>
              </div>
              <div style={{ fontSize: 11, color: P.muted, marginTop: 4 }}>
                Kurz ke dni začátku cesty (o víkendu páteční).
              </div>
            </div>
            <div>
              <label style={S.label}>Kapesné / den {maxKapesne ? `(max ${maxKapesne} ${info.mena})` : ""}</label>
              <input inputMode="decimal" value={form.kapesne} onChange={(e) => setF({ kapesne: e.target.value })} placeholder="0" style={{ ...S.input, ...S.mono }} />
            </div>
            <div>
              <label style={S.label}>Vyplacená záloha {info ? `(${info.mena})` : ""}</label>
              <input inputMode="decimal" value={form.zaloha} onChange={(e) => setF({ zaloha: e.target.value })} placeholder="0" style={{ ...S.input, ...S.mono }} />
            </div>
          </div>

          {/* Dny */}
          <h3 style={{ fontSize: 14, margin: "16px 0 8px" }}>Pobyt v zahraničí ({(form.days || []).length} dní)</h3>
          {(form.days || []).length > 14 && (
            <div style={{ color: P.warn, fontSize: 13, marginBottom: 8 }}>Formulář pojme jen 14 dní — zkraťte rozsah.</div>
          )}
          <div style={{ display: "grid", gap: 8 }}>
            {(form.days || []).map((d, i) => {
              const v = info ? vypocetDne(d.od, d.do, d.jidla, info.sazba, form.kapesne) : null;
              return (
                <div key={d.date} style={{ border: `1px solid ${P.line}`, borderRadius: 10, padding: 10, display: "grid", gridTemplateColumns: isMobile ? "1fr 1fr" : "110px 1.4fr 90px 90px 80px 1fr", gap: 8, alignItems: "end" }}>
                  <div style={{ gridColumn: isMobile ? "1 / -1" : "auto", ...S.mono, fontWeight: 700, fontSize: 13, paddingBottom: isMobile ? 0 : 10 }}>
                    {fmtDate(d.date)}
                  </div>
                  <div style={{ gridColumn: isMobile ? "1 / -1" : "auto" }}>
                    <label style={S.label}>Místo pobytu</label>
                    <input value={d.misto} onChange={(e) => setDay(i, { misto: e.target.value })} style={S.input} />
                  </div>
                  <div>
                    <label style={S.label}>Od</label>
                    <input value={d.od} onChange={(e) => setDay(i, { od: e.target.value })} placeholder="8:00" inputMode="numeric" style={{ ...S.input, ...S.mono }} />
                  </div>
                  <div>
                    <label style={S.label}>Do</label>
                    <input value={d.do} onChange={(e) => setDay(i, { do: e.target.value })} placeholder="18:00" inputMode="numeric" style={{ ...S.input, ...S.mono }} />
                  </div>
                  <div>
                    <label style={S.label}>Jídla</label>
                    <select value={d.jidla} onChange={(e) => setDay(i, { jidla: e.target.value })} style={S.input}>
                      <option value="">0</option>
                      {[1, 2, 3, 4].map((n) => <option key={n} value={n}>{n}</option>)}
                    </select>
                  </div>
                  <div style={{ gridColumn: isMobile ? "1 / -1" : "auto", fontSize: 13, color: P.muted, paddingBottom: isMobile ? 0 : 10, textAlign: isMobile ? "right" : "left" }}>
                    {v ? `${Math.floor(v.min / 60)} h ${v.min % 60} min → ` : ""}
                    {v ? <b style={{ ...S.mono, color: P.accent }}>{v.celkem.toLocaleString("cs-CZ")} {info.mena}</b> : "—"}
                  </div>
                </div>
              );
            })}
          </div>

          {/* Souhrn */}
          {souhrn && (
            <div style={{ marginTop: 14, padding: 12, border: `1px solid ${P.accent}`, borderRadius: 10, background: "#fff", fontSize: 14 }}>
              <div style={{ display: "flex", justifyContent: "space-between", flexWrap: "wrap", gap: 8 }}>
                <span style={{ color: P.muted }}>Sazba: <b>{souhrn.sazba} {souhrn.mena}/den</b></span>
                <span style={{ color: P.muted }}>Diety: <b style={S.mono}>{souhrn.diety.toFixed(2)} {souhrn.mena}</b></span>
                {souhrn.kap > 0 && <span style={{ color: P.muted }}>Kapesné: <b style={S.mono}>{souhrn.kap.toFixed(2)} {souhrn.mena}</b></span>}
                <span>Celkem: <b style={{ ...S.mono, color: P.accent }}>{souhrn.celkem.toFixed(2)} {souhrn.mena}</b></span>
              </div>
              <div style={{ marginTop: 8, paddingTop: 8, borderTop: `1px solid ${P.line}`, display: "flex", justifyContent: "space-between", flexWrap: "wrap", gap: 8 }}>
                <span>{souhrn.kProplaceni >= 0 ? "K proplacení zaměstnanci:" : "Zaměstnanec vrátí:"}{" "}
                  <b style={S.mono}>{Math.abs(souhrn.kProplaceni).toFixed(2)} {souhrn.mena}</b>
                </span>
                {souhrn.kc != null && <span>Převod na účet: <b style={{ ...S.mono, color: P.accent }}>{Math.abs(souhrn.kc).toLocaleString("cs-CZ")} Kč</b></span>}
              </div>
              <div style={{ marginTop: 8, fontSize: 12, color: P.muted }}>
                Ostatní výdaje (ubytování, PHM, doklady) se doplňují ručně v exportovaném souboru.
              </div>
            </div>
          )}

          <div style={{ display: "flex", gap: 10, marginTop: 14 }}>
            <button onClick={() => setForm(null)} style={{ ...S.btnGhost, flex: 1 }} disabled={busy}>Zrušit</button>
            <button onClick={ulozit} style={{ ...S.btn, flex: 2 }} disabled={busy}>{busy ? "Ukládám…" : "Uložit cestu"}</button>
          </div>
        </section>
      </div>
    );
  }

  // ================= SEZNAM =================
  return (
    <div style={{ display: "grid", gap: 14 }}>
      <section style={S.panel}>
        <div style={{ display: "flex", alignItems: "center", gap: 10, flexWrap: "wrap" }}>
          <h2 style={{ ...S.h2, marginBottom: 0, flex: 1 }}>Zahraniční pracovní cesty</h2>
          <button onClick={novaCesta} style={S.btn} disabled={people.length === 0 || !zahrTemplate}>+ Nová cesta</button>
        </div>
        <p style={{ margin: "10px 0 0", color: P.muted, fontSize: 13 }}>
          Samostatný formulář (příkaz + vyúčtování) pro každou cestu. Stravné, kapesné i přepočet na koruny
          počítá formulář sám podle země a časů; kurz ČNB umí aplikace stáhnout.
        </p>
        <div style={{ display: "flex", gap: 8, alignItems: "center", flexWrap: "wrap", marginTop: 12, fontSize: 13 }}>
          {zahrTemplate ? (
            <span style={{ background: P.accentSoft, color: P.accent, padding: "6px 12px", borderRadius: 8, fontWeight: 600 }}>
              Šablona: {zahrTemplate.name} · {zeme.length} zemí
            </span>
          ) : (
            <span style={{ color: P.warn, fontWeight: 600 }}>
              Šablona zahraniční cesty není nahraná — bez ní nelze zadávat ani exportovat.
            </span>
          )}
          {isAdmin && (
            <button onClick={() => sablonaRef.current?.click()} style={S.btnGhost} disabled={busy}>
              {zahrTemplate ? "Nahrát novější šablonu…" : "Nahrát šablonu (.xlsx)"}
            </button>
          )}
          <input ref={sablonaRef} type="file" accept=".xlsx" onChange={nahratSablonu} style={{ display: "none" }} />
        </div>
      </section>

      {sorted.length === 0 ? (
        <section style={S.panel}>
          <div style={{ color: P.muted }}>Zatím žádné cesty. Přidejte první tlačítkem „Nová cesta".</div>
        </section>
      ) : (
        <section style={S.panel}>
          <div style={{ display: "grid", gap: 8 }}>
            {sorted.map((t) => (
              <div key={t.id} style={{ border: `1px solid ${P.line}`, borderRadius: 10, padding: 12 }}>
                <div style={{ display: "flex", alignItems: "baseline", gap: 8, flexWrap: "wrap" }}>
                  <span style={{ fontWeight: 700, flex: 1, minWidth: 140 }}>{personName(t.personId)}</span>
                  <span style={{ ...S.mono, fontSize: 13, color: P.muted }}>
                    {fmtDate(t.zacatek)} – {fmtDate(t.konec)}
                  </span>
                </div>
                <div style={{ fontSize: 13, color: P.muted, marginTop: 4 }}>
                  {t.cil}, {t.zeme}
                  {t.ucel ? ` · ${t.ucel}` : ""}
                  {t.sazba ? ` · ${t.sazba} ${t.mena}/den` : ""}
                  {t.kurz ? ` · kurz ${t.kurz}` : ""}
                </div>
                <div style={{ fontSize: 12, color: t.vykazCinnost ? P.accent : P.warn, marginTop: 4 }}>
                  {t.vykazCinnost
                    ? `Ve výkazu jako „${t.vykazCinnost}"`
                    : "Do měsíčního výkazu nezapsáno — doplňte ručně"}
                </div>
                <div style={{ display: "flex", gap: 8, marginTop: 10, flexWrap: "wrap" }}>
                  <button onClick={() => exportCesta(t)} style={{ ...S.btn, padding: "9px 14px", minHeight: 40, fontSize: 14 }} disabled={busy}>
                    Stáhnout formulář
                  </button>
                  <button onClick={() => propsat(t)} style={{ ...S.btnGhost, minHeight: 40 }} disabled={busy} title="Znovu zapsat dny cesty do měsíčního výkazu">
                    Propsat do výkazu
                  </button>
                  <button onClick={() => editCesta(t)} style={{ ...S.btnGhost, minHeight: 40 }}>Upravit</button>
                  <button onClick={() => smazat(t)} style={{ ...S.btnGhost, minHeight: 40, color: P.warn, borderColor: P.warn }}>Smazat</button>
                </div>
              </div>
            ))}
          </div>
        </section>
      )}
    </div>
  );
}

// ---------- Záložka: Čas na projektu (podklad pro Odoo Timesheets) ----------

// "9:30" nebo "9.5" -> desetinné hodiny (Odoo formát: 9,5 = 9:30)
function casNaDesetinne(t) {
  if (t == null || t === "") return null;
  const s = String(t).trim().replace(",", ".");
  const m = s.match(/^(\d{1,2}):(\d{2})$/);
  if (m) {
    const h = parseInt(m[1]), mi = parseInt(m[2]);
    if (h > 23 || mi > 59) return null;
    return h + mi / 60;
  }
  const n = parseFloat(s);
  return isNaN(n) || n < 0 || n >= 24 ? null : n;
}

const desetinneNaCas = (d) => {
  if (d == null) return "";
  const h = Math.floor(d);
  const m = Math.round((d - h) * 60);
  return `${h}:${String(m).padStart(2, "0")}`;
};

// "Al-tamimi Zaid" -> "Zaid Al-tamimi" (Odoo používá pořadí jméno příjmení)
function odooJmeno(p) {
  if (p.odooName) return p.odooName;
  const casti = (p.name || "").trim().split(/\s+/);
  if (casti.length < 2) return p.name || "";
  return casti.slice(1).join(" ") + " " + casti[0];
}

const odooProjekt = (pr) => {
  if (!pr) return "";
  if (pr.odooName) return pr.odooName;
  return pr.code ? `${pr.code} - ${pr.name}` : pr.name;
};

function TabCas({ data, actions, showToast, isMobile, isAdmin }) {
  // Terénní zadávání — rozvržení je vždy mobilní, na počítači jen vycentrované
  const { people, projects, timesheet } = data;
  const [date, setDate] = useState(today());
  const [dateTo, setDateTo] = useState("");
  const [jenPracovni, setJenPracovni] = useState(true);
  const [projectId, setProjectId] = useState("");
  const [rezim, setRezim] = useState("trvani"); // trvani | odDo
  const [hod, setHod] = useState("");
  const [min, setMin] = useState("");
  const [od, setOd] = useState("");
  const [doC, setDoC] = useState("");
  const [popis, setPopis] = useState("");
  const [selected, setSelected] = useState([]);
  const [saving, setSaving] = useState(false);
  const [filterPerson, setFilterPerson] = useState("");
  const [filterMonth, setFilterMonth] = useState(monthKey(today()));
  const [uprava, setUprava] = useState(false);
  const [vybrane, setVybrane] = useState([]);
  const [expDialog, setExpDialog] = useState(false);

  const personName = (id) => people.find((p) => p.id === id)?.name || "(neznámý)";
  const projectName = (id) => {
    const pr = projects.find((p) => p.id === id);
    return pr ? (pr.code ? `${pr.code} – ${pr.name}` : pr.name) : "(smazaný projekt)";
  };
  const togglePerson = (id) => setSelected((s) => (s.includes(id) ? s.filter((x) => x !== id) : [...s, id]));

  // Vypočtené trvání podle režimu
  const trvani = useMemo(() => {
    if (rezim === "odDo") {
      const a = casNaDesetinne(od), b = casNaDesetinne(doC);
      if (a == null || b == null) return null;
      const d = b - a;
      return d > 0 ? Math.round(d * 10000) / 10000 : null;
    }
    const h = parseFloat(String(hod).replace(",", ".")) || 0;
    const m = parseFloat(String(min).replace(",", ".")) || 0;
    const d = h + m / 60;
    return d > 0 && d <= 24 ? Math.round(d * 10000) / 10000 : null;
  }, [rezim, hod, min, od, doC]);

  const ulozit = async () => {
    if (saving) return;
    if (!date) return showToast("Vyberte datum.");
    if (!projectId) return showToast("Vyberte projekt.");
    if (trvani == null) {
      return showToast(rezim === "odDo" ? "Zadejte platný čas od a do (např. 9:00 a 12:30)." : "Zadejte platné trvání.");
    }
    if (selected.length === 0) return showToast("Vyberte alespoň jednu osobu.");
    if (dateTo && dateTo < date) return showToast("Datum 'do' nesmí být před datem 'od'.");

    const dny = dateTo ? listDays(date, dateTo, jenPracovni) : [date];
    if (dny.length === 0) return showToast("Ve zvoleném rozsahu nejsou žádné pracovní dny.");
    if (dny.length * selected.length > 400) return showToast("Příliš velký rozsah — zmenšete počet dní nebo osob.");

    const zaznamy = dny.flatMap((den) => selected.map((pid) => ({
      date: den,
      personId: pid,
      projectId,
      popis: popis.trim(),
      od: rezim === "odDo" ? casNaDesetinne(od) : null,
      do: rezim === "odDo" ? casNaDesetinne(doC) : null,
      hodiny: trvani,
      createdAt: new Date().toISOString(),
    })));
    setSaving(true);
    const ok = await actions.addTimesheet(
      zaznamy,
      zaznamy.length === 1
        ? "Zapsáno."
        : dny.length === 1
        ? `Zapsáno pro ${selected.length} osob.`
        : `Zapsáno ${zaznamy.length} záznamů (${dny.length} dní × ${selected.length} os.).`
    );
    setSaving(false);
    if (ok) {
      setHod(""); setMin(""); setOd(""); setDoC(""); setPopis("");
    }
  };

  const filtered = useMemo(
    () =>
      (timesheet || [])
        .filter((t) => (!filterPerson || t.personId === filterPerson) && (!filterMonth || monthKey(t.date) === filterMonth))
        .sort((a, b) => (a.date < b.date ? 1 : a.date > b.date ? -1 : (b.od || 0) - (a.od || 0))),
    [timesheet, filterPerson, filterMonth]
  );

  const celkem = filtered.reduce((s, t) => s + (t.hodiny || 0), 0);
  const podleOsob = useMemo(() => {
    const m = {};
    filtered.forEach((t) => { m[t.personId] = (m[t.personId] || 0) + (t.hodiny || 0); });
    return m;
  }, [filtered]);

  const toggleVyber = (id) => setVybrane((s) => (s.includes(id) ? s.filter((x) => x !== id) : [...s, id]));
  const smazatVybrane = async () => {
    if (vybrane.length === 0) return showToast("Nic není vybráno.");
    if (!window.confirm(`Smazat ${vybrane.length} záznamů?`)) return;
    const ok = await actions.removeTimesheetBulk(vybrane, `Smazáno ${vybrane.length} záznamů.`);
    if (ok) { setVybrane([]); setUprava(false); }
  };

  const noSetup = people.length === 0 || projects.length === 0;

  // Rychlá volba data
  const posun = (dnu) => {
    const d = new Date(date + "T00:00:00Z");
    d.setUTCDate(d.getUTCDate() + dnu);
    setDate(d.toISOString().slice(0, 10));
  };
  const jeDnes = date === today();

  return (
    <div style={{ display: "grid", gap: 14, maxWidth: 600, margin: "0 auto", width: "100%" }}>
      {noSetup && (
        <div style={{ ...S.panel, borderLeft: `4px solid ${P.warn}`, fontSize: 14 }}>
          Nejdřív přidejte lidi a projekty, pak můžete vykazovat čas.
        </div>
      )}

      {/* Zadání */}
      <section style={S.panel}>
        <h2 style={S.h2}>Zapsat čas na projekt</h2>

        <div style={{ display: "grid", gap: 12, marginBottom: 12 }}>
          <div>
            <label style={S.label}>Datum</label>
            <div style={{ display: "flex", gap: 8, alignItems: "center" }}>
              <button onClick={() => posun(-1)} style={{ ...S.btnGhost, minWidth: 46, padding: "11px 0", fontSize: 18 }} title="Předchozí den">‹</button>
              <input type="date" value={date} onChange={(e) => setDate(e.target.value)} style={{ ...S.input, flex: 1, textAlign: "center" }} />
              <button onClick={() => posun(1)} style={{ ...S.btnGhost, minWidth: 46, padding: "11px 0", fontSize: 18 }} title="Další den">›</button>
            </div>
            <div style={{ display: "flex", gap: 12, alignItems: "center", flexWrap: "wrap" }}>
              {!jeDnes && (
                <button onClick={() => setDate(today())} style={{ ...S.btnDanger, color: P.accent, fontWeight: 600, fontSize: 13, padding: "6px 0" }}>
                  Zpět na dnešek
                </button>
              )}
              {!dateTo && (
                <button onClick={() => setDateTo(date)} style={{ ...S.btnDanger, color: P.accent, fontWeight: 600, fontSize: 13, padding: "6px 0" }}>
                  + Více dní
                </button>
              )}
            </div>
          </div>

          {dateTo && (
            <div>
              <label style={S.label}>Do</label>
              <div style={{ display: "flex", gap: 8, alignItems: "center" }}>
                <input type="date" value={dateTo} min={date} onChange={(e) => setDateTo(e.target.value)} style={{ ...S.input, flex: 1, textAlign: "center" }} />
                <button onClick={() => setDateTo("")} style={{ ...S.btnGhost, minWidth: 46, padding: "11px 0", fontSize: 16 }} title="Zrušit rozsah">✕</button>
              </div>
              <label style={{ display: "flex", alignItems: "center", gap: 8, fontSize: 14, cursor: "pointer", marginTop: 8 }}>
                <input type="checkbox" checked={jenPracovni} onChange={(e) => setJenPracovni(e.target.checked)} style={{ accentColor: P.accent, width: 18, height: 18 }} />
                Jen pracovní dny (bez víkendů a svátků)
              </label>
              <div style={{ fontSize: 13, color: P.accent, fontWeight: 600, marginTop: 6 }}>
                → {listDays(date, dateTo, jenPracovni).length} dní
              </div>
            </div>
          )}
          <div>
            <label style={S.label}>Projekt</label>
            <select value={projectId} onChange={(e) => setProjectId(e.target.value)} style={S.input}>
              <option value="">— vyberte projekt —</option>
              {projects.map((p) => (
                <option key={p.id} value={p.id}>{p.code ? `${p.code} – ${p.name}` : p.name}</option>
              ))}
            </select>
          </div>
        </div>

        {/* Přepínač způsobu zadání */}
        <div style={{ display: "flex", gap: 8, marginBottom: 12 }}>
          {[["trvani", "Trvání"], ["odDo", "Od – do"]].map(([v, l]) => (
            <button
              key={v}
              onClick={() => setRezim(v)}
              style={{
                ...S.chip(rezim === v),
                flex: isMobile ? 1 : "0 0 auto",
                minWidth: 110,
              }}
            >
              {l}
            </button>
          ))}
        </div>

        {rezim === "trvani" ? (
          <div style={{ marginBottom: 12 }}>
            <label style={S.label}>Rychlá volba</label>
            <div style={{ display: "grid", gridTemplateColumns: "repeat(4, 1fr)", gap: 6, marginBottom: 10 }}>
              {[[0, 30], [1, 0], [2, 0], [4, 0], [6, 0], [7, 30], [8, 0], [null, null]].map(([h, m], i) => (
                <button
                  key={i}
                  onClick={() => { if (h === null) { setHod(""); setMin(""); } else { setHod(String(h)); setMin(String(m)); } }}
                  style={{
                    ...S.chip(h !== null && String(hod) === String(h) && String(min || 0) === String(m)),
                    padding: "12px 4px", minHeight: 48, fontSize: 14, borderRadius: 10, textAlign: "center",
                  }}
                >
                  {h === null ? "smazat" : m ? `${h}:${String(m).padStart(2, "0")}` : `${h} h`}
                </button>
              ))}
            </div>
            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10 }}>
              <div>
                <label style={S.label}>Hodin</label>
                <input inputMode="numeric" value={hod} onChange={(e) => setHod(e.target.value)} placeholder="2" style={{ ...S.input, ...S.mono, fontSize: 18, textAlign: "center", padding: "14px 10px" }} />
              </div>
              <div>
                <label style={S.label}>Minut</label>
                <input inputMode="numeric" value={min} onChange={(e) => setMin(e.target.value)} placeholder="30" style={{ ...S.input, ...S.mono, fontSize: 18, textAlign: "center", padding: "14px 10px" }} />
              </div>
            </div>
            <div style={{ textAlign: "center", marginTop: 8, fontSize: 15 }}>
              {trvani != null
                ? <>Zapíše se <b style={{ ...S.mono, color: P.accent }}>{trvani.toLocaleString("cs-CZ", { maximumFractionDigits: 2 })} h</b></>
                : <span style={{ color: P.muted }}>Zadejte trvání</span>}
            </div>
          </div>
        ) : (
          <div style={{ marginBottom: 12 }}>
            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10 }}>
              <div>
                <label style={S.label}>Od</label>
                <input type="time" value={od} onChange={(e) => setOd(e.target.value)} style={{ ...S.input, ...S.mono, fontSize: 18, textAlign: "center", padding: "14px 10px" }} />
              </div>
              <div>
                <label style={S.label}>Do</label>
                <input type="time" value={doC} onChange={(e) => setDoC(e.target.value)} style={{ ...S.input, ...S.mono, fontSize: 18, textAlign: "center", padding: "14px 10px" }} />
              </div>
            </div>
            <div style={{ textAlign: "center", marginTop: 8, fontSize: 15 }}>
              {trvani != null
                ? <>Zapíše se <b style={{ ...S.mono, color: P.accent }}>{desetinneNaCas(trvani)}</b> ({trvani.toLocaleString("cs-CZ", { maximumFractionDigits: 2 })} h)</>
                : <span style={{ color: P.muted }}>Zadejte čas od a do</span>}
            </div>
          </div>
        )}

        <div style={{ marginBottom: 12 }}>
          <label style={S.label}>Popis práce</label>
          <input value={popis} onChange={(e) => setPopis(e.target.value)} placeholder="Co se dělalo…" style={S.input} />
        </div>

        <label style={S.label}>Zapsat pro</label>
        <div style={{ display: "flex", flexWrap: "wrap", gap: 8, marginBottom: 14 }}>
          {people.map((p) => (
            <button key={p.id} onClick={() => togglePerson(p.id)} style={S.chip(selected.includes(p.id))}>
              {selected.includes(p.id) ? "✓ " : ""}{p.name}
            </button>
          ))}
          {people.length > 1 && (
            <button
              onClick={() => setSelected(selected.length === people.length ? [] : people.map((p) => p.id))}
              style={{ ...S.chip(false), borderStyle: "dashed" }}
            >
              {selected.length === people.length ? "Zrušit výběr" : "Všichni"}
            </button>
          )}
        </div>

        <button onClick={ulozit} style={{ ...S.btn, width: "100%", padding: "16px 18px", fontSize: 17, minHeight: 54 }} disabled={saving || noSetup}>
          {(() => {
            if (saving) return "Ukládám…";
            const nd = dateTo ? listDays(date, dateTo, jenPracovni).length : 1;
            if (nd > 1 && selected.length > 1) return `Zapsat (${selected.length} os. × ${nd} dní)`;
            if (nd > 1) return `Zapsat (${nd} dní)`;
            if (selected.length > 1) return `Zapsat (${selected.length} osob)`;
            return "Zapsat";
          })()}
        </button>
      </section>

      {/* Přehled */}
      <section style={S.panel}>
        <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 12, gap: 8, flexWrap: "wrap" }}>
          <h2 style={{ ...S.h2, marginBottom: 0 }}>Přehled ({filtered.length})</h2>
          <div style={{ display: "flex", gap: 8, alignItems: "center" }}>
            <button onClick={() => setExpDialog(true)} style={{ ...S.btnGhost, minHeight: 36, padding: "7px 12px" }} disabled={filtered.length === 0}>
              Export pro Odoo
            </button>
            {filtered.length > 0 && (
              <button onClick={() => { setUprava(!uprava); setVybrane([]); }} style={{ ...S.btnDanger, color: P.accent, fontWeight: 600, fontSize: 13 }}>
                {uprava ? "Hotovo" : "Upravit"}
              </button>
            )}
          </div>
        </div>

        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10, marginBottom: 12 }}>
          <div>
            <label style={S.label}>Osoba</label>
            <select value={filterPerson} onChange={(e) => setFilterPerson(e.target.value)} style={S.input}>
              <option value="">Všichni</option>
              {people.map((p) => <option key={p.id} value={p.id}>{p.name}</option>)}
            </select>
          </div>
          <div>
            <label style={S.label}>Měsíc</label>
            <input type="month" value={filterMonth} onChange={(e) => setFilterMonth(e.target.value)} style={S.input} />
          </div>
        </div>


        {uprava && (
          <div style={{ display: "flex", gap: 8, marginBottom: 10, flexWrap: "wrap" }}>
            <button onClick={() => setVybrane(vybrane.length === filtered.length ? [] : filtered.map((t) => t.id))} style={S.btnGhost}>
              {vybrane.length === filtered.length ? "Zrušit výběr" : "Vybrat vše"}
            </button>
            <button onClick={smazatVybrane} style={{ ...S.btnGhost, color: P.warn, borderColor: P.warn, opacity: vybrane.length ? 1 : 0.5 }} disabled={vybrane.length === 0}>
              Smazat vybrané ({vybrane.length})
            </button>
          </div>
        )}

        {filtered.length === 0 ? (
          <div style={{ color: P.muted, padding: "12px 0" }}>Pro zvolený filtr nejsou žádné záznamy.</div>
        ) : (
          <>
            <div style={{ display: "grid", gap: 6 }}>
              {filtered.map((t) => (
                <div
                  key={t.id}
                  onClick={uprava ? () => toggleVyber(t.id) : undefined}
                  style={{
                    border: `1px solid ${uprava && vybrane.includes(t.id) ? P.warn : P.line}`,
                    background: uprava && vybrane.includes(t.id) ? "#FBEFE7" : "#fff",
                    borderRadius: 10, padding: "10px 12px", cursor: uprava ? "pointer" : "default",
                  }}
                >
                  <div style={{ display: "flex", alignItems: "baseline", gap: 8, flexWrap: "wrap" }}>
                    {uprava && (
                      <input type="checkbox" readOnly checked={vybrane.includes(t.id)} style={{ accentColor: P.warn, width: 16, height: 16, pointerEvents: "none" }} />
                    )}
                    <span style={{ ...S.mono, fontWeight: 700, fontSize: 13 }}>{fmtDateShort(t.date)}</span>
                    {t.od != null && (
                      <span style={{ ...S.mono, fontSize: 12, color: P.muted }}>{desetinneNaCas(t.od)}–{desetinneNaCas(t.do)}</span>
                    )}
                    <span style={{ fontSize: 13, flex: 1, minWidth: 120, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>
                      {projectName(t.projectId)}
                    </span>
                    <b style={{ ...S.mono, color: P.accent, fontSize: 13 }}>{(t.hodiny || 0).toLocaleString("cs-CZ", { maximumFractionDigits: 2 })} h</b>
                  </div>
                  <div style={{ fontSize: 12, color: P.muted, marginTop: 3 }}>
                    {personName(t.personId)}{t.popis ? ` · ${t.popis}` : ""}
                  </div>
                </div>
              ))}
            </div>

            <div style={{ display: "flex", gap: 14, flexWrap: "wrap", marginTop: 14, paddingTop: 12, borderTop: `2px solid ${P.line}`, fontSize: 13 }}>
              {Object.entries(podleOsob).map(([pid, h]) => (
                <div key={pid}>
                  <span style={{ color: P.muted }}>{personName(pid)}:</span>{" "}
                  <b style={S.mono}>{h.toLocaleString("cs-CZ", { maximumFractionDigits: 2 })} h</b>
                </div>
              ))}
              <div style={{ marginLeft: "auto" }}>
                Celkem: <b style={{ ...S.mono, color: P.accent }}>{celkem.toLocaleString("cs-CZ", { maximumFractionDigits: 2 })} h</b>
              </div>
            </div>
          </>
        )}
      </section>

      {expDialog && (
        <OdooExportDialog
          zaznamy={filtered}
          people={people}
          projects={projects}
          actions={actions}
          isAdmin={isAdmin}
          showToast={showToast}
          isMobile={isMobile}
          onClose={() => setExpDialog(false)}
        />
      )}
    </div>
  );
}

// ---------- Dialog: export do Odoo Timesheets ----------
function OdooExportDialog({ zaznamy, people, projects, actions, isAdmin, showToast, isMobile, onClose }) {
  const [busy, setBusy] = useState(false);
  const [mapa, setMapa] = useState(false); // zobrazit nastavení párování

  const osoby = useMemo(() => {
    const ids = [...new Set(zaznamy.map((z) => z.personId))];
    return ids.map((id) => people.find((p) => p.id === id)).filter(Boolean);
  }, [zaznamy, people]);

  const projekty = useMemo(() => {
    const ids = [...new Set(zaznamy.map((z) => z.projectId))];
    return ids.map((id) => projects.find((p) => p.id === id)).filter(Boolean);
  }, [zaznamy, projects]);

  const exportuj = async () => {
    setBusy(true);
    try {
      const ExcelJS = await getExcelJS();
      const wb = new ExcelJS.Workbook();
      const ws = wb.addWorksheet("Sheet1");
      ws.addRow(["Datum", "Projekty", "Zaměstnanec", "Úkol", "CRM příležitost", "Popis", "Čas začátku", "Čas konce", "Množství"]);

      const serazene = zaznamy.slice().sort((a, b) => (a.date === b.date ? (a.od || 0) - (b.od || 0) : a.date < b.date ? -1 : 1));
      serazene.forEach((z) => {
        const osoba = people.find((p) => p.id === z.personId);
        const projekt = projects.find((p) => p.id === z.projectId);
        ws.addRow([
          `${z.date} 00:00:00`,
          odooProjekt(projekt),
          osoba ? odooJmeno(osoba) : "",
          "",
          "",
          z.popis || "",
          z.od != null ? Math.round(z.od * 1e6) / 1e6 : "",
          z.do != null ? Math.round(z.do * 1e6) / 1e6 : "",
          Math.round((z.hodiny || 0) * 1e6) / 1e6,
        ]);
      });

      ws.getRow(1).font = { bold: true };
      ws.columns = [{ width: 20 }, { width: 32 }, { width: 20 }, { width: 10 }, { width: 14 }, { width: 46 }, { width: 12 }, { width: 12 }, { width: 12 }];

      const out = await wb.xlsx.writeBuffer();
      const dny = [...new Set(serazene.map((z) => z.date))].sort();
      const nazev = `Odoo_timesheet_${dny[0] || ""}_${dny[dny.length - 1] || ""}.xlsx`;
      downloadBlob(nazev, new Blob([out], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }));
      showToast(`Vygenerováno ${serazene.length} řádků pro Odoo.`);
      onClose();
    } catch (e) {
      console.error(e);
      showToast("Chyba při exportu: " + e.message);
    }
    setBusy(false);
  };

  return (
    <div
      style={{ position: "fixed", inset: 0, background: "rgba(28,37,48,0.55)", display: "flex", alignItems: isMobile ? "stretch" : "center", justifyContent: "center", zIndex: 100, padding: isMobile ? 0 : 16 }}
      onClick={(e) => { if (e.target === e.currentTarget && !busy) onClose(); }}
    >
      <div style={{ background: "#fff", borderRadius: isMobile ? 0 : 14, width: "100%", maxWidth: isMobile ? "100%" : 620, maxHeight: isMobile ? "100%" : "88vh", height: isMobile ? "100%" : "auto", display: "flex", flexDirection: "column", boxShadow: "0 12px 40px rgba(0,0,0,0.3)" }}>
        <div style={{ padding: "16px 18px", paddingTop: isMobile ? "max(16px, env(safe-area-inset-top))" : 16, borderBottom: `1px solid ${P.line}`, display: "flex", alignItems: "center" }}>
          <h3 style={{ margin: 0, fontSize: 17, flex: 1 }}>Export pro Odoo Timesheets</h3>
          <button onClick={() => !busy && onClose()} style={{ ...S.btnDanger, color: P.muted, fontSize: 20 }}>✕</button>
        </div>

        <div style={{ padding: "14px 18px", overflowY: "auto", flex: 1 }}>
          <p style={{ margin: "0 0 12px", color: P.muted, fontSize: 13 }}>
            Vytvoří soubor se sloupci, které Odoo očekává ({zaznamy.length} řádků). Zkontrolujte, že jména
            a projekty odpovídají tomu, jak se jmenují v Odoo — jinak import řádky nespáruje.
          </p>

          <div style={{ border: `1px solid ${P.line}`, borderRadius: 10, padding: 12, marginBottom: 12 }}>
            <div style={{ ...S.label, marginBottom: 8 }}>Zaměstnanci v exportu</div>
            {osoby.map((o) => (
              <div key={o.id} style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 6, fontSize: 13, flexWrap: "wrap" }}>
                <span style={{ color: P.muted, minWidth: 130 }}>{o.name}</span>
                <span style={{ color: P.muted }}>→</span>
                {mapa && isAdmin ? (
                  <input
                    defaultValue={odooJmeno(o)}
                    onBlur={(e) => {
                      const v = e.target.value.trim();
                      if (v && v !== odooJmeno(o)) actions.savePerson(o.id, { ...o, odooName: v }, "Uloženo.");
                    }}
                    style={{ ...S.input, flex: 1, minWidth: 160, padding: "6px 8px", fontSize: 14 }}
                  />
                ) : (
                  <b>{odooJmeno(o)}</b>
                )}
              </div>
            ))}
          </div>

          <div style={{ border: `1px solid ${P.line}`, borderRadius: 10, padding: 12 }}>
            <div style={{ ...S.label, marginBottom: 8 }}>Projekty v exportu</div>
            {projekty.map((pr) => (
              <div key={pr.id} style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 6, fontSize: 13, flexWrap: "wrap" }}>
                <span style={{ color: P.muted, minWidth: 130 }}>{pr.code || pr.name}</span>
                <span style={{ color: P.muted }}>→</span>
                {mapa && isAdmin ? (
                  <input
                    defaultValue={odooProjekt(pr)}
                    onBlur={(e) => {
                      const v = e.target.value.trim();
                      if (v && v !== odooProjekt(pr)) actions.saveProject(pr.id, { ...pr, odooName: v }, "Uloženo.");
                    }}
                    style={{ ...S.input, flex: 1, minWidth: 160, padding: "6px 8px", fontSize: 14 }}
                  />
                ) : (
                  <b style={{ wordBreak: "break-word" }}>{odooProjekt(pr)}</b>
                )}
              </div>
            ))}
          </div>

          {isAdmin && (
            <button onClick={() => setMapa(!mapa)} style={{ ...S.btnGhost, marginTop: 12, width: "100%" }}>
              {mapa ? "Skrýt úpravu názvů" : "Upravit názvy pro Odoo"}
            </button>
          )}
        </div>

        <div style={{ padding: "14px 18px", paddingBottom: isMobile ? "max(14px, env(safe-area-inset-bottom))" : 14, borderTop: `1px solid ${P.line}`, display: "flex", gap: 10 }}>
          <button onClick={onClose} style={{ ...S.btnGhost, flex: 1 }} disabled={busy}>Zrušit</button>
          <button onClick={exportuj} style={{ ...S.btn, flex: 2 }} disabled={busy}>
            {busy ? "Generuji…" : `Stáhnout (${zaznamy.length})`}
          </button>
        </div>
      </div>
    </div>
  );
}
