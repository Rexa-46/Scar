import React, { useState, useEffect, useMemo, useCallback, useRef } from "react";
import {
  Bell, Menu, Plus, X, Search, ChevronDown, ChevronLeft, ChevronRight,
  ArrowLeftRight, Landmark, TrendingUp, TrendingDown, Receipt,
  Trash2, Grid3x3, PieChart as PieChartIcon, Home as HomeIcon,
  Tag, Check, Star, CreditCard, Lock, Sun, Moon, Image as ImageIcon,
  Repeat, Download, Upload, Bitcoin, Landmark as Bank, CalendarDays,
  BellRing, FileSpreadsheet, Printer, Users, ShieldCheck, Palette, Save
} from "lucide-react";
import {
  PieChart, Pie, Cell, ResponsiveContainer, BarChart, Bar, XAxis, YAxis, Tooltip,
  LineChart, Line, CartesianGrid
} from "recharts";
import * as XLSX from "xlsx";

/* ---------------------------------------------------------
   Helpers
--------------------------------------------------------- */
const toFaInt = (n) => Math.round(Number(n || 0)).toLocaleString("fa-IR");
const FA_DIGIT_MAP = ["۰", "۱", "۲", "۳", "۴", "۵", "۶", "۷", "۸", "۹"];
const faDigits = (n) => String(n).split("").map((ch) => (/[0-9]/.test(ch) ? FA_DIGIT_MAP[+ch] : ch)).join("");
function jalaliYear(d) {
  try { return parseInt(new Intl.DateTimeFormat("en-US-u-ca-persian", { year: "numeric" }).format(d), 10); }
  catch { return new Date(d).getFullYear(); }
}
function jalaliParts(d) {
  try {
    const parts = new Intl.DateTimeFormat("en-US-u-ca-persian", { year: "numeric", month: "numeric", day: "numeric" }).formatToParts(d);
    return {
      y: +parts.find((p) => p.type === "year").value,
      m: +parts.find((p) => p.type === "month").value,
      day: +parts.find((p) => p.type === "day").value,
    };
  } catch { return { y: d.getFullYear(), m: d.getMonth() + 1, day: d.getDate() }; }
}
function faLongDate(d) {
  try { return new Intl.DateTimeFormat("fa-IR", { weekday: "long", day: "numeric", month: "long", year: "numeric" }).format(d); }
  catch { return d.toDateString(); }
}
function faMonthYear(d) {
  try { return new Intl.DateTimeFormat("fa-IR", { month: "long", year: "numeric" }).format(d); }
  catch { return ""; }
}
function faTime(d) {
  try { return new Intl.DateTimeFormat("fa-IR", { hour: "2-digit", minute: "2-digit" }).format(d); }
  catch { return ""; }
}
function addMonths(dateStr, n) {
  const d = new Date(dateStr);
  d.setMonth(d.getMonth() + n);
  return d.toISOString().slice(0, 10);
}
function addDays(dateStr, n) {
  const d = new Date(dateStr);
  d.setDate(d.getDate() + n);
  return d.toISOString().slice(0, 10);
}
function daysUntil(dateStr) {
  const now = new Date(); now.setHours(0, 0, 0, 0);
  const d = new Date(dateStr); d.setHours(0, 0, 0, 0);
  return Math.round((d - now) / 86400000);
}
const uid = () => Math.random().toString(36).slice(2) + Date.now().toString(36);
const todayISO = () => new Date().toISOString().slice(0, 10);

function findJalaliMonthStart(jYear, jMonth) {
  const anchor = new Date(Date.UTC(jYear + 621, jMonth - 1, 15));
  for (let delta = -45; delta <= 45; delta++) {
    const d = new Date(anchor); d.setUTCDate(d.getUTCDate() + delta);
    const p = jalaliParts(d);
    if (p.y === jYear && p.m === jMonth && p.day === 1) return d;
  }
  return null;
}
function getJalaliMonthCells(jYear, jMonth) {
  const start = findJalaliMonthStart(jYear, jMonth);
  if (!start) return [];
  const cells = []; let d = new Date(start);
  for (let i = 0; i < 32; i++) {
    const p = jalaliParts(d);
    if (p.y !== jYear || p.m !== jMonth) break;
    cells.push({ date: new Date(d), day: p.day, weekday: (d.getUTCDay() + 1) % 7 });
    d.setUTCDate(d.getUTCDate() + 1);
  }
  return cells;
}

function parseBankSms(text) {
  // Best-effort parser for common Iranian bank SMS wording. Not real SMS access —
  // the user pastes the message text and we extract what we can.
  const normalized = text.replace(/[۰-۹]/g, (d) => "0123456789"["۰۱۲۳۴۵۶۷۸۹".indexOf(d)]).replace(/,/g, "");
  const amountMatch = normalized.match(/(\d{4,})/);
  const amount = amountMatch ? parseInt(amountMatch[1], 10) : null;
  let type = "expense";
  if (/واریز|دریافت|credit|deposit/i.test(text)) type = "income";
  if (/برداشت|خرید|انتقال|پرداخت|debit|purchase/i.test(text)) type = "expense";
  return { amount, type, note: text.trim().slice(0, 140) };
}

function resizeImage(file, maxSize = 480) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      const img = new window.Image();
      img.onload = () => {
        let w = img.width, h = img.height;
        if (w > h) { if (w > maxSize) { h = h * maxSize / w; w = maxSize; } }
        else { if (h > maxSize) { w = w * maxSize / h; h = maxSize; } }
        const canvas = document.createElement("canvas");
        canvas.width = w; canvas.height = h;
        canvas.getContext("2d").drawImage(img, 0, 0, w, h);
        resolve(canvas.toDataURL("image/jpeg", 0.7));
      };
      img.onerror = reject;
      img.src = e.target.result;
    };
    reader.onerror = reject;
    reader.readAsDataURL(file);
  });
}

/* ---------------------------------------------------------
   Seed data
--------------------------------------------------------- */
const seedAccounts = () => ([
  { id: uid(), name: "بانک ملت", type: "bank", initial: 5000000 },
  { id: uid(), name: "صندوق نقدی", type: "fund", initial: 800000 },
  { id: uid(), name: "کارت عابر ملت", type: "card", initial: 0 },
]);
const seedCategories = () => ([
  { id: uid(), name: "بنزین", kind: "expense" },
  { id: uid(), name: "خوراک و بازار", kind: "expense" },
  { id: uid(), name: "قبوض", kind: "expense" },
  { id: uid(), name: "حمل و نقل", kind: "expense" },
  { id: uid(), name: "درمان", kind: "expense" },
  { id: uid(), name: "متفرقه", kind: "expense" },
  { id: uid(), name: "حقوق", kind: "income" },
  { id: uid(), name: "درآمد متفرقه", kind: "income" },
]);
const seedSettings = () => ({ theme: "light", pin: "", sharedFamily: false, themeColor: "purple", profile: { name: "alireza shadfar", phone: "", email: "" } });

/* ---------------------------------------------------------
   Theme / palette
   BRAND is intentionally mutable: App() re-assigns its keys from the
   chosen preset on every render, so every component below (which reads
   BRAND.xxx directly at render time) automatically reflects the user's
   color choice without needing a context/hook everywhere.
--------------------------------------------------------- */
const COLOR_PRESETS = {
  purple: { name: "بنفش کلاسیک", header: "#3E1461", mauve: "#A65475", darkgreen: "#1B6B2C", green: "#1E8449", violet: "#6C3FA0", teal: "#4E9AA0", gold: "#A98A3B", crimson: "#B01E4A", orange: "#C56A1F", fab: "#28C76F" },
  ocean: { name: "آبی اقیانوسی", header: "#0B4F6C", mauve: "#3D7EA6", darkgreen: "#0F7173", green: "#14919B", violet: "#145DA0", teal: "#4CC9F0", gold: "#B08968", crimson: "#D64550", orange: "#F2A65A", fab: "#2EC4B6" },
  forest: { name: "سبز جنگلی", header: "#1B4332", mauve: "#40916C", darkgreen: "#2D6A4F", green: "#40916C", violet: "#52796F", teal: "#74C69D", gold: "#B08968", crimson: "#BC4749", orange: "#DDA15E", fab: "#52B788" },
  rose: { name: "گلبهی", header: "#6D2148", mauve: "#B23A5D", darkgreen: "#2D6A4F", green: "#40916C", violet: "#8E3B6C", teal: "#C9738A", gold: "#B08968", crimson: "#C1121F", orange: "#E07A5F", fab: "#F4978E" },
  charcoal: { name: "زغالی تیره", header: "#22223B", mauve: "#4A4E69", darkgreen: "#22577A", green: "#38A3A5", violet: "#4A4E69", teal: "#5C677D", gold: "#9A8C98", crimson: "#C9184A", orange: "#C08552", fab: "#57CC99" },
};
let BRAND = { ...COLOR_PRESETS.purple };
const THEME = {
  light: { bg: "#F1EFF4", card: "#ffffff", text: "#241a30", sub: "#8a8194", border: "#f0eef3", input: "#faf9fb", inputBorder: "#e3e0ea" },
  dark: { bg: "#17131c", card: "#241d2c", text: "#f1eef5", sub: "#a79fb3", border: "#332b3d", input: "#2c2434", inputBorder: "#3d3348" },
};
const FONT = "'Vazirmatn', Tahoma, 'Segoe UI', sans-serif";

// Storage abstraction: uses the Claude-artifact window.storage API when present
// (always true inside claude.ai), and falls back to localStorage automatically
// when this same code runs as a standalone deployed app / APK build.
const hasCloudStorage = typeof window !== "undefined" && window.storage && typeof window.storage.get === "function";
async function loadKey(key, fallback, shared) {
  try {
    if (hasCloudStorage) {
      const res = await window.storage.get(key, !!shared);
      return res && res.value ? JSON.parse(res.value) : fallback;
    }
    const raw = window.localStorage.getItem(key);
    return raw ? JSON.parse(raw) : fallback;
  } catch { return fallback; }
}
async function saveKey(key, value, shared) {
  try {
    if (hasCloudStorage) { await window.storage.set(key, JSON.stringify(value), !!shared); }
    else { window.localStorage.setItem(key, JSON.stringify(value)); }
  } catch (e) { console.error("save fail", e); }
}

/* ---------------------------------------------------------
   Small UI atoms (theme aware)
--------------------------------------------------------- */
function useT() { return React.useContext(ThemeCtx); }
const ThemeCtx = React.createContext(THEME.light);

function GaugeCircle({ value, max, color, label }) {
  const t = useT();
  const size = 156, stroke = 10, r = (size - stroke) / 2, c = 2 * Math.PI * r;
  const frac = max > 0 ? Math.min(value / max, 1) : 0;
  const dash = Math.max(frac * c, value > 0 ? 6 : 0);
  return (
    <div style={{ display: "flex", flexDirection: "column", alignItems: "center", gap: 8 }}>
      <div style={{ position: "relative", width: size, height: size }}>
        <svg width={size} height={size} style={{ transform: "rotate(-90deg)" }}>
          <circle cx={size / 2} cy={size / 2} r={r} fill="none" stroke={t.border} strokeWidth={stroke} />
          <circle cx={size / 2} cy={size / 2} r={r} fill="none" stroke={color} strokeWidth={stroke}
            strokeDasharray={`${dash} ${c - dash}`} strokeLinecap="round" style={{ transition: "stroke-dasharray .6s" }} />
        </svg>
        <div style={{ position: "absolute", inset: 0, display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center" }}>
          <div style={{ fontSize: 18, fontWeight: 800, color: t.text }}>{toFaInt(value)}</div>
          <div style={{ fontSize: 10.5, color: t.sub, marginTop: 2 }}>ریال</div>
        </div>
      </div>
      <div style={{ fontWeight: 700, color, fontSize: 14.5 }}>{label}</div>
    </div>
  );
}

function CollapsibleSection({ color, title, open, onToggle, children, badge }) {
  const t = useT();
  return (
    <div style={{ borderRadius: 14, overflow: "hidden", marginBottom: 12, boxShadow: "0 1px 3px rgba(0,0,0,0.08)" }}>
      <button onClick={onToggle} style={{
        width: "100%", background: color, border: "none", color: "#fff", padding: "15px 16px",
        display: "flex", alignItems: "center", justifyContent: "space-between", cursor: "pointer", fontFamily: "inherit"
      }}>
        <span style={{ width: 32, height: 32, borderRadius: "50%", background: "rgba(255,255,255,.18)", display: "flex", alignItems: "center", justifyContent: "center" }}>
          <ChevronDown size={17} style={{ transform: open ? "rotate(180deg)" : "none", transition: "transform .2s" }} />
        </span>
        <span style={{ display: "flex", alignItems: "center", gap: 8, fontWeight: 700, fontSize: 15 }}>
          {badge}{title}
        </span>
      </button>
      {open && <div style={{ background: t.card, padding: "6px 14px" }}>{children}</div>}
    </div>
  );
}

function Row({ leftIcon, leftColor = "#eee", title, subtitle, value, valueColor, onClick, chevron = "left", extra }) {
  const t = useT();
  return (
    <div onClick={onClick} style={{
      display: "flex", alignItems: "center", justifyContent: "space-between", padding: "12px 4px",
      borderBottom: `1px solid ${t.border}`, cursor: onClick ? "pointer" : "default"
    }}>
      <div style={{ display: "flex", alignItems: "center", gap: 10, minWidth: 0 }}>
        {chevron && <ChevronLeft size={16} color={t.sub} style={{ flexShrink: 0, transform: chevron === "left" ? "none" : "rotate(180deg)" }} />}
        <div style={{ minWidth: 0 }}>
          <div style={{ fontSize: 14.5, fontWeight: 600, color: t.text }}>{title}</div>
          {subtitle && <div style={{ fontSize: 12, color: t.sub, marginTop: 2 }}>{subtitle}</div>}
        </div>
      </div>
      <div style={{ display: "flex", alignItems: "center", gap: 10, flexShrink: 0 }}>
        {extra}
        {value !== undefined && <span style={{ fontSize: 13.5, fontWeight: 700, color: valueColor || t.text }}>{value}</span>}
        {leftIcon && (
          <span style={{ width: 36, height: 36, borderRadius: 10, background: leftColor, display: "flex", alignItems: "center", justifyContent: "center", color: "#fff" }}>
            {leftIcon}
          </span>
        )}
      </div>
    </div>
  );
}
function EmptyRow({ text }) { const t = useT(); return <div style={{ padding: "16px 4px", textAlign: "center", color: t.sub, fontSize: 13 }}>{text}</div>; }
function AddLink({ text, onClick }) { return <div onClick={onClick} style={{ padding: "10px 4px", color: BRAND.header, fontSize: 13, fontWeight: 700, cursor: "pointer", textAlign: "center" }}>{text}</div>; }
function SectionTitle({ text }) { const t = useT(); return <div style={{ fontWeight: 700, fontSize: 14, color: t.text, margin: "6px 4px 10px" }}>{text}</div>; }
function StatusBadge({ text, color }) { return <span style={{ fontSize: 11, fontWeight: 700, color: "#fff", background: color, padding: "3px 9px", borderRadius: 20 }}>{text}</span>; }

function pillStyle(active) {
  return { flex: 1, padding: "8px 4px", borderRadius: 8, border: `1.5px solid ${active ? BRAND.header : "#ccc4d8"}`, background: active ? BRAND.header : "transparent", color: active ? "#fff" : "inherit", cursor: "pointer", fontWeight: 700, fontSize: 13 };
}
function useStyles() {
  const t = useT();
  return {
    input: { width: "100%", padding: "10px 12px", borderRadius: 9, border: `1.5px solid ${t.inputBorder}`, marginBottom: 10, fontSize: 14, outline: "none", background: t.input, color: t.text },
    primaryBtn: { width: "100%", padding: "11px", borderRadius: 9, border: "none", background: BRAND.fab, color: "#fff", fontWeight: 700, fontSize: 14, cursor: "pointer" },
    label: { display: "block", fontSize: 12.5, color: t.sub, fontWeight: 600, marginBottom: 6 },
    card: { background: t.card, borderRadius: 14, boxShadow: "0 1px 3px rgba(0,0,0,0.08)" },
  };
}

/* ---------------------------------------------------------
   Header + BottomNav
--------------------------------------------------------- */
function Header({ title = "alireza shadfar", onMenu, back, onBack }) {
  return (
    <div style={{ background: BRAND.header, color: "#fff", padding: "16px 16px", display: "flex", alignItems: "center", justifyContent: "space-between", position: "sticky", top: 0, zIndex: 20 }}>
      {back ? <button onClick={onBack} style={iconBtn}><ChevronRight size={24} /></button> : <Bell size={22} />}
      <div style={{ fontWeight: 700, fontSize: 17 }}>{title}</div>
      <button onClick={onMenu} style={iconBtn}><Menu size={22} /></button>
    </div>
  );
}
const iconBtn = { background: "none", border: "none", color: "#fff", cursor: "pointer" };

function BottomNav({ active, setActive, onAdd }) {
  const t = useT();
  const items = [
    { key: "operations", label: "عملیات", icon: Grid3x3 },
    { key: "reports", label: "گزارش ها", icon: PieChartIcon },
    { key: "transactions", label: "تراکنش ها", icon: Receipt },
    { key: "home", label: "خانه", icon: HomeIcon },
  ];
  return (
    <div style={{ position: "sticky", bottom: 0, background: t.card, borderTop: `1px solid ${t.border}`, display: "flex", alignItems: "center", justifyContent: "space-around", padding: "8px 4px 10px", zIndex: 20 }}>
      {items.slice(0, 2).map((it) => <NavBtn key={it.key} it={it} active={active} setActive={setActive} />)}
      <button onClick={onAdd} style={{ width: 54, height: 54, borderRadius: "50%", background: BRAND.fab, border: `4px solid ${t.card}`, marginTop: -26, boxShadow: "0 3px 10px rgba(0,0,0,0.25)", display: "flex", alignItems: "center", justifyContent: "center", color: "#fff", cursor: "pointer" }}>
        <Plus size={26} />
      </button>
      {items.slice(2).map((it) => <NavBtn key={it.key} it={it} active={active} setActive={setActive} />)}
    </div>
  );
}
function NavBtn({ it, active, setActive }) {
  const Icon = it.icon; const isActive = active === it.key;
  return (
    <button onClick={() => setActive(it.key)} style={{ background: "none", border: "none", display: "flex", flexDirection: "column", alignItems: "center", gap: 3, color: isActive ? BRAND.header : "#8a8a8a", cursor: "pointer", fontFamily: "inherit" }}>
      <Icon size={22} />
      <span style={{ fontSize: 11, fontWeight: isActive ? 700 : 500 }}>{it.label}</span>
    </button>
  );
}

/* ---------------------------------------------------------
   Lock screen
--------------------------------------------------------- */
function LockScreen({ pin, onUnlock }) {
  const [val, setVal] = useState("");
  const [err, setErr] = useState(false);
  return (
    <div dir="rtl" style={{ minHeight: "100vh", background: BRAND.header, display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center", fontFamily: FONT, color: "#fff", gap: 18, maxWidth: 480, margin: "0 auto" }}>
      <Lock size={40} />
      <div style={{ fontWeight: 700, fontSize: 16 }}>حسابدار من قفل است</div>
      <input
        type="password" inputMode="numeric" maxLength={6} value={val}
        onChange={(e) => { setVal(e.target.value.replace(/[^0-9]/g, "")); setErr(false); }}
        placeholder="رمز عبور"
        style={{ width: 180, textAlign: "center", fontSize: 22, letterSpacing: 6, padding: "10px", borderRadius: 10, border: "none", outline: "none" }}
      />
      {err && <div style={{ color: "#ffb3c1", fontSize: 13 }}>رمز اشتباه است</div>}
      <button onClick={() => (val === pin ? onUnlock() : setErr(true))}
        style={{ background: BRAND.fab, color: "#fff", border: "none", borderRadius: 9, padding: "10px 30px", fontWeight: 700, cursor: "pointer" }}>
        باز کردن
      </button>
      <div style={{ fontSize: 11.5, color: "#d8c9e8", marginTop: 6, maxWidth: 260, textAlign: "center" }}>
        قفل اثر انگشت در این محیط پشتیبانی نمی‌شود؛ به‌جای آن از رمز عددی استفاده می‌شود.
      </div>
    </div>
  );
}

/* ===========================================================
   MAIN APP
=========================================================== */
export default function App() {
  const [loaded, setLoaded] = useState(false);
  const [settings, setSettings] = useState(seedSettings());
  const [unlocked, setUnlocked] = useState(false);

  const [accounts, setAccounts] = useState([]);
  const [categories, setCategories] = useState([]);
  const [transactions, setTransactions] = useState([]);
  const [budgets, setBudgets] = useState([]);
  const [loans, setLoans] = useState([]);
  const [checks, setChecks] = useState([]);
  const [bills, setBills] = useState([]);
  const [assets, setAssets] = useState([]);
  const [recurring, setRecurring] = useState([]);
  const [favorites, setFavorites] = useState({ categories: [], accounts: [] });
  const [members, setMembers] = useState([]);
  const [events, setEvents] = useState([]);
  const [projects, setProjects] = useState([]);
  const [fiscalPeriods, setFiscalPeriods] = useState([]);

  const [tab, setTab] = useState("home");
  const [subView, setSubView] = useState(null);
  const [showAdd, setShowAdd] = useState(false);
  const [prefillTx, setPrefillTx] = useState(null);
  const [open, setOpen] = useState({});
  const [year, setYear] = useState(jalaliYear(new Date()));
  const [txFilter, setTxFilter] = useState("all");
  const [txSearch, setTxSearch] = useState("");
  const [menuOpen, setMenuOpen] = useState(false);

  const shared = settings.sharedFamily;

  const reloadAll = useCallback(async (sh) => {
    const [a, c, t, b, ln, ck, bl, as, rc, fv, mb, ev, pj, fp] = await Promise.all([
      loadKey("hs:accounts", null, sh), loadKey("hs:categories", null, sh),
      loadKey("hs:transactions", null, sh), loadKey("hs:budgets", null, sh),
      loadKey("hs:loans", [], sh), loadKey("hs:checks", [], sh),
      loadKey("hs:bills", [], sh), loadKey("hs:assets", [], sh),
      loadKey("hs:recurring", [], sh), loadKey("hs:favorites", { categories: [], accounts: [] }, sh),
      loadKey("hs:members", [], sh), loadKey("hs:events", [], sh),
      loadKey("hs:projects", [], sh), loadKey("hs:fiscalPeriods", [], sh),
    ]);
    setAccounts(a || seedAccounts());
    setCategories(c || seedCategories());
    setTransactions(t || []);
    setBudgets(b || []);
    setLoans(ln); setChecks(ck); setBills(bl); setAssets(as); setRecurring(rc); setFavorites(fv);
    setMembers(mb); setEvents(ev); setProjects(pj); setFiscalPeriods(fp);
  }, []);

  useEffect(() => {
    (async () => {
      const s = await loadKey("hs:settings", seedSettings(), false);
      setSettings(s);
      await reloadAll(s.sharedFamily);
      setLoaded(true);
    })();
    // eslint-disable-next-line
  }, []);

  useEffect(() => { if (loaded) saveKey("hs:settings", settings, false); }, [settings, loaded]);
  useEffect(() => { if (loaded) saveKey("hs:accounts", accounts, shared); }, [accounts, loaded, shared]);
  useEffect(() => { if (loaded) saveKey("hs:categories", categories, shared); }, [categories, loaded, shared]);
  useEffect(() => { if (loaded) saveKey("hs:transactions", transactions, shared); }, [transactions, loaded, shared]);
  useEffect(() => { if (loaded) saveKey("hs:budgets", budgets, shared); }, [budgets, loaded, shared]);
  useEffect(() => { if (loaded) saveKey("hs:loans", loans, shared); }, [loans, loaded, shared]);
  useEffect(() => { if (loaded) saveKey("hs:checks", checks, shared); }, [checks, loaded, shared]);
  useEffect(() => { if (loaded) saveKey("hs:bills", bills, shared); }, [bills, loaded, shared]);
  useEffect(() => { if (loaded) saveKey("hs:assets", assets, shared); }, [assets, loaded, shared]);
  useEffect(() => { if (loaded) saveKey("hs:recurring", recurring, shared); }, [recurring, loaded, shared]);
  useEffect(() => { if (loaded) saveKey("hs:favorites", favorites, shared); }, [favorites, loaded, shared]);
  useEffect(() => { if (loaded) saveKey("hs:members", members, shared); }, [members, loaded, shared]);
  useEffect(() => { if (loaded) saveKey("hs:events", events, shared); }, [events, loaded, shared]);
  useEffect(() => { if (loaded) saveKey("hs:projects", projects, shared); }, [projects, loaded, shared]);
  useEffect(() => { if (loaded) saveKey("hs:fiscalPeriods", fiscalPeriods, shared); }, [fiscalPeriods, loaded, shared]);

  // process recurring templates once after load
  useEffect(() => {
    if (!loaded) return;
    let changed = false;
    const newTx = [];
    const updated = recurring.map((r) => {
      if (!r.active) return r;
      let next = r.nextDate; let guard = 0; let rr = { ...r };
      while (next <= todayISO() && guard < 24) {
        newTx.push({ id: uid(), type: rr.type, amount: rr.amount, categoryId: rr.categoryId, accountId: rr.accountId, date: next, note: rr.note || "تراکنش تکرارشونده", tags: ["تکرارشونده"], createdAt: new Date().toISOString(), recurringId: rr.id });
        next = rr.interval === "weekly" ? addDays(next, 7) : addMonths(next, 1);
        guard++; changed = true;
      }
      return { ...rr, nextDate: next };
    });
    if (changed) {
      setRecurring(updated);
      setTransactions((prev) => [...newTx, ...prev]);
    }
    // eslint-disable-next-line
  }, [loaded]);

  const toggle = (key) => setOpen((o) => ({ ...o, [key]: !o[key] }));
  const catById = useCallback((id) => categories.find((c) => c.id === id), [categories]);
  const accById = useCallback((id) => accounts.find((a) => a.id === id), [accounts]);

  const accountBalance = useCallback((accId) => {
    const acc = accById(accId); if (!acc) return 0;
    let bal = acc.initial || 0;
    transactions.forEach((t) => {
      if (t.type === "expense" && t.accountId === accId) bal -= t.amount;
      if (t.type === "income" && t.accountId === accId) bal += t.amount;
      if (t.type === "transfer") { if (t.accountId === accId) bal -= t.amount; if (t.toAccountId === accId) bal += t.amount; }
    });
    return bal;
  }, [accounts, transactions, accById]);

  const totalBalance = useMemo(() => accounts.reduce((s, a) => s + accountBalance(a.id), 0), [accounts, accountBalance]);
  const totalAssets = useMemo(() => assets.reduce((s, a) => s + (a.quantity * a.currentPrice || 0), 0), [assets]);

  const yearTx = useMemo(() => transactions.filter((t) => jalaliYear(new Date(t.date)) === year), [transactions, year]);
  const totalIncomeYear = useMemo(() => yearTx.filter((t) => t.type === "income").reduce((s, t) => s + t.amount, 0), [yearTx]);
  const totalExpenseYear = useMemo(() => yearTx.filter((t) => t.type === "expense").reduce((s, t) => s + t.amount, 0), [yearTx]);
  const gaugeMax = Math.max(totalIncomeYear, totalExpenseYear, 1);

  const expenseByCategory = useMemo(() => {
    const map = {};
    yearTx.filter((t) => t.type === "expense").forEach((t) => { map[t.categoryId] = (map[t.categoryId] || 0) + t.amount; });
    return Object.entries(map).map(([catId, amount]) => ({ catId, amount, name: catById(catId)?.name || "بدون دسته" })).sort((a, b) => b.amount - a.amount);
  }, [yearTx, catById]);
  const incomeByCategory = useMemo(() => {
    const map = {};
    yearTx.filter((t) => t.type === "income").forEach((t) => { map[t.categoryId] = (map[t.categoryId] || 0) + t.amount; });
    return Object.entries(map).map(([catId, amount]) => ({ catId, amount, name: catById(catId)?.name || "بدون دسته" })).sort((a, b) => b.amount - a.amount);
  }, [yearTx, catById]);

  function groupBy(idField, list) {
    const map = {};
    yearTx.filter((tx) => tx.type === "expense" && tx[idField]).forEach((tx) => { map[tx[idField]] = (map[tx[idField]] || 0) + tx.amount; });
    return Object.entries(map).map(([id, amount]) => ({ id, amount, name: list.find((x) => x.id === id)?.name || "—" })).sort((a, b) => b.amount - a.amount);
  }
  const expenseByMember = useMemo(() => groupBy("memberId", members), [yearTx, members]);
  const expenseByEvent = useMemo(() => groupBy("eventId", events), [yearTx, events]);
  const expenseByProject = useMemo(() => groupBy("projectId", projects), [yearTx, projects]);

  // net worth trend: last 8 months
  const netWorthTrend = useMemo(() => {
    const points = [];
    for (let i = 7; i >= 0; i--) {
      const d = new Date(); d.setMonth(d.getMonth() - i);
      const cutoff = d.toISOString().slice(0, 10);
      let bal = accounts.reduce((s, a) => s + (a.initial || 0), 0);
      transactions.forEach((t) => {
        if (t.date > cutoff) return;
        if (t.type === "expense") bal -= t.amount;
        if (t.type === "income") bal += t.amount;
      });
      points.push({ name: faMonthYear(d).split(" ")[0], مانده: bal });
    }
    return points;
  }, [accounts, transactions]);

  function addTransaction(tx) { setTransactions((p) => [{ ...tx, id: uid(), createdAt: new Date().toISOString() }, ...p]); }
  function deleteTransaction(id) { setTransactions((p) => p.filter((t) => t.id !== id)); }
  function addAccount(a) { setAccounts((p) => [...p, { ...a, id: uid() }]); }
  function deleteAccount(id) { setAccounts((p) => p.filter((a) => a.id !== id)); }
  function addCategory(c) { setCategories((p) => [...p, { ...c, id: uid() }]); }
  function deleteCategory(id) { setCategories((p) => p.filter((c) => c.id !== id)); }
  function upsertBudget(categoryId, amount) {
    setBudgets((prev) => prev.find((b) => b.categoryId === categoryId)
      ? prev.map((b) => (b.categoryId === categoryId ? { ...b, amount } : b))
      : [...prev, { id: uid(), categoryId, amount }]);
  }
  function toggleFavorite(kind, id) {
    setFavorites((prev) => {
      const list = prev[kind];
      return { ...prev, [kind]: list.includes(id) ? list.filter((x) => x !== id) : [...list, id] };
    });
  }

  const backupState = { accounts, categories, transactions, budgets, loans, checks, bills, assets, recurring, favorites, settings };
  function exportBackup() {
    const blob = new Blob([JSON.stringify(backupState, null, 2)], { type: "application/json" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url; a.download = `hesabdari-backup-${todayISO()}.json`;
    document.body.appendChild(a); a.click(); document.body.removeChild(a);
    URL.revokeObjectURL(url);
  }
  function importBackup(file) {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = JSON.parse(e.target.result);
        if (data.accounts) setAccounts(data.accounts);
        if (data.categories) setCategories(data.categories);
        if (data.transactions) setTransactions(data.transactions);
        if (data.budgets) setBudgets(data.budgets);
        if (data.loans) setLoans(data.loans);
        if (data.checks) setChecks(data.checks);
        if (data.bills) setBills(data.bills);
        if (data.assets) setAssets(data.assets);
        if (data.recurring) setRecurring(data.recurring);
        if (data.favorites) setFavorites(data.favorites);
        alert("بازیابی اطلاعات با موفقیت انجام شد");
      } catch { alert("فایل پشتیبان نامعتبر است"); }
    };
    reader.readAsText(file);
  }
  function exportExcel() {
    const rows = transactions.map((t) => ({
      نوع: t.type === "expense" ? "پرداخت" : t.type === "income" ? "دریافت" : "انتقال",
      مبلغ: t.amount, تاریخ: t.date,
      دسته: catById(t.categoryId)?.name || "", حساب: accById(t.accountId)?.name || "",
      برچسب: (t.tags || []).join("، "), یادداشت: t.note || "",
    }));
    const ws = XLSX.utils.json_to_sheet(rows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "تراکنش‌ها");
    XLSX.writeFile(wb, `transactions-${todayISO()}.xlsx`);
  }

  const t = THEME[settings.theme] || THEME.light;
  Object.assign(BRAND, COLOR_PRESETS[settings.themeColor] || COLOR_PRESETS.purple);

  if (!loaded) {
    return <div style={{ minHeight: "100vh", display: "flex", alignItems: "center", justifyContent: "center", background: THEME.light.bg, fontFamily: FONT }}>در حال بارگذاری...</div>;
  }
  if (settings.pin && !unlocked) {
    return <LockScreen pin={settings.pin} onUnlock={() => setUnlocked(true)} />;
  }

  function openWithPrefill(data) {
    setPrefillTx(data);
    setShowAdd(true);
  }

  const ctx = {
    accounts, addAccount, deleteAccount, accountBalance,
    categories, addCategory, deleteCategory,
    budgets, upsertBudget, expenseByCategory, incomeByCategory,
    loans, setLoans, checks, setChecks, bills, setBills, assets, setAssets,
    recurring, setRecurring, favorites, toggleFavorite,
    settings, setSettings, exportBackup, importBackup, exportExcel,
    transactions, catById, accById, totalBalance, totalAssets,
    members, setMembers, events, setEvents, projects, setProjects,
    fiscalPeriods, setFiscalPeriods, openWithPrefill,
  };

  return (
    <ThemeCtx.Provider value={t}>
      <div dir="rtl" style={{ fontFamily: FONT, background: t.bg, color: t.text, minHeight: "100vh", maxWidth: 480, margin: "0 auto", display: "flex", flexDirection: "column", position: "relative", boxShadow: "0 0 30px rgba(0,0,0,0.08)" }}>
        <style>{`
          @import url('https://fonts.googleapis.com/css2?family=Vazirmatn:wght@400;500;600;700;800&display=swap');
          * { box-sizing: border-box; }
          input, select, textarea, button { font-family: inherit; }
          ::-webkit-scrollbar { width: 0; height: 0; }
          @media print { .no-print { display: none !important; } }
        `}</style>

        {subView ? (
          <SubViewRouter subView={subView} onBack={() => setSubView(null)} ctx={ctx} />
        ) : (
          <>
            <Header title={settings.profile?.name || "alireza shadfar"} onMenu={() => setMenuOpen(true)} />
            <div style={{ flex: 1, overflowY: "auto", paddingBottom: 8 }}>
              {tab === "home" && (
                <HomeView
                  year={year} setYear={setYear}
                  totalIncomeYear={totalIncomeYear} totalExpenseYear={totalExpenseYear} gaugeMax={gaugeMax}
                  open={open} toggle={toggle}
                  accounts={accounts} accountBalance={accountBalance}
                  expenseByCategory={expenseByCategory} incomeByCategory={incomeByCategory}
                  budgets={budgets} categories={categories} transactions={yearTx}
                  bills={bills} loans={loans} checks={checks} assets={assets} totalAssets={totalAssets}
                  openAccounts={() => setSubView("accounts")} openBudgets={() => setSubView("budgets")}
                  openBills={() => setSubView("bills")} openLoans={() => setSubView("loans")}
                  openChecks={() => setSubView("checks")} openAssets={() => setSubView("assets")}
                />
              )}
              {tab === "transactions" && (
                <TransactionsView
                  transactions={transactions} catById={catById} accById={accById}
                  filter={txFilter} setFilter={setTxFilter} onDelete={deleteTransaction}
                  search={txSearch} setSearch={setTxSearch}
                />
              )}
              {tab === "operations" && <OperationsView setSubView={setSubView} onAdd={() => setShowAdd(true)} />}
              {tab === "reports" && (
                <ReportsView
                  expenseByCategory={expenseByCategory} incomeByCategory={incomeByCategory}
                  totalIncomeYear={totalIncomeYear} totalExpenseYear={totalExpenseYear}
                  accounts={accounts} accountBalance={accountBalance}
                  netWorthTrend={netWorthTrend} exportExcel={exportExcel}
                  expenseByMember={expenseByMember} expenseByEvent={expenseByEvent} expenseByProject={expenseByProject}
                />
              )}
            </div>
            <BottomNav active={tab} setActive={setTab} onAdd={() => { setPrefillTx(null); setShowAdd(true); }} />
          </>
        )}

        {showAdd && (
          <AddTransactionSheet
            accounts={accounts} categories={categories} favorites={favorites}
            members={members} events={events} projects={projects}
            initial={prefillTx}
            onClose={() => { setShowAdd(false); setPrefillTx(null); }}
            onSubmit={(tx) => { addTransaction(tx); setShowAdd(false); setPrefillTx(null); }}
          />
        )}
        {menuOpen && <SideMenu onClose={() => setMenuOpen(false)} setSubView={(v) => { setSubView(v); setMenuOpen(false); }} profileName={settings.profile?.name} />}
      </div>
    </ThemeCtx.Provider>
  );
}

/* ---------------------------------------------------------
   Home View
--------------------------------------------------------- */
function HomeView({
  year, setYear, totalIncomeYear, totalExpenseYear, gaugeMax, open, toggle,
  accounts, accountBalance, expenseByCategory, incomeByCategory, budgets, categories,
  transactions, bills, loans, checks, assets, totalAssets,
  openAccounts, openBudgets, openBills, openLoans, openChecks, openAssets
}) {
  const t = useT();
  const banks = accounts.filter((a) => a.type === "bank");
  const funds = accounts.filter((a) => a.type === "fund");
  const cards = accounts.filter((a) => a.type === "card");
  const net = totalIncomeYear - totalExpenseYear;
  const upcomingBills = bills.filter((b) => !b.paid && daysUntil(b.dueDate) <= 5).sort((a, b) => a.dueDate.localeCompare(b.dueDate));
  const upcomingChecks = checks.filter((c) => c.status === "pending" && daysUntil(c.dueDate) <= 7).sort((a, b) => a.dueDate.localeCompare(b.dueDate));

  return (
    <div style={{ padding: "18px 16px 8px" }}>
      <div style={{ display: "flex", alignItems: "center", justifyContent: "center", gap: 14, marginBottom: 18 }}>
        <button onClick={() => setYear((y) => y - 1)} style={navArrowStyle(t)}><ChevronLeft size={16} /></button>
        <div style={{ background: t.card, borderRadius: 20, padding: "6px 18px", fontWeight: 700, color: BRAND.header, boxShadow: "0 1px 3px rgba(0,0,0,0.1)" }}>سال {faDigits(year)}</div>
        <button onClick={() => setYear((y) => y + 1)} style={navArrowStyle(t)}><ChevronRight size={16} /></button>
      </div>

      <div style={{ display: "flex", justifyContent: "space-around", marginBottom: 8 }}>
        <GaugeCircle value={totalIncomeYear} max={gaugeMax} color={BRAND.darkgreen} label="درآمد" />
        <GaugeCircle value={totalExpenseYear} max={gaugeMax} color={BRAND.header} label="هزینه" />
      </div>
      <div style={{ textAlign: "center", marginBottom: 12, fontSize: 13, color: net >= 0 ? BRAND.darkgreen : BRAND.crimson, fontWeight: 700 }}>
        مانده سالانه: {toFaInt(Math.abs(net))} ریال {net >= 0 ? "مثبت" : "منفی"}
      </div>
      {totalAssets > 0 && (
        <div style={{ textAlign: "center", marginBottom: 18, fontSize: 12.5, color: BRAND.orange, fontWeight: 700 }}>
          ارزش دارایی‌های دیجیتال/بورس: {toFaInt(totalAssets)} ریال
        </div>
      )}

      {(upcomingBills.length > 0 || upcomingChecks.length > 0) && (
        <div style={{ background: "#fff6ea", border: "1px solid #f0d9a8", borderRadius: 12, padding: 12, marginBottom: 16, display: "flex", flexDirection: "column", gap: 6 }}>
          <div style={{ display: "flex", alignItems: "center", gap: 6, color: BRAND.orange, fontWeight: 700, fontSize: 13 }}><BellRing size={16} /> یادآوری‌های نزدیک</div>
          {upcomingBills.map((b) => (
            <div key={b.id} style={{ fontSize: 12.5, color: "#6b4c14" }}>قبض {b.title} — {daysUntil(b.dueDate) < 0 ? "سررسید گذشته" : `${toFaInt(daysUntil(b.dueDate))} روز مانده`}</div>
          ))}
          {upcomingChecks.map((c) => (
            <div key={c.id} style={{ fontSize: 12.5, color: "#6b4c14" }}>چک {c.type === "received" ? "دریافتی" : "پرداختی"} {c.payee} — {daysUntil(c.dueDate) < 0 ? "سررسید گذشته" : `${toFaInt(daysUntil(c.dueDate))} روز مانده`}</div>
          ))}
        </div>
      )}

      <CollapsibleSection color={BRAND.mauve} title="میانبر تراکنش ها" open={!!open.shortcut} onToggle={() => toggle("shortcut")}>
        {transactions.slice(0, 5).length === 0 && <EmptyRow text="هنوز تراکنشی ثبت نشده" />}
        {transactions.slice(0, 5).map((tx) => (
          <Row key={tx.id} title={categories.find((c) => c.id === tx.categoryId)?.name || (tx.type === "transfer" ? "انتقال وجه" : "—")}
            subtitle={faLongDate(new Date(tx.date))} value={`${toFaInt(tx.amount)} ریال`}
            valueColor={tx.type === "expense" ? BRAND.crimson : tx.type === "income" ? BRAND.darkgreen : BRAND.violet} />
        ))}
      </CollapsibleSection>

      <CollapsibleSection color={BRAND.header} title="هزینه ها" open={!!open.expense} onToggle={() => toggle("expense")}>
        {expenseByCategory.length === 0 && <EmptyRow text="هزینه‌ای ثبت نشده" />}
        {expenseByCategory.map((e) => <Row key={e.catId} title={e.name} value={`${toFaInt(e.amount)} ریال`} valueColor={BRAND.crimson} />)}
      </CollapsibleSection>

      <CollapsibleSection color={BRAND.darkgreen} title="درآمدها" open={!!open.income} onToggle={() => toggle("income")}>
        {incomeByCategory.length === 0 && <EmptyRow text="درآمدی ثبت نشده" />}
        {incomeByCategory.map((e) => <Row key={e.catId} title={e.name} value={`${toFaInt(e.amount)} ریال`} valueColor={BRAND.darkgreen} />)}
      </CollapsibleSection>

      <CollapsibleSection color={BRAND.violet} title="بانک ها و کارت ها" open={!!open.banks} onToggle={() => toggle("banks")}>
        {[...banks, ...cards].length === 0 && <EmptyRow text="حسابی ثبت نشده" />}
        {[...banks, ...cards].map((a) => (
          <Row key={a.id} title={a.name} subtitle={a.type === "card" ? "کارت" : "بانک"} value={`${toFaInt(accountBalance(a.id))} ریال`}
            valueColor={accountBalance(a.id) >= 0 ? t.text : BRAND.crimson} />
        ))}
        <AddLink text="+ مدیریت حساب‌ها و کارت‌ها" onClick={openAccounts} />
      </CollapsibleSection>

      <CollapsibleSection color={BRAND.teal} title="صندوق ها" open={!!open.funds} onToggle={() => toggle("funds")}>
        {funds.length === 0 && <EmptyRow text="صندوقی ثبت نشده" />}
        {funds.map((a) => <Row key={a.id} title={a.name} value={`${toFaInt(accountBalance(a.id))} ریال`} valueColor={accountBalance(a.id) >= 0 ? t.text : BRAND.crimson} />)}
      </CollapsibleSection>

      <CollapsibleSection color={BRAND.gold} title="گزارش مانده حساب ها" open={!!open.balrep} onToggle={() => toggle("balrep")}>
        {accounts.map((a) => (
          <Row key={a.id} title={a.name} subtitle={a.type === "bank" ? "بانک" : a.type === "card" ? "کارت" : "صندوق"}
            value={`${toFaInt(accountBalance(a.id))} ریال`} valueColor={accountBalance(a.id) >= 0 ? BRAND.darkgreen : BRAND.crimson} />
        ))}
      </CollapsibleSection>

      <CollapsibleSection color={BRAND.green} title="بودجه بندی" open={!!open.budget} onToggle={() => toggle("budget")}>
        {budgets.length === 0 && <EmptyRow text="بودجه‌ای تعریف نشده" />}
        {budgets.map((b) => {
          const cat = categories.find((c) => c.id === b.categoryId);
          const spent = expenseByCategory.find((e) => e.catId === b.categoryId)?.amount || 0;
          const pct = Math.min(100, Math.round((spent / (b.amount || 1)) * 100));
          return (
            <div key={b.id} style={{ padding: "10px 4px", borderBottom: `1px solid ${t.border}` }}>
              <div style={{ display: "flex", justifyContent: "space-between", fontSize: 13.5, fontWeight: 600, marginBottom: 6 }}>
                <span>{cat?.name || "—"}</span>
                <span style={{ color: pct >= 100 ? BRAND.crimson : t.sub }}>{toFaInt(pct)}٪ — {toFaInt(spent)}/{toFaInt(b.amount)}</span>
              </div>
              <div style={{ height: 7, background: t.border, borderRadius: 4, overflow: "hidden" }}>
                <div style={{ width: `${pct}%`, height: "100%", background: pct >= 100 ? BRAND.crimson : BRAND.green }} />
              </div>
            </div>
          );
        })}
        <AddLink text="+ مدیریت بودجه‌بندی" onClick={openBudgets} />
      </CollapsibleSection>

      <CollapsibleSection color={BRAND.orange} title="وام ها و چک ها" open={!!open.loanchk} onToggle={() => toggle("loanchk")}>
        {loans.length === 0 && checks.length === 0 && <EmptyRow text="موردی ثبت نشده" />}
        {loans.map((l) => (
          <Row key={l.id} title={l.title} subtitle="وام" value={`${toFaInt(l.principal - (l.paidCount || 0) * l.monthlyPayment)} ریال باقی‌مانده`} valueColor={BRAND.crimson} />
        ))}
        {checks.filter((c) => c.status === "pending").map((c) => (
          <Row key={c.id} title={`${c.payee} (${c.type === "received" ? "دریافتی" : "پرداختی"})`} subtitle={faLongDate(new Date(c.dueDate))} value={`${toFaInt(c.amount)} ریال`} />
        ))}
        <div style={{ display: "flex", gap: 14, justifyContent: "center", marginTop: 4 }}>
          <AddLink text="+ وام‌ها" onClick={openLoans} />
          <AddLink text="+ چک‌ها" onClick={openChecks} />
          <AddLink text="+ دارایی‌ها" onClick={openAssets} />
        </div>
      </CollapsibleSection>

      <CollapsibleSection color={BRAND.crimson} title="یادآوری قبض ها" open={!!open.bills} onToggle={() => toggle("bills")}>
        {bills.length === 0 && <EmptyRow text="قبضی ثبت نشده" />}
        {bills.map((b) => (
          <Row key={b.id} title={b.title} subtitle={faLongDate(new Date(b.dueDate))}
            value={b.paid ? "پرداخت شده" : `${toFaInt(daysUntil(b.dueDate))} روز`}
            valueColor={b.paid ? BRAND.darkgreen : daysUntil(b.dueDate) < 0 ? BRAND.crimson : BRAND.orange} />
        ))}
        <AddLink text="+ مدیریت قبض‌ها" onClick={openBills} />
      </CollapsibleSection>
    </div>
  );
}
const navArrowStyle = (t) => ({ width: 30, height: 30, borderRadius: "50%", border: "none", background: t.card, color: BRAND.header, display: "flex", alignItems: "center", justifyContent: "center", boxShadow: "0 1px 3px rgba(0,0,0,0.1)", cursor: "pointer" });

/* ---------------------------------------------------------
   Transactions View
--------------------------------------------------------- */
function TransactionsView({ transactions, catById, accById, filter, setFilter, onDelete, search, setSearch }) {
  const t = useT();
  const [preview, setPreview] = useState(null);
  const filtered = transactions.filter((tx) => {
    if (filter !== "all" && tx.type !== filter) return false;
    if (search.trim()) {
      const q = search.trim();
      const hay = `${catById(tx.categoryId)?.name || ""} ${tx.note || ""} ${(tx.tags || []).join(" ")} ${tx.amount}`;
      if (!hay.includes(q)) return false;
    }
    return true;
  });
  const grouped = useMemo(() => {
    const map = {};
    filtered.forEach((tx) => { (map[tx.date] = map[tx.date] || []).push(tx); });
    return Object.entries(map).sort((a, b) => new Date(b[0]) - new Date(a[0]));
  }, [filtered]);
  const tabs = [{ key: "all", label: "همه" }, { key: "expense", label: "پرداخت ها" }, { key: "income", label: "دریافت ها" }, { key: "transfer", label: "انتقال ها" }];

  return (
    <div>
      <div style={{ padding: "10px 14px 0" }}>
        <div style={{ position: "relative" }}>
          <Search size={16} color={t.sub} style={{ position: "absolute", top: 12, right: 12 }} />
          <input value={search} onChange={(e) => setSearch(e.target.value)} placeholder="جستجو در تراکنش‌ها، برچسب یا یادداشت..."
            style={{ width: "100%", padding: "10px 36px 10px 12px", borderRadius: 10, border: `1.5px solid ${t.inputBorder}`, background: t.input, color: t.text, fontSize: 13.5, outline: "none", marginBottom: 10 }} />
        </div>
      </div>
      <div style={{ display: "flex", background: BRAND.header, padding: "0 8px 12px", gap: 4 }}>
        {tabs.map((tItem) => (
          <button key={tItem.key} onClick={() => setFilter(tItem.key)} style={{ flex: 1, padding: "8px 4px", border: "none", borderRadius: 8, cursor: "pointer", background: filter === tItem.key ? "rgba(255,255,255,0.18)" : "transparent", color: "#fff", fontWeight: filter === tItem.key ? 700 : 500, fontSize: 12.5 }}>{tItem.label}</button>
        ))}
      </div>
      <div style={{ padding: "12px 14px" }}>
        {grouped.length === 0 && <EmptyRow text="تراکنشی یافت نشد" />}
        {grouped.map(([date, txs]) => {
          const dayTotal = txs.reduce((s, tx) => s + (tx.type === "expense" ? -tx.amount : tx.type === "income" ? tx.amount : 0), 0);
          return (
            <div key={date} style={{ marginBottom: 16 }}>
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 8 }}>
                <span style={{ fontSize: 12.5, color: t.sub, fontWeight: 600 }}>{faLongDate(new Date(date))}</span>
                <span style={{ fontSize: 12.5, fontWeight: 700, color: dayTotal >= 0 ? BRAND.darkgreen : BRAND.crimson }}>{dayTotal >= 0 ? "+" : "-"}{toFaInt(Math.abs(dayTotal))} ریال</span>
              </div>
              {txs.map((tx) => {
                const cat = catById(tx.categoryId), acc = accById(tx.accountId), toAcc = accById(tx.toAccountId);
                const color = tx.type === "expense" ? BRAND.crimson : tx.type === "income" ? BRAND.darkgreen : BRAND.violet;
                const Icon = tx.type === "expense" ? TrendingDown : tx.type === "income" ? TrendingUp : ArrowLeftRight;
                return (
                  <div key={tx.id} style={{ background: t.card, borderRadius: 12, padding: 14, marginBottom: 10, boxShadow: "0 1px 3px rgba(0,0,0,0.08)" }}>
                    <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start" }}>
                      <div style={{ display: "flex", gap: 10 }}>
                        <span style={{ width: 38, height: 38, borderRadius: 10, background: color + "22", color, display: "flex", alignItems: "center", justifyContent: "center", flexShrink: 0 }}><Icon size={18} /></span>
                        <div>
                          <div style={{ fontWeight: 700, fontSize: 14 }}>{tx.type === "transfer" ? `انتقال به ${toAcc?.name || "—"}` : (cat?.name || "بدون دسته")}</div>
                          <div style={{ fontSize: 12, color: t.sub, marginTop: 3 }}>{tx.type === "expense" ? "از حساب" : tx.type === "income" ? "به حساب" : "از حساب"}: {acc?.name || "—"}</div>
                          {tx.note && <div style={{ fontSize: 12, color: t.sub, marginTop: 2 }}>{tx.note}</div>}
                          {(tx.tags || []).length > 0 && (
                            <div style={{ display: "flex", gap: 5, marginTop: 5, flexWrap: "wrap" }}>
                              {tx.tags.map((tag, i) => <span key={i} style={{ fontSize: 10.5, background: t.border, color: t.sub, padding: "2px 8px", borderRadius: 20 }}>{tag}</span>)}
                            </div>
                          )}
                          {tx.photo && <img src={tx.photo} onClick={() => setPreview(tx.photo)} style={{ width: 44, height: 44, borderRadius: 8, objectFit: "cover", marginTop: 6, cursor: "pointer" }} />}
                        </div>
                      </div>
                      <div style={{ textAlign: "left" }}>
                        <div style={{ fontWeight: 800, color, fontSize: 14.5 }}>{toFaInt(tx.amount)}</div>
                        <button onClick={() => onDelete(tx.id)} style={{ background: "none", border: "none", color: t.sub, cursor: "pointer", marginTop: 6 }}><Trash2 size={15} /></button>
                      </div>
                    </div>
                  </div>
                );
              })}
            </div>
          );
        })}
      </div>
      {preview && (
        <div onClick={() => setPreview(null)} style={{ position: "absolute", inset: 0, background: "rgba(0,0,0,0.7)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 70, maxWidth: 480, margin: "0 auto" }}>
          <img src={preview} style={{ maxWidth: "88%", maxHeight: "70%", borderRadius: 10 }} />
        </div>
      )}
    </div>
  );
}

/* ---------------------------------------------------------
   Operations View
--------------------------------------------------------- */
function OperationsView({ setSubView, onAdd }) {
  const st = useStyles();
  const items = [
    { title: "حساب‌ها و کارت‌ها", icon: <Landmark size={17} />, color: BRAND.violet, key: "accounts" },
    { title: "دسته‌بندی‌ها و برچسب‌ها", icon: <Tag size={17} />, color: BRAND.mauve, key: "categories" },
    { title: "بودجه‌بندی", icon: <Save size={17} />, color: BRAND.green, key: "budgets" },
    { title: "تراکنش‌های تکرارشونده", icon: <Repeat size={17} />, color: BRAND.teal, key: "recurring" },
    { title: "چک‌ها", icon: <FileSpreadsheet size={17} />, color: BRAND.gold, key: "checks" },
    { title: "وام و اقساط", icon: <Bank size={17} />, color: BRAND.crimson, key: "loans" },
    { title: "یادآوری قبض‌ها", icon: <BellRing size={17} />, color: BRAND.orange, key: "bills" },
    { title: "دارایی‌ها (ارز دیجیتال / بورس)", icon: <Bitcoin size={17} />, color: "#7a5cff", key: "assets" },
    { title: "اعضای منزل، رویداد و پروژه", icon: <Users size={17} />, color: BRAND.violet, key: "tags" },
    { title: "دوره مالی", icon: <CalendarDays size={17} />, color: BRAND.darkgreen, key: "periods" },
    { title: "پیامک بانکی (افزودن نیمه‌خودکار)", icon: <BellRing size={17} />, color: "#666", key: "sms" },
    { title: "تقویم شمسی", icon: <CalendarDays size={17} />, color: BRAND.header, key: "calendar" },
    { title: "اطلاعات کاربری", icon: <Users size={17} />, color: "#444", key: "profile" },
    { title: "تنظیمات و پشتیبان‌گیری", icon: <ShieldCheck size={17} />, color: "#555", key: "settings" },
  ];
  return (
    <div style={{ padding: "10px 16px" }}>
      <div style={{ ...st.card, padding: "4px 12px" }}>
        {items.map((it) => <Row key={it.key} title={it.title} leftIcon={it.icon} leftColor={it.color} onClick={() => setSubView(it.key)} />)}
        <Row title="ثبت تراکنش جدید" leftIcon={<Plus size={17} />} leftColor={BRAND.fab} onClick={onAdd} />
      </div>
    </div>
  );
}

/* ---------------------------------------------------------
   Reports View
--------------------------------------------------------- */
const PIE_COLORS = ["#B01E4A", "#6C3FA0", "#4E9AA0", "#A98A3B", "#3E1461", "#1E8449", "#A65475", "#555"];
function ReportsView({ expenseByCategory, incomeByCategory, totalIncomeYear, totalExpenseYear, accounts, accountBalance, netWorthTrend, exportExcel, expenseByMember, expenseByEvent, expenseByProject }) {
  const st = useStyles();
  const barData = [{ name: "درآمد", مقدار: totalIncomeYear, fill: BRAND.darkgreen }, { name: "هزینه", مقدار: totalExpenseYear, fill: BRAND.crimson }];
  const profit = totalIncomeYear - totalExpenseYear;
  return (
    <div style={{ padding: "16px" }}>
      <div className="no-print" style={{ display: "flex", gap: 8, marginBottom: 16 }}>
        <button onClick={exportExcel} style={{ ...st.primaryBtn, background: BRAND.green, display: "flex", alignItems: "center", justifyContent: "center", gap: 6 }}><FileSpreadsheet size={16} /> خروجی Excel</button>
        <button onClick={() => window.print()} style={{ ...st.primaryBtn, background: "#555", display: "flex", alignItems: "center", justifyContent: "center", gap: 6 }}><Printer size={16} /> چاپ / PDF</button>
      </div>

      <SectionTitle text="سود و زیان سالانه" />
      <div style={{ ...st.card, padding: 14, marginBottom: 18 }}>
        <Row title="مجموع درآمد" value={`${toFaInt(totalIncomeYear)} ریال`} valueColor={BRAND.darkgreen} chevron={null} />
        <Row title="مجموع هزینه" value={`${toFaInt(totalExpenseYear)} ریال`} valueColor={BRAND.crimson} chevron={null} />
        <Row title="سود / زیان خالص" value={`${toFaInt(Math.abs(profit))} ریال`} valueColor={profit >= 0 ? BRAND.darkgreen : BRAND.crimson} chevron={null} />
      </div>

      <SectionTitle text="روند دارایی خالص (۸ ماه اخیر)" />
      <div style={{ ...st.card, padding: 12, marginBottom: 18 }}>
        <ResponsiveContainer width="100%" height={180}>
          <LineChart data={netWorthTrend}>
            <CartesianGrid strokeDasharray="3 3" />
            <XAxis dataKey="name" tick={{ fontFamily: FONT, fontSize: 11 }} />
            <YAxis tick={{ fontFamily: FONT, fontSize: 10 }} width={44} />
            <Tooltip formatter={(v) => toFaInt(v)} contentStyle={{ fontFamily: FONT, direction: "rtl" }} />
            <Line type="monotone" dataKey="مانده" stroke={BRAND.violet} strokeWidth={2.5} dot={{ r: 3 }} />
          </LineChart>
        </ResponsiveContainer>
      </div>

      <SectionTitle text="درآمد و هزینه" />
      <div style={{ ...st.card, padding: 12, marginBottom: 18 }}>
        <ResponsiveContainer width="100%" height={180}>
          <BarChart data={barData}>
            <XAxis dataKey="name" tick={{ fontFamily: FONT, fontSize: 12 }} />
            <YAxis tick={{ fontFamily: FONT, fontSize: 10 }} width={44} />
            <Tooltip formatter={(v) => toFaInt(v)} contentStyle={{ fontFamily: FONT, direction: "rtl" }} />
            <Bar dataKey="مقدار" radius={[6, 6, 0, 0]} />
          </BarChart>
        </ResponsiveContainer>
      </div>

      <SectionTitle text="توزیع هزینه‌ها بر اساس دسته" />
      <div style={{ ...st.card, padding: 12, marginBottom: 18 }}>
        {expenseByCategory.length === 0 ? <EmptyRow text="داده‌ای برای نمایش نیست" /> : (
          <ResponsiveContainer width="100%" height={220}>
            <PieChart>
              <Pie data={expenseByCategory} dataKey="amount" nameKey="name" outerRadius={80} label={(e) => e.name}>
                {expenseByCategory.map((_, i) => <Cell key={i} fill={PIE_COLORS[i % PIE_COLORS.length]} />)}
              </Pie>
              <Tooltip formatter={(v) => toFaInt(v)} contentStyle={{ fontFamily: FONT, direction: "rtl" }} />
            </PieChart>
          </ResponsiveContainer>
        )}
      </div>

      <SectionTitle text="گزارش مانده حساب‌ها" />
      <div style={{ ...st.card, padding: "4px 12px", marginBottom: expenseByMember?.length || expenseByEvent?.length || expenseByProject?.length ? 18 : 0 }}>
        {accounts.map((a) => <Row key={a.id} title={a.name} subtitle={a.type === "bank" ? "بانک" : a.type === "card" ? "کارت" : "صندوق"} value={`${toFaInt(accountBalance(a.id))} ریال`} valueColor={accountBalance(a.id) >= 0 ? BRAND.darkgreen : BRAND.crimson} chevron={null} />)}
      </div>

      {expenseByMember?.length > 0 && (
        <>
          <SectionTitle text="گزارش هزینه به‌تفکیک اعضای خانواده" />
          <div style={{ ...st.card, padding: "4px 12px", marginBottom: 18 }}>
            {expenseByMember.map((m) => <Row key={m.id} title={m.name} value={`${toFaInt(m.amount)} ریال`} valueColor={BRAND.crimson} chevron={null} />)}
          </div>
        </>
      )}
      {expenseByEvent?.length > 0 && (
        <>
          <SectionTitle text="گزارش هزینه به‌تفکیک رویداد" />
          <div style={{ ...st.card, padding: "4px 12px", marginBottom: 18 }}>
            {expenseByEvent.map((m) => <Row key={m.id} title={m.name} value={`${toFaInt(m.amount)} ریال`} valueColor={BRAND.crimson} chevron={null} />)}
          </div>
        </>
      )}
      {expenseByProject?.length > 0 && (
        <>
          <SectionTitle text="گزارش هزینه به‌تفکیک پروژه" />
          <div style={{ ...st.card, padding: "4px 12px" }}>
            {expenseByProject.map((m) => <Row key={m.id} title={m.name} value={`${toFaInt(m.amount)} ریال`} valueColor={BRAND.crimson} chevron={null} />)}
          </div>
        </>
      )}
    </div>
  );
}

/* ---------------------------------------------------------
   SubView Router
--------------------------------------------------------- */
function SubViewRouter({ subView, onBack, ctx }) {
  const titles = { accounts: "حساب‌ها و کارت‌ها", categories: "دسته‌بندی‌ها و برچسب‌ها", budgets: "بودجه‌بندی", recurring: "تراکنش‌های تکرارشونده", checks: "چک‌ها", loans: "وام و اقساط", bills: "یادآوری قبض‌ها", assets: "دارایی‌ها", calendar: "تقویم شمسی", settings: "تنظیمات", profile: "اطلاعات کاربری", tags: "اعضا، رویداد و پروژه", periods: "دوره مالی", sms: "پیامک بانکی" };
  return (
    <>
      <Header title={titles[subView]} back onBack={onBack} />
      <div style={{ flex: 1, overflowY: "auto", padding: "16px", paddingBottom: 40 }}>
        {subView === "accounts" && <AccountsManager {...ctx} />}
        {subView === "categories" && <CategoriesManager {...ctx} />}
        {subView === "budgets" && <BudgetsManager {...ctx} />}
        {subView === "recurring" && <RecurringManager {...ctx} />}
        {subView === "checks" && <ChecksManager {...ctx} />}
        {subView === "loans" && <LoansManager {...ctx} />}
        {subView === "bills" && <BillsManager {...ctx} />}
        {subView === "assets" && <AssetsManager {...ctx} />}
        {subView === "calendar" && <CalendarViewSub {...ctx} />}
        {subView === "settings" && <SettingsView {...ctx} />}
        {subView === "profile" && <SettingsView {...ctx} />}
        {subView === "tags" && <MembersEventsProjectsManager {...ctx} />}
        {subView === "periods" && <FiscalPeriodsManager {...ctx} />}
        {subView === "sms" && <BankSmsManager {...ctx} onBack={onBack} />}
      </div>
    </>
  );
}

/* ---------------------------------------------------------
   Accounts / Categories / Budgets Managers
--------------------------------------------------------- */
function AccountsManager({ accounts, addAccount, deleteAccount, accountBalance, favorites, toggleFavorite }) {
  const st = useStyles();
  const [name, setName] = useState(""); const [type, setType] = useState("bank"); const [initial, setInitial] = useState("");
  return (
    <div>
      <div style={{ ...st.card, padding: 14, marginBottom: 16 }}>
        <div style={{ fontWeight: 700, marginBottom: 10, fontSize: 14 }}>افزودن حساب / کارت جدید</div>
        <div style={{ display: "flex", gap: 8, marginBottom: 10 }}>
          <button onClick={() => setType("bank")} style={pillStyle(type === "bank")}>بانک</button>
          <button onClick={() => setType("fund")} style={pillStyle(type === "fund")}>صندوق</button>
          <button onClick={() => setType("card")} style={pillStyle(type === "card")}>کارت</button>
        </div>
        <input placeholder="نام حساب" value={name} onChange={(e) => setName(e.target.value)} style={st.input} />
        <input placeholder="موجودی اولیه (ریال)" value={initial} onChange={(e) => setInitial(e.target.value.replace(/[^0-9]/g, ""))} style={st.input} inputMode="numeric" />
        <button onClick={() => { if (!name.trim()) return; addAccount({ name: name.trim(), type, initial: Number(initial || 0) }); setName(""); setInitial(""); }} style={st.primaryBtn}>افزودن</button>
      </div>
      <div style={{ ...st.card, padding: "4px 12px" }}>
        {accounts.length === 0 && <EmptyRow text="حسابی ثبت نشده" />}
        {accounts.map((a) => (
          <Row key={a.id} title={a.name} subtitle={`${a.type === "bank" ? "بانک" : a.type === "card" ? "کارت" : "صندوق"} · موجودی: ${toFaInt(accountBalance(a.id))} ریال`}
            extra={<button onClick={() => toggleFavorite("accounts", a.id)} style={{ background: "none", border: "none", cursor: "pointer" }}><Star size={16} fill={favorites.accounts.includes(a.id) ? "#f5b301" : "none"} color="#f5b301" /></button>}
            leftIcon={<Trash2 size={15} />} leftColor={BRAND.crimson} onClick={() => deleteAccount(a.id)} chevron={null} />
        ))}
      </div>
    </div>
  );
}

function CategoriesManager({ categories, addCategory, deleteCategory, favorites, toggleFavorite }) {
  const st = useStyles();
  const [name, setName] = useState(""); const [kind, setKind] = useState("expense");
  const expense = categories.filter((c) => c.kind === "expense"); const income = categories.filter((c) => c.kind === "income");
  return (
    <div>
      <div style={{ ...st.card, padding: 14, marginBottom: 16 }}>
        <div style={{ fontWeight: 700, marginBottom: 10, fontSize: 14 }}>افزودن دسته‌بندی</div>
        <div style={{ display: "flex", gap: 8, marginBottom: 10 }}>
          <button onClick={() => setKind("expense")} style={pillStyle(kind === "expense")}>هزینه</button>
          <button onClick={() => setKind("income")} style={pillStyle(kind === "income")}>درآمد</button>
        </div>
        <input placeholder="نام دسته" value={name} onChange={(e) => setName(e.target.value)} style={st.input} />
        <button onClick={() => { if (!name.trim()) return; addCategory({ name: name.trim(), kind }); setName(""); }} style={st.primaryBtn}>افزودن</button>
      </div>
      <SectionTitle text="دسته‌های هزینه" />
      <div style={{ ...st.card, padding: "4px 12px", marginBottom: 16 }}>
        {expense.length === 0 && <EmptyRow text="دسته‌ای ثبت نشده" />}
        {expense.map((c) => <CatRow key={c.id} c={c} onDelete={deleteCategory} fav={favorites.categories.includes(c.id)} onFav={() => toggleFavorite("categories", c.id)} />)}
      </div>
      <SectionTitle text="دسته‌های درآمد" />
      <div style={{ ...st.card, padding: "4px 12px" }}>
        {income.length === 0 && <EmptyRow text="دسته‌ای ثبت نشده" />}
        {income.map((c) => <CatRow key={c.id} c={c} onDelete={deleteCategory} fav={favorites.categories.includes(c.id)} onFav={() => toggleFavorite("categories", c.id)} />)}
      </div>
    </div>
  );
}
function CatRow({ c, onDelete, fav, onFav }) {
  const t = useT();
  return (
    <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", padding: "10px 4px", borderBottom: `1px solid ${t.border}` }}>
      <span style={{ fontSize: 14, fontWeight: 600 }}>{c.name}</span>
      <div style={{ display: "flex", gap: 10 }}>
        <button onClick={onFav} style={{ background: "none", border: "none", cursor: "pointer" }}><Star size={16} fill={fav ? "#f5b301" : "none"} color="#f5b301" /></button>
        <button onClick={() => onDelete(c.id)} style={{ background: "none", border: "none", color: BRAND.crimson, cursor: "pointer" }}><Trash2 size={16} /></button>
      </div>
    </div>
  );
}

function BudgetsManager({ categories, budgets, upsertBudget }) {
  const st = useStyles();
  const expenseCats = categories.filter((c) => c.kind === "expense");
  return (
    <div style={{ ...st.card, padding: "8px 12px" }}>
      {expenseCats.length === 0 && <EmptyRow text="ابتدا یک دسته هزینه بسازید" />}
      {expenseCats.map((c) => {
        const b = budgets.find((bb) => bb.categoryId === c.id);
        return (
          <div key={c.id} style={{ display: "flex", alignItems: "center", justifyContent: "space-between", padding: "12px 4px", borderBottom: "1px solid #f0eef3", gap: 10 }}>
            <span style={{ fontSize: 13.5, fontWeight: 600, flexShrink: 0 }}>{c.name}</span>
            <input placeholder="بودجه ماهانه" defaultValue={b ? b.amount : ""} onBlur={(e) => upsertBudget(c.id, Number(e.target.value.replace(/[^0-9]/g, "") || 0))} inputMode="numeric" style={{ ...st.input, margin: 0, width: 150, textAlign: "left" }} />
          </div>
        );
      })}
    </div>
  );
}

/* ---------------------------------------------------------
   Recurring Manager
--------------------------------------------------------- */
function RecurringManager({ recurring, setRecurring, categories, accounts }) {
  const st = useStyles();
  const [form, setForm] = useState({ type: "expense", amount: "", categoryId: "", accountId: accounts[0]?.id || "", interval: "monthly", startDate: todayISO(), note: "" });
  function add() {
    if (!form.amount || !form.categoryId || !form.accountId) return;
    setRecurring((p) => [...p, { id: uid(), ...form, amount: Number(form.amount), nextDate: form.startDate, active: true }]);
    setForm({ ...form, amount: "", note: "" });
  }
  return (
    <div>
      <div style={{ ...st.card, padding: 14, marginBottom: 16 }}>
        <div style={{ fontWeight: 700, marginBottom: 10, fontSize: 14 }}>افزودن تراکنش تکرارشونده (مثل حقوق یا اجاره)</div>
        <div style={{ display: "flex", gap: 8, marginBottom: 10 }}>
          <button onClick={() => setForm({ ...form, type: "expense", categoryId: "" })} style={pillStyle(form.type === "expense")}>هزینه</button>
          <button onClick={() => setForm({ ...form, type: "income", categoryId: "" })} style={pillStyle(form.type === "income")}>درآمد</button>
        </div>
        <input placeholder="مبلغ" value={form.amount} onChange={(e) => setForm({ ...form, amount: e.target.value.replace(/[^0-9]/g, "") })} style={st.input} inputMode="numeric" />
        <select value={form.categoryId} onChange={(e) => setForm({ ...form, categoryId: e.target.value })} style={st.input}>
          <option value="">دسته را انتخاب کنید</option>
          {categories.filter((c) => c.kind === form.type).map((c) => <option key={c.id} value={c.id}>{c.name}</option>)}
        </select>
        <select value={form.accountId} onChange={(e) => setForm({ ...form, accountId: e.target.value })} style={st.input}>
          {accounts.map((a) => <option key={a.id} value={a.id}>{a.name}</option>)}
        </select>
        <div style={{ display: "flex", gap: 8, marginBottom: 10 }}>
          <button onClick={() => setForm({ ...form, interval: "monthly" })} style={pillStyle(form.interval === "monthly")}>ماهانه</button>
          <button onClick={() => setForm({ ...form, interval: "weekly" })} style={pillStyle(form.interval === "weekly")}>هفتگی</button>
        </div>
        <label style={st.label}>تاریخ شروع</label>
        <input type="date" value={form.startDate} onChange={(e) => setForm({ ...form, startDate: e.target.value })} style={st.input} />
        <input placeholder="یادداشت (اختیاری)" value={form.note} onChange={(e) => setForm({ ...form, note: e.target.value })} style={st.input} />
        <button onClick={add} style={st.primaryBtn}>افزودن</button>
      </div>
      <div style={{ ...st.card, padding: "4px 12px" }}>
        {recurring.length === 0 && <EmptyRow text="موردی ثبت نشده" />}
        {recurring.map((r) => (
          <Row key={r.id} title={categories.find((c) => c.id === r.categoryId)?.name || "—"}
            subtitle={`${r.interval === "monthly" ? "ماهانه" : "هفتگی"} · تراکنش بعدی: ${faLongDate(new Date(r.nextDate))}`}
            value={`${toFaInt(r.amount)} ریال`}
            extra={<button onClick={() => setRecurring((p) => p.map((x) => x.id === r.id ? { ...x, active: !x.active } : x))} style={{ background: "none", border: "none", cursor: "pointer", color: r.active ? BRAND.green : "#aaa", fontSize: 11, fontWeight: 700 }}>{r.active ? "فعال" : "غیرفعال"}</button>}
            leftIcon={<Trash2 size={15} />} leftColor={BRAND.crimson} onClick={() => setRecurring((p) => p.filter((x) => x.id !== r.id))} chevron={null} />
        ))}
      </div>
    </div>
  );
}

/* ---------------------------------------------------------
   Checks Manager
--------------------------------------------------------- */
function ChecksManager({ checks, setChecks }) {
  const st = useStyles();
  const [form, setForm] = useState({ type: "received", payee: "", amount: "", dueDate: todayISO(), note: "" });
  function add() {
    if (!form.payee || !form.amount) return;
    setChecks((p) => [{ id: uid(), ...form, amount: Number(form.amount), status: "pending" }, ...p]);
    setForm({ ...form, payee: "", amount: "", note: "" });
  }
  const statusColor = { pending: BRAND.gold, cashed: BRAND.darkgreen, bounced: BRAND.crimson };
  const statusLabel = { pending: "در انتظار", cashed: "نقد شده", bounced: "برگشتی" };
  const sorted = [...checks].sort((a, b) => a.dueDate.localeCompare(b.dueDate));
  const dueSoon = sorted.filter((c) => c.status === "pending" && daysUntil(c.dueDate) <= 7);
  return (
    <div>
      {dueSoon.length > 0 && (
        <div style={{ background: "#fff6ea", border: "1px solid #f0d9a8", borderRadius: 10, padding: 12, marginBottom: 16 }}>
          <div style={{ display: "flex", alignItems: "center", gap: 6, color: BRAND.orange, fontWeight: 700, fontSize: 13, marginBottom: 6 }}><BellRing size={16} /> یادآوری سررسید چک‌ها</div>
          {dueSoon.map((c) => (
            <div key={c.id} style={{ fontSize: 12.5, color: "#6b4c14" }}>
              {c.payee} — {toFaInt(c.amount)} ریال — {daysUntil(c.dueDate) < 0 ? "سررسید گذشته" : daysUntil(c.dueDate) === 0 ? "امروز سررسید است" : `${toFaInt(daysUntil(c.dueDate))} روز مانده`}
            </div>
          ))}
        </div>
      )}
      <div style={{ ...st.card, padding: 14, marginBottom: 16 }}>
        <div style={{ fontWeight: 700, marginBottom: 10, fontSize: 14 }}>ثبت چک جدید</div>
        <div style={{ display: "flex", gap: 8, marginBottom: 10 }}>
          <button onClick={() => setForm({ ...form, type: "received" })} style={pillStyle(form.type === "received")}>دریافتی</button>
          <button onClick={() => setForm({ ...form, type: "paid" })} style={pillStyle(form.type === "paid")}>پرداختی</button>
        </div>
        <input placeholder="نام طرف حساب" value={form.payee} onChange={(e) => setForm({ ...form, payee: e.target.value })} style={st.input} />
        <input placeholder="مبلغ" value={form.amount} onChange={(e) => setForm({ ...form, amount: e.target.value.replace(/[^0-9]/g, "") })} style={st.input} inputMode="numeric" />
        <label style={st.label}>تاریخ سررسید</label>
        <input type="date" value={form.dueDate} onChange={(e) => setForm({ ...form, dueDate: e.target.value })} style={st.input} />
        <input placeholder="یادداشت (اختیاری)" value={form.note} onChange={(e) => setForm({ ...form, note: e.target.value })} style={st.input} />
        <button onClick={add} style={st.primaryBtn}>ثبت چک</button>
      </div>
      <div style={{ ...st.card, padding: "4px 12px" }}>
        {checks.length === 0 && <EmptyRow text="چکی ثبت نشده" />}
        {sorted.map((c) => {
          const d = daysUntil(c.dueDate);
          const urgent = c.status === "pending" && d <= 7;
          return (
          <div key={c.id} style={{ padding: "12px 4px", borderBottom: "1px solid #f0eef3" }}>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 6 }}>
              <span style={{ fontWeight: 700, fontSize: 14 }}>{c.payee} ({c.type === "received" ? "دریافتی" : "پرداختی"})</span>
              <span style={{ fontWeight: 700 }}>{toFaInt(c.amount)} ریال</span>
            </div>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
              <span style={{ fontSize: 12, color: urgent ? (d < 0 ? BRAND.crimson : BRAND.orange) : "#8a8194", fontWeight: urgent ? 700 : 400 }}>
                سررسید: {faLongDate(new Date(c.dueDate))}{c.status === "pending" ? ` (${d < 0 ? "گذشته" : d === 0 ? "امروز" : `${toFaInt(d)} روز مانده`})` : ""}
              </span>
              <div style={{ display: "flex", gap: 6, alignItems: "center" }}>
                <StatusBadge text={statusLabel[c.status]} color={statusColor[c.status]} />
                <select value={c.status} onChange={(e) => setChecks((p) => p.map((x) => x.id === c.id ? { ...x, status: e.target.value } : x))} style={{ fontSize: 11, borderRadius: 6, border: "1px solid #ddd" }}>
                  <option value="pending">در انتظار</option><option value="cashed">نقد شده</option><option value="bounced">برگشتی</option>
                </select>
                <button onClick={() => setChecks((p) => p.filter((x) => x.id !== c.id))} style={{ background: "none", border: "none", color: BRAND.crimson, cursor: "pointer" }}><Trash2 size={15} /></button>
              </div>
            </div>
          </div>
          );
        })}
      </div>
    </div>
  );
}

/* ---------------------------------------------------------
   Loans Manager
--------------------------------------------------------- */
function LoansManager({ loans, setLoans }) {
  const st = useStyles();
  const [form, setForm] = useState({ title: "", principal: "", installments: "", monthlyPayment: "", startDate: todayISO() });
  function add() {
    if (!form.title || !form.principal || !form.installments || !form.monthlyPayment) return;
    setLoans((p) => [...p, { id: uid(), title: form.title, principal: Number(form.principal), installments: Number(form.installments), monthlyPayment: Number(form.monthlyPayment), startDate: form.startDate, paidCount: 0 }]);
    setForm({ title: "", principal: "", installments: "", monthlyPayment: "", startDate: todayISO() });
  }
  return (
    <div>
      <div style={{ ...st.card, padding: 14, marginBottom: 16 }}>
        <div style={{ fontWeight: 700, marginBottom: 10, fontSize: 14 }}>ثبت وام جدید</div>
        <input placeholder="عنوان وام (مثلا وام خودرو)" value={form.title} onChange={(e) => setForm({ ...form, title: e.target.value })} style={st.input} />
        <input placeholder="مبلغ اصل وام" value={form.principal} onChange={(e) => setForm({ ...form, principal: e.target.value.replace(/[^0-9]/g, "") })} style={st.input} inputMode="numeric" />
        <input placeholder="تعداد اقساط" value={form.installments} onChange={(e) => setForm({ ...form, installments: e.target.value.replace(/[^0-9]/g, "") })} style={st.input} inputMode="numeric" />
        <input placeholder="مبلغ هر قسط" value={form.monthlyPayment} onChange={(e) => setForm({ ...form, monthlyPayment: e.target.value.replace(/[^0-9]/g, "") })} style={st.input} inputMode="numeric" />
        <label style={st.label}>تاریخ شروع</label>
        <input type="date" value={form.startDate} onChange={(e) => setForm({ ...form, startDate: e.target.value })} style={st.input} />
        <button onClick={add} style={st.primaryBtn}>ثبت وام</button>
      </div>
      <div style={{ ...st.card, padding: "4px 12px" }}>
        {loans.length === 0 && <EmptyRow text="وامی ثبت نشده" />}
        {loans.map((l) => {
          const remaining = l.principal - l.paidCount * l.monthlyPayment;
          const nextDue = addMonths(l.startDate, l.paidCount);
          return (
            <div key={l.id} style={{ padding: "12px 4px", borderBottom: "1px solid #f0eef3" }}>
              <div style={{ display: "flex", justifyContent: "space-between", marginBottom: 6 }}>
                <span style={{ fontWeight: 700, fontSize: 14 }}>{l.title}</span>
                <button onClick={() => setLoans((p) => p.filter((x) => x.id !== l.id))} style={{ background: "none", border: "none", color: BRAND.crimson, cursor: "pointer" }}><Trash2 size={15} /></button>
              </div>
              <div style={{ fontSize: 12.5, color: "#8a8194", marginBottom: 4 }}>قسط {toFaInt(l.paidCount)} از {toFaInt(l.installments)} پرداخت شده — سررسید بعدی: {l.paidCount < l.installments ? faLongDate(new Date(nextDue)) : "تسویه شده"}</div>
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                <span style={{ fontWeight: 700, color: remaining > 0 ? BRAND.crimson : BRAND.darkgreen }}>باقی‌مانده: {toFaInt(Math.max(remaining, 0))} ریال</span>
                {l.paidCount < l.installments && (
                  <button onClick={() => setLoans((p) => p.map((x) => x.id === l.id ? { ...x, paidCount: x.paidCount + 1 } : x))} style={{ background: BRAND.green, color: "#fff", border: "none", borderRadius: 8, padding: "6px 12px", fontSize: 12, fontWeight: 700, cursor: "pointer" }}>ثبت پرداخت قسط</button>
                )}
              </div>
            </div>
          );
        })}
      </div>
    </div>
  );
}

/* ---------------------------------------------------------
   Bills Manager
--------------------------------------------------------- */
const BILL_CATS = ["آب", "برق", "گاز", "اینترنت", "تلفن", "سایر"];
function BillsManager({ bills, setBills }) {
  const st = useStyles();
  const [form, setForm] = useState({ title: "آب", amount: "", dueDate: todayISO(), recurringMonthly: true });
  function add() {
    setBills((p) => [...p, { id: uid(), title: form.title, amount: Number(form.amount || 0), dueDate: form.dueDate, recurringMonthly: form.recurringMonthly, paid: false }]);
    setForm({ ...form, amount: "" });
  }
  function markPaid(b) {
    setBills((prev) => prev.map((x) => x.id === b.id ? (x.recurringMonthly ? { ...x, dueDate: addMonths(x.dueDate, 1), paid: false } : { ...x, paid: true }) : x));
  }
  return (
    <div>
      <div style={{ ...st.card, padding: 14, marginBottom: 16 }}>
        <div style={{ fontWeight: 700, marginBottom: 10, fontSize: 14 }}>افزودن قبض</div>
        <div style={{ display: "flex", gap: 6, marginBottom: 10, flexWrap: "wrap" }}>
          {BILL_CATS.map((c) => <button key={c} onClick={() => setForm({ ...form, title: c })} style={{ ...pillStyle(form.title === c), flex: "none", padding: "7px 12px" }}>{c}</button>)}
        </div>
        <input placeholder="مبلغ (اختیاری)" value={form.amount} onChange={(e) => setForm({ ...form, amount: e.target.value.replace(/[^0-9]/g, "") })} style={st.input} inputMode="numeric" />
        <label style={st.label}>تاریخ سررسید</label>
        <input type="date" value={form.dueDate} onChange={(e) => setForm({ ...form, dueDate: e.target.value })} style={st.input} />
        <label style={{ display: "flex", alignItems: "center", gap: 8, fontSize: 13, marginBottom: 12 }}>
          <input type="checkbox" checked={form.recurringMonthly} onChange={(e) => setForm({ ...form, recurringMonthly: e.target.checked })} /> یادآوری ماهانه تکرار شود
        </label>
        <button onClick={add} style={st.primaryBtn}>افزودن قبض</button>
      </div>
      <div style={{ ...st.card, padding: "4px 12px" }}>
        {bills.length === 0 && <EmptyRow text="قبضی ثبت نشده" />}
        {bills.sort((a, b) => a.dueDate.localeCompare(b.dueDate)).map((b) => {
          const d = daysUntil(b.dueDate);
          return (
            <div key={b.id} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "12px 4px", borderBottom: "1px solid #f0eef3" }}>
              <div>
                <div style={{ fontWeight: 700, fontSize: 14 }}>{b.title}{b.amount ? ` — ${toFaInt(b.amount)} ریال` : ""}</div>
                <div style={{ fontSize: 12, color: b.paid ? BRAND.darkgreen : d < 0 ? BRAND.crimson : "#8a8194" }}>
                  {faLongDate(new Date(b.dueDate))} · {b.paid ? "پرداخت شده" : d < 0 ? "سررسید گذشته" : `${toFaInt(d)} روز مانده`}
                </div>
              </div>
              <div style={{ display: "flex", gap: 8 }}>
                {!b.paid && <button onClick={() => markPaid(b)} style={{ background: BRAND.green, color: "#fff", border: "none", borderRadius: 8, padding: "6px 10px", fontSize: 11.5, fontWeight: 700, cursor: "pointer" }}>ثبت پرداخت</button>}
                <button onClick={() => setBills((p) => p.filter((x) => x.id !== b.id))} style={{ background: "none", border: "none", color: BRAND.crimson, cursor: "pointer" }}><Trash2 size={15} /></button>
              </div>
            </div>
          );
        })}
      </div>
    </div>
  );
}

/* ---------------------------------------------------------
   Assets Manager
--------------------------------------------------------- */
function AssetsManager({ assets, setAssets }) {
  const st = useStyles();
  const [form, setForm] = useState({ kind: "crypto", symbol: "", quantity: "", avgPrice: "", currentPrice: "" });
  const total = assets.reduce((s, a) => s + a.quantity * a.currentPrice, 0);
  function add() {
    if (!form.symbol || !form.quantity) return;
    setAssets((p) => [...p, { id: uid(), kind: form.kind, symbol: form.symbol, quantity: Number(form.quantity), avgPrice: Number(form.avgPrice || 0), currentPrice: Number(form.currentPrice || form.avgPrice || 0) }]);
    setForm({ ...form, symbol: "", quantity: "", avgPrice: "", currentPrice: "" });
  }
  return (
    <div>
      <div style={{ background: "#fff6ea", border: "1px solid #f0d9a8", borderRadius: 10, padding: 10, marginBottom: 14, fontSize: 12, color: "#6b4c14" }}>
        قیمت لحظه‌ای به‌صورت خودکار دریافت نمی‌شود؛ قیمت فعلی را خودتان وارد یا ویرایش کنید.
      </div>
      <div style={{ ...st.card, padding: 14, marginBottom: 16 }}>
        <div style={{ fontWeight: 700, marginBottom: 10, fontSize: 14 }}>افزودن دارایی</div>
        <div style={{ display: "flex", gap: 8, marginBottom: 10 }}>
          <button onClick={() => setForm({ ...form, kind: "crypto" })} style={pillStyle(form.kind === "crypto")}>ارز دیجیتال</button>
          <button onClick={() => setForm({ ...form, kind: "stock" })} style={pillStyle(form.kind === "stock")}>بورس</button>
        </div>
        <input placeholder="نماد (مثلا BTC یا فولاد)" value={form.symbol} onChange={(e) => setForm({ ...form, symbol: e.target.value })} style={st.input} />
        <input placeholder="تعداد / مقدار" value={form.quantity} onChange={(e) => setForm({ ...form, quantity: e.target.value })} style={st.input} inputMode="decimal" />
        <input placeholder="قیمت خرید (ریال)" value={form.avgPrice} onChange={(e) => setForm({ ...form, avgPrice: e.target.value.replace(/[^0-9]/g, "") })} style={st.input} inputMode="numeric" />
        <input placeholder="قیمت فعلی (ریال)" value={form.currentPrice} onChange={(e) => setForm({ ...form, currentPrice: e.target.value.replace(/[^0-9]/g, "") })} style={st.input} inputMode="numeric" />
        <button onClick={add} style={st.primaryBtn}>افزودن دارایی</button>
      </div>
      <SectionTitle text={`ارزش کل: ${toFaInt(total)} ریال`} />
      <div style={{ ...st.card, padding: "4px 12px" }}>
        {assets.length === 0 && <EmptyRow text="دارایی‌ای ثبت نشده" />}
        {assets.map((a) => (
          <div key={a.id} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "12px 4px", borderBottom: "1px solid #f0eef3" }}>
            <div>
              <div style={{ fontWeight: 700, fontSize: 14 }}>{a.symbol} <span style={{ fontSize: 11, color: "#8a8194" }}>({a.kind === "crypto" ? "ارز دیجیتال" : "بورس"})</span></div>
              <div style={{ fontSize: 12, color: "#8a8194" }}>تعداد: {a.quantity} · ارزش: {toFaInt(a.quantity * a.currentPrice)} ریال</div>
            </div>
            <div style={{ display: "flex", gap: 8, alignItems: "center" }}>
              <input defaultValue={a.currentPrice} onBlur={(e) => setAssets((p) => p.map((x) => x.id === a.id ? { ...x, currentPrice: Number(e.target.value.replace(/[^0-9]/g, "") || 0) } : x))} style={{ ...st.input, margin: 0, width: 100, textAlign: "left" }} inputMode="numeric" />
              <button onClick={() => setAssets((p) => p.filter((x) => x.id !== a.id))} style={{ background: "none", border: "none", color: BRAND.crimson, cursor: "pointer" }}><Trash2 size={15} /></button>
            </div>
          </div>
        ))}
      </div>
    </div>
  );
}

/* ---------------------------------------------------------
   Members / Events / Projects Manager (تگ‌گذاری اعضای خانواده، رویداد و پروژه)
--------------------------------------------------------- */
function MembersEventsProjectsManager({ members, setMembers, events, setEvents, projects, setProjects }) {
  const st = useStyles();
  const [tab, setTab] = useState("members");
  const [name, setName] = useState("");
  const map = {
    members: { list: members, setList: setMembers, label: "عضو خانواده", color: BRAND.violet },
    events: { list: events, setList: setEvents, label: "رویداد", color: BRAND.teal },
    projects: { list: projects, setList: setProjects, label: "پروژه", color: BRAND.gold },
  };
  const cur = map[tab];
  function add() {
    if (!name.trim()) return;
    cur.setList((p) => [...p, { id: uid(), name: name.trim() }]);
    setName("");
  }
  return (
    <div>
      <div style={{ display: "flex", gap: 8, marginBottom: 16 }}>
        <button onClick={() => setTab("members")} style={pillStyle(tab === "members")}>اعضای خانواده</button>
        <button onClick={() => setTab("events")} style={pillStyle(tab === "events")}>رویدادها</button>
        <button onClick={() => setTab("projects")} style={pillStyle(tab === "projects")}>پروژه‌ها</button>
      </div>
      <div style={{ ...st.card, padding: 14, marginBottom: 16 }}>
        <div style={{ fontWeight: 700, marginBottom: 10, fontSize: 14 }}>افزودن {cur.label}</div>
        <input placeholder={`نام ${cur.label}`} value={name} onChange={(e) => setName(e.target.value)} style={st.input} />
        <button onClick={add} style={st.primaryBtn}>افزودن</button>
      </div>
      <div style={{ ...st.card, padding: "4px 12px" }}>
        {cur.list.length === 0 && <EmptyRow text={`${cur.label}ی ثبت نشده`} />}
        {cur.list.map((m) => (
          <Row key={m.id} title={m.name} leftIcon={<Trash2 size={15} />} leftColor={BRAND.crimson}
            onClick={() => cur.setList((p) => p.filter((x) => x.id !== m.id))} chevron={null} />
        ))}
      </div>
      <div style={{ fontSize: 12, color: "#8a8194", marginTop: 12 }}>
        این موارد را می‌توانید هنگام ثبت تراکنش (بخش «بیشتر») به هر پرداخت یا دریافت نسبت دهید تا در گزارش‌ها به‌تفکیک عضو، رویداد یا پروژه ببینید.
      </div>
    </div>
  );
}

/* ---------------------------------------------------------
   Fiscal Periods Manager (دوره مالی)
--------------------------------------------------------- */
function FiscalPeriodsManager({ fiscalPeriods, setFiscalPeriods }) {
  const st = useStyles();
  const [form, setForm] = useState({ title: "", startDate: todayISO(), endDate: todayISO() });
  function add() {
    if (!form.title) return;
    setFiscalPeriods((p) => [...p, { id: uid(), ...form }]);
    setForm({ title: "", startDate: todayISO(), endDate: todayISO() });
  }
  return (
    <div>
      <div style={{ ...st.card, padding: 14, marginBottom: 16 }}>
        <div style={{ fontWeight: 700, marginBottom: 10, fontSize: 14 }}>تعریف دوره مالی جدید</div>
        <div style={{ fontSize: 12, color: "#8a8194", marginBottom: 10 }}>مثلا «سال مالی ۱۴۰۴» یا «فصل بهار». در گزارش‌ها می‌توانید بر اساس این بازه فیلتر کنید.</div>
        <input placeholder="عنوان دوره" value={form.title} onChange={(e) => setForm({ ...form, title: e.target.value })} style={st.input} />
        <label style={st.label}>تاریخ شروع</label>
        <input type="date" value={form.startDate} onChange={(e) => setForm({ ...form, startDate: e.target.value })} style={st.input} />
        <label style={st.label}>تاریخ پایان</label>
        <input type="date" value={form.endDate} onChange={(e) => setForm({ ...form, endDate: e.target.value })} style={st.input} />
        <button onClick={add} style={st.primaryBtn}>افزودن دوره</button>
      </div>
      <div style={{ ...st.card, padding: "4px 12px" }}>
        {fiscalPeriods.length === 0 && <EmptyRow text="دوره‌ای ثبت نشده" />}
        {fiscalPeriods.map((p) => (
          <Row key={p.id} title={p.title} subtitle={`${faLongDate(new Date(p.startDate))} تا ${faLongDate(new Date(p.endDate))}`}
            leftIcon={<Trash2 size={15} />} leftColor={BRAND.crimson}
            onClick={() => setFiscalPeriods((prev) => prev.filter((x) => x.id !== p.id))} chevron={null} />
        ))}
      </div>
    </div>
  );
}

/* ---------------------------------------------------------
   Bank SMS quick-add (پیامک بانکی)
--------------------------------------------------------- */
function BankSmsManager({ openWithPrefill, onBack }) {
  const st = useStyles();
  const [text, setText] = useState("");
  const parsed = text.trim() ? parseBankSms(text) : null;
  return (
    <div>
      <div style={{ background: "#fff6ea", border: "1px solid #f0d9a8", borderRadius: 10, padding: 12, marginBottom: 16, fontSize: 12.5, color: "#6b4c14" }}>
        خواندن خودکار پیامک‌های بانکی در این محیط (مرورگر / پیش‌نمایش) امکان‌پذیر نیست، چون دسترسی به پیامک‌های گوشی نیازمند مجوز واقعی اندروید است. اما می‌توانید متن پیامک بانک را اینجا Paste کنید تا مبلغ و نوع تراکنش به‌طور خودکار تشخیص داده شود و فرم ثبت تراکنش از قبل پر شود.
        اگر بعداً این برنامه به‌صورت اپلیکیشن اندروید (APK) ساخته شود، افزودن خواندن واقعی پیامک با یک افزونه Capacitor امکان‌پذیر خواهد بود.
      </div>
      <div style={{ ...st.card, padding: 14, marginBottom: 16 }}>
        <div style={{ fontWeight: 700, marginBottom: 10, fontSize: 14 }}>متن پیامک بانکی را وارد کنید</div>
        <textarea value={text} onChange={(e) => setText(e.target.value)} placeholder="مثلا: از حساب شما مبلغ 250,000 ریال بابت خرید کسر شد..." rows={4}
          style={{ ...st.input, resize: "vertical", fontFamily: "inherit" }} />
        {parsed && (
          <div style={{ fontSize: 12.5, marginBottom: 10, color: "#8a8194" }}>
            تشخیص داده شده: {parsed.amount ? `${toFaInt(parsed.amount)} ریال` : "مبلغی یافت نشد"} — نوع: {parsed.type === "income" ? "دریافت" : "پرداخت"}
          </div>
        )}
        <button
          disabled={!parsed || !parsed.amount}
          onClick={() => { openWithPrefill({ amount: parsed.amount, type: parsed.type, note: parsed.note }); onBack(); }}
          style={{ ...st.primaryBtn, opacity: parsed && parsed.amount ? 1 : 0.5 }}>
          ادامه و تکمیل ثبت تراکنش
        </button>
      </div>
    </div>
  );
}

/* ---------------------------------------------------------
   Calendar View
--------------------------------------------------------- */
function CalendarViewSub({ transactions, catById }) {
  const t = useT();
  const now = new Date();
  const [jy, setJy] = useState(jalaliYear(now));
  const [jm, setJm] = useState(jalaliParts(now).m);
  const [selDay, setSelDay] = useState(null);
  const cells = useMemo(() => getJalaliMonthCells(jy, jm), [jy, jm]);
  const dayTx = (d) => transactions.filter((tx) => tx.date === d.date.toISOString().slice(0, 10));
  function prevMonth() { if (jm === 1) { setJm(12); setJy((y) => y - 1); } else setJm((m) => m - 1); setSelDay(null); }
  function nextMonth() { if (jm === 12) { setJm(1); setJy((y) => y + 1); } else setJm((m) => m + 1); setSelDay(null); }
  const weekDays = ["ش", "ی", "د", "س", "چ", "پ", "ج"];
  return (
    <div>
      <div style={{ display: "flex", alignItems: "center", justifyContent: "center", gap: 14, marginBottom: 16 }}>
        <button onClick={prevMonth} style={navArrowStyle(t)}><ChevronLeft size={16} /></button>
        <div style={{ fontWeight: 700 }}>{cells[0] ? faMonthYear(cells[0].date) : ""}</div>
        <button onClick={nextMonth} style={navArrowStyle(t)}><ChevronRight size={16} /></button>
      </div>
      <div style={{ display: "grid", gridTemplateColumns: "repeat(7, 1fr)", gap: 6, marginBottom: 8 }}>
        {weekDays.map((w) => <div key={w} style={{ textAlign: "center", fontSize: 11, fontWeight: 700, color: t.sub }}>{w}</div>)}
        {cells[0] && Array.from({ length: cells[0].weekday }).map((_, i) => <div key={"b" + i} />)}
        {cells.map((c) => {
          const txs = dayTx(c);
          const isSel = selDay && selDay.day === c.day;
          return (
            <button key={c.day} onClick={() => setSelDay(c)} style={{
              aspectRatio: "1", borderRadius: 9, border: "none", cursor: "pointer",
              background: isSel ? BRAND.header : t.card, color: isSel ? "#fff" : t.text,
              display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center", position: "relative", fontSize: 13, fontWeight: 600
            }}>
              {toFaInt(c.day)}
              {txs.length > 0 && <span style={{ width: 5, height: 5, borderRadius: "50%", background: isSel ? "#fff" : BRAND.crimson, position: "absolute", bottom: 5 }} />}
            </button>
          );
        })}
      </div>
      {selDay && (
        <div style={{ marginTop: 16 }}>
          <SectionTitle text={faLongDate(selDay.date)} />
          {dayTx(selDay).length === 0 && <EmptyRow text="تراکنشی در این روز نیست" />}
          {dayTx(selDay).map((tx) => (
            <Row key={tx.id} title={catById(tx.categoryId)?.name || (tx.type === "transfer" ? "انتقال" : "—")} value={`${toFaInt(tx.amount)} ریال`}
              valueColor={tx.type === "expense" ? BRAND.crimson : tx.type === "income" ? BRAND.darkgreen : BRAND.violet} chevron={null} />
          ))}
        </div>
      )}
    </div>
  );
}

/* ---------------------------------------------------------
   Settings View
--------------------------------------------------------- */
function SettingsView({ settings, setSettings, exportBackup, importBackup }) {
  const st = useStyles();
  const t = useT();
  const fileRef = useRef();
  const [pinInput, setPinInput] = useState("");
  const [profile, setProfile] = useState(settings.profile || { name: "", phone: "", email: "" });
  return (
    <div>
      <SectionTitle text="اطلاعات کاربری" />
      <div style={{ ...st.card, padding: 14, marginBottom: 18 }}>
        <input placeholder="نام و نام خانوادگی" value={profile.name} onChange={(e) => setProfile({ ...profile, name: e.target.value })} style={st.input} />
        <input placeholder="شماره تماس" value={profile.phone} onChange={(e) => setProfile({ ...profile, phone: e.target.value })} style={st.input} />
        <input placeholder="ایمیل (اختیاری)" value={profile.email} onChange={(e) => setProfile({ ...profile, email: e.target.value })} style={st.input} />
        <button onClick={() => setSettings((s) => ({ ...s, profile }))} style={st.primaryBtn}>ذخیره اطلاعات کاربری</button>
      </div>

      <SectionTitle text="رنگ‌بندی برنامه" />
      <div style={{ ...st.card, padding: 14, marginBottom: 18 }}>
        <div style={{ display: "flex", gap: 10, flexWrap: "wrap" }}>
          {Object.entries(COLOR_PRESETS).map(([key, p]) => (
            <button key={key} onClick={() => setSettings((s) => ({ ...s, themeColor: key }))}
              style={{ display: "flex", flexDirection: "column", alignItems: "center", gap: 6, background: "none", border: "none", cursor: "pointer" }}>
              <span style={{
                width: 42, height: 42, borderRadius: "50%", background: p.header,
                border: settings.themeColor === key ? `3px solid ${p.fab}` : "3px solid transparent",
                boxShadow: settings.themeColor === key ? "0 0 0 2px #fff, 0 0 0 3px " + p.header : "none",
                display: "flex", alignItems: "center", justifyContent: "center"
              }}>
                {settings.themeColor === key && <Check size={18} color="#fff" />}
              </span>
              <span style={{ fontSize: 10.5, color: t.text }}>{p.name}</span>
            </button>
          ))}
        </div>
      </div>

      <SectionTitle text="ظاهر برنامه" />
      <div style={{ ...st.card, padding: "4px 12px", marginBottom: 18 }}>
        <Row title="تم روشن / تاریک" leftIcon={settings.theme === "dark" ? <Moon size={16} /> : <Sun size={16} />} leftColor={BRAND.header}
          extra={<button onClick={() => setSettings((s) => ({ ...s, theme: s.theme === "dark" ? "light" : "dark" }))} style={{ background: BRAND.green, color: "#fff", border: "none", borderRadius: 8, padding: "6px 12px", fontSize: 12, fontWeight: 700, cursor: "pointer" }}>{settings.theme === "dark" ? "روشن کن" : "تاریک کن"}</button>}
          chevron={null} />
      </div>

      <SectionTitle text="امنیت" />
      <div style={{ ...st.card, padding: 14, marginBottom: 18 }}>
        <div style={{ fontSize: 12.5, color: "#8a8194", marginBottom: 10 }}>قفل با اثر انگشت در این محیط پشتیبانی نمی‌شود؛ به‌جای آن یک رمز عددی تنظیم کنید.</div>
        <input placeholder="رمز عددی جدید (خالی = بدون قفل)" value={pinInput} onChange={(e) => setPinInput(e.target.value.replace(/[^0-9]/g, "").slice(0, 6))} style={st.input} inputMode="numeric" />
        <button onClick={() => setSettings((s) => ({ ...s, pin: pinInput }))} style={st.primaryBtn}>{pinInput ? "تنظیم رمز" : "حذف رمز"}</button>
      </div>

      <SectionTitle text="حساب مشترک خانوادگی" />
      <div style={{ ...st.card, padding: 14, marginBottom: 18 }}>
        <div style={{ fontSize: 12.5, color: "#8a8194", marginBottom: 10 }}>
          با روشن‌کردن این گزینه، اطلاعات مالی (حساب‌ها، تراکنش‌ها، بودجه و...) به‌صورت مشترک ذخیره می‌شود و هر کسی که به همین برنامه دسترسی داشته باشد آن را می‌بیند. این حالت جایگزین ورود واقعی چند کاربره نیست، صرفاً یک دفتر مشترک است.
        </div>
        <Row title="فعال‌سازی حساب مشترک" leftIcon={<Users size={16} />} leftColor={BRAND.violet}
          extra={<button onClick={() => setSettings((s) => ({ ...s, sharedFamily: !s.sharedFamily }))} style={{ background: settings.sharedFamily ? BRAND.crimson : BRAND.green, color: "#fff", border: "none", borderRadius: 8, padding: "6px 12px", fontSize: 12, fontWeight: 700, cursor: "pointer" }}>{settings.sharedFamily ? "غیرفعال کن" : "فعال کن"}</button>}
          chevron={null} />
      </div>

      <SectionTitle text="پشتیبان‌گیری و بازیابی" />
      <div style={{ ...st.card, padding: 14, marginBottom: 18 }}>
        <button onClick={exportBackup} style={{ ...st.primaryBtn, background: BRAND.header, display: "flex", alignItems: "center", justifyContent: "center", gap: 6, marginBottom: 10 }}><Download size={16} /> دریافت فایل پشتیبان</button>
        <input ref={fileRef} type="file" accept="application/json" style={{ display: "none" }} onChange={(e) => e.target.files[0] && importBackup(e.target.files[0])} />
        <button onClick={() => fileRef.current.click()} style={{ ...st.primaryBtn, background: BRAND.gold, display: "flex", alignItems: "center", justifyContent: "center", gap: 6 }}><Upload size={16} /> بازیابی از فایل پشتیبان</button>
      </div>

      <SectionTitle text="درباره همگام‌سازی با Google Drive" />
      <div style={{ ...st.card, padding: 14, fontSize: 12.5, color: "#8a8194" }}>
        اتصال مستقیم به Google Drive نیازمند ورود واقعی به حساب گوگل است که در این محیط در دسترس نیست. برای انتقال اطلاعات بین دستگاه‌ها از «دریافت فایل پشتیبان» استفاده کنید و همان فایل را در Drive خودتان نگه دارید یا در دستگاه دیگر «بازیابی» کنید.
      </div>
    </div>
  );
}

/* ---------------------------------------------------------
   Add Transaction Sheet
--------------------------------------------------------- */
function AddTransactionSheet({ accounts, categories, favorites, members = [], events = [], projects = [], initial, onClose, onSubmit }) {
  const st = useStyles();
  const [type, setType] = useState(initial?.type || "expense");
  const [amount, setAmount] = useState(initial?.amount ? String(initial.amount) : "");
  const [categoryId, setCategoryId] = useState("");
  const [accountId, setAccountId] = useState(accounts[0]?.id || "");
  const [toAccountId, setToAccountId] = useState(accounts[1]?.id || accounts[0]?.id || "");
  const [date, setDate] = useState(todayISO());
  const [note, setNote] = useState(initial?.note || "");
  const [tagsInput, setTagsInput] = useState("");
  const [photo, setPhoto] = useState(null);
  const [showMore, setShowMore] = useState(false);
  const [memberId, setMemberId] = useState("");
  const [eventId, setEventId] = useState("");
  const [projectId, setProjectId] = useState("");

  const filteredCats = categories.filter((c) => c.kind === type);
  const favCats = filteredCats.filter((c) => favorites.categories.includes(c.id));
  const canSubmit = amount && Number(amount) > 0 && accountId && (type === "transfer" ? toAccountId && toAccountId !== accountId : categoryId);

  async function handlePhoto(e) {
    const f = e.target.files[0];
    if (!f) return;
    const dataUrl = await resizeImage(f);
    setPhoto(dataUrl);
  }

  return (
    <div style={{ position: "absolute", inset: 0, background: "rgba(0,0,0,0.45)", display: "flex", alignItems: "flex-end", zIndex: 50, maxWidth: 480, margin: "0 auto" }}>
      <div style={{ background: "#fff", width: "100%", borderRadius: "18px 18px 0 0", padding: "18px 18px 26px", maxHeight: "88vh", overflowY: "auto" }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 14 }}>
          <button onClick={onClose} style={{ background: "none", border: "none", cursor: "pointer", color: "#8a8194" }}><X size={22} /></button>
          <div style={{ fontWeight: 700, fontSize: 16 }}>ثبت تراکنش جدید</div>
          <div style={{ width: 22 }} />
        </div>

        <div style={{ display: "flex", gap: 8, marginBottom: 14 }}>
          {[{ k: "expense", l: "پرداخت", c: BRAND.crimson }, { k: "income", l: "دریافت", c: BRAND.darkgreen }, { k: "transfer", l: "انتقال", c: BRAND.violet }].map((o) => (
            <button key={o.k} onClick={() => { setType(o.k); setCategoryId(""); }} style={{ flex: 1, padding: "10px 4px", borderRadius: 9, border: `1.5px solid ${type === o.k ? o.c : "#e3e0ea"}`, background: type === o.k ? o.c : "#fff", color: type === o.k ? "#fff" : "#241a30", fontWeight: 700, cursor: "pointer", fontSize: 13.5 }}>{o.l}</button>
          ))}
        </div>

        {favCats.length > 0 && (
          <div style={{ display: "flex", gap: 6, flexWrap: "wrap", marginBottom: 12 }}>
            {favCats.map((c) => (
              <button key={c.id} onClick={() => setCategoryId(c.id)} style={{ display: "flex", alignItems: "center", gap: 4, background: categoryId === c.id ? BRAND.header : "#f1eef4", color: categoryId === c.id ? "#fff" : "#3E1461", border: "none", borderRadius: 20, padding: "6px 12px", fontSize: 12, fontWeight: 700, cursor: "pointer" }}>
                <Star size={11} fill="#f5b301" color="#f5b301" /> {c.name}
              </button>
            ))}
          </div>
        )}

        <label style={st.label}>مبلغ (ریال)</label>
        <input value={amount} onChange={(e) => setAmount(e.target.value.replace(/[^0-9]/g, ""))} placeholder="0" inputMode="numeric" style={st.input} />

        {type !== "transfer" ? (
          <>
            <label style={st.label}>دسته‌بندی</label>
            <select value={categoryId} onChange={(e) => setCategoryId(e.target.value)} style={st.input}>
              <option value="">انتخاب کنید</option>
              {filteredCats.map((c) => <option key={c.id} value={c.id}>{c.name}</option>)}
            </select>
            <label style={st.label}>{type === "expense" ? "از حساب" : "به حساب"}</label>
            <select value={accountId} onChange={(e) => setAccountId(e.target.value)} style={st.input}>
              {accounts.map((a) => <option key={a.id} value={a.id}>{a.name}</option>)}
            </select>
          </>
        ) : (
          <>
            <label style={st.label}>از حساب</label>
            <select value={accountId} onChange={(e) => setAccountId(e.target.value)} style={st.input}>
              {accounts.map((a) => <option key={a.id} value={a.id}>{a.name}</option>)}
            </select>
            <label style={st.label}>به حساب</label>
            <select value={toAccountId} onChange={(e) => setToAccountId(e.target.value)} style={st.input}>
              {accounts.map((a) => <option key={a.id} value={a.id}>{a.name}</option>)}
            </select>
          </>
        )}

        <label style={st.label}>تاریخ</label>
        <input type="date" value={date} onChange={(e) => setDate(e.target.value)} style={st.input} />

        <label style={st.label}>برچسب‌ها (با کاما جدا کنید)</label>
        <input value={tagsInput} onChange={(e) => setTagsInput(e.target.value)} placeholder="مثلا: سفر, ضروری" style={st.input} />

        <label style={st.label}>یادداشت (اختیاری)</label>
        <input value={note} onChange={(e) => setNote(e.target.value)} placeholder="توضیحات..." style={st.input} />

        <label style={st.label}>پیوست عکس (اختیاری)</label>
        <label style={{ display: "flex", alignItems: "center", gap: 8, border: "1.5px dashed #d8d2e0", borderRadius: 9, padding: "10px 12px", cursor: "pointer", marginBottom: 12, fontSize: 13, color: "#8a8194" }}>
          <ImageIcon size={16} /> {photo ? "عکس انتخاب شد ✓" : "افزودن عکس رسید"}
          <input type="file" accept="image/*" onChange={handlePhoto} style={{ display: "none" }} />
        </label>
        {photo && <img src={photo} style={{ width: 60, height: 60, borderRadius: 8, objectFit: "cover", marginBottom: 12 }} />}

        {(members.length > 0 || events.length > 0 || projects.length > 0) && (
          <div style={{ marginBottom: 12 }}>
            <button onClick={() => setShowMore((v) => !v)} style={{ background: "none", border: "none", color: BRAND.header, fontWeight: 700, fontSize: 12.5, cursor: "pointer", padding: 0, marginBottom: showMore ? 10 : 0 }}>
              {showMore ? "بستن جزئیات بیشتر ▲" : "افزودن عضو / رویداد / پروژه ▼"}
            </button>
            {showMore && (
              <>
                {members.length > 0 && (<>
                  <label style={st.label}>عضو خانواده</label>
                  <select value={memberId} onChange={(e) => setMemberId(e.target.value)} style={st.input}>
                    <option value="">— بدون عضو —</option>
                    {members.map((m) => <option key={m.id} value={m.id}>{m.name}</option>)}
                  </select>
                </>)}
                {events.length > 0 && (<>
                  <label style={st.label}>رویداد</label>
                  <select value={eventId} onChange={(e) => setEventId(e.target.value)} style={st.input}>
                    <option value="">— بدون رویداد —</option>
                    {events.map((ev) => <option key={ev.id} value={ev.id}>{ev.name}</option>)}
                  </select>
                </>)}
                {projects.length > 0 && (<>
                  <label style={st.label}>پروژه</label>
                  <select value={projectId} onChange={(e) => setProjectId(e.target.value)} style={st.input}>
                    <option value="">— بدون پروژه —</option>
                    {projects.map((pr) => <option key={pr.id} value={pr.id}>{pr.name}</option>)}
                  </select>
                </>)}
              </>
            )}
          </div>
        )}

        <button disabled={!canSubmit}
          onClick={() => onSubmit({
            type, amount: Number(amount), categoryId, accountId,
            toAccountId: type === "transfer" ? toAccountId : undefined,
            date, note, photo: photo || undefined,
            tags: tagsInput.split(",").map((s) => s.trim()).filter(Boolean),
            memberId: memberId || undefined, eventId: eventId || undefined, projectId: projectId || undefined,
          })}
          style={{ ...st.primaryBtn, opacity: canSubmit ? 1 : 0.5, marginTop: 6 }}>
          <span style={{ display: "flex", alignItems: "center", justifyContent: "center", gap: 6 }}><Check size={17} /> ثبت تراکنش</span>
        </button>
      </div>
    </div>
  );
}

/* ---------------------------------------------------------
   Side menu
--------------------------------------------------------- */
function SideMenu({ onClose, setSubView, profileName }) {
  const items = [
    { label: "ویرایش اطلاعات کاربری", key: "profile" },
    { label: "تنظیمات و پشتیبان‌گیری", key: "settings" },
    { label: "اعضای منزل، رویداد و پروژه", key: "tags" },
    { label: "دوره مالی", key: "periods" },
    { label: "پیامک بانکی", key: "sms" },
    { label: "تقویم شمسی", key: "calendar" },
    { label: "دارایی‌ها", key: "assets" },
    { label: "چک‌ها", key: "checks" },
    { label: "وام و اقساط", key: "loans" },
    { label: "یادآوری قبض‌ها", key: "bills" },
  ];
  return (
    <div style={{ position: "absolute", inset: 0, background: "rgba(0,0,0,0.4)", zIndex: 60, maxWidth: 480, margin: "0 auto" }} onClick={onClose}>
      <div onClick={(e) => e.stopPropagation()} style={{ position: "absolute", top: 0, left: 0, width: "82%", height: "100%", background: "#fff", boxShadow: "3px 0 12px rgba(0,0,0,0.2)", overflowY: "auto" }}>
        <div style={{ background: BRAND.header, color: "#fff", padding: "18px 16px", display: "flex", justifyContent: "space-between", alignItems: "center" }}>
          <button onClick={onClose} style={{ background: "none", border: "none", color: "#fff", cursor: "pointer" }}><X size={22} /></button>
          <div style={{ fontWeight: 700 }}>{profileName || "alireza shadfar"}</div>
        </div>
        <div style={{ padding: "8px 16px" }}>
          {items.map((it) => (
            <div key={it.key} onClick={() => setSubView(it.key)} style={{ padding: "13px 4px", borderBottom: "1px solid #f0eef3", fontSize: 14, fontWeight: 600, color: "#241a30", cursor: "pointer" }}>{it.label}</div>
          ))}
        </div>
      </div>
    </div>
  );
}
