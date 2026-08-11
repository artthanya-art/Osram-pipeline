import React, { useState, useEffect, useMemo, useRef } from "react";
import {
  LogIn, LogOut, Plus, Search, Trash2, Edit3, Save, X, Filter,
  TrendingUp, Users, Lightbulb, Download, Upload, UserPlus,
  BarChart3, ListChecks, AlertCircle, CheckCircle2, ShieldCheck,
  UserCheck, UserX, Clock, ShieldAlert, Database, FileUp, FileDown,
  Crown, ChevronRight,
} from "lucide-react";
import Papa from "papaparse";
import * as XLSX from "xlsx";
import {
  fetchUsers, insertUser, updateUserFields,
  fetchZones, insertZone,
  fetchEntries, insertEntry, updateEntry as updateEntryDb, deleteEntryById,
  bulkInsertEntries, replaceAllEntries, replaceAllUsers,
} from "./db";
import {
  BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip, ResponsiveContainer, Cell, PieChart, Pie, Legend
} from "recharts";

// ---------- constants ----------
const RED = "#E2001A";
const INK = "#17171B";
const AMBER = "#FFB81C";
const PAPER = "#F6F5F2";
const GREEN = "#1E8E3E";

const DEFAULT_ZONES = ["Red Zone", "Blue Zone", "Green Zone", "Yellow Zone", "Orange Zone"];
const ADD_ZONE_VALUE = "__add_new_zone__";
const CUSTOMER_TYPE_OPTIONS = ["New Reg", "Existing"];
const SEGMENT_OPTIONS = ["Wholesale", "Retail", "Project", "OEM", "Government"];
const DELIVERY_OPTIONS = ["ส่งครั้งเดียว", "ทยอยส่ง"];
const PROGRESS_OPTIONS = [
  "0% - ยังไม่ได้เริ่มงาน",
  "20% - นำเสนอสินค้า",
  "40% - ทำใบเสนอราคา/ทดสอบสินค้า",
  "90% - ได้งาน - ทยอยส่ง",
  "100% - จบงาน - ได้งาน",
  "100% - จบงาน - ไม่ได้งาน",
  "100% - ลูกค้าขอปฏิเสธ",
];
const GRADE_OPTIONS = ["A", "B", "C"];
const MONTH_OPTIONS = ["ม.ค.","ก.พ.","มี.ค.","เม.ย.","พ.ค.","มิ.ย.","ก.ค.","ส.ค.","ก.ย.","ต.ค.","พ.ย.","ธ.ค."];

const ENTRY_COLS = [
  "salesId","salesName","zone","customerId","customerName","customerType","customerSegment",
  "productType","itemCode","itemDescription","qty","uom","price","competitorName","competitorPrice",
  "projectCloseYear","projectCloseMonth","quotationNumber","deliveryMethod","kpiRegister",
  "actionPlanMonth","progress","deliverInMonths","deliveryStartYear","deliveryStartMonth",
  "startWorkingMonth","visitDate","grade",
];

// Maps each internal field to every column-header spelling we should recognize on import —
// both our own CSV export headers (camelCase) and the original OSRAM Excel report headers.
const HEADER_ALIASES = {
  salesId: ["salesId", "sales id"],
  salesName: ["salesName", "sales name"],
  zone: ["zone", "โซน"],
  customerId: ["customerId", "customer id"],
  customerName: ["customerName", "customer name"],
  customerType: ["customerType", "customer type"],
  customerSegment: ["customerSegment", "customer segment"],
  productType: ["productType", "product type"],
  itemCode: ["itemCode", "item code", "full code"],
  itemDescription: ["itemDescription", "item description"],
  qty: ["qty", "quantity"],
  uom: ["uom", "unit"],
  price: ["price"],
  competitorName: ["competitorName", "competitor name"],
  competitorPrice: ["competitorPrice", "competitor price"],
  projectCloseYear: ["projectCloseYear", "project close year"],
  projectCloseMonth: ["projectCloseMonth", "project close month"],
  quotationNumber: ["quotationNumber", "quotation number", "quotation no", "quotation no."],
  deliveryMethod: ["deliveryMethod", "delivery method", "วิธีการจัดส่ง"],
  kpiRegister: ["kpiRegister", "kpi register"],
  actionPlanMonth: ["actionPlanMonth", "action plan (month)", "action plan"],
  progress: ["progress", "% progresssion", "% progression", "progression"],
  deliverInMonths: ["deliverInMonths", "90% ทยอยส่ง (กี่เดือน)", "deliver in months"],
  deliveryStartYear: ["deliveryStartYear", "deliverly start year", "delivery start year"],
  deliveryStartMonth: ["deliveryStartMonth", "deliverly start month", "delivery start month"],
  startWorkingMonth: ["startWorkingMonth", "start working month"],
  visitDate: ["visitDate", "visit date"],
  grade: ["grade"],
};

function normalizeHeader(h) {
  return String(h ?? "").trim().toLowerCase().replace(/[\s_\-./]+/g, "");
}
// normalized-alias -> internal field name, built once
const HEADER_LOOKUP = Object.entries(HEADER_ALIASES).reduce((acc, [field, aliases]) => {
  aliases.forEach((a) => { acc[normalizeHeader(a)] = field; });
  return acc;
}, {});

function pad2(n) { return String(n).padStart(2, "0"); }
function isoFromDate(d) {
  return `${d.getFullYear()}-${pad2(d.getMonth() + 1)}-${pad2(d.getDate())}`;
}
function isoFromParts(y, mo, d) {
  if (!y || !mo || !d) return "";
  const dt = new Date(y, mo - 1, d);
  if (isNaN(dt)) return "";
  return isoFromDate(dt);
}
// Excel's serial date system: day 1 = 1900-01-01 (with Excel's well-known 1900 leap-year bug,
// which this offset already accounts for). This is what lets us recover a real date from a
// cell that was typed as a plain number instead of being formatted as a date.
function excelSerialToDate(serial) {
  const utcDays = Math.floor(serial - 25569);
  return new Date(utcDays * 86400 * 1000);
}

// Normalizes messy visit-date input into a consistent yyyy-mm-dd string. Handles:
//  - a real JS Date (Excel cell was formatted as a date)
//  - a raw Excel serial number (Excel cell was NOT formatted as a date, just a plain number)
//  - text dates in yyyy-mm-dd, yyyy/mm/dd, dd/mm/yyyy, or dd-mm-yyyy
//  - Thai Buddhist-era years (e.g. 2569) are converted to the Gregorian year
//  - anything unrecognized is returned as-is so no data silently disappears
function normalizeDateValue(v) {
  if (v === null || v === undefined || v === "") return "";

  if (v instanceof Date && !isNaN(v)) return isoFromDate(v);

  if (typeof v === "number") {
    if (v > 20000 && v < 80000) {
      const d = excelSerialToDate(v);
      if (!isNaN(d)) return isoFromDate(d);
    }
    return String(v);
  }

  const s = String(v).trim();
  if (!s) return "";

  let m = s.match(/^(\d{4})[-/](\d{1,2})[-/](\d{1,2})$/);
  if (m) return isoFromParts(+m[1], +m[2], +m[3]);

  m = s.match(/^(\d{1,2})[-/](\d{1,2})[-/](\d{2,4})$/);
  if (m) {
    let [, d, mo, y] = m.map(Number);
    if (y < 100) y += 2000;
    if (y > 2400) y -= 543; // Buddhist calendar -> Gregorian
    return isoFromParts(y, mo, d);
  }

  if (/^\d{4,6}$/.test(s)) {
    const n = Number(s);
    if (n > 20000 && n < 80000) {
      const d = excelSerialToDate(n);
      if (!isNaN(d)) return isoFromDate(d);
    }
  }

  const parsed = new Date(s);
  if (!isNaN(parsed)) return isoFromDate(parsed);

  return s; // couldn't recognize it — keep the original text rather than losing it
}

// Turns one raw row (arbitrary header names, from CSV or Excel) into our internal entry shape.
function mapRowToEntry(row) {
  const entry = {};
  Object.entries(row).forEach(([rawHeader, rawValue]) => {
    const field = HEADER_LOOKUP[normalizeHeader(rawHeader)];
    if (!field) return;
    entry[field] = field === "visitDate" ? normalizeDateValue(rawValue) : String(rawValue ?? "").trim();
  });
  return entry;
}

const EMPTY_FORM = {
  zone: "Red Zone",
  customerId: "New",
  customerName: "",
  customerType: "New Reg",
  customerSegment: "Wholesale",
  productType: "",
  itemCode: "",
  itemDescription: "",
  qty: "",
  uom: "",
  price: "",
  competitorName: "",
  competitorPrice: "",
  projectCloseYear: "",
  projectCloseMonth: "",
  quotationNumber: "",
  deliveryMethod: "ส่งครั้งเดียว",
  kpiRegister: "",
  actionPlanMonth: "",
  progress: "0% - ยังไม่ได้เริ่มงาน",
  deliverInMonths: "",
  deliveryStartYear: "",
  deliveryStartMonth: "",
  startWorkingMonth: "",
  visitDate: "",
  grade: "B",
};


function gradeColor(g) {
  if (g === "A") return GREEN;
  if (g === "B") return AMBER;
  return "#9AA0A6";
}
function progressPercent(p) {
  const m = /^(\d+)%/.exec(p || "");
  return m ? parseInt(m[1], 10) : 0;
}
function money(n) {
  const v = Number(n) || 0;
  return v.toLocaleString("th-TH", { maximumFractionDigits: 0 });
}
function dealValue(e) {
  return (Number(e.qty) || 0) * (Number(e.price) || 0);
}

// ---------- UI atoms ----------
function Field({ label, children, required }) {
  return (
    <label style={{ display: "flex", flexDirection: "column", gap: 6, fontSize: 13 }}>
      <span style={{ fontWeight: 600, color: "#4B4B52" }}>
        {label} {required && <span style={{ color: RED }}>*</span>}
      </span>
      {children}
    </label>
  );
}
const inputStyle = {
  border: "1px solid #DEDCD6", borderRadius: 8, padding: "9px 11px", fontSize: 14,
  outline: "none", fontFamily: "inherit", background: "#fff", color: INK, width: "100%", boxSizing: "border-box",
};
function TextInput(props) { return <input {...props} style={{ ...inputStyle, ...(props.style || {}) }} />; }
function Select({ children, ...props }) { return <select {...props} style={{ ...inputStyle, ...(props.style || {}) }}>{children}</select>; }
function SectionTitle({ children, icon: Icon }) {
  return (
    <div style={{ display: "flex", alignItems: "center", gap: 8, margin: "22px 0 12px" }}>
      {Icon && <Icon size={16} color={RED} />}
      <h3 style={{ fontSize: 13, letterSpacing: "0.06em", textTransform: "uppercase", color: RED, margin: 0, fontWeight: 800 }}>{children}</h3>
      <div style={{ flex: 1, height: 1, background: "#E7E5DF" }} />
    </div>
  );
}
function Badge({ children, color, bg }) {
  return (
    <span style={{ display: "inline-flex", alignItems: "center", gap: 4, padding: "3px 8px", borderRadius: 999, fontSize: 11, fontWeight: 700, color, background: bg }}>
      {children}
    </span>
  );
}
function Th({ children }) {
  return <th style={{ padding: "10px 12px", fontSize: 11.5, color: "#8A8A90", fontWeight: 700, textTransform: "uppercase", letterSpacing: "0.04em", whiteSpace: "nowrap" }}>{children}</th>;
}
function Td({ children, style }) { return <td style={{ padding: "10px 12px", verticalAlign: "top", ...style }}>{children}</td>; }
function StatCard({ label, value, sub, icon: Icon }) {
  return (
    <div style={{ background: INK, borderRadius: 14, padding: 18, color: "#fff" }}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start" }}>
        <div style={{ fontSize: 12, color: "#B9B9C0", fontWeight: 600 }}>{label}</div>
        {Icon && <Icon size={15} color={AMBER} />}
      </div>
      <div style={{ fontSize: 30, fontWeight: 800, margin: "6px 0 2px" }}>{value}</div>
      <div style={{ fontSize: 11.5, color: "#9A9AA0" }}>{sub}</div>
    </div>
  );
}

// Zone select that lets the user pick an existing zone or type a brand-new one inline.
function ZoneSelect({ zones, value, onChange, onAddZone, style }) {
  const [adding, setAdding] = useState(false);
  const [draft, setDraft] = useState("");

  function commitNewZone() {
    const name = draft.trim();
    if (!name) { setAdding(false); return; }
    onAddZone(name);
    onChange(name);
    setAdding(false);
    setDraft("");
  }

  if (adding) {
    return (
      <div style={{ display: "flex", gap: 6 }}>
        <TextInput
          autoFocus
          value={draft}
          onChange={(e) => setDraft(e.target.value)}
          placeholder="ชื่อโซนใหม่ เช่น Purple Zone"
          onKeyDown={(e) => {
            if (e.key === "Enter") { e.preventDefault(); commitNewZone(); }
            if (e.key === "Escape") { setAdding(false); setDraft(""); }
          }}
          style={style}
        />
        <button type="button" onClick={commitNewZone} title="เพิ่มโซนนี้" style={{ border: "none", background: RED, color: "#fff", borderRadius: 8, padding: "0 12px", cursor: "pointer", fontWeight: 700, fontSize: 13 }}>+</button>
        <button type="button" onClick={() => { setAdding(false); setDraft(""); }} title="ยกเลิก" style={{ border: "1px solid #DEDCD6", background: "#fff", color: "#6B6B70", borderRadius: 8, padding: "0 10px", cursor: "pointer", fontSize: 13 }}>
          <X size={13} />
        </button>
      </div>
    );
  }

  return (
    <Select
      value={value}
      onChange={(e) => {
        if (e.target.value === ADD_ZONE_VALUE) { setAdding(true); return; }
        onChange(e.target.value);
      }}
      style={style}
    >
      {zones.map((z) => <option key={z} value={z}>{z}</option>)}
      <option value={ADD_ZONE_VALUE}>+ เพิ่มโซนใหม่...</option>
    </Select>
  );
}

// ================= APP =================
export default function App() {
  const [ready, setReady] = useState(false);
  const [users, setUsers] = useState([]);
  const [entries, setEntries] = useState([]);
  const [zones, setZones] = useState(DEFAULT_ZONES);
  const [currentUser, setCurrentUser] = useState(null);

  const [authMode, setAuthMode] = useState("login");
  const [authForm, setAuthForm] = useState({ salesId: "", password: "", name: "", zone: "Red Zone" });
  const [authError, setAuthError] = useState("");
  const [authNotice, setAuthNotice] = useState("");
  const [showForgotHint, setShowForgotHint] = useState(false);

  const [tab, setTab] = useState("new");
  const [adminSubTab, setAdminSubTab] = useState("pending");
  const [form, setForm] = useState(EMPTY_FORM);
  const [editingId, setEditingId] = useState(null);
  const [toast, setToast] = useState("");
  const [importing, setImporting] = useState(false);

  const [search, setSearch] = useState("");
  const [filterZone, setFilterZone] = useState("");
  const [filterGrade, setFilterGrade] = useState("");
  const [filterRep, setFilterRep] = useState("");

  const importFileRef = useRef(null);
  const backupFileRef = useRef(null);

  useEffect(() => {
    (async () => {
      try {
        let u = await fetchUsers();
        if (u.length === 0) {
          // Seed demo accounts so the app is usable immediately (also seeded by supabase/schema.sql).
          await insertUser({ salesId: "admin", password: "admin123", name: "ผู้ดูแลระบบ (Demo)", zone: "Red Zone", role: "admin", status: "approved" });
          await insertUser({ salesId: "demo", password: "demo123", name: "พนักงานขาย (Demo)", zone: "Red Zone", role: "sales", status: "approved" });
          u = await fetchUsers();
        }
        setUsers(u);

        const e = await fetchEntries();
        setEntries(e);

        const z = await fetchZones();
        setZones(z.length ? z : DEFAULT_ZONES);
      } catch (err) {
        console.error(err);
        showToast("เชื่อมต่อฐานข้อมูล Supabase ไม่สำเร็จ ตรวจสอบไฟล์ .env");
      }
      setReady(true);
    })();
  }, []);

  function showToast(msg) { setToast(msg); setTimeout(() => setToast(""), 2600); }

  async function addZone(name) {
    const trimmed = name.trim();
    if (!trimmed) return;
    if (zones.some((z) => z.toLowerCase() === trimmed.toLowerCase())) { showToast(`มีโซน "${trimmed}" อยู่แล้ว`); return; }
    try {
      await insertZone(trimmed);
      setZones((prev) => [...prev, trimmed]);
      showToast(`เพิ่มโซน "${trimmed}" เรียบร้อย`);
    } catch (err) { showToast("เพิ่มโซนไม่สำเร็จ: " + err.message); }
  }

  // ---------- auth ----------
  function handleAuthSubmit(e) {
    e.preventDefault();
    setAuthError(""); setAuthNotice("");
    const salesId = authForm.salesId.trim();
    if (!salesId || !authForm.password) { setAuthError("กรุณากรอก Sales ID และรหัสผ่าน"); return; }

    if (authMode === "register") {
      if (!authForm.name.trim()) { setAuthError("กรุณากรอกชื่อพนักงานขาย"); return; }
      if (users.some((u) => u.salesId.toLowerCase() === salesId.toLowerCase())) {
        setAuthError("Sales ID นี้มีผู้ใช้งานแล้ว กรุณาเข้าสู่ระบบแทน"); return;
      }
      const isFirstUser = users.length === 0;
      const newUser = {
        salesId, password: authForm.password, name: authForm.name.trim(), zone: authForm.zone,
        role: isFirstUser ? "admin" : "sales",
        status: isFirstUser ? "approved" : "pending",
      };
      (async () => {
        try {
          await insertUser(newUser);
          const savedUser = { ...newUser, createdAt: new Date().toISOString() };
          setUsers((prev) => [...prev, savedUser]);
          setAuthForm({ salesId: "", password: "", name: "", zone: "Red Zone" });
          if (isFirstUser) {
            setCurrentUser(savedUser);
            showToast(`สร้างบัญชีผู้ดูแลระบบสำเร็จ ยินดีต้อนรับคุณ ${savedUser.name}`);
          } else {
            setAuthMode("login");
            setAuthNotice("สมัครสมาชิกสำเร็จ กรุณารอผู้ดูแลระบบอนุมัติบัญชีก่อนเข้าสู่ระบบ");
          }
        } catch (err) {
          setAuthError("สมัครสมาชิกไม่สำเร็จ: " + err.message);
        }
      })();
    } else {
      const found = users.find((u) => u.salesId.toLowerCase() === salesId.toLowerCase());
      if (!found || found.password !== authForm.password) { setAuthError("Sales ID หรือรหัสผ่านไม่ถูกต้อง"); return; }
      if (found.status === "pending") { setAuthError("บัญชีนี้ยังรอการอนุมัติจากผู้ดูแลระบบ"); return; }
      if (found.status === "rejected") { setAuthError("บัญชีนี้ถูกระงับการใช้งาน กรุณาติดต่อผู้ดูแลระบบ"); return; }
      setCurrentUser(found);
      setAuthForm({ salesId: "", password: "", name: "", zone: "Red Zone" });
    }
  }

  function handleLogout() {
    setCurrentUser(null); setTab("new"); setForm(EMPTY_FORM); setEditingId(null);
  }

  // ---------- entry CRUD ----------
  function updateField(key, value) { setForm((f) => ({ ...f, [key]: value })); }
  function resetForm() { setForm({ ...EMPTY_FORM, zone: currentUser?.zone || "Red Zone" }); setEditingId(null); }

  async function handleSaveEntry(e) {
    e.preventDefault();
    if (!form.customerName.trim()) { showToast("กรุณากรอกชื่อลูกค้า"); return; }
    try {
      if (editingId) {
        await updateEntryDb(editingId, form);
        setEntries((prev) => prev.map((en) => (en.id === editingId ? { ...en, ...form, updatedAt: new Date().toISOString() } : en)));
        showToast("บันทึกการแก้ไขเรียบร้อย");
      } else {
        const created = await insertEntry({ salesId: currentUser.salesId, salesName: currentUser.name, ...form });
        setEntries((prev) => [created, ...prev]);
        showToast("เพิ่มรายการ Pipeline เรียบร้อย");
      }
      resetForm(); setTab("mine");
    } catch (err) {
      showToast("บันทึกไม่สำเร็จ: " + err.message);
    }
  }
  function startEdit(entry) { setForm({ ...EMPTY_FORM, ...entry }); setEditingId(entry.id); setTab("new"); }
  async function deleteEntry(id) {
    try {
      await deleteEntryById(id);
      setEntries((prev) => prev.filter((e) => e.id !== id));
      showToast("ลบรายการแล้ว");
    } catch (err) { showToast("ลบไม่สำเร็จ: " + err.message); }
  }

  // ---------- admin: user management ----------
  async function approveUser(salesId) {
    try {
      await updateUserFields(salesId, { status: "approved" });
      setUsers((prev) => prev.map((u) => (u.salesId === salesId ? { ...u, status: "approved" } : u)));
      showToast(`อนุมัติบัญชี ${salesId} แล้ว`);
    } catch (err) { showToast("อนุมัติไม่สำเร็จ: " + err.message); }
  }
  async function rejectUser(salesId) {
    try {
      await updateUserFields(salesId, { status: "rejected" });
      setUsers((prev) => prev.map((u) => (u.salesId === salesId ? { ...u, status: "rejected" } : u)));
      showToast(`ระงับบัญชี ${salesId} แล้ว`);
    } catch (err) { showToast("ดำเนินการไม่สำเร็จ: " + err.message); }
  }
  async function reactivateUser(salesId) {
    try {
      await updateUserFields(salesId, { status: "approved" });
      setUsers((prev) => prev.map((u) => (u.salesId === salesId ? { ...u, status: "approved" } : u)));
      showToast(`เปิดใช้งานบัญชี ${salesId} แล้ว`);
    } catch (err) { showToast("ดำเนินการไม่สำเร็จ: " + err.message); }
  }
  async function toggleAdmin(salesId) {
    const target = users.find((u) => u.salesId === salesId);
    if (!target) return;
    const newRole = target.role === "admin" ? "sales" : "admin";
    try {
      await updateUserFields(salesId, { role: newRole });
      setUsers((prev) => prev.map((u) => (u.salesId === salesId ? { ...u, role: newRole } : u)));
      showToast(`อัปเดตสิทธิ์ผู้ใช้ ${salesId} แล้ว`);
    } catch (err) { showToast("ดำเนินการไม่สำเร็จ: " + err.message); }
  }
  async function resetPassword(salesId) {
    const newPw = window.prompt(`ตั้งรหัสผ่านใหม่ให้ ${salesId} (อย่างน้อย 4 ตัวอักษร):`);
    if (newPw === null) return; // cancelled
    if (newPw.trim().length < 4) { showToast("รหัสผ่านสั้นเกินไป (อย่างน้อย 4 ตัวอักษร)"); return; }
    try {
      await updateUserFields(salesId, { password: newPw.trim() });
      setUsers((prev) => prev.map((u) => (u.salesId === salesId ? { ...u, password: newPw.trim() } : u)));
      showToast(`ตั้งรหัสผ่านใหม่ให้ ${salesId} แล้ว แจ้งพนักงานให้ทราบรหัสใหม่`);
    } catch (err) { showToast("รีเซ็ตรหัสผ่านไม่สำเร็จ: " + err.message); }
  }

  // ---------- import / export ----------
  function exportCsv(list, filename, cols = ENTRY_COLS) {
    const rows = [cols.join(",")].concat(
      list.map((e) => cols.map((c) => `"${String(e[c] ?? "").replace(/"/g, '""')}"`).join(","))
    );
    const blob = new Blob(["\uFEFF" + rows.join("\n")], { type: "text/csv;charset=utf-8;" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url; a.download = filename;
    document.body.appendChild(a); a.click(); document.body.removeChild(a);
    URL.revokeObjectURL(url);
  }

  function exportUsersCsv() {
    const cols = ["salesId", "name", "zone", "role", "status", "createdAt"];
    exportCsv(users, "osram_users_backup.csv", cols);
  }

  function exportFullBackupJson() {
    const payload = { users, entries, exportedAt: new Date().toISOString() };
    const blob = new Blob([JSON.stringify(payload, null, 2)], { type: "application/json" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url; a.download = `osram_pipeline_backup_${Date.now()}.json`;
    document.body.appendChild(a); a.click(); document.body.removeChild(a);
    URL.revokeObjectURL(url);
  }

  function parseCsvFile(file) {
    return new Promise((resolve, reject) => {
      Papa.parse(file, {
        header: true,
        skipEmptyLines: true,
        complete: (results) => resolve(results.data),
        error: reject,
      });
    });
  }
  function parseExcelFile(file) {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = (e) => {
        try {
          const workbook = XLSX.read(new Uint8Array(e.target.result), { type: "array", cellDates: true });
          const sheet = workbook.Sheets[workbook.SheetNames[0]];
          resolve(XLSX.utils.sheet_to_json(sheet, { defval: "" }));
        } catch (err) { reject(err); }
      };
      reader.onerror = () => reject(new Error("อ่านไฟล์ไม่สำเร็จ"));
      reader.readAsArrayBuffer(file);
    });
  }

  function entrySignature(e) {
    // A row "looks the same" if these key fields match — used to skip re-importing
    // rows that are already in the system (e.g. the same file imported twice).
    return [e.salesId, e.customerName, e.itemDescription, e.qty, e.price, e.visitDate]
      .map((v) => String(v ?? "").trim().toLowerCase())
      .join("|");
  }

  async function handleImportFile(file) {
    if (!file || importing) return; // guard against double-click while an import is already running
    setImporting(true);
    const isExcel = /\.(xlsx|xls)$/i.test(file.name);
    const userBySalesId = new Map(users.map((u) => [u.salesId.toLowerCase(), u]));
    let unknownCount = 0;
    try {
      const rows = isExcel ? await parseExcelFile(file) : await parseCsvFile(file);
      if (!rows.length) { showToast("ไม่พบข้อมูลในไฟล์"); return; }
      const prepared = rows
        .map((r) => {
          const entry = mapRowToEntry(r);
          // Never trust salesName straight from the file — old exports or the original
          // Excel report may contain sales codes tied to names that are outdated or don't
          // match any registered account. Resolve the name against real registered users;
          // if the code isn't registered (or missing), mark the row "Unknown" instead of
          // silently attributing it to whoever happens to be doing the import.
          const matched = entry.salesId ? userBySalesId.get(entry.salesId.toLowerCase()) : null;
          if (matched) {
            entry.salesId = matched.salesId;
            entry.salesName = matched.name;
          } else {
            entry.salesId = entry.salesId || "unknown";
            entry.salesName = "Unknown";
            unknownCount++;
          }
          if (!entry.grade) entry.grade = "C";
          if (!entry.progress) entry.progress = "0% - ยังไม่ได้เริ่มงาน";
          if (!entry.zone) entry.zone = "Red Zone";
          return entry;
        })
        .filter((e) => e.customerName);
      if (!prepared.length) {
        showToast("ไม่พบแถวที่มีชื่อลูกค้า ตรวจสอบหัวคอลัมน์ในไฟล์อีกครั้ง");
        return;
      }

      // Skip rows that already exist (same file imported twice, or overlapping with
      // rows entered manually) — compared against currently-loaded entries plus
      // duplicates within this same file.
      const existingSignatures = new Set(entries.map(entrySignature));
      const seenInBatch = new Set();
      let duplicateCount = 0;
      const toInsert = prepared.filter((e) => {
        const sig = entrySignature(e);
        if (existingSignatures.has(sig) || seenInBatch.has(sig)) { duplicateCount++; return false; }
        seenInBatch.add(sig);
        return true;
      });

      if (!toInsert.length) {
        showToast("ทุกรายการในไฟล์นี้มีอยู่ในระบบแล้ว ไม่มีการเพิ่มข้อมูลซ้ำ");
        return;
      }

      const inserted = await bulkInsertEntries(toInsert);
      setEntries((prev) => [...inserted, ...prev]);

      const parts = [`นำเข้าข้อมูลสำเร็จ ${inserted.length} รายการ`];
      if (duplicateCount > 0) parts.push(`ข้าม ${duplicateCount} รายการที่ซ้ำกับข้อมูลเดิม`);
      if (unknownCount > 0) parts.push(`${unknownCount} รายการรหัสเซลไม่ตรง/ไม่ได้ลงทะเบียน แสดงเป็น "Unknown"`);
      showToast(parts.join(" — "));
    } catch (err) {
      showToast("นำเข้าข้อมูลไม่สำเร็จ: " + err.message);
    } finally {
      setImporting(false);
    }
  }

  function handleRestoreBackup(file) {
    if (!file) return;
    const reader = new FileReader();
    reader.onload = async (e) => {
      try {
        const payload = JSON.parse(e.target.result);
        if (!Array.isArray(payload.users) || !Array.isArray(payload.entries)) throw new Error("invalid");
        if (!window.confirm(`กู้คืนข้อมูลจากไฟล์สำรอง จะแทนที่ผู้ใช้ ${payload.users.length} คน และรายการ ${payload.entries.length} รายการทั้งหมด ยืนยันหรือไม่?`)) return;
        await replaceAllUsers(payload.users);
        await replaceAllEntries(payload.entries);
        const [u, en] = await Promise.all([fetchUsers(), fetchEntries()]);
        setUsers(u); setEntries(en);
        showToast("กู้คืนข้อมูลจากไฟล์สำรองเรียบร้อย");
      } catch (err) {
        showToast("ไฟล์สำรองไม่ถูกต้อง หรือกู้คืนไม่สำเร็จ: " + err.message);
      }
    };
    reader.readAsText(file);
  }

  // ---------- derived ----------
  const myEntries = useMemo(() => entries.filter((e) => currentUser && e.salesId === currentUser.salesId), [entries, currentUser]);
  const pendingUsers = useMemo(() => users.filter((u) => u.status === "pending"), [users]);
  const isAdmin = currentUser?.role === "admin";

  function applyFilters(list) {
    return list.filter((e) => {
      if (filterZone && e.zone !== filterZone) return false;
      if (filterGrade && e.grade !== filterGrade) return false;
      if (filterRep && e.salesId !== filterRep) return false;
      if (search) {
        const s = search.toLowerCase();
        const hay = `${e.customerName} ${e.itemDescription} ${e.productType} ${e.salesName}`.toLowerCase();
        if (!hay.includes(s)) return false;
      }
      return true;
    });
  }
  const visibleMine = applyFilters(myEntries);
  const visibleTeam = applyFilters(entries);

  const teamStats = useMemo(() => {
    const byGrade = { A: 0, B: 0, C: 0 };
    const byProgress = {};
    const byZone = {};
    let totalValue = 0;
    entries.forEach((e) => {
      if (byGrade[e.grade] !== undefined) byGrade[e.grade]++;
      const key = `${progressPercent(e.progress)}%`;
      byProgress[key] = (byProgress[key] || 0) + 1;
      byZone[e.zone] = (byZone[e.zone] || 0) + 1;
      totalValue += dealValue(e);
    });
    const gradeData = Object.entries(byGrade).map(([k, v]) => ({ name: k, value: v }));
    const progressData = [0, 20, 40, 90, 100].map((p) => ({ name: `${p}%`, value: byProgress[`${p}%`] || 0 }));
    const zoneData = Object.entries(byZone).map(([k, v]) => ({ name: k, value: v }));

    const repMap = {};
    entries.forEach((e) => {
      if (!repMap[e.salesId]) repMap[e.salesId] = { salesId: e.salesId, salesName: e.salesName, deals: 0, value: 0, won: 0 };
      repMap[e.salesId].deals++;
      repMap[e.salesId].value += dealValue(e);
      if (progressPercent(e.progress) >= 90 && e.progress.includes("จบงาน") === false) repMap[e.salesId].won++;
      if (e.progress.includes("จบงาน - ได้งาน")) repMap[e.salesId].won++;
    });
    const leaderboard = Object.values(repMap).sort((a, b) => b.value - a.value);

    return { gradeData, progressData, zoneData, leaderboard, total: entries.length, reps: new Set(entries.map((e) => e.salesId)).size, totalValue };
  }, [entries]);

  // ================= RENDER =================
  if (!ready) {
    return (
      <div style={{ minHeight: "100vh", display: "flex", alignItems: "center", justifyContent: "center", background: INK }}>
        <div style={{ color: "#fff", fontFamily: "system-ui" }}>กำลังโหลด...</div>
      </div>
    );
  }

  if (!currentUser) {
    return (
      <div style={{
        minHeight: "100vh", background: `radial-gradient(1200px 600px at 80% -10%, #2a2a30, ${INK} 60%)`,
        display: "flex", alignItems: "center", justifyContent: "center", fontFamily: "system-ui, -apple-system, 'Segoe UI', sans-serif", padding: 20,
      }}>
        <div style={{ width: "100%", maxWidth: 420 }}>
          <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 22, justifyContent: "center" }}>
            <div style={{ width: 40, height: 40, borderRadius: 10, background: RED, display: "flex", alignItems: "center", justifyContent: "center" }}>
              <Lightbulb size={22} color={AMBER} strokeWidth={2.3} />
            </div>
            <div>
              <div style={{ color: "#fff", fontWeight: 800, fontSize: 20, letterSpacing: "-0.01em" }}>OSRAM</div>
              <div style={{ color: "#B9B9C0", fontSize: 11, letterSpacing: "0.12em", textTransform: "uppercase" }}>Sales Pipeline</div>
            </div>
          </div>

          <div style={{ background: "#fff", borderRadius: 16, padding: 28, boxShadow: "0 30px 60px -20px rgba(0,0,0,0.5)" }}>
            <div style={{ display: "flex", gap: 6, marginBottom: 20, background: "#F1F0EC", borderRadius: 10, padding: 4 }}>
              <button onClick={() => { setAuthMode("login"); setAuthError(""); setAuthNotice(""); setShowForgotHint(false); }} style={{
                flex: 1, padding: "9px 0", borderRadius: 8, border: "none", cursor: "pointer", fontWeight: 700, fontSize: 13,
                background: authMode === "login" ? "#fff" : "transparent", color: authMode === "login" ? INK : "#8A8A90",
                boxShadow: authMode === "login" ? "0 1px 4px rgba(0,0,0,0.12)" : "none",
              }}>เข้าสู่ระบบ</button>
              <button onClick={() => { setAuthMode("register"); setAuthError(""); setAuthNotice(""); setShowForgotHint(false); }} style={{
                flex: 1, padding: "9px 0", borderRadius: 8, border: "none", cursor: "pointer", fontWeight: 700, fontSize: 13,
                background: authMode === "register" ? "#fff" : "transparent", color: authMode === "register" ? INK : "#8A8A90",
                boxShadow: authMode === "register" ? "0 1px 4px rgba(0,0,0,0.12)" : "none",
              }}>สมัครใหม่</button>
            </div>

            <form onSubmit={handleAuthSubmit} style={{ display: "flex", flexDirection: "column", gap: 14 }}>
              {authMode === "register" && (
                <Field label="ชื่อพนักงานขาย" required>
                  <TextInput value={authForm.name} onChange={(e) => setAuthForm((f) => ({ ...f, name: e.target.value }))} placeholder="เช่น สมชาย ใจดี" />
                </Field>
              )}
              <Field label="Sales ID" required>
                <TextInput value={authForm.salesId} onChange={(e) => setAuthForm((f) => ({ ...f, salesId: e.target.value }))} placeholder="เช่น S173" />
              </Field>
              {authMode === "register" && (
                <Field label="โซนที่รับผิดชอบ">
                  <ZoneSelect
                    zones={zones}
                    value={authForm.zone}
                    onChange={(v) => setAuthForm((f) => ({ ...f, zone: v }))}
                    onAddZone={addZone}
                  />
                </Field>
              )}
              <Field label="รหัสผ่าน" required>
                <TextInput type="password" value={authForm.password} onChange={(e) => setAuthForm((f) => ({ ...f, password: e.target.value }))} placeholder="••••••••" />
              </Field>

              {authError && (
                <div style={{ display: "flex", gap: 6, alignItems: "center", color: RED, fontSize: 12.5, background: "#FDEAEA", padding: "8px 10px", borderRadius: 8 }}>
                  <AlertCircle size={14} /> {authError}
                </div>
              )}
              {authNotice && (
                <div style={{ display: "flex", gap: 6, alignItems: "center", color: "#8A6D00", fontSize: 12.5, background: "#FFF6DC", padding: "8px 10px", borderRadius: 8 }}>
                  <Clock size={14} /> {authNotice}
                </div>
              )}

              <button type="submit" style={{
                marginTop: 4, background: RED, color: "#fff", border: "none", borderRadius: 10, padding: "12px 0",
                fontWeight: 800, fontSize: 14, cursor: "pointer", display: "flex", alignItems: "center", justifyContent: "center", gap: 8,
              }}>
                {authMode === "login" ? <LogIn size={16} /> : <UserPlus size={16} />}
                {authMode === "login" ? "เข้าสู่ระบบ" : "สมัครสมาชิก"}
              </button>

              {authMode === "login" && (
                <div style={{ textAlign: "center" }}>
                  <button
                    type="button"
                    onClick={() => setShowForgotHint((v) => !v)}
                    style={{ background: "none", border: "none", color: "#6B6B70", fontSize: 12.5, cursor: "pointer", textDecoration: "underline", padding: 0 }}
                  >
                    ลืมรหัสผ่าน?
                  </button>
                  {showForgotHint && (
                    <div style={{ marginTop: 10, background: "#F1F0EC", borderRadius: 8, padding: "10px 12px", fontSize: 12, color: "#4B4B52", lineHeight: 1.7, textAlign: "left" }}>
                      ระบบนี้ยังไม่รองรับการรีเซ็ตรหัสผ่านด้วยตัวเอง กรุณาติดต่อ<b>ผู้ดูแลระบบ</b>ของทีมคุณ
                      เพื่อให้ช่วยตั้งรหัสผ่านใหม่ให้ (Admin ทำได้จากหน้า "ผู้ดูแลระบบ → จัดการผู้ใช้")
                    </div>
                  )}
                </div>
              )}
            </form>
          </div>
          <p style={{ color: "#8B8B93", fontSize: 11.5, textAlign: "center", marginTop: 16, lineHeight: 1.6 }}>
            บัญชีสมัครใหม่ต้องรอการอนุมัติจากผู้ดูแลระบบก่อนจึงจะเข้าใช้งานได้ (ผู้สมัครคนแรกของระบบจะเป็นผู้ดูแลระบบโดยอัตโนมัติ)
            ข้อมูลเก็บแบบใช้ร่วมกันทั้งทีม ไม่ควรใช้รหัสผ่านที่สำคัญ
          </p>
        </div>
      </div>
    );
  }

  const tabs = [
    { id: "new", label: editingId ? "แก้ไขรายการ" : "เพิ่มรายการใหม่", icon: Plus },
    { id: "mine", label: `รายการของฉัน (${myEntries.length})`, icon: ListChecks },
    { id: "team", label: `ภาพรวมทีม (${entries.length})`, icon: BarChart3 },
  ];
  if (isAdmin) tabs.push({ id: "admin", label: `ผู้ดูแลระบบ${pendingUsers.length ? ` (${pendingUsers.length})` : ""}`, icon: ShieldCheck });

  return (
    <div style={{ minHeight: "100vh", background: PAPER, fontFamily: "system-ui, -apple-system, 'Segoe UI', sans-serif", color: INK, overflowX: "hidden" }}>
      <div style={{ background: INK, color: "#fff", padding: "14px 16px", display: "flex", alignItems: "center", justifyContent: "space-between", gap: 10, position: "sticky", top: 0, zIndex: 20 }}>
        <div style={{ display: "flex", alignItems: "center", gap: 10, minWidth: 0 }}>
          <div style={{ width: 32, height: 32, borderRadius: 8, background: RED, display: "flex", alignItems: "center", justifyContent: "center", flexShrink: 0 }}>
            <Lightbulb size={17} color={AMBER} />
          </div>
          <div style={{ minWidth: 0 }}>
            <div style={{ fontWeight: 800, fontSize: 15, lineHeight: 1.1, display: "flex", alignItems: "center", gap: 6, whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis" }}>
              OSRAM Sales Pipeline {isAdmin && <Crown size={13} color={AMBER} style={{ flexShrink: 0 }} />}
            </div>
            <div style={{ fontSize: 11, color: "#A9A9B0", whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis" }}>{currentUser.name} · {currentUser.salesId} · {currentUser.zone}</div>
          </div>
        </div>
        <button onClick={handleLogout} style={{
          background: "transparent", border: "1px solid #3B3B42", color: "#EDEDEF", borderRadius: 8, flexShrink: 0,
          padding: "8px 12px", fontSize: 12.5, display: "flex", alignItems: "center", gap: 6, cursor: "pointer", fontWeight: 600,
        }}><LogOut size={14} /> ออกจากระบบ</button>
      </div>

      <div style={{ background: "#fff", borderBottom: "1px solid #E7E5DF", padding: "0 20px", display: "flex", gap: 4, position: "sticky", top: 60, zIndex: 19, overflowX: "auto" }}>
        {tabs.map((t) => (
          <button key={t.id} onClick={() => setTab(t.id)} style={{
            border: "none", background: "transparent", padding: "14px 10px", cursor: "pointer",
            display: "flex", alignItems: "center", gap: 6, fontSize: 13.5, fontWeight: 700,
            color: tab === t.id ? RED : "#8A8A90", borderBottom: tab === t.id ? `2px solid ${RED}` : "2px solid transparent",
            whiteSpace: "nowrap",
          }}><t.icon size={15} /> {t.label}</button>
        ))}
      </div>

      <div style={{ maxWidth: 1150, margin: "0 auto", padding: "24px 20px 80px" }}>
        {tab === "new" && (
          <form onSubmit={handleSaveEntry} style={{ background: "#fff", borderRadius: 14, padding: 24, border: "1px solid #E7E5DF" }}>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
              <h2 style={{ margin: 0, fontSize: 18 }}>{editingId ? "แก้ไขรายการ Pipeline" : "กรอกรายการ Pipeline ใหม่"}</h2>
              {editingId && (
                <button type="button" onClick={resetForm} style={{ border: "none", background: "#F1F0EC", borderRadius: 8, padding: "6px 10px", fontSize: 12, cursor: "pointer", display: "flex", gap: 6, alignItems: "center" }}>
                  <X size={13} /> ยกเลิกแก้ไข
                </button>
              )}
            </div>

            <SectionTitle icon={Filter}>ข้อมูลพื้นฐาน</SectionTitle>
            <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(180px, 1fr))", gap: 14 }}>
              <Field label="โซน (Zone)"><ZoneSelect zones={zones} value={form.zone} onChange={(v) => updateField("zone", v)} onAddZone={addZone} /></Field>
              <Field label="วันที่เยี่ยมลูกค้า (Visit Date)"><TextInput type="date" value={form.visitDate} onChange={(e) => updateField("visitDate", e.target.value)} /></Field>
              <Field label="เกรดลูกค้า (Grade)"><Select value={form.grade} onChange={(e) => updateField("grade", e.target.value)}>
                {GRADE_OPTIONS.map((g) => <option key={g}>{g}</option>)}
              </Select></Field>
            </div>

            <SectionTitle icon={Users}>ข้อมูลลูกค้า</SectionTitle>
            <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(180px, 1fr))", gap: 14 }}>
              <Field label="Customer ID"><TextInput value={form.customerId} onChange={(e) => updateField("customerId", e.target.value)} placeholder="New / รหัสลูกค้า" /></Field>
              <Field label="ชื่อลูกค้า (Customer Name)" required><TextInput value={form.customerName} onChange={(e) => updateField("customerName", e.target.value)} placeholder="ชื่อร้าน/บริษัท" /></Field>
              <Field label="ประเภทลูกค้า (Customer Type)"><Select value={form.customerType} onChange={(e) => updateField("customerType", e.target.value)}>
                {CUSTOMER_TYPE_OPTIONS.map((o) => <option key={o}>{o}</option>)}
              </Select></Field>
              <Field label="กลุ่มลูกค้า (Segment)"><Select value={form.customerSegment} onChange={(e) => updateField("customerSegment", e.target.value)}>
                {SEGMENT_OPTIONS.map((o) => <option key={o}>{o}</option>)}
              </Select></Field>
            </div>

            <SectionTitle icon={Lightbulb}>ข้อมูลสินค้า</SectionTitle>
            <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(180px, 1fr))", gap: 14 }}>
              <Field label="ประเภทสินค้า (Product Type)"><TextInput value={form.productType} onChange={(e) => updateField("productType", e.target.value)} placeholder="เช่น OSRAM LED" /></Field>
              <Field label="รหัสสินค้า (Item Code)"><TextInput value={form.itemCode} onChange={(e) => updateField("itemCode", e.target.value)} /></Field>
              <Field label="รายละเอียดสินค้า (Item Description)"><TextInput value={form.itemDescription} onChange={(e) => updateField("itemDescription", e.target.value)} /></Field>
              <Field label="จำนวน (QTY)"><TextInput type="number" value={form.qty} onChange={(e) => updateField("qty", e.target.value)} /></Field>
              <Field label="หน่วย (Uom)"><TextInput value={form.uom} onChange={(e) => updateField("uom", e.target.value)} placeholder="ชิ้น / กล่อง / เซต" /></Field>
              <Field label="ราคา (Price)"><TextInput type="number" value={form.price} onChange={(e) => updateField("price", e.target.value)} /></Field>
            </div>

            <SectionTitle icon={AlertCircle}>ข้อมูลคู่แข่ง</SectionTitle>
            <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(180px, 1fr))", gap: 14 }}>
              <Field label="ชื่อคู่แข่ง (Competitor Name)"><TextInput value={form.competitorName} onChange={(e) => updateField("competitorName", e.target.value)} /></Field>
              <Field label="ราคาคู่แข่ง (Competitor Price)"><TextInput type="number" value={form.competitorPrice} onChange={(e) => updateField("competitorPrice", e.target.value)} /></Field>
            </div>

            <SectionTitle icon={TrendingUp}>แผนงานและกำหนดการ</SectionTitle>
            <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(180px, 1fr))", gap: 14 }}>
              <Field label="ปีที่คาดว่าจะปิดงาน (Project Close Year)"><TextInput type="number" value={form.projectCloseYear} onChange={(e) => updateField("projectCloseYear", e.target.value)} /></Field>
              <Field label="เดือนที่คาดว่าจะปิดงาน (Close Month)"><Select value={form.projectCloseMonth} onChange={(e) => updateField("projectCloseMonth", e.target.value)}>
                <option value="">- เลือกเดือน -</option>{MONTH_OPTIONS.map((m) => <option key={m}>{m}</option>)}
              </Select></Field>
              <Field label="เลขที่ใบเสนอราคา (Quotation Number)"><TextInput value={form.quotationNumber} onChange={(e) => updateField("quotationNumber", e.target.value)} /></Field>
              <Field label="วิธีการจัดส่ง"><Select value={form.deliveryMethod} onChange={(e) => updateField("deliveryMethod", e.target.value)}>
                {DELIVERY_OPTIONS.map((o) => <option key={o}>{o}</option>)}
              </Select></Field>
              <Field label="จำนวนเดือนที่ทยอยส่ง (90%)"><TextInput type="number" value={form.deliverInMonths} onChange={(e) => updateField("deliverInMonths", e.target.value)} /></Field>
              <Field label="ปีที่เริ่มส่งมอบ (Delivery Start Year)"><TextInput type="number" value={form.deliveryStartYear} onChange={(e) => updateField("deliveryStartYear", e.target.value)} /></Field>
              <Field label="เดือนที่เริ่มส่งมอบ (Delivery Start Month)"><Select value={form.deliveryStartMonth} onChange={(e) => updateField("deliveryStartMonth", e.target.value)}>
                <option value="">- เลือกเดือน -</option>{MONTH_OPTIONS.map((m) => <option key={m}>{m}</option>)}
              </Select></Field>
              <Field label="เดือนที่เริ่มดำเนินงาน (Start Working Month)"><Select value={form.startWorkingMonth} onChange={(e) => updateField("startWorkingMonth", e.target.value)}>
                <option value="">- เลือกเดือน -</option>{MONTH_OPTIONS.map((m) => <option key={m}>{m}</option>)}
              </Select></Field>
            </div>

            <SectionTitle icon={ListChecks}>ความคืบหน้า (KPI)</SectionTitle>
            <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(180px, 1fr))", gap: 14 }}>
              <Field label="KPI Register"><TextInput value={form.kpiRegister} onChange={(e) => updateField("kpiRegister", e.target.value)} /></Field>
              <Field label="แผนปฏิบัติงาน (Action Plan)"><TextInput value={form.actionPlanMonth} onChange={(e) => updateField("actionPlanMonth", e.target.value)} placeholder="เช่น แนะนำตัว + เสนอสินค้า" /></Field>
              <Field label="% ความคืบหน้า (Progression)" required><Select value={form.progress} onChange={(e) => updateField("progress", e.target.value)}>
                {PROGRESS_OPTIONS.map((o) => <option key={o}>{o}</option>)}
              </Select></Field>
            </div>

            <div style={{ display: "flex", gap: 10, marginTop: 26 }}>
              <button type="submit" style={{ background: RED, color: "#fff", border: "none", borderRadius: 10, padding: "12px 22px", fontWeight: 800, fontSize: 14, cursor: "pointer", display: "flex", alignItems: "center", gap: 8 }}>
                <Save size={16} /> {editingId ? "บันทึกการแก้ไข" : "บันทึกรายการ"}
              </button>
              <button type="button" onClick={resetForm} style={{ background: "#F1F0EC", color: INK, border: "none", borderRadius: 10, padding: "12px 18px", fontWeight: 700, fontSize: 14, cursor: "pointer" }}>ล้างฟอร์ม</button>
            </div>
          </form>
        )}

        {(tab === "mine" || tab === "team") && (
          <div>
            <div style={{ display: "flex", gap: 10, flexWrap: "wrap", marginBottom: 16, alignItems: "center" }}>
              <div style={{ position: "relative", flex: "1 1 220px" }}>
                <Search size={15} style={{ position: "absolute", left: 10, top: 11, color: "#9A9A9F" }} />
                <TextInput value={search} onChange={(e) => setSearch(e.target.value)} placeholder="ค้นหาลูกค้า / สินค้า / พนักงาน" style={{ paddingLeft: 32 }} />
              </div>
              <Select value={filterZone} onChange={(e) => setFilterZone(e.target.value)} style={{ maxWidth: 160 }}>
                <option value="">ทุกโซน</option>{zones.map((z) => <option key={z}>{z}</option>)}
              </Select>
              <Select value={filterGrade} onChange={(e) => setFilterGrade(e.target.value)} style={{ maxWidth: 130 }}>
                <option value="">ทุกเกรด</option>{GRADE_OPTIONS.map((g) => <option key={g}>{g}</option>)}
              </Select>
              {tab === "team" && (
                <Select value={filterRep} onChange={(e) => setFilterRep(e.target.value)} style={{ maxWidth: 180 }}>
                  <option value="">ทุกพนักงานขาย</option>{users.map((u) => <option key={u.salesId} value={u.salesId}>{u.name} ({u.salesId})</option>)}
                </Select>
              )}
              <button onClick={() => exportCsv(tab === "mine" ? visibleMine : visibleTeam, tab === "mine" ? "my_pipeline.csv" : "team_pipeline.csv")}
                style={{ border: "1px solid #E7E5DF", background: "#fff", borderRadius: 8, padding: "9px 12px", fontSize: 12.5, display: "flex", alignItems: "center", gap: 6, cursor: "pointer", fontWeight: 700 }}>
                <Download size={14} /> Export CSV
              </button>
            </div>

            {tab === "team" && (
              <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(240px,1fr))", gap: 14, marginBottom: 20 }}>
                <StatCard label="รายการทั้งหมด" value={teamStats.total} sub={`${teamStats.reps} พนักงานขาย · มูลค่ารวม ฿${money(teamStats.totalValue)}`} icon={Database} />
                <div style={{ background: "#fff", border: "1px solid #E7E5DF", borderRadius: 14, padding: 16 }}>
                  <div style={{ fontSize: 12, fontWeight: 700, color: "#8A8A90", marginBottom: 8 }}>สัดส่วนตามเกรดลูกค้า</div>
                  <ResponsiveContainer width="100%" height={140}>
                    <PieChart>
                      <Pie data={teamStats.gradeData} dataKey="value" nameKey="name" innerRadius={32} outerRadius={56} paddingAngle={2}>
                        {teamStats.gradeData.map((d) => <Cell key={d.name} fill={gradeColor(d.name)} />)}
                      </Pie>
                      <Legend iconSize={8} wrapperStyle={{ fontSize: 11 }} /><Tooltip />
                    </PieChart>
                  </ResponsiveContainer>
                </div>
                <div style={{ background: "#fff", border: "1px solid #E7E5DF", borderRadius: 14, padding: 16 }}>
                  <div style={{ fontSize: 12, fontWeight: 700, color: "#8A8A90", marginBottom: 8 }}>จำนวนดีลตามความคืบหน้า</div>
                  <ResponsiveContainer width="100%" height={140}>
                    <BarChart data={teamStats.progressData}>
                      <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#EDEBE5" />
                      <XAxis dataKey="name" tick={{ fontSize: 11 }} /><YAxis tick={{ fontSize: 11 }} allowDecimals={false} />
                      <Tooltip /><Bar dataKey="value" fill={RED} radius={[4,4,0,0]} />
                    </BarChart>
                  </ResponsiveContainer>
                </div>
              </div>
            )}

            <div style={{ background: "#fff", border: "1px solid #E7E5DF", borderRadius: 14, overflow: "hidden" }}>
              <div style={{ overflowX: "auto" }}>
                <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
                  <thead>
                    <tr style={{ background: "#FAFAF8", textAlign: "left" }}>
                      {tab === "team" && <Th>พนักงานขาย</Th>}
                      <Th>โซน</Th><Th>ลูกค้า</Th><Th>สินค้า</Th><Th>มูลค่า</Th><Th>ความคืบหน้า</Th><Th>เกรด</Th><Th>วันที่เยี่ยม</Th><Th></Th>
                    </tr>
                  </thead>
                  <tbody>
                    {(tab === "mine" ? visibleMine : visibleTeam).length === 0 && (
                      <tr><td colSpan={9} style={{ padding: 28, textAlign: "center", color: "#9A9A9F" }}>ยังไม่มีรายการ — เริ่มเพิ่มรายการ Pipeline แรกของคุณ</td></tr>
                    )}
                    {(tab === "mine" ? visibleMine : visibleTeam).map((en) => (
                      <tr key={en.id} style={{ borderTop: "1px solid #F0EFEA" }}>
                        {tab === "team" && <Td>{en.salesName} <span style={{ color: "#9A9A9F" }}>({en.salesId})</span></Td>}
                        <Td>{en.zone}</Td>
                        <Td style={{ maxWidth: 220 }}>
                          <div style={{ fontWeight: 700 }}>{en.customerName}</div>
                          <div style={{ fontSize: 11.5, color: "#9A9A9F" }}>{en.customerSegment} · {en.customerType}</div>
                        </Td>
                        <Td style={{ maxWidth: 200 }}>
                          <div>{en.productType || "-"}</div>
                          <div style={{ fontSize: 11.5, color: "#9A9A9F" }}>{en.itemDescription}</div>
                        </Td>
                        <Td>{dealValue(en) ? `฿${money(dealValue(en))}` : "-"}</Td>
                        <Td>
                          <div style={{ display: "flex", alignItems: "center", gap: 6 }}>
                            <div style={{ width: 60, height: 6, borderRadius: 4, background: "#EEECE6", overflow: "hidden" }}>
                              <div style={{ width: `${progressPercent(en.progress)}%`, height: "100%", background: RED }} />
                            </div>
                            <span style={{ fontSize: 11 }}>{progressPercent(en.progress)}%</span>
                          </div>
                        </Td>
                        <Td><span style={{ display: "inline-flex", width: 22, height: 22, borderRadius: 999, alignItems: "center", justifyContent: "center", background: gradeColor(en.grade) + "22", color: gradeColor(en.grade), fontWeight: 800, fontSize: 11 }}>{en.grade}</span></Td>
                        <Td>{en.visitDate || "-"}</Td>
                        <Td>
                          {(tab === "mine" || en.salesId === currentUser.salesId || isAdmin) && (
                            <div style={{ display: "flex", gap: 6 }}>
                              <button onClick={() => startEdit(en)} title="แก้ไข" style={{ border: "none", background: "#F1F0EC", borderRadius: 6, padding: 6, cursor: "pointer" }}><Edit3 size={13} /></button>
                              <button onClick={() => deleteEntry(en.id)} title="ลบ" style={{ border: "none", background: "#FDEAEA", color: RED, borderRadius: 6, padding: 6, cursor: "pointer" }}><Trash2 size={13} /></button>
                            </div>
                          )}
                        </Td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>
          </div>
        )}

        {tab === "admin" && isAdmin && (
          <div>
            <div style={{ display: "flex", gap: 4, marginBottom: 18, background: "#fff", border: "1px solid #E7E5DF", borderRadius: 10, padding: 4, maxWidth: "100%", overflowX: "auto", WebkitOverflowScrolling: "touch" }}>
              {[
                { id: "pending", label: `รออนุมัติ (${pendingUsers.length})`, icon: Clock },
                { id: "users", label: "จัดการผู้ใช้", icon: Users },
                { id: "reports", label: "รายงาน", icon: BarChart3 },
                { id: "data", label: "นำเข้า/ส่งออกข้อมูล", icon: Database },
              ].map((s) => (
                <button key={s.id} onClick={() => setAdminSubTab(s.id)} style={{
                  border: "none", padding: "8px 14px", borderRadius: 7, cursor: "pointer", fontSize: 12.5, fontWeight: 700,
                  display: "flex", alignItems: "center", gap: 6, whiteSpace: "nowrap", flexShrink: 0,
                  background: adminSubTab === s.id ? INK : "transparent", color: adminSubTab === s.id ? "#fff" : "#6B6B70",
                }}><s.icon size={13} /> {s.label}</button>
              ))}
            </div>

            {adminSubTab === "pending" && (
              <div style={{ background: "#fff", border: "1px solid #E7E5DF", borderRadius: 14, overflow: "hidden" }}>
                {pendingUsers.length === 0 ? (
                  <div style={{ padding: 40, textAlign: "center", color: "#9A9A9F" }}>
                    <CheckCircle2 size={26} style={{ marginBottom: 8 }} />
                    <div>ไม่มีบัญชีที่รอการอนุมัติ</div>
                  </div>
                ) : (
                  <div style={{ overflowX: "auto" }}>
                  <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
                    <thead><tr style={{ background: "#FAFAF8", textAlign: "left" }}>
                      <Th>ชื่อ</Th><Th>Sales ID</Th><Th>โซน</Th><Th>วันที่สมัคร</Th><Th></Th>
                    </tr></thead>
                    <tbody>
                      {pendingUsers.map((u) => (
                        <tr key={u.salesId} style={{ borderTop: "1px solid #F0EFEA" }}>
                          <Td style={{ fontWeight: 700 }}>{u.name}</Td>
                          <Td>{u.salesId}</Td>
                          <Td>{u.zone}</Td>
                          <Td>{u.createdAt ? new Date(u.createdAt).toLocaleDateString("th-TH") : "-"}</Td>
                          <Td>
                            <div style={{ display: "flex", gap: 6 }}>
                              <button onClick={() => approveUser(u.salesId)} style={{ border: "none", background: "#E7F6EC", color: GREEN, borderRadius: 6, padding: "6px 10px", cursor: "pointer", fontWeight: 700, fontSize: 12, display: "flex", alignItems: "center", gap: 5 }}>
                                <UserCheck size={13} /> อนุมัติ
                              </button>
                              <button onClick={() => rejectUser(u.salesId)} style={{ border: "none", background: "#FDEAEA", color: RED, borderRadius: 6, padding: "6px 10px", cursor: "pointer", fontWeight: 700, fontSize: 12, display: "flex", alignItems: "center", gap: 5 }}>
                                <UserX size={13} /> ปฏิเสธ
                              </button>
                            </div>
                          </Td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                  </div>
                )}
              </div>
            )}

            {adminSubTab === "users" && (
              <div style={{ background: "#fff", border: "1px solid #E7E5DF", borderRadius: 14, overflow: "hidden" }}>
                <div style={{ overflowX: "auto" }}>
                <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
                  <thead><tr style={{ background: "#FAFAF8", textAlign: "left" }}>
                    <Th>ชื่อ</Th><Th>Sales ID</Th><Th>โซน</Th><Th>สิทธิ์</Th><Th>สถานะ</Th><Th></Th>
                  </tr></thead>
                  <tbody>
                    {users.map((u) => (
                      <tr key={u.salesId} style={{ borderTop: "1px solid #F0EFEA" }}>
                        <Td style={{ fontWeight: 700 }}>{u.name}</Td>
                        <Td>{u.salesId}</Td>
                        <Td>{u.zone}</Td>
                        <Td>{u.role === "admin" ? <Badge color="#8A6D00" bg="#FFF6DC"><Crown size={11}/> ผู้ดูแลระบบ</Badge> : <Badge color="#4B4B52" bg="#F1F0EC">พนักงานขาย</Badge>}</Td>
                        <Td>
                          {u.status === "approved" && <Badge color={GREEN} bg="#E7F6EC"><CheckCircle2 size={11}/> อนุมัติแล้ว</Badge>}
                          {u.status === "pending" && <Badge color="#8A6D00" bg="#FFF6DC"><Clock size={11}/> รออนุมัติ</Badge>}
                          {u.status === "rejected" && <Badge color={RED} bg="#FDEAEA"><ShieldAlert size={11}/> ระงับใช้งาน</Badge>}
                        </Td>
                        <Td>
                          <div style={{ display: "flex", gap: 6, flexWrap: "wrap" }}>
                            {u.status !== "approved" && (
                              <button onClick={() => reactivateUser(u.salesId)} style={{ border: "1px solid #E7E5DF", background: "#fff", borderRadius: 6, padding: "5px 9px", cursor: "pointer", fontSize: 11.5, fontWeight: 700 }}>เปิดใช้งาน</button>
                            )}
                            {u.status === "approved" && u.salesId !== currentUser.salesId && (
                              <button onClick={() => rejectUser(u.salesId)} style={{ border: "1px solid #E7E5DF", background: "#fff", color: RED, borderRadius: 6, padding: "5px 9px", cursor: "pointer", fontSize: 11.5, fontWeight: 700 }}>ระงับ</button>
                            )}
                            {u.salesId !== currentUser.salesId && (
                              <button onClick={() => toggleAdmin(u.salesId)} style={{ border: "1px solid #E7E5DF", background: "#fff", borderRadius: 6, padding: "5px 9px", cursor: "pointer", fontSize: 11.5, fontWeight: 700 }}>
                                {u.role === "admin" ? "ถอดสิทธิ์แอดมิน" : "ตั้งเป็นแอดมิน"}
                              </button>
                            )}
                            <button onClick={() => resetPassword(u.salesId)} style={{ border: "1px solid #E7E5DF", background: "#fff", borderRadius: 6, padding: "5px 9px", cursor: "pointer", fontSize: 11.5, fontWeight: 700 }}>
                              รีเซ็ตรหัสผ่าน
                            </button>
                          </div>
                        </Td>
                      </tr>
                    ))}
                  </tbody>
                </table>
                </div>
              </div>
            )}

            {adminSubTab === "reports" && (
              <div style={{ display: "flex", flexDirection: "column", gap: 16 }}>
                <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(220px,1fr))", gap: 14 }}>
                  <StatCard label="รายการ Pipeline ทั้งหมด" value={teamStats.total} sub={`${teamStats.reps} พนักงานขาย`} icon={Database} />
                  <StatCard label="มูลค่ารวมทั้งหมด" value={`฿${money(teamStats.totalValue)}`} sub="QTY × Price ทุกรายการ" icon={TrendingUp} />
                  <StatCard label="บัญชีผู้ใช้ทั้งหมด" value={users.length} sub={`${pendingUsers.length} รออนุมัติ`} icon={Users} />
                  <StatCard label="Grade A" value={teamStats.gradeData.find(g=>g.name==="A")?.value || 0} sub="ลูกค้าเกรดสูงสุด" icon={CheckCircle2} />
                </div>

                <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(240px,1fr))", gap: 14 }}>
                  <div style={{ background: "#fff", border: "1px solid #E7E5DF", borderRadius: 14, padding: 16 }}>
                    <div style={{ fontSize: 12, fontWeight: 700, color: "#8A8A90", marginBottom: 8 }}>จำนวนรายการตามโซน</div>
                    <ResponsiveContainer width="100%" height={200}>
                      <BarChart data={teamStats.zoneData} layout="vertical">
                        <CartesianGrid strokeDasharray="3 3" horizontal={false} stroke="#EDEBE5" />
                        <XAxis type="number" tick={{ fontSize: 11 }} allowDecimals={false} />
                        <YAxis type="category" dataKey="name" tick={{ fontSize: 11 }} width={80} />
                        <Tooltip /><Bar dataKey="value" fill={RED} radius={[0,4,4,0]} />
                      </BarChart>
                    </ResponsiveContainer>
                  </div>
                  <div style={{ background: "#fff", border: "1px solid #E7E5DF", borderRadius: 14, padding: 16 }}>
                    <div style={{ fontSize: 12, fontWeight: 700, color: "#8A8A90", marginBottom: 8 }}>สัดส่วนตามเกรดลูกค้า</div>
                    <ResponsiveContainer width="100%" height={200}>
                      <PieChart>
                        <Pie data={teamStats.gradeData} dataKey="value" nameKey="name" innerRadius={40} outerRadius={72} paddingAngle={2}>
                          {teamStats.gradeData.map((d) => <Cell key={d.name} fill={gradeColor(d.name)} />)}
                        </Pie>
                        <Legend iconSize={8} wrapperStyle={{ fontSize: 11 }} /><Tooltip />
                      </PieChart>
                    </ResponsiveContainer>
                  </div>
                </div>

                <div style={{ background: "#fff", border: "1px solid #E7E5DF", borderRadius: 14, padding: 16 }}>
                  <div style={{ fontSize: 12, fontWeight: 700, color: "#8A8A90", marginBottom: 8 }}>อันดับพนักงานขาย (ตามมูลค่า Pipeline)</div>
                  <div style={{ overflowX: "auto" }}>
                  <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
                    <thead><tr style={{ textAlign: "left" }}><Th>#</Th><Th>พนักงานขาย</Th><Th>จำนวนดีล</Th><Th>มูลค่ารวม</Th></tr></thead>
                    <tbody>
                      {teamStats.leaderboard.length === 0 && <tr><td colSpan={4} style={{ padding: 20, textAlign: "center", color: "#9A9A9F" }}>ยังไม่มีข้อมูล</td></tr>}
                      {teamStats.leaderboard.map((r, i) => (
                        <tr key={r.salesId} style={{ borderTop: "1px solid #F0EFEA" }}>
                          <Td>{i + 1}{i === 0 && <Crown size={12} color={AMBER} style={{ marginLeft: 4, verticalAlign: "middle" }} />}</Td>
                          <Td style={{ fontWeight: 700 }}>{r.salesName} <span style={{ color: "#9A9A9F", fontWeight: 400 }}>({r.salesId})</span></Td>
                          <Td>{r.deals}</Td>
                          <Td>฿{money(r.value)}</Td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                  </div>
                </div>
              </div>
            )}

            {adminSubTab === "data" && (
              <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(240px,1fr))", gap: 14 }}>
                <div style={{ background: "#fff", border: "1px solid #E7E5DF", borderRadius: 14, padding: 18 }}>
                  <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 8 }}><FileDown size={16} color={RED} /><b style={{ fontSize: 13.5 }}>ส่งออกข้อมูล Pipeline</b></div>
                  <p style={{ fontSize: 12, color: "#8A8A90", margin: "0 0 12px" }}>ส่งออกรายการ Pipeline ทั้งหมดของทุกทีมเป็นไฟล์ CSV</p>
                  <button onClick={() => exportCsv(entries, "all_pipeline_data.csv")} style={{ border: "none", background: RED, color: "#fff", borderRadius: 8, padding: "9px 14px", fontSize: 12.5, fontWeight: 700, cursor: "pointer", display: "flex", alignItems: "center", gap: 6 }}>
                    <Download size={13} /> Export ทุกรายการ ({entries.length})
                  </button>
                </div>

                <div style={{ background: "#fff", border: "1px solid #E7E5DF", borderRadius: 14, padding: 18 }}>
                  <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 8 }}><FileUp size={16} color={RED} /><b style={{ fontSize: 13.5 }}>นำเข้าข้อมูล Pipeline (CSV / Excel)</b></div>
                  <p style={{ fontSize: 12, color: "#8A8A90", margin: "0 0 12px" }}>
                    รองรับไฟล์ .csv, .xlsx, .xls — ใช้ได้ทั้งไฟล์ที่ Export ออกจากระบบนี้ และไฟล์ Excel pipeline
                    ต้นฉบับ (ระบบจะจับคู่หัวคอลัมน์ให้อัตโนมัติ) อ่านเฉพาะชีตแรกในไฟล์
                  </p>
                  <input
                    ref={importFileRef} type="file" accept=".csv,.xlsx,.xls" style={{ display: "none" }}
                    onChange={(e) => { const f = e.target.files?.[0]; e.target.value = ""; handleImportFile(f); }}
                  />
                  <button
                    onClick={() => importFileRef.current?.click()}
                    disabled={importing}
                    style={{ border: "1px solid #DEDCD6", background: "#fff", color: INK, borderRadius: 8, padding: "9px 14px", fontSize: 12.5, fontWeight: 700, cursor: importing ? "not-allowed" : "pointer", display: "flex", alignItems: "center", gap: 6, opacity: importing ? 0.6 : 1 }}
                  >
                    <Upload size={13} /> {importing ? "กำลังนำเข้า..." : "เลือกไฟล์ CSV / Excel"}
                  </button>
                </div>

                <div style={{ background: "#fff", border: "1px solid #E7E5DF", borderRadius: 14, padding: 18 }}>
                  <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 8 }}><Users size={16} color={RED} /><b style={{ fontSize: 13.5 }}>ส่งออกรายชื่อผู้ใช้</b></div>
                  <p style={{ fontSize: 12, color: "#8A8A90", margin: "0 0 12px" }}>สำรองรายชื่อพนักงานขายและสถานะบัญชีเป็น CSV</p>
                  <button onClick={exportUsersCsv} style={{ border: "1px solid #DEDCD6", background: "#fff", color: INK, borderRadius: 8, padding: "9px 14px", fontSize: 12.5, fontWeight: 700, cursor: "pointer", display: "flex", alignItems: "center", gap: 6 }}>
                    <Download size={13} /> Export รายชื่อผู้ใช้ ({users.length})
                  </button>
                </div>

                <div style={{ background: "#fff", border: "1px solid #E7E5DF", borderRadius: 14, padding: 18 }}>
                  <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 8 }}><Database size={16} color={RED} /><b style={{ fontSize: 13.5 }}>สำรอง / กู้คืนข้อมูลทั้งหมด</b></div>
                  <p style={{ fontSize: 12, color: "#8A8A90", margin: "0 0 12px" }}>สำรองผู้ใช้และ Pipeline ทั้งหมดเป็นไฟล์ JSON เดียว หรือกู้คืนจากไฟล์สำรอง (จะแทนที่ข้อมูลปัจจุบันทั้งหมด)</p>
                  <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
                    <button onClick={exportFullBackupJson} style={{ border: "1px solid #DEDCD6", background: "#fff", color: INK, borderRadius: 8, padding: "9px 14px", fontSize: 12.5, fontWeight: 700, cursor: "pointer", display: "flex", alignItems: "center", gap: 6 }}>
                      <Download size={13} /> สำรองข้อมูล (JSON)
                    </button>
                    <input ref={backupFileRef} type="file" accept=".json" style={{ display: "none" }} onChange={(e) => handleRestoreBackup(e.target.files?.[0])} />
                    <button onClick={() => backupFileRef.current?.click()} style={{ border: "1px solid #DEDCD6", background: "#fff", color: RED, borderRadius: 8, padding: "9px 14px", fontSize: 12.5, fontWeight: 700, cursor: "pointer", display: "flex", alignItems: "center", gap: 6 }}>
                      <Upload size={13} /> กู้คืนจากไฟล์สำรอง
                    </button>
                  </div>
                </div>
              </div>
            )}
          </div>
        )}
      </div>

      {toast && (
        <div style={{
          position: "fixed", bottom: 24, left: "50%", transform: "translateX(-50%)",
          background: INK, color: "#fff", padding: "12px 18px", borderRadius: 10, fontSize: 13,
          display: "flex", alignItems: "center", gap: 8, boxShadow: "0 12px 30px rgba(0,0,0,0.3)", zIndex: 50,
        }}><CheckCircle2 size={15} color={AMBER} /> {toast}</div>
      )}
    </div>
  );
}
