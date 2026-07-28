import React, { useState, useEffect, useMemo, useCallback } from "react";
import {
  LayoutDashboard, Users, FileSignature, Wallet, Landmark, BookOpen,
  ClipboardList, Settings, Plus, X, Search, Printer, Download, Upload,
  Edit2, Trash2, ChevronRight, ChevronDown, Check, AlertCircle, Calendar,
  Building2, ArrowUpRight, ArrowDownRight, ArrowLeftRight, FileText, Save,
  RotateCcw, Phone, Briefcase, Image as ImageIcon, Loader2, TrendingUp,
  TrendingDown, ChevronLeft, UserPlus, ShieldCheck, DatabaseBackup, Eye,
  FileSpreadsheet, FileOutput,
} from "lucide-react";
import {
  BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip, ResponsiveContainer,
} from "recharts";
import * as XLSX from "xlsx";

/* ============================== SUPABASE CONFIG ============================== */
const SUPABASE_URL = "https://jkyongfvxotfsjfhcnzr.supabase.co";
const SUPABASE_ANON_KEY = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImpreW9uZ2Z2eG90ZnNqZmhjbnpyIiwicm9sZSI6ImFub24iLCJpYXQiOjE3ODQ4MDgxNjUsImV4cCI6MjEwMDM4NDE2NX0.ypS3-7M_JyDveROnm8YSM2e9IAKRfgPTvznTTgvPkAY";

/* ---- auth session (module-level; React components read/write via functions below) ---- */
let authSession = { accessToken: null, refreshToken: null, expiresAt: 0 };
let onSessionChanged = null;
function setAuthSession(s) {
  authSession = s || { accessToken: null, refreshToken: null, expiresAt: 0 };
  if (onSessionChanged) onSessionChanged(authSession);
}
function setSessionChangedHandler(fn) { onSessionChanged = fn; }

async function authRequest(path, body) {
  const res = await fetch(`${SUPABASE_URL}/auth/v1/${path}`, {
    method: "POST",
    headers: { apikey: SUPABASE_ANON_KEY, "Content-Type": "application/json" },
    body: JSON.stringify(body),
  });
  const data = await res.json().catch(() => ({}));
  if (!res.ok) throw new Error(data.error_description || data.msg || data.error || data.message || "Autentikasi gagal");
  return data;
}
const authSignUp = (email, password) => authRequest("signup", { email, password });
const authSignIn = (email, password) => authRequest("token?grant_type=password", { email, password });
const authRefresh = (refresh_token) => authRequest("token?grant_type=refresh_token", { refresh_token });
const toSession = (tok) => ({ accessToken: tok.access_token, refreshToken: tok.refresh_token, expiresAt: Date.now() + (tok.expires_in || 3600) * 1000 });

async function sb(path, options = {}, allowRetry = true) {
  const token = authSession.accessToken || SUPABASE_ANON_KEY;
  const res = await fetch(`${SUPABASE_URL}/rest/v1/${path}`, {
    ...options,
    headers: {
      apikey: SUPABASE_ANON_KEY,
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json",
      ...(options.headers || {}),
    },
  });
  if (res.status === 401 && allowRetry && authSession.refreshToken) {
    try {
      const tok = await authRefresh(authSession.refreshToken);
      setAuthSession(toSession(tok));
      return sb(path, options, false);
    } catch (e) { /* fall through */ }
  }
  if (!res.ok) {
    const text = await res.text().catch(() => "");
    let msg = text;
    try { const j = JSON.parse(text); msg = j.message || j.error_description || j.msg || text; } catch (e) { /* not json */ }
    throw new Error(msg || `${res.status}`);
  }
  const text = await res.text();
  return text ? JSON.parse(text) : null;
}
const sbSelectAll = (table) => sb(`${table}?select=*`);
const sbUpsert = (table, row, onConflict = "id") =>
  sb(`${table}?on_conflict=${onConflict}`, { method: "POST", headers: { Prefer: "resolution=merge-duplicates,return=representation" }, body: JSON.stringify(row) });
const sbInsert = (table, row) => sb(`${table}`, { method: "POST", headers: { Prefer: "return=representation" }, body: JSON.stringify(row) });
const sbDelete = (table, id) => sb(`${table}?id=eq.${encodeURIComponent(id)}`, { method: "DELETE" });
const fetchProfile = async (id) => {
  const rows = await sb(`mfs_profiles?id=eq.${id}&select=*`);
  return rows && rows[0] ? { id: rows[0].id, nama: rows[0].nama, email: rows[0].email, role: rows[0].role } : null;
};

/* ---- mapping helpers: app (camelCase) <-> supabase (snake_case) ---- */
const toCompanyDb = (c) => ({ id: 1, nama: c.nama, alamat: c.alamat, telp: c.telp, npwp: c.npwp, modal_awal: +c.modalAwal || 0, tgl_modal: c.tglModal || null });
const fromCompanyDb = (r) => ({ nama: r.nama || "Mutis Finance", alamat: r.alamat || "", telp: r.telp || "", npwp: r.npwp || "", modalAwal: r.modal_awal || 0, tglModal: r.tgl_modal || todayStr() });

const toUserDb = (u) => ({ id: u.id, nama: u.nama, role: u.role });
const fromUserDb = (r) => ({ id: r.id, nama: r.nama, role: r.role });

const toNasabahDb = (n) => ({
  id: n.id, nama: n.nama, nik: n.nik, no_kk: n.noKK, alamat: n.alamat, telp: n.telp, pekerjaan: n.pekerjaan,
  instansi: n.instansi, penghasilan: n.penghasilan ? +n.penghasilan : null, penjamin_nama: n.penjaminNama,
  penjamin_telp: n.penjaminTelp, penjamin_hubungan: n.penjaminHubungan, foto_ktp: n.fotoKTP, foto_kk: n.fotoKK, catatan: n.catatan,
});
const fromNasabahDb = (r) => ({
  id: r.id, nama: r.nama, nik: r.nik || "", noKK: r.no_kk || "", alamat: r.alamat || "", telp: r.telp || "", pekerjaan: r.pekerjaan || "",
  instansi: r.instansi || "", penghasilan: r.penghasilan || "", penjaminNama: r.penjamin_nama || "", penjaminTelp: r.penjamin_telp || "",
  penjaminHubungan: r.penjamin_hubungan || "", fotoKTP: r.foto_ktp || "", fotoKK: r.foto_kk || "", catatan: r.catatan || "", createdAt: r.created_at,
});

const toAkadDb = (a) => ({
  id: a.id, no_akad: a.noAkad, nasabah_id: a.nasabahId, nasabah_nama: a.nasabahNama, jenis_akad: a.jenisAkad,
  nilai_pembiayaan: +a.nilaiPembiayaan || 0, margin_persen: +a.marginPersen || 0, biaya_admin: +a.biayaAdmin || 0, dp: +a.dp || 0,
  tenor: +a.tenor || 0, tgl_akad: a.tglAkad, metode_pencairan: a.metodePencairan, status: a.status, catatan: a.catatan,
  total_tagihan: a.totalTagihan, cicilan_per_bulan: a.cicilanPerBulan, jadwal: a.jadwal,
});
const fromAkadDb = (r) => ({
  id: r.id, noAkad: r.no_akad, nasabahId: r.nasabah_id, nasabahNama: r.nasabah_nama, jenisAkad: r.jenis_akad,
  nilaiPembiayaan: r.nilai_pembiayaan, marginPersen: r.margin_persen, biayaAdmin: r.biaya_admin, dp: r.dp, tenor: r.tenor,
  tglAkad: r.tgl_akad, metodePencairan: r.metode_pencairan, status: r.status, catatan: r.catatan || "",
  totalTagihan: r.total_tagihan, cicilanPerBulan: r.cicilan_per_bulan, jadwal: r.jadwal || [], createdAt: r.created_at,
});

const toPembayaranDb = (p) => ({
  id: p.id, no_bukti: p.noBukti, akad_id: p.akadId, no_akad: p.noAkad, nasabah_nama: p.nasabahNama, angsuran_ke: p.angsuranKe,
  tgl_bayar: p.tglBayar, jumlah_bayar: p.jumlahBayar, denda: p.denda || 0, metode: p.metode,
  margin_admin_portion: p.marginAdminPortion || 0, margin_portion: p.marginPortion || 0, admin_portion: p.adminPortion || 0,
});
const fromPembayaranDb = (r) => ({
  id: r.id, noBukti: r.no_bukti, akadId: r.akad_id, noAkad: r.no_akad, nasabahNama: r.nasabah_nama, angsuranKe: r.angsuran_ke,
  tglBayar: r.tgl_bayar, jumlahBayar: r.jumlah_bayar, denda: r.denda || 0, metode: r.metode,
  marginAdminPortion: r.margin_admin_portion || 0, marginPortion: r.margin_portion || 0, adminPortion: r.admin_portion || 0,
});

const toKasbankDb = (k) => ({ id: k.id, no_bukti: k.noBukti, tgl: k.tgl, tipe: k.tipe, akun: k.akun, akun_tujuan: k.akunTujuan, kategori: k.kategori, jumlah: k.jumlah, keterangan: k.keterangan });
const fromKasbankDb = (r) => ({ id: r.id, noBukti: r.no_bukti, tgl: r.tgl, tipe: r.tipe, akun: r.akun, akunTujuan: r.akun_tujuan, kategori: r.kategori, jumlah: r.jumlah, keterangan: r.keterangan || "" });

/* ============================== PALETTE / TOKENS ============================== */
const C = {
  bg: "#F5F2EA",
  surface: "#FFFFFF",
  surfaceAlt: "#ECE6D6",
  ink: "#1E2521",
  inkSoft: "#5B655F",
  inkFaint: "#8B9089",
  primary: "#0F3D3A",
  primaryDark: "#0A2B29",
  primarySoft: "#DCE7E3",
  accent: "#B9863E",
  accentSoft: "#F1E3C7",
  danger: "#A8503E",
  dangerSoft: "#F3E1DC",
  success: "#3D7A5C",
  successSoft: "#DEEAE1",
  border: "#DDD5BE",
  borderSoft: "#E9E3D2",
};

const FONTS = `
@import url('https://fonts.googleapis.com/css2?family=Fraunces:ital,opsz,wght@0,9..144,400;0,9..144,600;0,9..144,700;1,9..144,500&family=Inter:wght@400;500;600;700&family=IBM+Plex+Mono:wght@400;500;600&display=swap');
`;

const JENIS_AKAD = ["Murabahah", "Qardh", "Mudharabah", "Musyarakah", "Ijarah", "Multijasa"];
const KATEGORI_PENGELUARAN = ["Beban Gaji", "Beban Listrik & Air", "Beban Sewa", "Beban ATK", "Beban Transport", "Beban Pemeliharaan", "Beban Lain-lain"];
const KATEGORI_PENERIMAAN = ["Pendapatan Lain", "Setoran Modal Tambahan", "Penerimaan Lain-lain"];
const ROLE_OPTIONS = ["Owner", "Admin", "Kasir", "Marketing", "Akunting"];

/* ============================== UTILITIES ============================== */
const uid = () => Date.now().toString(36) + Math.random().toString(36).slice(2, 8);
const todayStr = () => new Date().toISOString().slice(0, 10);
const clampNum = (v) => (isFinite(v) ? v : 0);
const fmtRp = (n) => "Rp " + Math.round(clampNum(n)).toLocaleString("id-ID");
const fmtNum = (n) => Math.round(clampNum(n)).toLocaleString("id-ID");
const fmtDate = (s) => {
  if (!s) return "-";
  const d = new Date(s + (s.length === 10 ? "T00:00:00" : ""));
  if (isNaN(d)) return s;
  return d.toLocaleDateString("id-ID", { day: "2-digit", month: "short", year: "numeric" });
};
const fmtMonthYear = (s) => {
  const d = new Date(s + "-01T00:00:00");
  return d.toLocaleDateString("id-ID", { month: "long", year: "numeric" });
};
const addMonths = (dateStr, n) => {
  const d = new Date(dateStr + "T00:00:00");
  d.setMonth(d.getMonth() + n);
  return d.toISOString().slice(0, 10);
};
const daysBetween = (a, b) => Math.round((new Date(b) - new Date(a)) / 86400000);

function generateNoAkad(jenis, existingAkad, tglAkad) {
  const code = { Murabahah: "MRB", Qardh: "QRD", Mudharabah: "MDB", Musyarakah: "MSY", Ijarah: "IJR", Multijasa: "MLJ" }[jenis] || "FIN";
  const ym = tglAkad.slice(0, 7).replace("-", "");
  const count = existingAkad.filter((a) => a.noAkad && a.noAkad.includes(ym)).length + 1;
  return `MF/${code}/${ym}/${String(count).padStart(4, "0")}`;
}

function computeJadwal(akad) {
  const nilai = clampNum(+akad.nilaiPembiayaan);
  const marginNominal = clampNum(nilai * (+akad.marginPersen / 100));
  const admin = clampNum(+akad.biayaAdmin);
  const dp = clampNum(+akad.dp);
  const tenor = Math.max(1, Math.round(+akad.tenor));
  const totalTagihan = nilai + marginNominal + admin - dp;
  const cicilanPerBulan = totalTagihan / tenor;
  const pokokPerBulan = nilai / tenor;
  const marginAdminPerBulan = (marginNominal + admin) / tenor;
  let sisa = totalTagihan;
  const jadwal = [];
  for (let i = 1; i <= tenor; i++) {
    sisa = Math.max(0, sisa - cicilanPerBulan);
    jadwal.push({
      no: i,
      tglJatuhTempo: addMonths(akad.tglAkad, i),
      pokok: pokokPerBulan,
      marginAdmin: marginAdminPerBulan,
      angsuran: cicilanPerBulan,
      sisaTagihan: sisa,
      status: "belum",
      tglBayar: null,
      denda: 0,
      pembayaranId: null,
    });
  }
  return { jadwal, totalTagihan, cicilanPerBulan, marginNominal };
}

function outstandingPokok(akad) {
  if (!akad.jadwal) return 0;
  const belumBayar = akad.jadwal.filter((j) => j.status !== "lunas");
  return belumBayar.reduce((s, j) => s + j.pokok, 0);
}

/* Derive full general journal from source data (single source of truth). */
function computeJournal({ company, akadList, pembayaranList, kasbankList }) {
  const rows = [];
  let seq = 0;
  const push = (tgl, noBukti, keterangan, akun, debit, kredit, kategoriKas) => {
    seq++;
    rows.push({ id: "j" + seq, tgl, noBukti, keterangan, akun, debit: clampNum(debit), kredit: clampNum(kredit), kategoriKas: kategoriKas || null });
  };

  if (company.modalAwal && +company.modalAwal > 0) {
    const tgl = company.tglModal || todayStr();
    push(tgl, "MODAL-AWAL", "Setoran modal awal usaha", "Kas", +company.modalAwal, 0, "pendanaan");
    push(tgl, "MODAL-AWAL", "Setoran modal awal usaha", "Modal", 0, +company.modalAwal, null);
  }

  akadList.forEach((akad) => {
    if (akad.status === "diajukan") return;
    const nilai = clampNum(+akad.nilaiPembiayaan);
    const dp = clampNum(+akad.dp);
    const marginAdmin = clampNum(akad.totalTagihan - nilai + dp);
    const akunKas = akad.metodePencairan === "Bank" ? "Bank" : "Kas";
    push(akad.tglAkad, akad.noAkad, `Pencairan pembiayaan ${akad.noAkad}`, "Piutang Pembiayaan", akad.totalTagihan, 0, null);
    push(akad.tglAkad, akad.noAkad, `Pencairan pembiayaan ${akad.noAkad}`, akunKas, 0, Math.max(0, nilai - dp), "investasi");
    if (marginAdmin > 0) push(akad.tglAkad, akad.noAkad, `Margin & admin ditangguhkan ${akad.noAkad}`, "Margin & Admin Ditangguhkan", 0, marginAdmin, null);
  });

  pembayaranList.forEach((p) => {
    const akunKas = p.metode === "Bank" ? "Bank" : "Kas";
    const total = clampNum(p.jumlahBayar) + clampNum(p.denda);
    push(p.tglBayar, p.noBukti, `Pembayaran angsuran ke-${p.angsuranKe} (${p.noAkad})`, akunKas, total, 0, "operasional");
    push(p.tglBayar, p.noBukti, `Pembayaran angsuran ke-${p.angsuranKe} (${p.noAkad})`, "Piutang Pembiayaan", 0, p.jumlahBayar, null);
    if (p.denda > 0) push(p.tglBayar, p.noBukti, `Denda keterlambatan (${p.noAkad})`, "Pendapatan Denda", 0, p.denda, null);
    if (p.marginAdminPortion > 0) {
      push(p.tglBayar, p.noBukti, `Pengakuan margin & admin (${p.noAkad})`, "Margin & Admin Ditangguhkan", p.marginAdminPortion, 0, null);
      if (p.marginPortion > 0) push(p.tglBayar, p.noBukti, `Pengakuan pendapatan margin (${p.noAkad})`, "Pendapatan Margin", 0, p.marginPortion, null);
      if (p.adminPortion > 0) push(p.tglBayar, p.noBukti, `Pengakuan pendapatan admin (${p.noAkad})`, "Pendapatan Admin", 0, p.adminPortion, null);
    }
  });

  kasbankList.forEach((k) => {
    const noBukti = k.noBukti || k.id;
    if (k.tipe === "penerimaan") {
      push(k.tgl, noBukti, k.keterangan || k.kategori, k.akun, k.jumlah, 0, "operasional");
      push(k.tgl, noBukti, k.keterangan || k.kategori, k.kategori, 0, k.jumlah, null);
    } else if (k.tipe === "pengeluaran") {
      push(k.tgl, noBukti, k.keterangan || k.kategori, k.kategori, k.jumlah, 0, null);
      push(k.tgl, noBukti, k.keterangan || k.kategori, k.akun, 0, k.jumlah, "operasional");
    } else if (k.tipe === "transfer") {
      push(k.tgl, noBukti, k.keterangan || "Transfer antar rekening", k.akunTujuan, k.jumlah, 0, "operasional");
      push(k.tgl, noBukti, k.keterangan || "Transfer antar rekening", k.akun, 0, k.jumlah, "operasional");
    }
  });

  rows.sort((a, b) => (a.tgl < b.tgl ? -1 : a.tgl > b.tgl ? 1 : 0));
  return rows;
}

function exportExcel(filename, rows, sheetName = "Sheet1") {
  try {
    if (!rows || rows.length === 0) { alert("Tidak ada data untuk diekspor."); return; }
    const ws = XLSX.utils.json_to_sheet(rows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, sheetName.slice(0, 31));
    XLSX.writeFile(wb, filename);
  } catch (e) { console.error("Export gagal", e); alert("Gagal mengekspor ke Excel."); }
}

function ReportToolbar({ onExcel, pdfLabel }) {
  return (
    <div className="mfs-noprint" style={{ display: "flex", justifyContent: "flex-end", gap: 8, marginBottom: 12 }}>
      {onExcel && <Btn variant="ghost" onClick={onExcel}><FileSpreadsheet size={14} /> Export Excel</Btn>}
      <Btn variant="ghost" onClick={() => window.print()}><FileOutput size={14} /> {pdfLabel || "Cetak / Simpan PDF"}</Btn>
    </div>
  );
}

function PrintHeader({ company, title, subtitle }) {
  return (
    <div className="mfs-printonly" style={{ display: "none", marginBottom: 14 }}>
      <div style={{ fontFamily: "Fraunces, serif", fontWeight: 700, fontSize: 17 }}>{company?.nama || "Mutis Finance"}</div>
      <div style={{ fontSize: 13, fontWeight: 600, marginTop: 2 }}>{title}</div>
      {subtitle && <div style={{ fontSize: 11, color: "#555" }}>{subtitle}</div>}
      <div style={{ fontSize: 10, color: "#777" }}>Dicetak: {fmtDate(todayStr())}</div>
      <div style={{ borderTop: "1px solid #999", margin: "8px 0" }} />
    </div>
  );
}

const KREDIT_NORMAL = new Set(["Margin & Admin Ditangguhkan", "Modal", "Pendapatan Margin", "Pendapatan Admin", "Pendapatan Denda", "Pendapatan Lain", "Setoran Modal Tambahan", "Penerimaan Lain-lain"]);
const isPendapatan = (akun) => akun.startsWith("Pendapatan") || akun === "Setoran Modal Tambahan" || akun === "Penerimaan Lain-lain";
const isBeban = (akun) => akun.startsWith("Beban");

/* ============================== SMALL UI PRIMITIVES ============================== */
const Card = ({ children, style, ...rest }) => (
  <div style={{ background: C.surface, border: `1px solid ${C.border}`, borderRadius: 10, ...style }} {...rest}>{children}</div>
);

const Btn = ({ children, variant = "primary", onClick, type = "button", style, disabled, ...rest }) => {
  const base = {
    display: "inline-flex", alignItems: "center", gap: 6, fontFamily: "Inter, sans-serif",
    fontSize: 13.5, fontWeight: 600, padding: "9px 14px", borderRadius: 7, cursor: disabled ? "not-allowed" : "pointer",
    border: "1px solid transparent", transition: "all .15s", opacity: disabled ? 0.5 : 1,
  };
  const variants = {
    primary: { background: C.primary, color: "#fff" },
    accent: { background: C.accent, color: "#fff" },
    ghost: { background: "transparent", color: C.primary, border: `1px solid ${C.border}` },
    danger: { background: "transparent", color: C.danger, border: `1px solid ${C.dangerSoft}` },
    subtle: { background: C.surfaceAlt, color: C.ink },
  };
  return (
    <button type={type} disabled={disabled} onClick={onClick} style={{ ...base, ...variants[variant], ...style }}>
      {children}
    </button>
  );
};

const inputStyle = {
  width: "100%", fontFamily: "Inter, sans-serif", fontSize: 13.5, padding: "8px 10px",
  borderRadius: 6, border: `1px solid ${C.border}`, background: "#fff", color: C.ink, outline: "none", boxSizing: "border-box",
};
const Field = ({ label, children, hint }) => (
  <label style={{ display: "block", marginBottom: 12 }}>
    <div style={{ fontSize: 12, fontWeight: 600, color: C.inkSoft, marginBottom: 4 }}>{label}</div>
    {children}
    {hint && <div style={{ fontSize: 11, color: C.inkFaint, marginTop: 3 }}>{hint}</div>}
  </label>
);
const Input = (props) => <input style={inputStyle} {...props} />;
const Select = ({ children, ...props }) => <select style={{ ...inputStyle, appearance: "auto" }} {...props}>{children}</select>;
const TextArea = (props) => <textarea style={{ ...inputStyle, minHeight: 60, resize: "vertical" }} {...props} />;

const Badge = ({ children, tone = "neutral" }) => {
  const tones = {
    neutral: { bg: C.surfaceAlt, fg: C.inkSoft },
    success: { bg: C.successSoft, fg: C.success },
    danger: { bg: C.dangerSoft, fg: C.danger },
    accent: { bg: C.accentSoft, fg: "#8A6425" },
    primary: { bg: C.primarySoft, fg: C.primary },
  };
  const t = tones[tone];
  return <span style={{ background: t.bg, color: t.fg, fontSize: 11, fontWeight: 700, padding: "3px 9px", borderRadius: 20, whiteSpace: "nowrap" }}>{children}</span>;
};

const SectionTitle = ({ icon: Icon, title, subtitle, right }) => (
  <div style={{ display: "flex", alignItems: "flex-end", justifyContent: "space-between", marginBottom: 18, flexWrap: "wrap", gap: 10 }}>
    <div>
      <div style={{ display: "flex", alignItems: "center", gap: 9 }}>
        {Icon && <Icon size={19} color={C.accent} />}
        <h1 style={{ fontFamily: "Fraunces, serif", fontWeight: 600, fontSize: 24, color: C.primaryDark, margin: 0 }}>{title}</h1>
      </div>
      {subtitle && <div style={{ fontSize: 12.5, color: C.inkFaint, marginTop: 4 }}>{subtitle}</div>}
    </div>
    {right}
  </div>
);

const Empty = ({ text, icon: Icon }) => (
  <div style={{ textAlign: "center", padding: "48px 20px", color: C.inkFaint }}>
    {Icon && <Icon size={30} style={{ marginBottom: 8, opacity: 0.5 }} />}
    <div style={{ fontSize: 13 }}>{text}</div>
  </div>
);

function Modal({ title, onClose, children, width = 560 }) {
  return (
    <div style={{ position: "fixed", inset: 0, background: "rgba(20,24,20,0.5)", display: "flex", alignItems: "flex-start", justifyContent: "center", zIndex: 100, padding: "40px 16px", overflowY: "auto" }} onClick={onClose}>
      <div onClick={(e) => e.stopPropagation()} style={{ background: C.surface, borderRadius: 12, width: "100%", maxWidth: width, boxShadow: "0 20px 50px rgba(0,0,0,0.25)" }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "16px 20px", borderBottom: `1px solid ${C.borderSoft}` }}>
          <div style={{ fontFamily: "Fraunces, serif", fontWeight: 600, fontSize: 17, color: C.primaryDark }}>{title}</div>
          <button onClick={onClose} style={{ background: "none", border: "none", cursor: "pointer", color: C.inkFaint, padding: 4 }}><X size={18} /></button>
        </div>
        <div style={{ padding: 20 }}>{children}</div>
      </div>
    </div>
  );
}

function StatCard({ label, value, icon: Icon, tone = "primary", note }) {
  const tones = {
    primary: { bg: C.primary, fg: "#fff" },
    accent: { bg: C.accent, fg: "#fff" },
    light: { bg: C.surface, fg: C.ink },
  };
  const t = tones[tone];
  return (
    <div style={{ background: t.bg, color: t.fg, borderRadius: 10, padding: "16px 18px", border: tone === "light" ? `1px solid ${C.border}` : "none", minWidth: 0 }}>
      <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 10 }}>
        <div style={{ fontSize: 11.5, fontWeight: 600, opacity: 0.8, textTransform: "uppercase", letterSpacing: 0.4 }}>{label}</div>
        {Icon && <Icon size={16} style={{ opacity: 0.75 }} />}
      </div>
      <div style={{ fontFamily: "IBM Plex Mono, monospace", fontWeight: 600, fontSize: 20, fontVariantNumeric: "tabular-nums" }}>{value}</div>
      {note && <div style={{ fontSize: 11, opacity: 0.75, marginTop: 4 }}>{note}</div>}
    </div>
  );
}

const Th = ({ children, style }) => <th style={{ textAlign: "left", fontSize: 11, fontWeight: 700, color: C.inkFaint, textTransform: "uppercase", letterSpacing: 0.3, padding: "9px 12px", borderBottom: `1.5px solid ${C.border}`, ...style }}>{children}</th>;
const Td = ({ children, style }) => <td style={{ padding: "10px 12px", fontSize: 13, color: C.ink, borderBottom: `1px solid ${C.borderSoft}`, ...style }}>{children}</td>;
const TdNum = ({ children, style }) => <Td style={{ fontFamily: "IBM Plex Mono, monospace", fontVariantNumeric: "tabular-nums", textAlign: "right", ...style }}>{children}</Td>;

/* ============================== MENU CONFIG ============================== */
const MENUS = [
  { id: "dashboard", label: "Dashboard", icon: LayoutDashboard },
  { id: "nasabah", label: "Data Nasabah", icon: Users },
  { id: "pembiayaan", label: "Pengajuan Pembiayaan", icon: FileSignature },
  { id: "pembayaran", label: "Pembayaran Angsuran", icon: Wallet },
  { id: "kasbank", label: "Kas & Bank", icon: Landmark },
  { id: "akuntansi", label: "Akuntansi", icon: BookOpen },
  { id: "laporan", label: "Laporan", icon: ClipboardList },
  { id: "pengaturan", label: "Pengaturan", icon: Settings },
];

/* ============================== LOGIN SCREEN ============================== */
function LoginScreen({ onLoggedIn }) {
  const [mode, setMode] = useState("login");
  const [form, setForm] = useState({ nama: "", email: "", password: "", role: "admin" });
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState("");
  const set = (k) => (e) => setForm({ ...form, [k]: e.target.value });

  const submit = async () => {
    setError(""); setBusy(true);
    try {
      if (mode === "login") {
        const tok = await authSignIn(form.email, form.password);
        const session = toSession(tok);
        setAuthSession(session);
        const profile = await fetchProfile(tok.user.id);
        if (!profile) { setAuthSession(null); throw new Error("Akun Anda belum atau tidak lagi aktif. Hubungi Owner untuk mengaktifkan akses."); }
        onLoggedIn({ session, profile });
      } else {
        if (!form.nama || !form.email || !form.password) throw new Error("Lengkapi semua kolom.");
        if (form.password.length < 6) throw new Error("Kata sandi minimal 6 karakter.");
        const tok = await authSignUp(form.email, form.password);
        if (!tok.access_token) throw new Error("Pendaftaran dibuat, tapi belum bisa masuk otomatis. Periksa pengaturan konfirmasi email di Supabase Dashboard.");
        const session = toSession(tok);
        setAuthSession(session);
        await sbInsert("mfs_profiles", { id: tok.user.id, nama: form.nama, email: form.email.toLowerCase(), role: form.role });
        onLoggedIn({ session, profile: { id: tok.user.id, nama: form.nama, email: form.email.toLowerCase(), role: form.role } });
      }
    } catch (e) {
      setError(e.message || "Terjadi kesalahan.");
    }
    setBusy(false);
  };

  return (
    <div style={{ fontFamily: "Inter, sans-serif", minHeight: 640, display: "flex", alignItems: "center", justifyContent: "center", background: C.bg, borderRadius: 14, border: `1px solid ${C.border}` }}>
      <style>{FONTS}</style>
      <div style={{ width: "100%", maxWidth: 380, padding: 28 }}>
        <div style={{ textAlign: "center", marginBottom: 22 }}>
          <div style={{ width: 44, height: 44, borderRadius: 10, background: C.accent, display: "flex", alignItems: "center", justifyContent: "center", margin: "0 auto 12px" }}>
            <svg width="22" height="22" viewBox="0 0 16 16"><path d="M8 0 L16 8 L8 16 L0 8 Z" fill="#fff" opacity="0.9" /><path d="M8 4 L12 8 L8 12 L4 8 Z" fill={C.accent} /></svg>
          </div>
          <div style={{ fontFamily: "Fraunces, serif", fontWeight: 700, fontSize: 20, color: C.primaryDark }}>Mutis Finance</div>
          <div style={{ fontSize: 11, letterSpacing: 1, color: C.inkFaint }}>SYSTEM · MFS</div>
        </div>

        <Card style={{ padding: 22 }}>
          <div style={{ display: "flex", gap: 4, marginBottom: 18, background: C.surfaceAlt, borderRadius: 8, padding: 3 }}>
            <button onClick={() => { setMode("login"); setError(""); }} style={{ flex: 1, padding: "7px 0", borderRadius: 6, border: "none", cursor: "pointer", fontSize: 12.5, fontWeight: 600, background: mode === "login" ? "#fff" : "transparent", color: mode === "login" ? C.primaryDark : C.inkFaint }}>Masuk</button>
            <button onClick={() => { setMode("register"); setError(""); }} style={{ flex: 1, padding: "7px 0", borderRadius: 6, border: "none", cursor: "pointer", fontSize: 12.5, fontWeight: 600, background: mode === "register" ? "#fff" : "transparent", color: mode === "register" ? C.primaryDark : C.inkFaint }}>Daftar Akun</button>
          </div>

          {mode === "register" && (
            <>
              <Field label="Nama Lengkap"><Input value={form.nama} onChange={set("nama")} /></Field>
              <Field label="Peran">
                <Select value={form.role} onChange={set("role")}><option value="admin">Admin</option><option value="owner">Owner</option></Select>
              </Field>
            </>
          )}
          <Field label="Email"><Input type="email" value={form.email} onChange={set("email")} autoComplete="username" /></Field>
          <Field label="Kata Sandi"><Input type="password" value={form.password} onChange={set("password")} autoComplete={mode === "login" ? "current-password" : "new-password"} onKeyDown={(e) => e.key === "Enter" && submit()} /></Field>

          {error && <div style={{ background: C.dangerSoft, color: C.danger, fontSize: 12, borderRadius: 6, padding: "8px 10px", marginBottom: 12 }}>{error}</div>}

          <Btn variant="accent" onClick={submit} disabled={busy} style={{ width: "100%", justifyContent: "center" }}>
            {busy ? <Loader2 size={14} /> : mode === "login" ? <Check size={14} /> : <UserPlus size={14} />}
            {mode === "login" ? "Masuk" : "Daftar & Masuk"}
          </Btn>
        </Card>
        <div style={{ textAlign: "center", fontSize: 11, color: C.inkFaint, marginTop: 14 }}>
          Akses Owner (penuh) atau Admin (operasional harian).
        </div>
      </div>
    </div>
  );
}


export default function App() {
  const [loading, setLoading] = useState(true);
  const [loadError, setLoadError] = useState(null);
  const [active, setActive] = useState("dashboard");
  const [mobileNavOpen, setMobileNavOpen] = useState(false);
  const [currentUser, setCurrentUser] = useState(null);
  const [authChecked, setAuthChecked] = useState(false);

  useEffect(() => {
    setSessionChangedHandler((s) => {
      if (s && s.refreshToken) window.storage.set("mfs_session", JSON.stringify({ refreshToken: s.refreshToken })).catch((e) => console.error(e));
    });
    (async () => {
      try {
        const res = await window.storage.get("mfs_session");
        const refreshToken = res && res.value ? JSON.parse(res.value).refreshToken : null;
        if (refreshToken) {
          const tok = await authRefresh(refreshToken);
          setAuthSession(toSession(tok));
          const profile = await fetchProfile(tok.user.id);
          if (profile) setCurrentUser(profile);
          else { setAuthSession(null); await window.storage.delete("mfs_session").catch(() => {}); }
        }
      } catch (e) {
        console.error("Sesi kedaluwarsa atau tidak valid", e);
        try { await window.storage.delete("mfs_session"); } catch (e2) { /* ignore */ }
      }
      setAuthChecked(true);
    })();
  }, []);

  const handleLoggedIn = ({ session, profile }) => {
    setAuthSession(session);
    setCurrentUser(profile);
    window.storage.set("mfs_session", JSON.stringify({ refreshToken: session.refreshToken })).catch((e) => console.error(e));
  };
  const handleLogout = () => {
    setAuthSession(null);
    setCurrentUser(null);
    window.storage.delete("mfs_session").catch((e) => console.error(e));
  };

  const [company, setCompany] = useState({ nama: "Mutis Finance", alamat: "", telp: "", npwp: "", modalAwal: 0, tglModal: todayStr() });
  const [users, setUsers] = useState([]);
  const [nasabahList, setNasabahList] = useState([]);
  const [akadList, setAkadList] = useState([]);
  const [pembayaranList, setPembayaranList] = useState([]);
  const [kasbankList, setKasbankList] = useState([]);

  useEffect(() => {
    if (!currentUser) return;
    (async () => {
      try {
        const [companyRows, userRows, nasabahRows, akadRows, pembayaranRows, kasbankRows] = await Promise.all([
          sbSelectAll("mfs_company"), sbSelectAll("mfs_users"), sbSelectAll("mfs_nasabah"),
          sbSelectAll("mfs_akad"), sbSelectAll("mfs_pembayaran"), sbSelectAll("mfs_kasbank"),
        ]);
        if (companyRows && companyRows[0]) setCompany(fromCompanyDb(companyRows[0]));
        setUsers((userRows || []).map(fromUserDb));
        setNasabahList((nasabahRows || []).map(fromNasabahDb));
        setAkadList((akadRows || []).map(fromAkadDb));
        setPembayaranList((pembayaranRows || []).map(fromPembayaranDb));
        setKasbankList((kasbankRows || []).map(fromKasbankDb));
      } catch (e) {
        console.error("Gagal memuat data dari Supabase", e);
        setLoadError(e.message);
      }
      setLoading(false);
    })();
  }, [currentUser]);

  const saveCompany = async (v) => { setCompany(v); try { await sbUpsert("mfs_company", toCompanyDb(v)); } catch (e) { console.error(e); alert("Gagal menyimpan data perusahaan ke server."); } };

  const upsertUser = async (u) => {
    setUsers((prev) => { const exists = prev.some((x) => x.id === u.id); return exists ? prev.map((x) => (x.id === u.id ? u : x)) : [...prev, u]; });
    try { await sbUpsert("mfs_users", toUserDb(u)); } catch (e) { console.error(e); alert("Gagal menyimpan pengguna."); }
  };
  const deleteUser = async (id) => { setUsers((prev) => prev.filter((x) => x.id !== id)); try { await sbDelete("mfs_users", id); } catch (e) { console.error(e); } };

  const upsertNasabah = async (n) => {
    setNasabahList((prev) => { const exists = prev.some((x) => x.id === n.id); return exists ? prev.map((x) => (x.id === n.id ? n : x)) : [...prev, n]; });
    try { await sbUpsert("mfs_nasabah", toNasabahDb(n)); } catch (e) { console.error(e); alert("Gagal menyimpan data nasabah ke server."); }
  };
  const deleteNasabah = async (id) => {
    setNasabahList((prev) => prev.filter((x) => x.id !== id));
    try { await sbDelete("mfs_nasabah", id); } catch (e) { console.error(e); alert("Gagal menghapus nasabah di server."); }
  };

  const upsertAkad = async (a) => {
    setAkadList((prev) => { const exists = prev.some((x) => x.id === a.id); return exists ? prev.map((x) => (x.id === a.id ? a : x)) : [...prev, a]; });
    try { await sbUpsert("mfs_akad", toAkadDb(a)); } catch (e) { console.error(e); alert("Gagal menyimpan akad ke server."); }
  };

  const addPembayaran = async (payment, updatedAkad) => {
    setPembayaranList((prev) => [...prev, payment]);
    setAkadList((prev) => prev.map((a) => (a.id === updatedAkad.id ? updatedAkad : a)));
    try {
      await sbUpsert("mfs_pembayaran", toPembayaranDb(payment));
      await sbUpsert("mfs_akad", toAkadDb(updatedAkad));
    } catch (e) { console.error(e); alert("Gagal menyimpan pembayaran ke server."); }
  };
  const updatePembayaran = async (payment, updatedAkad) => {
    setPembayaranList((prev) => prev.map((p) => (p.id === payment.id ? payment : p)));
    setAkadList((prev) => prev.map((a) => (a.id === updatedAkad.id ? updatedAkad : a)));
    try {
      await sbUpsert("mfs_pembayaran", toPembayaranDb(payment));
      await sbUpsert("mfs_akad", toAkadDb(updatedAkad));
    } catch (e) { console.error(e); alert("Gagal memperbarui pembayaran di server."); }
  };

  const upsertKasbank = async (k) => {
    setKasbankList((prev) => { const exists = prev.some((x) => x.id === k.id); return exists ? prev.map((x) => (x.id === k.id ? k : x)) : [...prev, k]; });
    try { await sbUpsert("mfs_kasbank", toKasbankDb(k)); } catch (e) { console.error(e); alert("Gagal menyimpan transaksi ke server."); }
  };
  const deleteKasbank = async (id) => {
    setKasbankList((prev) => prev.filter((x) => x.id !== id));
    try { await sbDelete("mfs_kasbank", id); } catch (e) { console.error(e); alert("Gagal menghapus transaksi di server."); }
  };

  const journal = useMemo(() => computeJournal({ company, akadList, pembayaranList, kasbankList }), [company, akadList, pembayaranList, kasbankList]);

  const restoreAll = async (data) => {
    if (data.company) await saveCompany(data.company);
    if (data.users) for (const u of data.users) await upsertUser(u);
    if (data.nasabahList) for (const n of data.nasabahList) await upsertNasabah(n);
    if (data.akadList) for (const a of data.akadList) await upsertAkad(a);
    if (data.pembayaranList) for (const p of data.pembayaranList) await sbUpsert("mfs_pembayaran", toPembayaranDb(p));
    if (data.kasbankList) for (const k of data.kasbankList) await upsertKasbank(k);
    setPembayaranList(data.pembayaranList || []);
  };

  const ctx = {
    company, saveCompany, users, upsertUser, deleteUser, nasabahList, upsertNasabah, deleteNasabah,
    akadList, upsertAkad, pembayaranList, addPembayaran, updatePembayaran, kasbankList, upsertKasbank, deleteKasbank,
    journal, restoreAll, currentUser, handleLogout,
  };

  if (!authChecked) {
    return (
      <div style={{ minHeight: 400, display: "flex", alignItems: "center", justifyContent: "center", color: C.inkFaint, fontFamily: "Inter, sans-serif" }}>
        <style>{FONTS}</style>
        <Loader2 size={24} />
      </div>
    );
  }

  if (!currentUser) {
    return <LoginScreen onLoggedIn={handleLoggedIn} />;
  }

  if (loadError) {
    return (
      <div style={{ minHeight: 300, display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center", color: C.danger, fontFamily: "Inter, sans-serif", padding: 30, textAlign: "center" }}>
        <AlertCircle size={26} style={{ marginBottom: 10 }} />
        <div style={{ fontSize: 13, fontWeight: 600, marginBottom: 4 }}>Tidak dapat terhubung ke database Supabase.</div>
        <div style={{ fontSize: 12, color: C.inkFaint }}>{loadError}</div>
      </div>
    );
  }

  if (loading) {
    return (
      <div style={{ minHeight: 400, display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center", color: C.inkFaint, fontFamily: "Inter, sans-serif" }}>
        <style>{FONTS}</style>
        <Loader2 size={26} className="mfs-spin" />
        <style>{`.mfs-spin{animation:mfsspin 1s linear infinite}@keyframes mfsspin{to{transform:rotate(360deg)}}`}</style>
        <div style={{ marginTop: 10, fontSize: 13 }}>Memuat data Mutis Finance System…</div>
      </div>
    );
  }

  return (
    <div style={{ fontFamily: "Inter, sans-serif", background: C.bg, color: C.ink, minHeight: 640, display: "flex", borderRadius: 14, overflow: "hidden", border: `1px solid ${C.border}` }}>
      <style>{FONTS}</style>
      <style>{`
        @media print {
          .mfs-noprint { display: none !important; }
          .mfs-app { display: block !important; max-height: none !important; overflow: visible !important; padding: 10px !important; }
          .mfs-printonly { display: block !important; }
          body { background: #fff !important; }
        }
        .mfs-scroll::-webkit-scrollbar { width: 8px; height: 8px; }
        .mfs-scroll::-webkit-scrollbar-thumb { background: ${C.border}; border-radius: 8px; }
        input:focus, select:focus, textarea:focus { border-color: ${C.accent} !important; box-shadow: 0 0 0 3px ${C.accentSoft}; }
      `}</style>

      {/* SIDEBAR */}
      <div className="mfs-noprint" style={{
        width: 224, minWidth: 224, background: C.primaryDark, color: "#EDEFE9", display: "flex", flexDirection: "column",
      }}>
        <div style={{ padding: "20px 18px 16px", borderBottom: "1px dashed rgba(233,201,138,0.35)" }}>
          <div style={{ display: "flex", alignItems: "center", gap: 9 }}>
            <div style={{ width: 30, height: 30, borderRadius: 7, background: C.accent, display: "flex", alignItems: "center", justifyContent: "center", flexShrink: 0 }}>
              <svg width="16" height="16" viewBox="0 0 16 16"><path d="M8 0 L16 8 L8 16 L0 8 Z" fill="#fff" opacity="0.9" /><path d="M8 4 L12 8 L8 12 L4 8 Z" fill={C.accent} /></svg>
            </div>
            <div>
              <div style={{ fontFamily: "Fraunces, serif", fontWeight: 700, fontSize: 15.5, lineHeight: 1.1 }}>Mutis Finance</div>
              <div style={{ fontSize: 10, letterSpacing: 1, opacity: 0.65, marginTop: 2 }}>SYSTEM · MFS</div>
            </div>
          </div>
        </div>
        <div className="mfs-scroll" style={{ flex: 1, overflowY: "auto", padding: "10px 10px" }}>
          {MENUS.filter((m) => m.id !== "pengaturan" || currentUser.role === "owner").map((m) => {
            const isActive = active === m.id;
            const Icon = m.icon;
            return (
              <button key={m.id} onClick={() => setActive(m.id)} style={{
                width: "100%", display: "flex", alignItems: "center", gap: 10, padding: "9px 11px", marginBottom: 3,
                borderRadius: 7, border: "none", cursor: "pointer", textAlign: "left",
                background: isActive ? "rgba(185,134,62,0.22)" : "transparent",
                borderLeft: isActive ? `2.5px solid ${C.accent}` : "2.5px solid transparent",
                color: isActive ? "#F6E9CE" : "#C9D2CB", fontSize: 13, fontWeight: isActive ? 600 : 500,
              }}>
                <Icon size={16} style={{ flexShrink: 0 }} />
                <span>{m.label}</span>
              </button>
            );
          })}
        </div>
        <div style={{ padding: "12px 16px", borderTop: "1px dashed rgba(233,201,138,0.25)" }}>
          <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 8 }}>
            <div style={{ minWidth: 0 }}>
              <div style={{ fontSize: 12.5, fontWeight: 600, color: "#EDEFE9", whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis" }}>{currentUser.nama}</div>
              <div style={{ fontSize: 10, color: C.accent, fontWeight: 700, textTransform: "uppercase", letterSpacing: 0.4 }}>{currentUser.role}</div>
            </div>
            <button onClick={handleLogout} title="Keluar" style={{ background: "none", border: "none", cursor: "pointer", color: "#C9D2CB", padding: 4, flexShrink: 0 }}>
              <RotateCcw size={15} />
            </button>
          </div>
          <div style={{ fontSize: 10, opacity: 0.5 }}>{company.nama || "Mutis Finance"} · MFS v1.0</div>
        </div>
      </div>

      {/* MAIN */}
      <div className="mfs-app mfs-scroll" style={{ flex: 1, minWidth: 0, overflowY: "auto", maxHeight: 900, padding: "26px 30px", background: C.bg }}>
        {active === "dashboard" && <Dashboard ctx={ctx} />}
        {active === "nasabah" && <NasabahPage ctx={ctx} />}
        {active === "pembiayaan" && <PembiayaanPage ctx={ctx} />}
        {active === "pembayaran" && <PembayaranPage ctx={ctx} />}
        {active === "kasbank" && <KasBankPage ctx={ctx} />}
        {active === "akuntansi" && <AkuntansiPage ctx={ctx} />}
        {active === "laporan" && <LaporanPage ctx={ctx} />}
        {active === "pengaturan" && currentUser.role === "owner" && <PengaturanPage ctx={ctx} />}
      </div>
    </div>
  );
}

/* ============================== 1. DASHBOARD ============================== */
function Dashboard({ ctx }) {
  const { company, nasabahList, akadList, pembayaranList, journal } = ctx;
  const today = todayStr();

  const kasSaldo = useMemo(() => journal.filter((j) => j.akun === "Kas").reduce((s, j) => s + j.debit - j.kredit, 0), [journal]);
  const bankSaldo = useMemo(() => journal.filter((j) => j.akun === "Bank").reduce((s, j) => s + j.debit - j.kredit, 0), [journal]);
  const piutangGross = useMemo(() => journal.filter((j) => j.akun === "Piutang Pembiayaan").reduce((s, j) => s + j.debit - j.kredit, 0), [journal]);
  const ditangguhkan = useMemo(() => journal.filter((j) => j.akun === "Margin & Admin Ditangguhkan").reduce((s, j) => s + j.kredit - j.debit, 0), [journal]);
  const piutangNeto = piutangGross - ditangguhkan;

  const bayarHariIni = pembayaranList.filter((p) => p.tglBayar === today).reduce((s, p) => s + p.jumlahBayar + p.denda, 0);

  const jatuhTempoList = useMemo(() => {
    const list = [];
    akadList.forEach((a) => (a.jadwal || []).forEach((j) => {
      if (j.status !== "lunas" && daysBetween(today, j.tglJatuhTempo) <= 7) {
        list.push({ ...j, noAkad: a.noAkad, nasabahNama: a.nasabahNama });
      }
    }));
    return list.sort((a, b) => (a.tglJatuhTempo < b.tglJatuhTempo ? -1 : 1));
  }, [akadList, today]);

  const bulanIni = today.slice(0, 7);
  const labaBulanIni = useMemo(() => {
    const rows = journal.filter((j) => j.tgl.slice(0, 7) === bulanIni);
    const pendapatan = rows.filter((j) => isPendapatan(j.akun)).reduce((s, j) => s + j.kredit - j.debit, 0);
    const beban = rows.filter((j) => isBeban(j.akun)).reduce((s, j) => s + j.debit - j.kredit, 0);
    return pendapatan - beban;
  }, [journal, bulanIni]);

  const chartData = useMemo(() => {
    const out = [];
    for (let i = 5; i >= 0; i--) {
      const d = new Date(); d.setMonth(d.getMonth() - i);
      const key = d.toISOString().slice(0, 7);
      const rows = journal.filter((j) => j.tgl.slice(0, 7) === key && j.akun === "Pendapatan Margin");
      const total = rows.reduce((s, j) => s + j.kredit, 0);
      out.push({ bulan: d.toLocaleDateString("id-ID", { month: "short" }), margin: Math.round(total) });
    }
    return out;
  }, [journal]);

  return (
    <div>
      <SectionTitle icon={LayoutDashboard} title="Dashboard" subtitle={`Ringkasan operasional ${company.nama || "Mutis Finance"} · ${fmtDate(today)}`} />
      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(190px, 1fr))", gap: 14, marginBottom: 22 }}>
        <StatCard label="Total Modal" value={fmtRp(company.modalAwal)} icon={Building2} tone="primary" />
        <StatCard label="Saldo Kas" value={fmtRp(kasSaldo)} icon={Wallet} tone="light" />
        <StatCard label="Saldo Bank" value={fmtRp(bankSaldo)} icon={Landmark} tone="light" />
        <StatCard label="Piutang Pembiayaan" value={fmtRp(piutangNeto)} icon={FileSignature} tone="accent" />
        <StatCard label="Pembayaran Hari Ini" value={fmtRp(bayarHariIni)} icon={TrendingUp} tone="light" />
        <StatCard label="Laba Bulan Berjalan" value={fmtRp(labaBulanIni)} icon={TrendingUp} tone={labaBulanIni >= 0 ? "primary" : "light"} note={fmtMonthYear(bulanIni)} />
      </div>

      <div style={{ display: "grid", gridTemplateColumns: "1.3fr 1fr", gap: 16 }}>
        <Card style={{ padding: 18 }}>
          <div style={{ fontSize: 13, fontWeight: 700, color: C.primaryDark, marginBottom: 12 }}>Pendapatan margin · 6 bulan terakhir</div>
          <div style={{ width: "100%", height: 210 }}>
            <ResponsiveContainer>
              <BarChart data={chartData}>
                <CartesianGrid strokeDasharray="3 3" stroke={C.borderSoft} vertical={false} />
                <XAxis dataKey="bulan" tick={{ fontSize: 11, fill: C.inkFaint }} axisLine={{ stroke: C.border }} tickLine={false} />
                <YAxis tick={{ fontSize: 10, fill: C.inkFaint }} axisLine={false} tickLine={false} width={70} tickFormatter={(v) => (v >= 1000000 ? (v / 1000000).toFixed(1) + "jt" : v)} />
                <Tooltip formatter={(v) => fmtRp(v)} contentStyle={{ fontSize: 12, borderRadius: 8, border: `1px solid ${C.border}` }} />
                <Bar dataKey="margin" fill={C.primary} radius={[4, 4, 0, 0]} />
              </BarChart>
            </ResponsiveContainer>
          </div>
        </Card>

        <Card style={{ padding: 18 }}>
          <div style={{ fontSize: 13, fontWeight: 700, color: C.primaryDark, marginBottom: 4 }}>Angsuran jatuh tempo (7 hari)</div>
          <div style={{ fontSize: 11, color: C.inkFaint, marginBottom: 10 }}>{jatuhTempoList.length} item</div>
          <div style={{ maxHeight: 190, overflowY: "auto" }} className="mfs-scroll">
            {jatuhTempoList.length === 0 && <Empty text="Tidak ada angsuran jatuh tempo dalam 7 hari." />}
            {jatuhTempoList.map((j, i) => {
              const telat = daysBetween(j.tglJatuhTempo, today) > 0;
              return (
                <div key={i} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "8px 0", borderBottom: `1px solid ${C.borderSoft}` }}>
                  <div>
                    <div style={{ fontSize: 12.5, fontWeight: 600 }}>{j.nasabahNama}</div>
                    <div style={{ fontSize: 11, color: C.inkFaint }}>{j.noAkad} · ke-{j.no} · {fmtDate(j.tglJatuhTempo)}</div>
                  </div>
                  <Badge tone={telat ? "danger" : "accent"}>{telat ? "Telat" : fmtRp(j.angsuran)}</Badge>
                </div>
              );
            })}
          </div>
        </Card>
      </div>

      <div style={{ marginTop: 16, fontSize: 11, color: C.inkFaint }}>
        {nasabahList.length} nasabah terdaftar · {akadList.filter((a) => a.status !== "diajukan").length} akad aktif
      </div>
    </div>
  );
}

/* ============================== 2. DATA NASABAH ============================== */
const emptyNasabah = () => ({
  id: null, nama: "", nik: "", noKK: "", alamat: "", telp: "", pekerjaan: "", instansi: "", penghasilan: "",
  penjaminNama: "", penjaminTelp: "", penjaminHubungan: "", fotoKTP: "", fotoKK: "", catatan: "",
});

function NasabahPage({ ctx }) {
  const { nasabahList, upsertNasabah, deleteNasabah, akadList } = ctx;
  const [modal, setModal] = useState(null);
  const [detail, setDetail] = useState(null);
  const [q, setQ] = useState("");

  const filtered = nasabahList.filter((n) => (n.nama + n.nik + n.telp).toLowerCase().includes(q.toLowerCase()));

  const handleSave = (form) => {
    if (form.id) upsertNasabah(form);
    else upsertNasabah({ ...form, id: uid(), createdAt: todayStr() });
    setModal(null);
  };
  const handleDelete = (id) => {
    if (akadList.some((a) => a.nasabahId === id)) { alert("Nasabah ini memiliki riwayat pembiayaan dan tidak dapat dihapus."); return; }
    if (confirm("Hapus data nasabah ini?")) deleteNasabah(id);
  };

  return (
    <div>
      <SectionTitle icon={Users} title="Data Nasabah" subtitle={`${nasabahList.length} nasabah terdaftar`}
        right={<Btn variant="accent" onClick={() => setModal(emptyNasabah())}><Plus size={15} /> Tambah Nasabah</Btn>} />

      <div style={{ marginBottom: 14, maxWidth: 320, position: "relative" }}>
        <Search size={14} style={{ position: "absolute", left: 10, top: 10, color: C.inkFaint }} />
        <input style={{ ...inputStyle, paddingLeft: 30 }} placeholder="Cari nama, NIK, atau telepon…" value={q} onChange={(e) => setQ(e.target.value)} />
      </div>

      <Card>
        <table style={{ width: "100%", borderCollapse: "collapse" }}>
          <thead><tr><Th>Nama</Th><Th>NIK</Th><Th>Telepon</Th><Th>Pekerjaan</Th><Th>Penjamin</Th><Th style={{ textAlign: "right" }}>Aksi</Th></tr></thead>
          <tbody>
            {filtered.map((n) => (
              <tr key={n.id}>
                <Td style={{ fontWeight: 600, cursor: "pointer", color: C.primary }} onClick={() => setDetail(n)}>{n.nama}</Td>
                <Td>{n.nik || "-"}</Td>
                <Td>{n.telp || "-"}</Td>
                <Td>{n.pekerjaan || "-"}</Td>
                <Td>{n.penjaminNama || "-"}</Td>
                <Td style={{ textAlign: "right" }}>
                  <button onClick={() => setDetail(n)} style={iconBtnStyle}><Eye size={14} /></button>
                  <button onClick={() => setModal(n)} style={iconBtnStyle}><Edit2 size={14} /></button>
                  <button onClick={() => handleDelete(n.id)} style={{ ...iconBtnStyle, color: C.danger }}><Trash2 size={14} /></button>
                </Td>
              </tr>
            ))}
          </tbody>
        </table>
        {filtered.length === 0 && <Empty text="Belum ada data nasabah." icon={Users} />}
      </Card>

      {modal && <NasabahForm initial={modal} onSave={handleSave} onClose={() => setModal(null)} />}
      {detail && <NasabahDetail nasabah={detail} akadList={akadList.filter((a) => a.nasabahId === detail.id)} onClose={() => setDetail(null)} />}
    </div>
  );
}

const iconBtnStyle = { background: "none", border: "none", cursor: "pointer", color: C.inkSoft, padding: 5, marginLeft: 2 };

function fileToBase64(file, cb) {
  if (!file) return;
  if (file.size > 2.5 * 1024 * 1024) { alert("Ukuran file maksimal 2.5MB."); return; }
  const reader = new FileReader();
  reader.onload = () => cb(reader.result);
  reader.readAsDataURL(file);
}

function NasabahForm({ initial, onSave, onClose }) {
  const [form, setForm] = useState(initial);
  const set = (k) => (e) => setForm({ ...form, [k]: e.target.value });
  return (
    <Modal title={form.id ? "Edit Nasabah" : "Tambah Nasabah"} onClose={onClose} width={640}>
      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: "0 16px" }}>
        <Field label="Nama Lengkap *"><Input value={form.nama} onChange={set("nama")} placeholder="Nama sesuai KTP" /></Field>
        <Field label="No. KTP / NIK"><Input value={form.nik} onChange={set("nik")} placeholder="16 digit NIK" /></Field>
        <Field label="No. Kartu Keluarga"><Input value={form.noKK} onChange={set("noKK")} /></Field>
        <Field label="No. Telepon"><Input value={form.telp} onChange={set("telp")} /></Field>
      </div>
      <Field label="Alamat"><TextArea value={form.alamat} onChange={set("alamat")} /></Field>
      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: "0 16px" }}>
        <Field label="Pekerjaan"><Input value={form.pekerjaan} onChange={set("pekerjaan")} /></Field>
        <Field label="Instansi / Tempat Kerja"><Input value={form.instansi} onChange={set("instansi")} /></Field>
        <Field label="Penghasilan per Bulan"><Input type="number" value={form.penghasilan} onChange={set("penghasilan")} /></Field>
      </div>
      <div style={{ borderTop: `1px dashed ${C.border}`, margin: "8px 0 14px", paddingTop: 12, fontSize: 12, fontWeight: 700, color: C.primaryDark }}>Kontak Penjamin</div>
      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: "0 16px" }}>
        <Field label="Nama Penjamin"><Input value={form.penjaminNama} onChange={set("penjaminNama")} /></Field>
        <Field label="Telepon Penjamin"><Input value={form.penjaminTelp} onChange={set("penjaminTelp")} /></Field>
        <Field label="Hubungan"><Input value={form.penjaminHubungan} onChange={set("penjaminHubungan")} placeholder="Suami/istri, anak, dsb." /></Field>
      </div>
      <div style={{ borderTop: `1px dashed ${C.border}`, margin: "8px 0 14px", paddingTop: 12, fontSize: 12, fontWeight: 700, color: C.primaryDark }}>Foto Dokumen (opsional)</div>
      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: "0 16px" }}>
        <Field label="Foto KTP">
          <input type="file" accept="image/*" onChange={(e) => fileToBase64(e.target.files[0], (b64) => setForm({ ...form, fotoKTP: b64 }))} style={{ fontSize: 12 }} />
          {form.fotoKTP && <img src={form.fotoKTP} alt="KTP" style={{ marginTop: 6, maxWidth: "100%", maxHeight: 110, borderRadius: 6, border: `1px solid ${C.border}` }} />}
        </Field>
        <Field label="Foto KK">
          <input type="file" accept="image/*" onChange={(e) => fileToBase64(e.target.files[0], (b64) => setForm({ ...form, fotoKK: b64 }))} style={{ fontSize: 12 }} />
          {form.fotoKK && <img src={form.fotoKK} alt="KK" style={{ marginTop: 6, maxWidth: "100%", maxHeight: 110, borderRadius: 6, border: `1px solid ${C.border}` }} />}
        </Field>
      </div>
      <div style={{ display: "flex", justifyContent: "flex-end", gap: 8, marginTop: 8 }}>
        <Btn variant="ghost" onClick={onClose}>Batal</Btn>
        <Btn variant="accent" onClick={() => { if (!form.nama) { alert("Nama wajib diisi."); return; } onSave(form); }}><Save size={14} /> Simpan</Btn>
      </div>
    </Modal>
  );
}

function NasabahDetail({ nasabah, akadList, onClose }) {
  return (
    <Modal title={nasabah.nama} onClose={onClose} width={560}>
      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: "4px 20px", fontSize: 13, marginBottom: 14 }}>
        <div><b>NIK:</b> {nasabah.nik || "-"}</div>
        <div><b>No. KK:</b> {nasabah.noKK || "-"}</div>
        <div><b>Telepon:</b> {nasabah.telp || "-"}</div>
        <div><b>Pekerjaan:</b> {nasabah.pekerjaan || "-"} {nasabah.instansi ? `(${nasabah.instansi})` : ""}</div>
        <div style={{ gridColumn: "1/3" }}><b>Alamat:</b> {nasabah.alamat || "-"}</div>
        <div style={{ gridColumn: "1/3" }}><b>Penjamin:</b> {nasabah.penjaminNama || "-"} {nasabah.penjaminTelp ? `· ${nasabah.penjaminTelp}` : ""} {nasabah.penjaminHubungan ? `(${nasabah.penjaminHubungan})` : ""}</div>
      </div>
      {(nasabah.fotoKTP || nasabah.fotoKK) && (
        <div style={{ display: "flex", gap: 10, marginBottom: 16 }}>
          {nasabah.fotoKTP && <img src={nasabah.fotoKTP} style={{ maxHeight: 100, borderRadius: 6, border: `1px solid ${C.border}` }} />}
          {nasabah.fotoKK && <img src={nasabah.fotoKK} style={{ maxHeight: 100, borderRadius: 6, border: `1px solid ${C.border}` }} />}
        </div>
      )}
      <div style={{ fontSize: 12, fontWeight: 700, color: C.primaryDark, marginBottom: 8 }}>Riwayat Pembiayaan ({akadList.length})</div>
      {akadList.length === 0 && <div style={{ fontSize: 12, color: C.inkFaint }}>Belum ada riwayat pembiayaan.</div>}
      {akadList.map((a) => (
        <div key={a.id} style={{ display: "flex", justifyContent: "space-between", padding: "7px 0", borderBottom: `1px solid ${C.borderSoft}`, fontSize: 12.5 }}>
          <div>{a.noAkad} · {a.jenisAkad}</div>
          <div>{fmtRp(a.nilaiPembiayaan)}</div>
        </div>
      ))}
    </Modal>
  );
}

/* ============================== 3. PENGAJUAN PEMBIAYAAN ============================== */
const emptyAkad = (nasabahList) => ({
  id: null, nasabahId: nasabahList[0]?.id || "", jenisAkad: "Murabahah", nilaiPembiayaan: "", marginPersen: "",
  biayaAdmin: "", dp: "", tenor: "", tglAkad: todayStr(), metodePencairan: "Kas", status: "aktif", catatan: "",
});

function PembiayaanPage({ ctx }) {
  const { akadList, upsertAkad, nasabahList } = ctx;
  const [modal, setModal] = useState(null);
  const [editAkad, setEditAkad] = useState(null);
  const [detail, setDetail] = useState(null);

  const handleCreate = (akad) => {
    const nasabah = nasabahList.find((n) => n.id === akad.nasabahId);
    const noAkad = generateNoAkad(akad.jenisAkad, akadList, akad.tglAkad);
    const { jadwal, totalTagihan, cicilanPerBulan } = computeJadwal(akad);
    const newAkad = { ...akad, id: uid(), noAkad, nasabahNama: nasabah?.nama || "-", jadwal, totalTagihan, cicilanPerBulan, createdAt: todayStr() };
    upsertAkad(newAkad);
    setModal(null);
    setDetail(newAkad);
  };

  const handleUpdate = (akad, regenerateJadwal) => {
    const nasabah = nasabahList.find((n) => n.id === akad.nasabahId);
    let updated = { ...akad, nasabahNama: nasabah?.nama || akad.nasabahNama };
    if (regenerateJadwal) {
      const { jadwal, totalTagihan, cicilanPerBulan } = computeJadwal(akad);
      updated = { ...updated, jadwal, totalTagihan, cicilanPerBulan };
    }
    upsertAkad(updated);
    setEditAkad(null);
  };

  const sudahAdaPembayaran = (akad) => (akad.jadwal || []).some((j) => j.status === "lunas");

  return (
    <div>
      <SectionTitle icon={FileSignature} title="Pengajuan Pembiayaan" subtitle={`${akadList.length} akad tercatat`}
        right={<Btn variant="accent" onClick={() => setModal(true)} disabled={nasabahList.length === 0}><Plus size={15} /> Akad Baru</Btn>} />
      {nasabahList.length === 0 && <div style={{ fontSize: 12.5, color: C.danger, marginBottom: 10 }}>Tambahkan data nasabah terlebih dahulu sebelum membuat akad.</div>}

      <Card>
        <table style={{ width: "100%", borderCollapse: "collapse" }}>
          <thead><tr><Th>No. Akad</Th><Th>Nasabah</Th><Th>Jenis</Th><Th style={{ textAlign: "right" }}>Nilai</Th><Th style={{ textAlign: "right" }}>Cicilan/bln</Th><Th>Tenor</Th><Th>Status</Th><Th style={{ textAlign: "right" }}>Aksi</Th></tr></thead>
          <tbody>
            {akadList.map((a) => {
              const lunas = (a.jadwal || []).every((j) => j.status === "lunas");
              return (
                <tr key={a.id}>
                  <Td style={{ fontFamily: "IBM Plex Mono, monospace", fontWeight: 600, color: C.primary, cursor: "pointer" }} onClick={() => setDetail(a)}>{a.noAkad}</Td>
                  <Td>{a.nasabahNama}</Td>
                  <Td>{a.jenisAkad}</Td>
                  <TdNum>{fmtRp(a.nilaiPembiayaan)}</TdNum>
                  <TdNum>{fmtRp(a.cicilanPerBulan)}</TdNum>
                  <Td>{a.tenor} bln</Td>
                  <Td><Badge tone={lunas ? "success" : "accent"}>{lunas ? "Lunas" : "Berjalan"}</Badge></Td>
                  <Td style={{ textAlign: "right" }}>
                    <button onClick={() => setDetail(a)} style={iconBtnStyle}><Eye size={14} /></button>
                    <button onClick={() => setEditAkad(a)} style={iconBtnStyle}><Edit2 size={14} /></button>
                  </Td>
                </tr>
              );
            })}
          </tbody>
        </table>
        {akadList.length === 0 && <Empty text="Belum ada akad pembiayaan." icon={FileSignature} />}
      </Card>

      {modal && <AkadForm nasabahList={nasabahList} akadList={akadList} onSave={handleCreate} onClose={() => setModal(false)} />}
      {editAkad && (
        <AkadEditForm akad={editAkad} nasabahList={nasabahList} locked={sudahAdaPembayaran(editAkad)}
          onSave={handleUpdate} onClose={() => setEditAkad(null)} />
      )}
      {detail && <AkadDetail akad={detail} onClose={() => setDetail(null)} />}
    </div>
  );
}

function AkadForm({ nasabahList, akadList, onSave, onClose }) {
  const [form, setForm] = useState(emptyAkad(nasabahList));
  const set = (k) => (e) => setForm({ ...form, [k]: e.target.value });

  const preview = useMemo(() => {
    if (!form.nilaiPembiayaan || !form.tenor) return null;
    return computeJadwal(form);
  }, [form.nilaiPembiayaan, form.marginPersen, form.biayaAdmin, form.dp, form.tenor, form.tglAkad]);

  return (
    <Modal title="Pengajuan Akad Pembiayaan Baru" onClose={onClose} width={640}>
      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: "0 16px" }}>
        <Field label="Nasabah *">
          <Select value={form.nasabahId} onChange={set("nasabahId")}>
            {nasabahList.map((n) => <option key={n.id} value={n.id}>{n.nama}</option>)}
          </Select>
        </Field>
        <Field label="Jenis Akad">
          <Select value={form.jenisAkad} onChange={set("jenisAkad")}>{JENIS_AKAD.map((j) => <option key={j}>{j}</option>)}</Select>
        </Field>
        <Field label="Tanggal Akad"><Input type="date" value={form.tglAkad} onChange={set("tglAkad")} /></Field>
        <Field label="Metode Pencairan">
          <Select value={form.metodePencairan} onChange={set("metodePencairan")}><option>Kas</option><option>Bank</option></Select>
        </Field>
        <Field label="Nilai Pembiayaan (Rp) *"><Input type="number" value={form.nilaiPembiayaan} onChange={set("nilaiPembiayaan")} /></Field>
        <Field label="Margin (%)"><Input type="number" value={form.marginPersen} onChange={set("marginPersen")} placeholder="mis. 12" /></Field>
        <Field label="Biaya Admin (Rp)"><Input type="number" value={form.biayaAdmin} onChange={set("biayaAdmin")} /></Field>
        <Field label="DP / Uang Muka (Rp)"><Input type="number" value={form.dp} onChange={set("dp")} /></Field>
        <Field label="Lama Angsuran (bulan) *"><Input type="number" value={form.tenor} onChange={set("tenor")} /></Field>
      </div>
      <Field label="Catatan"><TextArea value={form.catatan} onChange={set("catatan")} /></Field>

      {preview && (
        <div style={{ background: C.primarySoft, borderRadius: 8, padding: 14, marginTop: 6, marginBottom: 14 }}>
          <div style={{ fontSize: 11.5, color: C.inkSoft, marginBottom: 6 }}>Cicilan = (Nilai + Margin + Admin − DP) ÷ Lama Angsuran</div>
          <div style={{ display: "flex", justifyContent: "space-between", fontSize: 13 }}>
            <span>Total tagihan:</span><b style={{ fontFamily: "IBM Plex Mono, monospace" }}>{fmtRp(preview.totalTagihan)}</b>
          </div>
          <div style={{ display: "flex", justifyContent: "space-between", fontSize: 15, marginTop: 4 }}>
            <span>Cicilan per bulan:</span><b style={{ fontFamily: "IBM Plex Mono, monospace", color: C.primaryDark }}>{fmtRp(preview.cicilanPerBulan)}</b>
          </div>
        </div>
      )}

      <div style={{ display: "flex", justifyContent: "flex-end", gap: 8 }}>
        <Btn variant="ghost" onClick={onClose}>Batal</Btn>
        <Btn variant="accent" onClick={() => {
          if (!form.nasabahId || !form.nilaiPembiayaan || !form.tenor) { alert("Lengkapi nasabah, nilai pembiayaan, dan lama angsuran."); return; }
          onSave(form);
        }}><Save size={14} /> Buat Akad & Jadwal</Btn>
      </div>
    </Modal>
  );
}

function AkadEditForm({ akad, nasabahList, locked, onSave, onClose }) {
  const [form, setForm] = useState(akad);
  const set = (k) => (e) => setForm({ ...form, [k]: e.target.value });

  const preview = useMemo(() => {
    if (locked || !form.nilaiPembiayaan || !form.tenor) return null;
    return computeJadwal(form);
  }, [form.nilaiPembiayaan, form.marginPersen, form.biayaAdmin, form.dp, form.tenor, form.tglAkad, locked]);

  return (
    <Modal title={`Edit Akad ${akad.noAkad}`} onClose={onClose} width={640}>
      {locked && (
        <div style={{ background: C.accentSoft, color: "#8A6425", fontSize: 12, borderRadius: 8, padding: 10, marginBottom: 14 }}>
          Sebagian angsuran pada akad ini sudah dibayar, sehingga nilai pembiayaan, margin, biaya admin, DP, dan tenor tidak dapat diubah agar jadwal & riwayat pembayaran tetap konsisten. Anda tetap dapat mengubah status, metode pencairan, dan catatan.
        </div>
      )}
      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: "0 16px" }}>
        <Field label="Nasabah *">
          <Select value={form.nasabahId} onChange={set("nasabahId")}>{nasabahList.map((n) => <option key={n.id} value={n.id}>{n.nama}</option>)}</Select>
        </Field>
        <Field label="Jenis Akad"><Select value={form.jenisAkad} onChange={set("jenisAkad")}>{JENIS_AKAD.map((j) => <option key={j}>{j}</option>)}</Select></Field>
        <Field label="Tanggal Akad"><Input type="date" value={form.tglAkad} onChange={set("tglAkad")} disabled={locked} /></Field>
        <Field label="Metode Pencairan"><Select value={form.metodePencairan} onChange={set("metodePencairan")}><option>Kas</option><option>Bank</option></Select></Field>
        <Field label="Nilai Pembiayaan (Rp) *"><Input type="number" value={form.nilaiPembiayaan} onChange={set("nilaiPembiayaan")} disabled={locked} /></Field>
        <Field label="Margin (%)"><Input type="number" value={form.marginPersen} onChange={set("marginPersen")} disabled={locked} /></Field>
        <Field label="Biaya Admin (Rp)"><Input type="number" value={form.biayaAdmin} onChange={set("biayaAdmin")} disabled={locked} /></Field>
        <Field label="DP / Uang Muka (Rp)"><Input type="number" value={form.dp} onChange={set("dp")} disabled={locked} /></Field>
        <Field label="Lama Angsuran (bulan) *"><Input type="number" value={form.tenor} onChange={set("tenor")} disabled={locked} /></Field>
        <Field label="Status"><Select value={form.status} onChange={set("status")}><option value="aktif">Aktif</option><option value="lunas">Lunas</option><option value="bermasalah">Bermasalah</option></Select></Field>
      </div>
      <Field label="Catatan"><TextArea value={form.catatan} onChange={set("catatan")} /></Field>

      {preview && (
        <div style={{ background: C.primarySoft, borderRadius: 8, padding: 14, marginBottom: 14 }}>
          <div style={{ fontSize: 11.5, color: C.inkSoft, marginBottom: 6 }}>Jadwal angsuran akan dihitung ulang sesuai perubahan.</div>
          <div style={{ display: "flex", justifyContent: "space-between", fontSize: 15 }}><span>Cicilan per bulan:</span><b style={{ fontFamily: "IBM Plex Mono, monospace", color: C.primaryDark }}>{fmtRp(preview.cicilanPerBulan)}</b></div>
        </div>
      )}

      <div style={{ display: "flex", justifyContent: "flex-end", gap: 8 }}>
        <Btn variant="ghost" onClick={onClose}>Batal</Btn>
        <Btn variant="accent" onClick={() => onSave(form, !locked)}><Save size={14} /> Simpan Perubahan</Btn>
      </div>
    </Modal>
  );
}

function AkadDetail({ akad, onClose }) {
  const totalDibayar = (akad.jadwal || []).filter((j) => j.status === "lunas").length;
  return (
    <Modal title={`Akad ${akad.noAkad}`} onClose={onClose} width={680}>
      <div style={{ display: "grid", gridTemplateColumns: "repeat(3, 1fr)", gap: "6px 16px", fontSize: 12.5, marginBottom: 16 }}>
        <div><b>Nasabah:</b> {akad.nasabahNama}</div>
        <div><b>Jenis Akad:</b> {akad.jenisAkad}</div>
        <div><b>Tanggal:</b> {fmtDate(akad.tglAkad)}</div>
        <div><b>Nilai:</b> {fmtRp(akad.nilaiPembiayaan)}</div>
        <div><b>Margin:</b> {akad.marginPersen || 0}%</div>
        <div><b>Admin/DP:</b> {fmtRp(akad.biayaAdmin)} / {fmtRp(akad.dp)}</div>
        <div><b>Tenor:</b> {akad.tenor} bulan</div>
        <div><b>Cicilan/bln:</b> {fmtRp(akad.cicilanPerBulan)}</div>
        <div><b>Terbayar:</b> {totalDibayar}/{akad.tenor}</div>
      </div>
      <div style={{ maxHeight: 320, overflowY: "auto" }} className="mfs-scroll">
        <table style={{ width: "100%", borderCollapse: "collapse" }}>
          <thead><tr><Th>Ke</Th><Th>Jatuh Tempo</Th><Th style={{ textAlign: "right" }}>Angsuran</Th><Th>Status</Th><Th>Tgl Bayar</Th></tr></thead>
          <tbody>
            {(akad.jadwal || []).map((j) => (
              <tr key={j.no}>
                <Td>{j.no}</Td><Td>{fmtDate(j.tglJatuhTempo)}</Td><TdNum>{fmtRp(j.angsuran)}</TdNum>
                <Td><Badge tone={j.status === "lunas" ? "success" : "neutral"}>{j.status === "lunas" ? "Lunas" : "Belum"}</Badge></Td>
                <Td>{j.tglBayar ? fmtDate(j.tglBayar) : "-"}</Td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </Modal>
  );
}

/* ============================== 4. PEMBAYARAN ANGSURAN ============================== */
function PembayaranPage({ ctx }) {
  const { akadList, pembayaranList, addPembayaran, updatePembayaran } = ctx;
  const [q, setQ] = useState("");
  const [selectedAkad, setSelectedAkad] = useState(null);
  const [payModal, setPayModal] = useState(null);
  const [editPayModal, setEditPayModal] = useState(null);
  const [receipt, setReceipt] = useState(null);

  const berjalan = akadList.filter((a) => a.status !== "diajukan" && (a.jadwal || []).some((j) => j.status !== "lunas"));
  const filtered = berjalan.filter((a) => (a.nasabahNama + a.noAkad).toLowerCase().includes(q.toLowerCase()));

  const splitMarginAdmin = (akad, marginAdminNominal) => {
    const totalMarginAdmin = (akad.marginPersen / 100 * akad.nilaiPembiayaan) + (+akad.biayaAdmin || 0);
    const marginShare = totalMarginAdmin > 0 ? ((akad.marginPersen / 100 * akad.nilaiPembiayaan) / totalMarginAdmin) : 0;
    const mPortion = marginAdminNominal * marginShare;
    return { mPortion, aPortion: marginAdminNominal - mPortion };
  };

  const submitPayment = (akad, jadwalItem, jumlahBayar, denda, metode, tglBayar) => {
    const noBukti = "PAY-" + uid().toUpperCase();
    const { mPortion, aPortion } = splitMarginAdmin(akad, jadwalItem.marginAdmin);

    const payment = {
      id: uid(), noBukti, akadId: akad.id, noAkad: akad.noAkad, nasabahNama: akad.nasabahNama,
      angsuranKe: jadwalItem.no, tglBayar, jumlahBayar, denda: clampNum(denda), metode,
      marginAdminPortion: jadwalItem.marginAdmin, marginPortion: mPortion, adminPortion: aPortion,
    };
    const updatedJadwal = akad.jadwal.map((j) => j.no === jadwalItem.no ? { ...j, status: "lunas", tglBayar, denda: clampNum(denda), pembayaranId: payment.id } : j);
    const updatedAkad = { ...akad, jadwal: updatedJadwal };
    addPembayaran(payment, updatedAkad);
    setSelectedAkad(updatedAkad);
    setPayModal(null);
    setReceipt({ payment, akad: updatedAkad });
  };

  const submitEditPayment = (payment, akad, jumlahBayar, denda, metode, tglBayar) => {
    const { mPortion, aPortion } = splitMarginAdmin(akad, payment.marginAdminPortion);
    const updatedPayment = { ...payment, jumlahBayar, denda: clampNum(denda), metode, tglBayar, marginPortion: mPortion, adminPortion: aPortion };
    const updatedJadwal = akad.jadwal.map((j) => j.no === payment.angsuranKe ? { ...j, tglBayar, denda: clampNum(denda) } : j);
    const updatedAkad = { ...akad, jadwal: updatedJadwal };
    updatePembayaran(updatedPayment, updatedAkad);
    setSelectedAkad(updatedAkad);
    setEditPayModal(null);
  };

  return (
    <div>
      <SectionTitle icon={Wallet} title="Pembayaran Angsuran" subtitle="Catat pembayaran, denda, dan cetak bukti" />
      <div style={{ display: "grid", gridTemplateColumns: "320px 1fr", gap: 18 }}>
        <div>
          <div style={{ marginBottom: 10, position: "relative" }}>
            <Search size={14} style={{ position: "absolute", left: 10, top: 10, color: C.inkFaint }} />
            <input style={{ ...inputStyle, paddingLeft: 30 }} placeholder="Cari nasabah / no akad…" value={q} onChange={(e) => setQ(e.target.value)} />
          </div>
          <Card style={{ maxHeight: 520, overflowY: "auto" }} className="mfs-scroll">
            {filtered.map((a) => (
              <div key={a.id} onClick={() => setSelectedAkad(a)} style={{
                padding: "11px 14px", borderBottom: `1px solid ${C.borderSoft}`, cursor: "pointer",
                background: selectedAkad?.id === a.id ? C.primarySoft : "transparent",
              }}>
                <div style={{ fontSize: 13, fontWeight: 600 }}>{a.nasabahNama}</div>
                <div style={{ fontSize: 11, color: C.inkFaint, fontFamily: "IBM Plex Mono, monospace" }}>{a.noAkad}</div>
              </div>
            ))}
            {filtered.length === 0 && <Empty text="Tidak ada akad berjalan." />}
          </Card>
        </div>

        <div>
          {!selectedAkad && <Card style={{ padding: 24 }}><Empty text="Pilih nasabah / akad di sebelah kiri untuk melihat jadwal angsuran." icon={Wallet} /></Card>}
          {selectedAkad && (
            <Card style={{ padding: 16 }}>
              <div style={{ display: "flex", justifyContent: "space-between", marginBottom: 12 }}>
                <div>
                  <div style={{ fontFamily: "Fraunces, serif", fontWeight: 600, fontSize: 16 }}>{selectedAkad.nasabahNama}</div>
                  <div style={{ fontSize: 11.5, color: C.inkFaint }}>{selectedAkad.noAkad} · Sisa pokok: {fmtRp(outstandingPokok(selectedAkad))}</div>
                </div>
              </div>
              <table style={{ width: "100%", borderCollapse: "collapse" }}>
                <thead><tr><Th>Ke</Th><Th>Jatuh Tempo</Th><Th style={{ textAlign: "right" }}>Angsuran</Th><Th>Status</Th><Th style={{ textAlign: "right" }}>Aksi</Th></tr></thead>
                <tbody>
                  {selectedAkad.jadwal.map((j) => {
                    const telat = j.status !== "lunas" && daysBetween(j.tglJatuhTempo, todayStr()) > 0;
                    return (
                      <tr key={j.no}>
                        <Td>{j.no}</Td><Td>{fmtDate(j.tglJatuhTempo)}</Td><TdNum>{fmtRp(j.angsuran)}</TdNum>
                        <Td><Badge tone={j.status === "lunas" ? "success" : telat ? "danger" : "neutral"}>{j.status === "lunas" ? "Lunas" : telat ? `Telat ${daysBetween(j.tglJatuhTempo, todayStr())}h` : "Belum"}</Badge></Td>
                        <Td style={{ textAlign: "right" }}>
                          {j.status !== "lunas"
                            ? <Btn variant="accent" style={{ padding: "5px 10px", fontSize: 12 }} onClick={() => setPayModal(j)}>Bayar</Btn>
                            : <>
                                <button style={iconBtnStyle} onClick={() => setReceipt({ payment: pembayaranList.find((p) => p.id === j.pembayaranId), akad: selectedAkad })}><Printer size={14} /></button>
                                <button style={iconBtnStyle} onClick={() => setEditPayModal(pembayaranList.find((p) => p.id === j.pembayaranId))}><Edit2 size={14} /></button>
                              </>}
                        </Td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </Card>
          )}
        </div>
      </div>

      {payModal && selectedAkad && (
        <PaymentForm akad={selectedAkad} jadwalItem={payModal} onClose={() => setPayModal(null)}
          onSubmit={(jumlahBayar, denda, metode, tglBayar) => submitPayment(selectedAkad, payModal, jumlahBayar, denda, metode, tglBayar)} />
      )}
      {editPayModal && selectedAkad && (
        <PaymentEditForm payment={editPayModal} akad={selectedAkad} onClose={() => setEditPayModal(null)}
          onSubmit={(jumlahBayar, denda, metode, tglBayar) => submitEditPayment(editPayModal, selectedAkad, jumlahBayar, denda, metode, tglBayar)} />
      )}
      {receipt && receipt.payment && <ReceiptModal data={receipt} onClose={() => setReceipt(null)} />}
    </div>
  );
}

function PaymentForm({ akad, jadwalItem, onSubmit, onClose }) {
  const [jumlahBayar, setJumlahBayar] = useState(jadwalItem.angsuran);
  const [denda, setDenda] = useState(0);
  const [metode, setMetode] = useState("Kas");
  const [tglBayar, setTglBayar] = useState(todayStr());
  return (
    <Modal title={`Bayar Angsuran ke-${jadwalItem.no}`} onClose={onClose} width={420}>
      <div style={{ fontSize: 12, color: C.inkFaint, marginBottom: 10 }}>{akad.nasabahNama} · {akad.noAkad}</div>
      <Field label="Tanggal Bayar"><Input type="date" value={tglBayar} onChange={(e) => setTglBayar(e.target.value)} /></Field>
      <Field label="Jumlah Angsuran (Rp)"><Input type="number" value={jumlahBayar} onChange={(e) => setJumlahBayar(+e.target.value)} /></Field>
      <Field label="Denda (opsional, Rp)"><Input type="number" value={denda} onChange={(e) => setDenda(+e.target.value)} /></Field>
      <Field label="Metode Pembayaran"><Select value={metode} onChange={(e) => setMetode(e.target.value)}><option>Kas</option><option>Bank</option></Select></Field>
      <div style={{ background: C.primarySoft, borderRadius: 8, padding: 12, display: "flex", justifyContent: "space-between", marginBottom: 14 }}>
        <span style={{ fontSize: 13 }}>Total diterima</span>
        <b style={{ fontFamily: "IBM Plex Mono, monospace" }}>{fmtRp((+jumlahBayar || 0) + (+denda || 0))}</b>
      </div>
      <div style={{ display: "flex", justifyContent: "flex-end", gap: 8 }}>
        <Btn variant="ghost" onClick={onClose}>Batal</Btn>
        <Btn variant="accent" onClick={() => onSubmit(+jumlahBayar || 0, +denda || 0, metode, tglBayar)}><Check size={14} /> Konfirmasi Bayar</Btn>
      </div>
    </Modal>
  );
}

function PaymentEditForm({ payment, akad, onSubmit, onClose }) {
  const [jumlahBayar, setJumlahBayar] = useState(payment.jumlahBayar);
  const [denda, setDenda] = useState(payment.denda);
  const [metode, setMetode] = useState(payment.metode);
  const [tglBayar, setTglBayar] = useState(payment.tglBayar);
  return (
    <Modal title={`Edit Pembayaran ke-${payment.angsuranKe}`} onClose={onClose} width={420}>
      <div style={{ fontSize: 12, color: C.inkFaint, marginBottom: 10 }}>{akad.nasabahNama} · {akad.noAkad} · {payment.noBukti}</div>
      <Field label="Tanggal Bayar"><Input type="date" value={tglBayar} onChange={(e) => setTglBayar(e.target.value)} /></Field>
      <Field label="Jumlah Angsuran (Rp)"><Input type="number" value={jumlahBayar} onChange={(e) => setJumlahBayar(+e.target.value)} /></Field>
      <Field label="Denda (opsional, Rp)"><Input type="number" value={denda} onChange={(e) => setDenda(+e.target.value)} /></Field>
      <Field label="Metode Pembayaran"><Select value={metode} onChange={(e) => setMetode(e.target.value)}><option>Kas</option><option>Bank</option></Select></Field>
      <div style={{ display: "flex", justifyContent: "flex-end", gap: 8 }}>
        <Btn variant="ghost" onClick={onClose}>Batal</Btn>
        <Btn variant="accent" onClick={() => onSubmit(+jumlahBayar || 0, +denda || 0, metode, tglBayar)}><Save size={14} /> Simpan Perubahan</Btn>
      </div>
    </Modal>
  );
}

function ReceiptModal({ data, onClose }) {
  const { payment, akad } = data;
  return (
    <Modal title="Bukti Pembayaran" onClose={onClose} width={400}>
      <div id="mfs-receipt" style={{ border: `1px dashed ${C.border}`, borderRadius: 8, padding: 18, fontFamily: "IBM Plex Mono, monospace" }}>
        <div style={{ textAlign: "center", marginBottom: 10 }}>
          <div style={{ fontFamily: "Fraunces, serif", fontWeight: 700, fontSize: 15 }}>MUTIS FINANCE</div>
          <div style={{ fontSize: 10, color: C.inkFaint }}>Bukti Pembayaran Angsuran</div>
        </div>
        <div style={{ fontSize: 12, lineHeight: 1.9 }}>
          <div style={{ display: "flex", justifyContent: "space-between" }}><span>No. Bukti</span><span>{payment.noBukti}</span></div>
          <div style={{ display: "flex", justifyContent: "space-between" }}><span>Tanggal</span><span>{fmtDate(payment.tglBayar)}</span></div>
          <div style={{ display: "flex", justifyContent: "space-between" }}><span>No. Akad</span><span>{payment.noAkad}</span></div>
          <div style={{ display: "flex", justifyContent: "space-between" }}><span>Nasabah</span><span>{payment.nasabahName || akad.nasabahNama}</span></div>
          <div style={{ display: "flex", justifyContent: "space-between" }}><span>Angsuran ke</span><span>{payment.angsuranKe} / {akad.tenor}</span></div>
          <div style={{ borderTop: `1px dashed ${C.border}`, margin: "6px 0" }} />
          <div style={{ display: "flex", justifyContent: "space-between" }}><span>Angsuran</span><span>{fmtRp(payment.jumlahBayar)}</span></div>
          <div style={{ display: "flex", justifyContent: "space-between" }}><span>Denda</span><span>{fmtRp(payment.denda)}</span></div>
          <div style={{ display: "flex", justifyContent: "space-between", fontWeight: 700 }}><span>Total</span><span>{fmtRp(payment.jumlahBayar + payment.denda)}</span></div>
        </div>
      </div>
      <div style={{ display: "flex", justifyContent: "flex-end", gap: 8, marginTop: 14 }}>
        <Btn variant="ghost" onClick={onClose}>Tutup</Btn>
        <Btn variant="accent" onClick={() => window.print()}><Printer size={14} /> Cetak</Btn>
      </div>
    </Modal>
  );
}

/* ============================== 5. KAS & BANK ============================== */
const emptyKasbank = () => ({ id: null, tgl: todayStr(), tipe: "penerimaan", akun: "Kas", akunTujuan: "Bank", kategori: KATEGORI_PENERIMAAN[0], jumlah: "", keterangan: "" });

function KasBankPage({ ctx }) {
  const { kasbankList, upsertKasbank, deleteKasbank, journal } = ctx;
  const [modal, setModal] = useState(null);
  const [recon, setRecon] = useState(false);
  const [saldoBank, setSaldoBank] = useState("");

  const kasSaldo = journal.filter((j) => j.akun === "Kas").reduce((s, j) => s + j.debit - j.kredit, 0);
  const bankSaldo = journal.filter((j) => j.akun === "Bank").reduce((s, j) => s + j.debit - j.kredit, 0);

  const handleSave = (form) => {
    const rec = { ...form, id: form.id || uid(), jumlah: +form.jumlah, noBukti: form.noBukti || "KB-" + uid().toUpperCase() };
    upsertKasbank(rec);
    setModal(null);
  };
  const handleDelete = (id) => { if (confirm("Hapus transaksi ini?")) deleteKasbank(id); };

  const sorted = [...kasbankList].sort((a, b) => (a.tgl < b.tgl ? 1 : -1));

  return (
    <div>
      <SectionTitle icon={Landmark} title="Kas & Bank" subtitle="Penerimaan, pengeluaran, transfer antar rekening, dan rekonsiliasi"
        right={<div style={{ display: "flex", gap: 8 }}><Btn variant="ghost" onClick={() => setRecon(!recon)}><ShieldCheck size={14} /> Rekonsiliasi</Btn><Btn variant="accent" onClick={() => setModal(emptyKasbank())}><Plus size={15} /> Transaksi Baru</Btn></div>} />

      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(180px,1fr))", gap: 12, marginBottom: 18 }}>
        <StatCard label="Saldo Kas (sistem)" value={fmtRp(kasSaldo)} icon={Wallet} tone="light" />
        <StatCard label="Saldo Bank (sistem)" value={fmtRp(bankSaldo)} icon={Landmark} tone="light" />
      </div>

      {recon && (
        <Card style={{ padding: 16, marginBottom: 16 }}>
          <div style={{ fontSize: 13, fontWeight: 700, color: C.primaryDark, marginBottom: 10 }}>Rekonsiliasi Bank</div>
          <div style={{ display: "flex", gap: 14, alignItems: "flex-end", flexWrap: "wrap" }}>
            <Field label="Saldo menurut rekening koran (Rp)"><Input type="number" value={saldoBank} onChange={(e) => setSaldoBank(e.target.value)} /></Field>
            <div style={{ paddingBottom: 12 }}>
              <div style={{ fontSize: 12, color: C.inkFaint }}>Saldo sistem: {fmtRp(bankSaldo)}</div>
              <div style={{ fontSize: 14, fontWeight: 700, color: (+saldoBank || 0) - bankSaldo === 0 ? C.success : C.danger }}>
                Selisih: {fmtRp((+saldoBank || 0) - bankSaldo)}
              </div>
            </div>
          </div>
        </Card>
      )}

      <Card>
        <table style={{ width: "100%", borderCollapse: "collapse" }}>
          <thead><tr><Th>Tanggal</Th><Th>Tipe</Th><Th>Akun</Th><Th>Kategori/Tujuan</Th><Th>Keterangan</Th><Th style={{ textAlign: "right" }}>Jumlah</Th><Th style={{ textAlign: "right" }}>Aksi</Th></tr></thead>
          <tbody>
            {sorted.map((k) => (
              <tr key={k.id}>
                <Td>{fmtDate(k.tgl)}</Td>
                <Td><Badge tone={k.tipe === "penerimaan" ? "success" : k.tipe === "pengeluaran" ? "danger" : "primary"}>
                  {k.tipe === "penerimaan" ? <ArrowDownRight size={11} /> : k.tipe === "pengeluaran" ? <ArrowUpRight size={11} /> : <ArrowLeftRight size={11} />} {k.tipe}
                </Badge></Td>
                <Td>{k.akun}</Td>
                <Td>{k.tipe === "transfer" ? `→ ${k.akunTujuan}` : k.kategori}</Td>
                <Td>{k.keterangan || "-"}</Td>
                <TdNum>{fmtRp(k.jumlah)}</TdNum>
                <Td style={{ textAlign: "right" }}>
                  <button style={iconBtnStyle} onClick={() => setModal(k)}><Edit2 size={14} /></button>
                  <button style={{ ...iconBtnStyle, color: C.danger }} onClick={() => handleDelete(k.id)}><Trash2 size={14} /></button>
                </Td>
              </tr>
            ))}
          </tbody>
        </table>
        {sorted.length === 0 && <Empty text="Belum ada transaksi kas & bank." icon={Landmark} />}
      </Card>

      {modal && <KasbankForm initial={modal} onSave={handleSave} onClose={() => setModal(null)} />}
    </div>
  );
}

function KasbankForm({ initial, onSave, onClose }) {
  const [form, setForm] = useState(initial);
  const set = (k) => (e) => setForm({ ...form, [k]: e.target.value });
  useEffect(() => {
    if (form.tipe === "penerimaan" && !KATEGORI_PENERIMAAN.includes(form.kategori)) setForm((f) => ({ ...f, kategori: KATEGORI_PENERIMAAN[0] }));
    if (form.tipe === "pengeluaran" && !KATEGORI_PENGELUARAN.includes(form.kategori)) setForm((f) => ({ ...f, kategori: KATEGORI_PENGELUARAN[0] }));
    // eslint-disable-next-line
  }, [form.tipe]);
  return (
    <Modal title={form.id ? "Edit Transaksi" : "Transaksi Kas & Bank Baru"} onClose={onClose} width={460}>
      <Field label="Tipe Transaksi">
        <Select value={form.tipe} onChange={set("tipe")}><option value="penerimaan">Penerimaan</option><option value="pengeluaran">Pengeluaran</option><option value="transfer">Transfer Antar Rekening</option></Select>
      </Field>
      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: "0 14px" }}>
        <Field label="Tanggal"><Input type="date" value={form.tgl} onChange={set("tgl")} /></Field>
        <Field label={form.tipe === "transfer" ? "Dari Akun" : "Akun"}>
          <Select value={form.akun} onChange={set("akun")}><option>Kas</option><option>Bank</option></Select>
        </Field>
      </div>
      {form.tipe === "transfer"
        ? <Field label="Ke Akun"><Select value={form.akunTujuan} onChange={set("akunTujuan")}><option>Kas</option><option>Bank</option></Select></Field>
        : <Field label="Kategori"><Select value={form.kategori} onChange={set("kategori")}>{(form.tipe === "penerimaan" ? KATEGORI_PENERIMAAN : KATEGORI_PENGELUARAN).map((k) => <option key={k}>{k}</option>)}</Select></Field>}
      <Field label="Jumlah (Rp)"><Input type="number" value={form.jumlah} onChange={set("jumlah")} /></Field>
      <Field label="Keterangan"><TextArea value={form.keterangan} onChange={set("keterangan")} /></Field>
      <div style={{ display: "flex", justifyContent: "flex-end", gap: 8 }}>
        <Btn variant="ghost" onClick={onClose}>Batal</Btn>
        <Btn variant="accent" onClick={() => { if (!form.jumlah) { alert("Isi jumlah."); return; } onSave(form); }}><Save size={14} /> Simpan</Btn>
      </div>
    </Modal>
  );
}

/* ============================== 6. AKUNTANSI ============================== */
function AkuntansiPage({ ctx }) {
  const { journal } = ctx;
  const [tab, setTab] = useState("jurnal");
  const [akun, setAkun] = useState("Kas");
  const [periode, setPeriode] = useState({ start: todayStr().slice(0, 8) + "01", end: todayStr() });

  const akunList = useMemo(() => [...new Set(journal.map((j) => j.akun))].sort(), [journal]);
  const subTabs = [
    { id: "jurnal", label: "Jurnal Umum" },
    { id: "besar", label: "Buku Besar" },
    { id: "necasaldo", label: "Neraca Saldo" },
    { id: "labarugi", label: "Laba Rugi" },
    { id: "neraca", label: "Neraca" },
    { id: "aruskas", label: "Arus Kas" },
  ];

  return (
    <div>
      <SectionTitle icon={BookOpen} title="Akuntansi" subtitle="Jurnal umum otomatis dari seluruh transaksi bisnis" />
      <div className="mfs-noprint" style={{ display: "flex", gap: 6, marginBottom: 16, flexWrap: "wrap" }}>
        {subTabs.map((t) => (
          <button key={t.id} onClick={() => setTab(t.id)} style={{
            padding: "7px 13px", borderRadius: 20, fontSize: 12.5, fontWeight: 600, cursor: "pointer",
            border: `1px solid ${tab === t.id ? C.primary : C.border}`, background: tab === t.id ? C.primary : "#fff", color: tab === t.id ? "#fff" : C.inkSoft,
          }}>{t.label}</button>
        ))}
      </div>

      {tab === "jurnal" && (
        <div>
          <PrintHeader company={ctx.company} title="Jurnal Umum" />
          <ReportToolbar onExcel={() => exportExcel(`jurnal-umum-${todayStr()}.xlsx`, journal.map((j) => ({
            Tanggal: fmtDate(j.tgl), "No. Bukti": j.noBukti, Keterangan: j.keterangan, Akun: j.akun, Debit: j.debit, Kredit: j.kredit,
          })), "Jurnal Umum")} />
          <Card>
            <table style={{ width: "100%", borderCollapse: "collapse" }}>
              <thead><tr><Th>Tanggal</Th><Th>No. Bukti</Th><Th>Keterangan</Th><Th>Akun</Th><Th style={{ textAlign: "right" }}>Debit</Th><Th style={{ textAlign: "right" }}>Kredit</Th></tr></thead>
              <tbody>
                {journal.map((j) => (
                  <tr key={j.id}>
                    <Td>{fmtDate(j.tgl)}</Td><Td style={{ fontFamily: "IBM Plex Mono, monospace", fontSize: 11.5 }}>{j.noBukti}</Td>
                    <Td>{j.keterangan}</Td><Td>{j.akun}</Td>
                    <TdNum>{j.debit ? fmtNum(j.debit) : ""}</TdNum><TdNum>{j.kredit ? fmtNum(j.kredit) : ""}</TdNum>
                  </tr>
                ))}
              </tbody>
            </table>
            {journal.length === 0 && <Empty text="Belum ada transaksi." />}
          </Card>
        </div>
      )}

      {tab === "besar" && (
        <div>
          <PrintHeader company={ctx.company} title="Buku Besar" subtitle={`Akun: ${akun}`} />
          <div className="mfs-noprint" style={{ marginBottom: 12, maxWidth: 260 }}>
            <Select value={akun} onChange={(e) => setAkun(e.target.value)}>{akunList.map((a) => <option key={a}>{a}</option>)}</Select>
          </div>
          <ReportToolbar onExcel={() => {
            const rows = journal.filter((j) => j.akun === akun);
            const kreditNormal = KREDIT_NORMAL.has(akun);
            let saldo = 0;
            exportExcel(`buku-besar-${akun.replace(/\s+/g, "-").toLowerCase()}-${todayStr()}.xlsx`, rows.map((j) => {
              saldo += kreditNormal ? j.kredit - j.debit : j.debit - j.kredit;
              return { Tanggal: fmtDate(j.tgl), Keterangan: j.keterangan, Debit: j.debit, Kredit: j.kredit, Saldo: Math.round(saldo) };
            }), "Buku Besar");
          }} />
          <BukuBesarTable journal={journal} akun={akun} />
        </div>
      )}

      {tab === "necasaldo" && (
        <div>
          <PrintHeader company={ctx.company} title="Neraca Saldo" />
          <ReportToolbar onExcel={() => {
            const akunList2 = [...new Set(journal.map((j) => j.akun))].sort();
            const rows = akunList2.map((a) => {
              const r = journal.filter((j) => j.akun === a);
              const debit = r.reduce((s, j) => s + j.debit, 0);
              const kredit = r.reduce((s, j) => s + j.kredit, 0);
              const kreditNormal = KREDIT_NORMAL.has(a);
              const saldo = kreditNormal ? kredit - debit : debit - kredit;
              return { Akun: a, Debit: kreditNormal ? 0 : Math.round(saldo), Kredit: kreditNormal ? Math.round(saldo) : 0 };
            });
            exportExcel(`neraca-saldo-${todayStr()}.xlsx`, rows, "Neraca Saldo");
          }} />
          <NeracaSaldoTable journal={journal} />
        </div>
      )}

      {tab === "labarugi" && (
        <div>
          <PrintHeader company={ctx.company} title="Laporan Laba Rugi" subtitle={`Periode ${fmtDate(periode.start)} s.d. ${fmtDate(periode.end)}`} />
          <ReportToolbar onExcel={() => {
            const rows = journal.filter((j) => j.tgl >= periode.start && j.tgl <= periode.end);
            const pAkun = [...new Set(rows.filter((j) => isPendapatan(j.akun)).map((j) => j.akun))];
            const bAkun = [...new Set(rows.filter((j) => isBeban(j.akun)).map((j) => j.akun))];
            const data = [
              ...pAkun.map((a) => ({ Kelompok: "Pendapatan", Akun: a, Nilai: Math.round(rows.filter((j) => j.akun === a).reduce((s, j) => s + j.kredit - j.debit, 0)) })),
              ...bAkun.map((a) => ({ Kelompok: "Beban", Akun: a, Nilai: Math.round(rows.filter((j) => j.akun === a).reduce((s, j) => s + j.debit - j.kredit, 0)) })),
            ];
            exportExcel(`laba-rugi-${periode.start}_${periode.end}.xlsx`, data, "Laba Rugi");
          }} />
          <LabaRugiTable journal={journal} periode={periode} setPeriode={setPeriode} />
        </div>
      )}

      {tab === "neraca" && (
        <div>
          <PrintHeader company={ctx.company} title="Neraca" subtitle={`Per tanggal ${fmtDate(periode.end)}`} />
          <ReportToolbar onExcel={() => {
            const rows = journal.filter((j) => j.tgl <= periode.end);
            const bal = (a) => rows.filter((j) => j.akun === a).reduce((s, j) => s + j.debit - j.kredit, 0);
            const balK = (a) => rows.filter((j) => j.akun === a).reduce((s, j) => s + j.kredit - j.debit, 0);
            const piutangNeto = bal("Piutang Pembiayaan") - balK("Margin & Admin Ditangguhkan");
            const pendapatan = rows.filter((j) => isPendapatan(j.akun)).reduce((s, j) => s + j.kredit - j.debit, 0);
            const beban = rows.filter((j) => isBeban(j.akun)).reduce((s, j) => s + j.debit - j.kredit, 0);
            const data = [
              { Kelompok: "Aset", Akun: "Kas", Nilai: Math.round(bal("Kas")) },
              { Kelompok: "Aset", Akun: "Bank", Nilai: Math.round(bal("Bank")) },
              { Kelompok: "Aset", Akun: "Piutang Pembiayaan (neto)", Nilai: Math.round(piutangNeto) },
              { Kelompok: "Modal", Akun: "Modal Disetor", Nilai: Math.round(balK("Modal")) },
              { Kelompok: "Modal", Akun: "Laba Ditahan", Nilai: Math.round(pendapatan - beban) },
            ];
            exportExcel(`neraca-${periode.end}.xlsx`, data, "Neraca");
          }} />
          <NeracaTable journal={journal} asOf={periode.end} setAsOf={(v) => setPeriode({ ...periode, end: v })} />
        </div>
      )}

      {tab === "aruskas" && (
        <div>
          <PrintHeader company={ctx.company} title="Laporan Arus Kas" subtitle={`Periode ${fmtDate(periode.start)} s.d. ${fmtDate(periode.end)}`} />
          <ReportToolbar onExcel={() => {
            const rows = journal.filter((j) => (j.akun === "Kas" || j.akun === "Bank") && j.kategoriKas && j.tgl >= periode.start && j.tgl <= periode.end);
            exportExcel(`arus-kas-${periode.start}_${periode.end}.xlsx`, rows.map((j) => ({
              Tanggal: fmtDate(j.tgl), Kategori: j.kategoriKas, Keterangan: j.keterangan, Nilai: Math.round(j.debit - j.kredit),
            })), "Arus Kas");
          }} />
          <ArusKasTable journal={journal} periode={periode} setPeriode={setPeriode} />
        </div>
      )}
    </div>
  );
}

function BukuBesarTable({ journal, akun }) {
  const rows = journal.filter((j) => j.akun === akun);
  const kreditNormal = KREDIT_NORMAL.has(akun);
  let saldo = 0;
  return (
    <Card>
      <table style={{ width: "100%", borderCollapse: "collapse" }}>
        <thead><tr><Th>Tanggal</Th><Th>Keterangan</Th><Th style={{ textAlign: "right" }}>Debit</Th><Th style={{ textAlign: "right" }}>Kredit</Th><Th style={{ textAlign: "right" }}>Saldo</Th></tr></thead>
        <tbody>
          {rows.map((j) => {
            saldo += kreditNormal ? j.kredit - j.debit : j.debit - j.kredit;
            return (
              <tr key={j.id}><Td>{fmtDate(j.tgl)}</Td><Td>{j.keterangan}</Td><TdNum>{j.debit ? fmtNum(j.debit) : ""}</TdNum><TdNum>{j.kredit ? fmtNum(j.kredit) : ""}</TdNum><TdNum style={{ fontWeight: 700 }}>{fmtNum(saldo)}</TdNum></tr>
            );
          })}
        </tbody>
      </table>
      {rows.length === 0 && <Empty text="Belum ada mutasi pada akun ini." />}
    </Card>
  );
}

function NeracaSaldoTable({ journal }) {
  const akunList = [...new Set(journal.map((j) => j.akun))].sort();
  const totals = akunList.map((a) => {
    const rows = journal.filter((j) => j.akun === a);
    const debit = rows.reduce((s, j) => s + j.debit, 0);
    const kredit = rows.reduce((s, j) => s + j.kredit, 0);
    const kreditNormal = KREDIT_NORMAL.has(a);
    const saldo = kreditNormal ? kredit - debit : debit - kredit;
    return { akun: a, debit: kreditNormal ? 0 : saldo, kredit: kreditNormal ? saldo : 0 };
  });
  const totalD = totals.reduce((s, t) => s + t.debit, 0);
  const totalK = totals.reduce((s, t) => s + t.kredit, 0);
  return (
    <Card>
      <table style={{ width: "100%", borderCollapse: "collapse" }}>
        <thead><tr><Th>Akun</Th><Th style={{ textAlign: "right" }}>Debit</Th><Th style={{ textAlign: "right" }}>Kredit</Th></tr></thead>
        <tbody>
          {totals.map((t) => <tr key={t.akun}><Td>{t.akun}</Td><TdNum>{t.debit ? fmtNum(t.debit) : ""}</TdNum><TdNum>{t.kredit ? fmtNum(t.kredit) : ""}</TdNum></tr>)}
          <tr><Td style={{ fontWeight: 700 }}>Total</Td><TdNum style={{ fontWeight: 700, borderTop: `2px solid ${C.border}` }}>{fmtNum(totalD)}</TdNum><TdNum style={{ fontWeight: 700, borderTop: `2px solid ${C.border}` }}>{fmtNum(totalK)}</TdNum></tr>
        </tbody>
      </table>
    </Card>
  );
}

function PeriodePicker({ periode, setPeriode }) {
  return (
    <div className="mfs-noprint" style={{ display: "flex", gap: 10, marginBottom: 14, alignItems: "center" }}>
      <Input type="date" value={periode.start} onChange={(e) => setPeriode({ ...periode, start: e.target.value })} style={{ width: 160 }} />
      <span style={{ fontSize: 12, color: C.inkFaint }}>s.d.</span>
      <Input type="date" value={periode.end} onChange={(e) => setPeriode({ ...periode, end: e.target.value })} style={{ width: 160 }} />
    </div>
  );
}

function LabaRugiTable({ journal, periode, setPeriode }) {
  const rows = journal.filter((j) => j.tgl >= periode.start && j.tgl <= periode.end);
  const pendapatanAkun = [...new Set(rows.filter((j) => isPendapatan(j.akun)).map((j) => j.akun))];
  const bebanAkun = [...new Set(rows.filter((j) => isBeban(j.akun)).map((j) => j.akun))];
  const pendapatanRows = pendapatanAkun.map((a) => ({ akun: a, nilai: rows.filter((j) => j.akun === a).reduce((s, j) => s + j.kredit - j.debit, 0) }));
  const bebanRows = bebanAkun.map((a) => ({ akun: a, nilai: rows.filter((j) => j.akun === a).reduce((s, j) => s + j.debit - j.kredit, 0) }));
  const totalPendapatan = pendapatanRows.reduce((s, r) => s + r.nilai, 0);
  const totalBeban = bebanRows.reduce((s, r) => s + r.nilai, 0);
  return (
    <div>
      <PeriodePicker periode={periode} setPeriode={setPeriode} />
      <Card style={{ padding: 18, maxWidth: 480 }}>
        <div style={{ fontSize: 12, fontWeight: 700, color: C.primaryDark, marginBottom: 8 }}>Pendapatan</div>
        {pendapatanRows.map((r) => <div key={r.akun} style={{ display: "flex", justifyContent: "space-between", fontSize: 13, padding: "4px 0" }}><span>{r.akun}</span><span style={{ fontFamily: "IBM Plex Mono, monospace" }}>{fmtNum(r.nilai)}</span></div>)}
        <div style={{ display: "flex", justifyContent: "space-between", fontSize: 13, fontWeight: 700, borderTop: `1px solid ${C.border}`, padding: "6px 0", marginBottom: 14 }}><span>Total Pendapatan</span><span>{fmtNum(totalPendapatan)}</span></div>
        <div style={{ fontSize: 12, fontWeight: 700, color: C.primaryDark, marginBottom: 8 }}>Beban</div>
        {bebanRows.map((r) => <div key={r.akun} style={{ display: "flex", justifyContent: "space-between", fontSize: 13, padding: "4px 0" }}><span>{r.akun}</span><span style={{ fontFamily: "IBM Plex Mono, monospace" }}>{fmtNum(r.nilai)}</span></div>)}
        <div style={{ display: "flex", justifyContent: "space-between", fontSize: 13, fontWeight: 700, borderTop: `1px solid ${C.border}`, padding: "6px 0" }}><span>Total Beban</span><span>{fmtNum(totalBeban)}</span></div>
        <div style={{ display: "flex", justifyContent: "space-between", fontSize: 15, fontWeight: 700, marginTop: 12, padding: "10px 0", borderTop: `2px solid ${C.primary}`, color: C.primaryDark }}><span>Laba / Rugi Bersih</span><span>{fmtRp(totalPendapatan - totalBeban)}</span></div>
      </Card>
    </div>
  );
}

function NeracaTable({ journal, asOf, setAsOf }) {
  const rows = journal.filter((j) => j.tgl <= asOf);
  const bal = (akun) => rows.filter((j) => j.akun === akun).reduce((s, j) => s + j.debit - j.kredit, 0);
  const balK = (akun) => rows.filter((j) => j.akun === akun).reduce((s, j) => s + j.kredit - j.debit, 0);
  const kas = bal("Kas"), bank = bal("Bank"), piutangGross = bal("Piutang Pembiayaan"), ditangguhkan = balK("Margin & Admin Ditangguhkan");
  const piutangNeto = piutangGross - ditangguhkan;
  const totalAset = kas + bank + piutangNeto;
  const modal = balK("Modal");
  const pendapatan = rows.filter((j) => isPendapatan(j.akun)).reduce((s, j) => s + j.kredit - j.debit, 0);
  const beban = rows.filter((j) => isBeban(j.akun)).reduce((s, j) => s + j.debit - j.kredit, 0);
  const labaDitahan = pendapatan - beban;
  const totalModal = modal + labaDitahan;
  return (
    <div>
      <div className="mfs-noprint" style={{ marginBottom: 14 }}><Field label="Per tanggal"><Input type="date" value={asOf} onChange={(e) => setAsOf(e.target.value)} style={{ width: 180 }} /></Field></div>
      <Card style={{ padding: 18, maxWidth: 480 }}>
        <div style={{ fontSize: 12, fontWeight: 700, color: C.primaryDark, marginBottom: 8 }}>Aset</div>
        <Row label="Kas" val={kas} /><Row label="Bank" val={bank} /><Row label="Piutang Pembiayaan (neto)" val={piutangNeto} />
        <div style={{ display: "flex", justifyContent: "space-between", fontSize: 13, fontWeight: 700, borderTop: `1px solid ${C.border}`, padding: "6px 0", marginBottom: 14 }}><span>Total Aset</span><span>{fmtNum(totalAset)}</span></div>
        <div style={{ fontSize: 12, fontWeight: 700, color: C.primaryDark, marginBottom: 8 }}>Modal</div>
        <Row label="Modal Disetor" val={modal} /><Row label="Laba Ditahan" val={labaDitahan} />
        <div style={{ display: "flex", justifyContent: "space-between", fontSize: 13, fontWeight: 700, borderTop: `1px solid ${C.border}`, padding: "6px 0" }}><span>Total Modal</span><span>{fmtNum(totalModal)}</span></div>
        <div style={{ marginTop: 10, fontSize: 11, color: Math.abs(totalAset - totalModal) < 1 ? C.success : C.danger }}>
          {Math.abs(totalAset - totalModal) < 1 ? "✓ Neraca seimbang" : `Selisih: ${fmtNum(totalAset - totalModal)}`}
        </div>
      </Card>
    </div>
  );
}
const Row = ({ label, val }) => <div style={{ display: "flex", justifyContent: "space-between", fontSize: 13, padding: "4px 0" }}><span>{label}</span><span style={{ fontFamily: "IBM Plex Mono, monospace" }}>{fmtNum(val)}</span></div>;

function ArusKasTable({ journal, periode, setPeriode }) {
  const rows = journal.filter((j) => (j.akun === "Kas" || j.akun === "Bank") && j.kategoriKas && j.tgl >= periode.start && j.tgl <= periode.end);
  const groups = { operasional: "Aktivitas Operasional", investasi: "Aktivitas Pembiayaan (Investasi)", pendanaan: "Aktivitas Pendanaan" };
  return (
    <div>
      <PeriodePicker periode={periode} setPeriode={setPeriode} />
      <Card style={{ padding: 18, maxWidth: 520 }}>
        {Object.entries(groups).map(([key, label]) => {
          const grouped = rows.filter((j) => j.kategoriKas === key);
          const total = grouped.reduce((s, j) => s + j.debit - j.kredit, 0);
          if (grouped.length === 0) return null;
          return (
            <div key={key} style={{ marginBottom: 14 }}>
              <div style={{ fontSize: 12, fontWeight: 700, color: C.primaryDark, marginBottom: 6 }}>{label}</div>
              {grouped.map((j) => <div key={j.id} style={{ display: "flex", justifyContent: "space-between", fontSize: 12.5, padding: "3px 0", color: C.inkSoft }}><span>{j.keterangan}</span><span style={{ fontFamily: "IBM Plex Mono, monospace" }}>{j.debit - j.kredit >= 0 ? "+" : ""}{fmtNum(j.debit - j.kredit)}</span></div>)}
              <div style={{ display: "flex", justifyContent: "space-between", fontSize: 13, fontWeight: 700, borderTop: `1px solid ${C.borderSoft}`, padding: "4px 0" }}><span>Subtotal</span><span>{fmtNum(total)}</span></div>
            </div>
          );
        })}
        <div style={{ display: "flex", justifyContent: "space-between", fontSize: 15, fontWeight: 700, borderTop: `2px solid ${C.primary}`, padding: "10px 0", color: C.primaryDark }}>
          <span>Kenaikan (Penurunan) Kas & Bank</span><span>{fmtNum(rows.reduce((s, j) => s + j.debit - j.kredit, 0))}</span>
        </div>
        {rows.length === 0 && <Empty text="Tidak ada arus kas pada periode ini." />}
      </Card>
    </div>
  );
}

/* ============================== 7. LAPORAN ============================== */
function LaporanPage({ ctx }) {
  const { nasabahList, akadList, pembayaranList, kasbankList, journal } = ctx;
  const [tab, setTab] = useState("nasabah");
  const [bulan, setBulan] = useState(todayStr().slice(0, 7));
  const today = todayStr();

  const subTabs = [
    { id: "nasabah", label: "Daftar Nasabah" },
    { id: "berjalan", label: "Piutang Berjalan" },
    { id: "macet", label: "Piutang Macet" },
    { id: "jatuhtempo", label: "Angsuran Jatuh Tempo" },
    { id: "margin", label: "Pendapatan Margin" },
    { id: "rekapkas", label: "Rekap Kas" },
    { id: "bulanan", label: "Laporan Bulanan" },
  ];

  const macetList = useMemo(() => {
    const list = [];
    akadList.forEach((a) => (a.jadwal || []).forEach((j) => {
      const telat = j.status !== "lunas" ? daysBetween(j.tglJatuhTempo, today) : 0;
      if (telat > 7) list.push({ ...j, noAkad: a.noAkad, nasabahNama: a.nasabahNama, telat });
    }));
    return list.sort((a, b) => b.telat - a.telat);
  }, [akadList, today]);

  const jatuhTempoList = useMemo(() => {
    const list = [];
    akadList.forEach((a) => (a.jadwal || []).forEach((j) => {
      if (j.status !== "lunas") list.push({ ...j, noAkad: a.noAkad, nasabahNama: a.nasabahNama });
    }));
    return list.sort((a, b) => (a.tglJatuhTempo < b.tglJatuhTempo ? -1 : 1));
  }, [akadList]);

  const berjalanAkad = akadList.filter((a) => a.status !== "diajukan" && (a.jadwal || []).some((j) => j.status !== "lunas"));

  const marginBulan = pembayaranList.filter((p) => p.tglBayar.slice(0, 7) === bulan);
  const totalMargin = marginBulan.reduce((s, p) => s + p.marginPortion, 0);
  const totalAdmin = marginBulan.reduce((s, p) => s + p.adminPortion, 0);

  const rekapRows = journal.filter((j) => (j.akun === "Kas" || j.akun === "Bank") && j.tgl.slice(0, 7) === bulan);
  const masuk = rekapRows.reduce((s, j) => s + j.debit, 0);
  const keluar = rekapRows.reduce((s, j) => s + j.kredit, 0);

  const bulanRows = journal.filter((j) => j.tgl.slice(0, 7) === bulan);
  const pencairanBulan = akadList.filter((a) => a.tglAkad.slice(0, 7) === bulan).reduce((s, a) => s + (+a.nilaiPembiayaan), 0);
  const diterimaBulan = pembayaranList.filter((p) => p.tglBayar.slice(0, 7) === bulan).reduce((s, p) => s + p.jumlahBayar + p.denda, 0);
  const marginBulanTotal = bulanRows.filter((j) => isPendapatan(j.akun)).reduce((s, j) => s + j.kredit - j.debit, 0);
  const bebanBulanTotal = bulanRows.filter((j) => isBeban(j.akun)).reduce((s, j) => s + j.debit - j.kredit, 0);

  return (
    <div>
      <SectionTitle icon={ClipboardList} title="Laporan" subtitle="Ringkasan operasional dan piutang" />
      <div className="mfs-noprint" style={{ display: "flex", gap: 6, marginBottom: 16, flexWrap: "wrap" }}>
        {subTabs.map((t) => (
          <button key={t.id} onClick={() => setTab(t.id)} style={{
            padding: "7px 13px", borderRadius: 20, fontSize: 12.5, fontWeight: 600, cursor: "pointer",
            border: `1px solid ${tab === t.id ? C.primary : C.border}`, background: tab === t.id ? C.primary : "#fff", color: tab === t.id ? "#fff" : C.inkSoft,
          }}>{t.label}</button>
        ))}
      </div>

      {tab === "nasabah" && (
        <div>
          <PrintHeader company={ctx.company} title="Daftar Nasabah" />
          <ReportToolbar onExcel={() => exportExcel(`daftar-nasabah-${todayStr()}.xlsx`, nasabahList.map((n) => ({
            Nama: n.nama, NIK: n.nik || "", Telepon: n.telp || "", Pekerjaan: n.pekerjaan || "", Alamat: n.alamat || "",
          })), "Daftar Nasabah")} />
          <Card><table style={{ width: "100%", borderCollapse: "collapse" }}>
            <thead><tr><Th>Nama</Th><Th>NIK</Th><Th>Telepon</Th><Th>Pekerjaan</Th></tr></thead>
            <tbody>{nasabahList.map((n) => <tr key={n.id}><Td>{n.nama}</Td><Td>{n.nik || "-"}</Td><Td>{n.telp || "-"}</Td><Td>{n.pekerjaan || "-"}</Td></tr>)}</tbody>
          </table>{nasabahList.length === 0 && <Empty text="Belum ada nasabah." />}</Card>
        </div>
      )}

      {tab === "berjalan" && (
        <div>
          <PrintHeader company={ctx.company} title="Piutang Berjalan" />
          <ReportToolbar onExcel={() => exportExcel(`piutang-berjalan-${todayStr()}.xlsx`, berjalanAkad.map((a) => ({
            "No. Akad": a.noAkad, Nasabah: a.nasabahNama, "Sisa Pokok": Math.round(outstandingPokok(a)), "Tenor Sisa (bln)": a.jadwal.filter((j) => j.status !== "lunas").length,
          })), "Piutang Berjalan")} />
          <Card><table style={{ width: "100%", borderCollapse: "collapse" }}>
            <thead><tr><Th>No. Akad</Th><Th>Nasabah</Th><Th style={{ textAlign: "right" }}>Sisa Pokok</Th><Th>Tenor Sisa</Th></tr></thead>
            <tbody>{berjalanAkad.map((a) => <tr key={a.id}><Td>{a.noAkad}</Td><Td>{a.nasabahNama}</Td><TdNum>{fmtRp(outstandingPokok(a))}</TdNum><Td>{a.jadwal.filter((j) => j.status !== "lunas").length} bln</Td></tr>)}</tbody>
          </table>{berjalanAkad.length === 0 && <Empty text="Tidak ada piutang berjalan." />}</Card>
        </div>
      )}

      {tab === "macet" && (
        <div>
          <PrintHeader company={ctx.company} title="Piutang Macet" subtitle="Keterlambatan lebih dari 7 hari" />
          <ReportToolbar onExcel={() => exportExcel(`piutang-macet-${todayStr()}.xlsx`, macetList.map((j) => ({
            "No. Akad": j.noAkad, Nasabah: j.nasabahNama, "Angsuran Ke": j.no, "Jatuh Tempo": fmtDate(j.tglJatuhTempo), Nilai: Math.round(j.angsuran), "Terlambat (hari)": j.telat,
          })), "Piutang Macet")} />
          <Card><table style={{ width: "100%", borderCollapse: "collapse" }}>
            <thead><tr><Th>No. Akad</Th><Th>Nasabah</Th><Th>Angsuran ke</Th><Th>Jatuh Tempo</Th><Th style={{ textAlign: "right" }}>Nilai</Th><Th>Keterlambatan</Th></tr></thead>
            <tbody>{macetList.map((j, i) => <tr key={i}><Td>{j.noAkad}</Td><Td>{j.nasabahNama}</Td><Td>{j.no}</Td><Td>{fmtDate(j.tglJatuhTempo)}</Td><TdNum>{fmtRp(j.angsuran)}</TdNum><Td><Badge tone="danger">{j.telat} hari</Badge></Td></tr>)}</tbody>
          </table>{macetList.length === 0 && <Empty text="Tidak ada piutang macet (lebih dari 7 hari terlambat)." />}</Card>
        </div>
      )}

      {tab === "jatuhtempo" && (
        <div>
          <PrintHeader company={ctx.company} title="Angsuran Jatuh Tempo" />
          <ReportToolbar onExcel={() => exportExcel(`angsuran-jatuh-tempo-${todayStr()}.xlsx`, jatuhTempoList.map((j) => ({
            "No. Akad": j.noAkad, Nasabah: j.nasabahNama, "Angsuran Ke": j.no, "Jatuh Tempo": fmtDate(j.tglJatuhTempo), Nilai: Math.round(j.angsuran),
          })), "Jatuh Tempo")} />
          <Card><table style={{ width: "100%", borderCollapse: "collapse" }}>
            <thead><tr><Th>No. Akad</Th><Th>Nasabah</Th><Th>Angsuran ke</Th><Th>Jatuh Tempo</Th><Th style={{ textAlign: "right" }}>Nilai</Th></tr></thead>
            <tbody>{jatuhTempoList.map((j, i) => <tr key={i}><Td>{j.noAkad}</Td><Td>{j.nasabahNama}</Td><Td>{j.no}</Td><Td>{fmtDate(j.tglJatuhTempo)}</Td><TdNum>{fmtRp(j.angsuran)}</TdNum></tr>)}</tbody>
          </table>{jatuhTempoList.length === 0 && <Empty text="Tidak ada angsuran mendatang." />}</Card>
        </div>
      )}

      {(tab === "margin" || tab === "rekapkas" || tab === "bulanan") && (
        <div className="mfs-noprint" style={{ marginBottom: 14 }}><Field label="Bulan"><Input type="month" value={bulan} onChange={(e) => setBulan(e.target.value)} style={{ width: 180 }} /></Field></div>
      )}

      {tab === "margin" && (
        <div>
          <PrintHeader company={ctx.company} title="Pendapatan Margin" subtitle={fmtMonthYear(bulan)} />
          <ReportToolbar onExcel={() => exportExcel(`pendapatan-margin-${bulan}.xlsx`, marginBulan.map((p) => ({
            Tanggal: fmtDate(p.tglBayar), "No. Akad": p.noAkad, Nasabah: p.nasabahNama, Margin: Math.round(p.marginPortion), Admin: Math.round(p.adminPortion),
          })), "Pendapatan Margin")} />
          <Card><table style={{ width: "100%", borderCollapse: "collapse" }}>
            <thead><tr><Th>Tanggal</Th><Th>No. Akad</Th><Th>Nasabah</Th><Th style={{ textAlign: "right" }}>Margin</Th><Th style={{ textAlign: "right" }}>Admin</Th></tr></thead>
            <tbody>
              {marginBulan.map((p) => <tr key={p.id}><Td>{fmtDate(p.tglBayar)}</Td><Td>{p.noAkad}</Td><Td>{p.nasabahNama}</Td><TdNum>{fmtNum(p.marginPortion)}</TdNum><TdNum>{fmtNum(p.adminPortion)}</TdNum></tr>)}
              <tr><Td style={{ fontWeight: 700 }} colSpan={3}>Total {fmtMonthYear(bulan)}</Td><TdNum style={{ fontWeight: 700 }}>{fmtNum(totalMargin)}</TdNum><TdNum style={{ fontWeight: 700 }}>{fmtNum(totalAdmin)}</TdNum></tr>
            </tbody>
          </table>{marginBulan.length === 0 && <Empty text="Tidak ada pendapatan margin pada bulan ini." />}</Card>
        </div>
      )}

      {tab === "rekapkas" && (
        <div>
          <PrintHeader company={ctx.company} title="Rekap Kas" subtitle={fmtMonthYear(bulan)} />
          <ReportToolbar onExcel={() => exportExcel(`rekap-kas-${bulan}.xlsx`, [
            { Keterangan: "Total Penerimaan", Nilai: Math.round(masuk) },
            { Keterangan: "Total Pengeluaran", Nilai: Math.round(keluar) },
            { Keterangan: "Selisih (Net)", Nilai: Math.round(masuk - keluar) },
          ], "Rekap Kas")} />
          <Card style={{ padding: 18, maxWidth: 420 }}>
            <Row label="Total Penerimaan" val={masuk} /><Row label="Total Pengeluaran" val={keluar} />
            <div style={{ display: "flex", justifyContent: "space-between", fontSize: 14, fontWeight: 700, borderTop: `1px solid ${C.border}`, padding: "8px 0", marginTop: 6 }}><span>Selisih (Net)</span><span>{fmtNum(masuk - keluar)}</span></div>
          </Card>
        </div>
      )}

      {tab === "bulanan" && (
        <div>
          <PrintHeader company={ctx.company} title="Laporan Bulanan" subtitle={fmtMonthYear(bulan)} />
          <ReportToolbar onExcel={() => exportExcel(`laporan-bulanan-${bulan}.xlsx`, [
            { Keterangan: "Pencairan Pembiayaan", Nilai: Math.round(pencairanBulan) },
            { Keterangan: "Angsuran Diterima", Nilai: Math.round(diterimaBulan) },
            { Keterangan: "Pendapatan", Nilai: Math.round(marginBulanTotal) },
            { Keterangan: "Beban", Nilai: Math.round(bebanBulanTotal) },
            { Keterangan: "Laba Bersih", Nilai: Math.round(marginBulanTotal - bebanBulanTotal) },
          ], "Laporan Bulanan")} />
          <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(190px,1fr))", gap: 12 }}>
            <StatCard label="Pencairan Pembiayaan" value={fmtRp(pencairanBulan)} tone="light" />
            <StatCard label="Angsuran Diterima" value={fmtRp(diterimaBulan)} tone="light" />
            <StatCard label="Pendapatan" value={fmtRp(marginBulanTotal)} tone="primary" />
            <StatCard label="Beban" value={fmtRp(bebanBulanTotal)} tone="light" />
            <StatCard label="Laba Bersih" value={fmtRp(marginBulanTotal - bebanBulanTotal)} tone="accent" note={fmtMonthYear(bulan)} />
          </div>
        </div>
      )}
    </div>
  );
}

/* ============================== 8. PENGATURAN ============================== */
function PengaturanPage({ ctx }) {
  const [tab, setTab] = useState("perusahaan");
  const subTabs = [
    { id: "perusahaan", label: "Data Perusahaan" },
    { id: "user", label: "User & Hak Akses" },
    { id: "backup", label: "Backup & Restore" },
  ];
  return (
    <div>
      <SectionTitle icon={Settings} title="Pengaturan" subtitle="Konfigurasi perusahaan, pengguna, dan cadangan data" />
      <div style={{ display: "flex", gap: 6, marginBottom: 16 }}>
        {subTabs.map((t) => (
          <button key={t.id} onClick={() => setTab(t.id)} style={{
            padding: "7px 13px", borderRadius: 20, fontSize: 12.5, fontWeight: 600, cursor: "pointer",
            border: `1px solid ${tab === t.id ? C.primary : C.border}`, background: tab === t.id ? C.primary : "#fff", color: tab === t.id ? "#fff" : C.inkSoft,
          }}>{t.label}</button>
        ))}
      </div>
      {tab === "perusahaan" && <CompanySettings ctx={ctx} />}
      {tab === "user" && <UserSettings ctx={ctx} />}
      {tab === "backup" && <BackupSettings ctx={ctx} />}
    </div>
  );
}

function CompanySettings({ ctx }) {
  const { company, saveCompany } = ctx;
  const [form, setForm] = useState(company);
  const set = (k) => (e) => setForm({ ...form, [k]: e.target.value });
  return (
    <Card style={{ padding: 20, maxWidth: 520 }}>
      <Field label="Nama Perusahaan"><Input value={form.nama} onChange={set("nama")} /></Field>
      <Field label="Alamat"><TextArea value={form.alamat} onChange={set("alamat")} /></Field>
      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: "0 14px" }}>
        <Field label="Telepon"><Input value={form.telp} onChange={set("telp")} /></Field>
        <Field label="NPWP"><Input value={form.npwp} onChange={set("npwp")} /></Field>
      </div>
      <div style={{ borderTop: `1px dashed ${C.border}`, margin: "8px 0 14px", paddingTop: 12, fontSize: 12, fontWeight: 700, color: C.primaryDark }}>Modal Usaha</div>
      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: "0 14px" }}>
        <Field label="Modal Awal (Rp)" hint="Dicatat sebagai setoran modal di jurnal & kas."><Input type="number" value={form.modalAwal} onChange={set("modalAwal")} /></Field>
        <Field label="Tanggal Setoran Modal"><Input type="date" value={form.tglModal} onChange={set("tglModal")} /></Field>
      </div>
      <Btn variant="accent" onClick={() => saveCompany({ ...form, modalAwal: +form.modalAwal || 0 })}><Save size={14} /> Simpan Pengaturan</Btn>
    </Card>
  );
}

function UserSettings({ ctx }) {
  const [profiles, setProfiles] = useState(null);
  const [showAdd, setShowAdd] = useState(false);
  const [form, setForm] = useState({ nama: "", email: "", password: "", role: "admin" });
  const [error, setError] = useState("");
  const [busy, setBusy] = useState(false);

  const load = () => { sbSelectAll("mfs_profiles").then(setProfiles).catch((e) => { console.error(e); setProfiles([]); }); };
  useEffect(load, []);

  const add = async () => {
    setError(""); setBusy(true);
    try {
      if (!form.nama || !form.email || !form.password) throw new Error("Lengkapi semua kolom.");
      if (form.password.length < 6) throw new Error("Kata sandi minimal 6 karakter.");
      const tok = await authSignUp(form.email, form.password); // does not affect the currently logged-in session
      if (!tok.user) throw new Error("Pendaftaran gagal.");
      await sbInsert("mfs_profiles", { id: tok.user.id, nama: form.nama, email: form.email.toLowerCase(), role: form.role });
      setForm({ nama: "", email: "", password: "", role: "admin" });
      setShowAdd(false);
      load();
    } catch (e) { setError(e.message); }
    setBusy(false);
  };

  const remove = async (id) => {
    if (id === ctx.currentUser.id) { alert("Tidak dapat menghapus akun yang sedang digunakan."); return; }
    if (!confirm("Hapus akun pengguna ini? Akses login pengguna tersebut akan dinonaktifkan.")) return;
    try { await sbDelete("mfs_profiles", id); load(); }
    catch (e) { alert(e.message); }
  };

  return (
    <div>
      <Card style={{ padding: 16, marginBottom: 16, maxWidth: 520 }}>
        {!showAdd ? (
          <Btn variant="accent" onClick={() => setShowAdd(true)}><UserPlus size={14} /> Tambah Pengguna</Btn>
        ) : (
          <div>
            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: "0 14px" }}>
              <Field label="Nama"><Input value={form.nama} onChange={(e) => setForm({ ...form, nama: e.target.value })} /></Field>
              <Field label="Peran"><Select value={form.role} onChange={(e) => setForm({ ...form, role: e.target.value })}><option value="admin">Admin</option><option value="owner">Owner</option></Select></Field>
              <Field label="Email"><Input type="email" value={form.email} onChange={(e) => setForm({ ...form, email: e.target.value })} /></Field>
              <Field label="Kata Sandi"><Input type="password" value={form.password} onChange={(e) => setForm({ ...form, password: e.target.value })} /></Field>
            </div>
            {error && <div style={{ background: C.dangerSoft, color: C.danger, fontSize: 12, borderRadius: 6, padding: "8px 10px", marginBottom: 10 }}>{error}</div>}
            <div style={{ display: "flex", gap: 8 }}>
              <Btn variant="ghost" onClick={() => { setShowAdd(false); setError(""); }}>Batal</Btn>
              <Btn variant="accent" onClick={add} disabled={busy}><Save size={14} /> Simpan</Btn>
            </div>
          </div>
        )}
      </Card>
      <Card>
        <table style={{ width: "100%", borderCollapse: "collapse" }}>
          <thead><tr><Th>Nama</Th><Th>Email</Th><Th>Peran</Th><Th style={{ textAlign: "right" }}>Aksi</Th></tr></thead>
          <tbody>{(profiles || []).map((u) => (
            <tr key={u.id}>
              <Td>{u.nama} {u.id === ctx.currentUser.id && <span style={{ fontSize: 10, color: C.inkFaint }}>(Anda)</span>}</Td>
              <Td>{u.email}</Td>
              <Td><Badge tone="primary">{u.role}</Badge></Td>
              <Td style={{ textAlign: "right" }}><button style={{ ...iconBtnStyle, color: C.danger }} onClick={() => remove(u.id)}><Trash2 size={14} /></button></Td>
            </tr>
          ))}</tbody>
        </table>
        {profiles === null && <Empty text="Memuat pengguna…" icon={Loader2} />}
        {profiles && profiles.length === 0 && <Empty text="Belum ada pengguna terdaftar." icon={ShieldCheck} />}
      </Card>
    </div>
  );
}

function BackupSettings({ ctx }) {
  const { company, users, nasabahList, akadList, pembayaranList, kasbankList, restoreAll } = ctx;
  const [importing, setImporting] = useState(false);

  const doExport = () => {
    const data = { company, users, nasabahList, akadList, pembayaranList, kasbankList, exportedAt: new Date().toISOString() };
    const blob = new Blob([JSON.stringify(data, null, 2)], { type: "application/json" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url; a.download = `mfs-backup-${todayStr()}.json`; a.click();
    URL.revokeObjectURL(url);
  };

  const doImport = (e) => {
    const file = e.target.files[0];
    if (!file) return;
    setImporting(true);
    const reader = new FileReader();
    reader.onload = () => {
      try { restoreAll(JSON.parse(reader.result)); alert("Data berhasil dipulihkan."); }
      catch (err) { alert("File backup tidak valid."); }
      setImporting(false);
    };
    reader.readAsText(file);
  };

  return (
    <Card style={{ padding: 20, maxWidth: 480 }}>
      <div style={{ display: "flex", gap: 10, alignItems: "flex-start", marginBottom: 18 }}>
        <DatabaseBackup size={20} color={C.accent} style={{ marginTop: 2 }} />
        <div>
          <div style={{ fontWeight: 700, fontSize: 13.5, marginBottom: 3 }}>Cadangkan seluruh data</div>
          <div style={{ fontSize: 12, color: C.inkFaint }}>Unduh seluruh data nasabah, akad, pembayaran, kas & bank, serta pengaturan sebagai satu berkas JSON.</div>
        </div>
      </div>
      <Btn variant="accent" onClick={doExport} style={{ marginBottom: 22 }}><Download size={14} /> Unduh Backup (.json)</Btn>

      <div style={{ borderTop: `1px dashed ${C.border}`, paddingTop: 18 }}>
        <div style={{ fontWeight: 700, fontSize: 13.5, marginBottom: 3 }}>Pulihkan data</div>
        <div style={{ fontSize: 12, color: C.inkFaint, marginBottom: 10 }}>Mengunggah berkas backup akan menimpa data yang sedang tersimpan.</div>
        <label style={{ display: "inline-flex", alignItems: "center", gap: 6, padding: "9px 14px", borderRadius: 7, border: `1px solid ${C.border}`, cursor: "pointer", fontSize: 13, fontWeight: 600, color: C.primary }}>
          {importing ? <Loader2 size={14} /> : <Upload size={14} />} Pilih Berkas Backup
          <input type="file" accept="application/json" onChange={doImport} style={{ display: "none" }} />
        </label>
      </div>
    </Card>
  );
}
