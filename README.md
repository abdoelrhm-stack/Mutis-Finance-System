# Mutis Finance System (MFS)

Aplikasi internal manajemen pembiayaan/kredit untuk **Mutis Finance** — multi-user,
role-based access (Owner/Admin), real-time data via Supabase, dan akuntansi
double-entry otomatis.

## ✨ Modul

1. **Dashboard** — KPI & grafik pendapatan margin
2. **Data Nasabah** — manajemen nasabah + upload dokumen (KTP/KK)
3. **Pengajuan Pembiayaan** — akad & generate jadwal angsuran otomatis
4. **Pembayaran Angsuran** — pencatatan pembayaran + cetak bukti
5. **Kas & Bank** — penerimaan, pengeluaran, transfer, rekonsiliasi bank
6. **Akuntansi** — jurnal umum, buku besar, neraca saldo, laba rugi, neraca, arus kas (otomatis dari seluruh transaksi)
7. **Laporan** — 7 jenis laporan dengan export Excel & cetak/PDF
8. **Pengaturan** — data perusahaan, user & hak akses, backup/restore (khusus Owner)

Formula angsuran:
```
cicilan = (nilai_pembiayaan + margin + biaya_admin − DP) ÷ lama_angsuran
```

## 🧱 Struktur Folder

```
mutis-finance-system/
├── index.html                 # HTML entry Vite
├── package.json
├── vite.config.js
├── jsconfig.json              # path alias & intellisense editor
├── vercel.json                # konfigurasi deploy Vercel (SPA rewrite)
├── .env.example                # contoh environment variable
├── .gitignore
├── README.md
├── public/
│   └── favicon.svg
├── supabase/
│   └── schema.sql             # skema tabel + RLS + seed owner (opsional)
└── src/
    ├── main.jsx                # entry point React
    ├── App.jsx                 # seluruh aplikasi (8 modul MFS)
    ├── index.css                # style global
    └── lib/
        ├── config.js            # baca env var Supabase
        └── storage.js           # shim window.storage (localStorage) untuk sesi login
```

## 🚀 Menjalankan secara lokal

Prasyarat: Node.js 18+ dan npm.

```bash
# 1. Install dependencies
npm install

# 2. Salin file environment
cp .env.example .env
# lalu isi VITE_SUPABASE_URL dan VITE_SUPABASE_ANON_KEY

# 3. Jalankan dev server
npm run dev
```

Aplikasi akan berjalan di `http://localhost:5173`.

## 🗄️ Setup Database Supabase

1. Buat project baru di [Supabase](https://supabase.com) (atau gunakan project yang sudah ada).
2. Buka **SQL Editor**, jalankan isi file `supabase/schema.sql` — ini membuat semua tabel
   (`mfs_profiles`, `mfs_company`, `mfs_users`, `mfs_nasabah`, `mfs_akad`, `mfs_pembayaran`, `mfs_kasbank`)
   beserta Row Level Security (RLS) yang membatasi akses hanya untuk pengguna yang sudah login (`authenticated`).
3. **Aktifkan Email Auth** di Supabase Dashboard → Authentication → Providers.
4. Untuk membuat akun **Owner** pertama kali, ada dua cara:
   - **Cara mudah:** daftar lewat halaman "Daftar Akun" di aplikasi, pilih peran **Owner**.
   - **Cara manual (SQL):** gunakan blok SQL yang di-comment di bagian bawah `supabase/schema.sql`
     (isi `auth.users` + `auth.identities` + `mfs_profiles` sekaligus).
5. Salin **Project URL** dan **anon public key** dari Settings → API ke file `.env` Anda.

> ⚠️ Catatan penting terkait password hashing: fungsi `extensions.crypt()` dan
> `extensions.gen_salt('bf')` memerlukan prefix `extensions.` secara eksplisit.
> Jika Anda membuat RPC function kustom yang perlu bypass RLS, gunakan pola
> `SECURITY DEFINER` dengan `search_path` diset ke `public, extensions`.

## ☁️ Deploy ke Vercel

### Opsi A — via Vercel CLI
```bash
npm i -g vercel
vercel
```

### Opsi B — via Git + Vercel Dashboard
1. Push folder ini ke repository Git (GitHub/GitLab/Bitbucket).
2. Import project di [vercel.com/new](https://vercel.com/new).
3. Vercel akan otomatis mendeteksi framework **Vite** (build command `npm run build`, output `dist`).
4. Tambahkan environment variables di **Project Settings → Environment Variables**:
   - `VITE_SUPABASE_URL`
   - `VITE_SUPABASE_ANON_KEY`
5. Deploy.

`vercel.json` sudah menyertakan rewrite SPA (`/* -> /index.html`) agar routing client-side tidak 404 saat refresh.

## 🔐 Role & Akses

| Fitur                         | Owner | Admin |
|-------------------------------|:-----:|:-----:|
| Dashboard, Nasabah, Pembiayaan, Pembayaran, Kas & Bank, Akuntansi, Laporan | ✅ | ✅ |
| Pengaturan (perusahaan, user, backup/restore) | ✅ | ❌ |

## 🛠️ Tech Stack

- **Frontend:** React 18, Vite, [lucide-react](https://lucide.dev) (ikon), [recharts](https://recharts.org) (grafik), [SheetJS/xlsx](https://sheetjs.com) (export Excel)
- **Backend:** Supabase (Postgres + Auth/GoTrue + Row Level Security)
- **Deployment:** Vercel

## 📦 Backup & Restore Data

Menu **Pengaturan → Backup & Restore** (khusus Owner) memungkinkan:
- **Export**: mengunduh seluruh data (perusahaan, user, nasabah, akad, pembayaran, kas & bank) sebagai satu file `.json`.
- **Import**: memulihkan data dari file backup `.json` (akan menimpa data yang sedang tersimpan di Supabase).

## 📝 Catatan Pengembangan

- Semua nilai akuntansi (jurnal umum, buku besar, neraca saldo, laba rugi, neraca, arus kas)
  diturunkan secara otomatis dan real-time dari data sumber (akad, pembayaran, kas & bank) melalui
  fungsi `computeJournal()` di `src/App.jsx` — bukan disimpan sebagai tabel terpisah, sehingga
  selalu konsisten (single source of truth).
- `src/lib/storage.js` menyediakan shim `window.storage` (berbasis `localStorage`) yang meniru API
  penyimpanan sandbox Claude.ai Artifacts, digunakan untuk menyimpan refresh token sesi login antar kunjungan browser.
