# AGE BALI — LPJ Report Builder (Serverless)

Aplikasi web serverless untuk membuat laporan LPJ (Laporan Pertanggungjawaban) **Ayam Guling Enakko**. Upload/paste bukti transfer & nota, lalu export ke **DOCX, XLSX, atau PDF** dengan kualitas profesional.

Versi 2.0 — di-deploy di [Vercel](https://vercel.com) dengan serverless functions, auto-save LocalStorage, dan output yang jauh lebih rapi.

---

## ✨ Fitur

- **Serverless** — DOCX & XLSX dihasilkan oleh Vercel Functions, PDF di-generate di browser.
- **Auto-Save** — Semua data otomatis tersimpan di LocalStorage browser. Refresh halaman tidak akan hilangkan data.
- **Backup / Restore JSON** — Simpan snapshot data ke file untuk pindah perangkat.
- **Smart Image Compression** — Gambar otomatis di-resize ke max 1400px JPEG saat upload (hemat storage & bandwidth).
- **Paste dari Clipboard** — Copy gambar dari Word/Docs lalu Ctrl+V langsung ke zona upload.
- **Drag & Drop** — File gambar bisa di-drop, dan thumbnail bisa di-reorder via drag.
- **DOCX Export** — Tabel rapi dengan zebra-stripe, header berwarna brand, kolom No/Jenis/Transfer/Nota, baris TOTAL, header brand, page numbers.
- **PDF Export** — Multi-column image grid (3 kolom), page header/footer, halaman ringkasan, brand bar di tiap halaman.
- **XLSX Export** — Spreadsheet ringkasan dengan styling profesional dan baris TOTAL.

---

## 🚀 Deploy ke Vercel

### Cara cepat (Recommended)

1. Push repo ke GitHub (sudah selesai jika kamu baca ini di repo `Kaiser117450/lpj`).
2. Buka [vercel.com/new](https://vercel.com/new), pilih repo `Kaiser117450/lpj`.
3. Framework Preset: **Other**. Klik **Deploy**. Selesai.

### Atau via CLI

```bash
npm i -g vercel
vercel login
vercel --prod
```

---

## 💻 Development Lokal

### Prasyarat
- **Node.js** v20+
- **Vercel CLI** (opsional, untuk dev mode)

### Setup

```bash
# Install dependencies
npm install

# Jalankan dev server (membutuhkan Vercel CLI)
npm run dev
# Buka http://localhost:3000
```

Untuk testing tanpa Vercel CLI, kamu bisa serve `public/` via static server biasa — tapi endpoint `/api/export/docx` dan `/api/export/xlsx` tidak akan jalan. PDF export tetap berfungsi karena 100% client-side.

---

## 📁 Struktur Project

```
lpj-report-builder/
├── api/
│   └── export/
│       ├── docx.js       ← Serverless function: generate Word doc
│       └── xlsx.js       ← Serverless function: generate Excel sheet
├── public/
│   └── index.html        ← Frontend (single-page, no build step)
├── package.json
├── vercel.json           ← Vercel runtime config (60s timeout untuk DOCX)
└── README.md
```

---

## 🎨 Kustomisasi Brand

Ganti warna brand di:
- **DOCX**: `api/export/docx.js` → konstanta `BRAND_PRIMARY` / `BRAND_DARK`
- **XLSX**: `api/export/xlsx.js` → konstanta yang sama
- **PDF & UI**: `public/index.html` → CSS `--accent` dan konstanta `C_BRAND` dalam JS `exportPDF`

---

## ⚙️ Konfigurasi

### Batas ukuran upload
Vercel Hobby tier: 4.5MB per request. Gambar di-resize otomatis ke 1400px JPEG quality 0.82 (~80-150KB per gambar), jadi 30+ gambar masih muat. Untuk laporan sangat besar:
- Upgrade ke **Vercel Pro** (50MB limit), atau
- Split menjadi beberapa laporan.

### Batas timeout
`vercel.json` set DOCX timeout ke 60 detik (max Hobby tier). Cukup untuk ratusan gambar.

---

## 🧰 Tech Stack

- **Frontend**: Vanilla HTML + CSS + JS, jsPDF (CDN)
- **Backend**: Vercel Serverless Functions (Node.js 20)
- **Libraries**: `docx` v9, `exceljs` v4
- **Hosting**: Vercel

---

## 📝 Lisensi

MIT — Internal AGE BALI.
