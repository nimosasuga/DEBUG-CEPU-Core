# C.E.P.U (Catatan Evaluasi Pekerja Utama) Enterprise OS

**C.E.P.U Enterprise Operating System** adalah solusi HRIS (Human Resources Information System) terintegrasi berbasis Google Apps Script (GAS) yang dirancang khusus untuk manajemen perjalanan dinas, pelacakan kinerja mekanik/sales, dan audit log operasional secara *real-time*.

Sistem ini menggunakan arsitektur **Serverless** dengan Google Sheets sebagai basis data utama dan antarmuka **Glassmorphism UI** yang dibangun dengan Vanilla JavaScript murni.

---

## 🚀 Fitur Utama
- **Hardware-Based Security:** Penguncian akun berdasarkan *Browser Fingerprinting* (satu HP satu akun).
- **Hybrid Tracking System:** Validasi keberangkatan menggunakan kombinasi QR-Code Scanner dan Geolocation (GPS) dengan *Haversine Formula*.
- **Smart Dynamic UPD:** Kalkulasi otomatis uang jalan berdasarkan 3 lapis validasi: Jabatan, Golongan, dan Status Proyek.
- **HRD Command Center:** Fitur "Sidang HRD" untuk rekalkulasi anomali jam kerja, *Bulk Approval*, dan sistem gembok cetak (*Print Lock*).
- **Enterprise Reporting:** Ekspor laporan otomatis ke format PDF (Premium Layout) dan Excel (.xlsx) untuk penggajian.
- **PWA Ready:** Mendukung instalasi di layar utama (Add to Home Screen) untuk stabilitas memori di iOS/Safari dan Samsung Browser.

---

## 🛠️ Stack Teknologi
- **Frontend:** HTML5, Tailwind CSS (via CDN), FontAwesome.
- **Logic:** Vanilla JavaScript (ES6+).
- **Backend:** Google Apps Script (GAS).
- **Database:** Google Sheets API.
- **Libraries:**
  - `html2pdf.js` (PDF Generation)
  - `xlsx.js` / `SheetJS` (Excel Processing)
  - `html5-qrcode` (QR/Barcode Scanner)

---

## 🏗️ Aturan Emas Arsitektur (Golden Rules)
Untuk menjaga stabilitas dan skalabilitas sistem, pengembang wajib mengikuti aturan berikut:
1. **Pure DOM Engine:** Dilarang menggunakan `innerHTML` untuk membangun UI utama. Semua elemen harus dibuat menggunakan fungsi `DOM_Engine.create()`.
2. **Zero Regression:** Modifikasi fitur baru tidak boleh merusak logika *auth* dan *ODOC (One Day One Checkin)* yang sudah stabil.
3. **No Frameworks:** Dilarang menggunakan React, Vue, atau jQuery. Gunakan manipulasi DOM murni.
4. **Isolated Config:** Semua ID Spreadsheet harus dikelola secara terpusat di dalam file `BE_Config.gs`.

---

## 📊 Struktur Database
Sistem ini membagi beban kerja ke dalam 4 Spreadsheet utama:
- **DB_MASTER:** Data karyawan, koordinat GPS (Master_Latlong), dan master rate UPD.
- **DB_UPD:** Log harian perjalanan dinas mekanik (RFMC, FMC, Satelite).
- **DB_SALES:** Log khusus aktivitas sales.
- **DB_REKAP:** Matriks absensi bulanan untuk pelaporan HRD.

---

## 📦 Instruksi Instalasi (Lingkungan Pengembangan)
1. **Clone Repository:** Unduh semua file `.html` dan `.gs`.
2. **Duplikasi Database:** Copy file database Master dan UPD di Google Drive Anda.
3. **Konfigurasi ID:** Masukkan ID Spreadsheet baru Anda ke dalam file `BE_Config.gs`.
4. **Deploy:** - Buka Google Apps Script Editor.
   - Klik **Deploy** > **New Deployment**.
   - Pilih **Web App**, set akses ke **Anyone**.
5. **PWA Setup:** Buka URL Web App di HP, gunakan fitur "Add to Home Screen".

---

## 📄 Lisensi & Kontributor
Sistem ini dikembangkan secara internal untuk kebutuhan operasional perusahaan.
- **Lead Architect:** Partner Coding AI & Enterprise Developer.
- **Status:** Production / Stable.
