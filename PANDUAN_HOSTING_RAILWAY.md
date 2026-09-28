# 🚀 Panduan Lengkap Deploy / Hosting ke Railway (Gratis & Cepat)

Aplikasi ini sudah dipersiapkan secara lengkap dengan **Node.js + Express**, konfigurasi `package.json`, `Procfile`, dan `railway.json`. Siap di-deploy ke **Railway.app** hanya dalam 3 menit!

---

## 🌟 Opsi Hasil yang Didapat Pengguna
Setelah website ini online di Railway, siapapun (anggota RT, bendahara, panitia, dll.) bisa:
1. Membuka link website lewat browser di HP / Laptop.
2. Mengisi bulan, saldo masuk, dan daftar pengeluaran.
3. Melihat Live Preview kertas A4 langsung di layar.
4. Memilih tombol unduh:
   - 📄 **Download Word (.docx)** : Menghasilkan file Microsoft Word yang rapi, dapat diedit, dan tabel sejajar.
   - 📕 **Download PDF (.pdf)** : Menghasilkan file PDF resolusi tinggi ukuran A4 siap cetak/kirim WhatsApp.
   - 🖨️ **Cetak** : Cetak langsung ke mesin printer.

---

## 🛠️ Langkah-Langkah Deploy ke Railway

### Cara 1: Menggunakan GitHub (Sangat Direkomendasikan & Otomatis)

1. **Buat Repository Baru di GitHub**:
   - Buka [github.com/new](https://github.com/new).
   - Beri nama repository, misalnya: `laporan-keuangan-word-pdf`.
   - Pilih *Public* atau *Private*, lalu klik **Create repository**.

2. **Upload Folder ini ke GitHub**:
   Buka terminal (PowerShell atau Command Prompt) di dalam folder `Aplikasi_Laporan_Keuangan_Word`, lalu jalankan:
   ```bash
   git init
   git add .
   git commit -m "Upload generator laporan keuangan"
   git branch -M main
   git remote add origin https://github.com/USERNAME-ANDA/laporan-keuangan-word-pdf.git
   git push -u origin main
   ```
   *(Ganti `USERNAME-ANDA` dengan username akun GitHub Anda)*.

3. **Deploy di Railway**:
   - Buka situs [railway.app](https://railway.app) dan login dengan akun GitHub Anda.
   - Klik tombol **"New Project"** (atau **"+ Create a new project"**).
   - Pilih opsi **"Deploy from GitHub repo"**.
   - Pilih repository `laporan-keuangan-word-pdf` yang tadi Anda upload.
   - Klik **Deploy Now**.

4. **Buat Link / Domain Publik**:
   - Tunggu proses build selesai (biasanya hanya 10-20 detik).
   - Klik card layanan aplikasi Anda di dashboard Railway.
   - Buka tab **Settings**.
   - Gulir ke bawah ke bagian **Networking**, lalu klik tombol **"Generate Domain"**.
   - Railway akan memberikan URL publik gratis, contohnya:
     `https://laporan-keuangan-production.up.railway.app`
   - Selesai! Link tersebut sekarang bisa dibagikan dan diakses oleh siapapun!

---

### Cara 2: Deploy Langsung Menggunakan Railway CLI (Tanpa Lewat GitHub)

Jika Anda tidak ingin mengunggah ke GitHub terlebih dahulu, Anda bisa mendeploy langsung dari komputer Anda menggunakan **Railway CLI**:

1. Buka PowerShell di folder `Aplikasi_Laporan_Keuangan_Word`.
2. Install Railway CLI:
   ```bash
   npm install -g @railway/cli
   ```
3. Login ke akun Railway:
   ```bash
   railway login
   ```
   *(Browser akan terbuka untuk konfirmasi login)*.

4. Inisialisasi proyek dan deploy:
   ```bash
   railway init
   railway up
   ```
5. Buat domain publik:
   ```bash
   railway domain
   ```
   Railway akan langsung menampilkan URL website aktif Anda!

---

## ⚙️ Mengapa Arsitektur Ini Sangat Cocok untuk Railway?
- **Sangat Hemat Resource (Gratis)**: Proses konversi Word (.docx) dan PDF (.pdf) dijalankan secara *client-side* di browser pengguna, sehingga server Railway hanya bertugas menyajikan halaman web. Server menggunakan RAM sangat kecil (< 30 MB), anti lelet, dan tidak akan memakan kuota hosting gratis Railway Anda!
- **Auto-Detect**: Railway secara otomatis mengenali file `package.json` dan menjalankan perintah `npm start` (`node server.js`).
- **Support Mobile**: Tampilan responsif otomatis menyesuaikan jika dibuka dari smartphone/tablet.
