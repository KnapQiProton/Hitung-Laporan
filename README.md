# Aplikasi Generator Laporan Keuangan Bulanan ke Word (.docx)

Aplikasi ini dibuat khusus untuk mempermudah pembuatan laporan keuangan bulanan secara otomatis dengan format rapi persis seperti dokumen pembukuan kas yang biasa Anda buat.

---

## 🌟 Fitur Utama

1. **Input Sangat Mudah & Cepat**:
   - Pilihan bulan & tahun laporan.
   - Sisa saldo bulan kemarin (bisa bernilai minus `-` jika kas tekor).
   - Saldo masuk bulan berjalan.
   - Daftar pengeluaran dinamis (bisa tambah/hapus baris sesuka hati).
   - Format rupiah otomatis (ketik angka biasa, langsung berubah rapi menjadi `Rp. 1.600.000`).
   - Tekan **Enter** di kolom nominal untuk langsung menambah baris baru.

2. **Kalkulasi Otomatis 100% Akurat**:
   - **Total Pemasukan** = `Sisa Saldo Bulan Kemarin + Saldo Masuk Bulan Ini`
   - **Total Pengeluaran** = Jumlah semua pos pengeluaran
   - **Sisa Saldo Akhir** = `Total Pemasukan - Total Pengeluaran`

3. **Live Preview (Pratinjau Kertas Langsung)**:
   - Di sebelah kanan layar, Anda langsung bisa melihat bentuk laporan di atas kertas A4 sebelum mencetak atau mendownload.

4. **1-Klik Download ke Word (.docx)**:
   - Dokumen Word asli yang dapat dibuka dan diedit di Microsoft Word, WPS Office, Google Docs, atau LibreOffice.
   - Menggunakan tabel tanpa garis tepi (borderless), dengan garis bawah khusus pada bagian penjumlahan/rekapitulasi sehingga angka rata kanan dengan rapi dan tidak akan berantakan.

5. **Fitur "Lanjut Bulan Berikutnya"**:
   - Saat berganti bulan, cukup klik tombol **Lanjut Bulan Berikutnya**.
   - Saldo akhir bulan ini akan otomatis dipindahkan menjadi saldo kemarin untuk bulan baru!
   - Nama Ketua dan Bendahara tetap tersimpan tanpa perlu diketik ulang.

6. **100% Offline**:
   - Tidak memerlukan koneksi internet untuk menghasilkan dokumen Word.

---

## 🚀 Cara Menggunakan

### Cara 1: Menggunakan Aplikasi Web (Paling Praktis & Direkomendasikan)
1. Cukup klik ganda (double-click) file:
   **`Buka_Aplikasi.bat`** (atau buka langsung file `index.html` dengan Google Chrome / Microsoft Edge).
2. Masukkan data keuangan bulan Anda.
3. Klik tombol biru **"Download File Word (.docx)"**.
4. Dokumen Word siap dicetak atau dikirim!

### Cara 2: Menggunakan Aplikasi Desktop Python
Jika Anda lebih suka tampilan window desktop software:
1. Klik ganda file:
   **`Buka_Aplikasi_Desktop.bat`** (atau jalankan perintah `python gui_app.py`).
2. Masukkan data pengeluaran.
3. Klik **"Buat & Buka File Word (.docx)"**.

### Cara 3: Menggunakan Perintah Python Langsung
Jika ingin membuat file Word langsung lewat script Python:
```bash
python generator_docx.py
```

---

## 📁 Struktur File
- `index.html` : Halaman aplikasi web utama dengan Live Preview dan export Word.
- `Buka_Aplikasi.bat` : Shortcut 1-klik untuk membuka aplikasi di browser.
- `docx.umd.js` & `FileSaver.min.js` : Library offline untuk membuat file Word.
- `gui_app.py` : Aplikasi desktop Python berbasis GUI.
- `Buka_Aplikasi_Desktop.bat` : Shortcut 1-klik untuk aplikasi desktop.
- `generator_docx.py` : Modul pembuat file Word menggunakan pustaka python-docx.
