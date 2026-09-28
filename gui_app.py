"""
Aplikasi GUI Desktop Generator Laporan Keuangan Word
Menggunakan Tkinter dan python-docx.
"""

import os
import sys
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from generator_docx import generate_report_docx, format_rupiah

class LaporanApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Generator Laporan Keuangan Bulanan ke Word")
        self.geometry("900x720")
        self.minsize(800, 650)

        # Style
        self.style = ttk.Style(self)
        self.style.theme_use("clam")

        self.list_pengeluaran = [
            {"nama": "Aceng", "nominal": 700000},
            {"nama": "Didin", "nominal": 150000},
            {"nama": "Eko", "nominal": 150000},
            {"nama": "Bu Yam", "nominal": 200000},
            {"nama": "10 Dus Tens", "nominal": 270000},
            {"nama": "Bu Hamina", "nominal": 200000},
            {"nama": "Listrik", "nominal": 352000},
        ]

        self.daftar_bulan = [
            "Januari", "Februari", "Maret", "April", "Mei", "Juni",
            "Juli", "Agustus", "September", "Oktober", "November", "Desember"
        ]

        self.setup_ui()
        self.update_hitung()

    def setup_ui(self):
        # Header banner
        header = tk.Frame(self, bg="#2563eb", height=60)
        header.pack(fill=tk.X)
        header_lbl = tk.Label(
            header, text="Generator Laporan Keuangan ke Word",
            bg="#2563eb", fg="white", font=("Arial", 14, "bold")
        )
        header_lbl.pack(side=tk.LEFT, padx=20, pady=12)

        # Main scrollable canvas / container
        container = tk.Frame(self, padx=20, pady=15)
        container.pack(fill=tk.BOTH, expand=True)

        # SECTION 1: Periode
        f_periode = ttk.LabelFrame(container, text=" 1. Periode Laporan ", padding=10)
        f_periode.pack(fill=tk.X, pady=(0, 10))

        ttk.Label(f_periode, text="Bulan Sekarang:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=4)
        self.var_bulan = tk.StringVar(value="Agustus")
        cb_bulan = ttk.Combobox(f_periode, textvariable=self.var_bulan, values=self.daftar_bulan, state="readonly", width=14)
        cb_bulan.grid(row=0, column=1, sticky=tk.W, padx=5, pady=4)
        cb_bulan.bind("<<ComboboxSelected>>", self.on_bulan_change)

        ttk.Label(f_periode, text="Tahun:").grid(row=0, column=2, sticky=tk.W, padx=(20, 5), pady=4)
        self.var_tahun = tk.StringVar(value="2026")
        ent_tahun = ttk.Entry(f_periode, textvariable=self.var_tahun, width=8)
        ent_tahun.grid(row=0, column=3, sticky=tk.W, padx=5, pady=4)

        ttk.Label(f_periode, text="Bulan Kemarin:").grid(row=0, column=4, sticky=tk.W, padx=(20, 5), pady=4)
        self.var_bulan_kemarin = tk.StringVar(value="juni")
        ent_kemarin = ttk.Entry(f_periode, textvariable=self.var_bulan_kemarin, width=12)
        ent_kemarin.grid(row=0, column=5, sticky=tk.W, padx=5, pady=4)

        ttk.Label(f_periode, text="Sub-Judul Lembaga:").grid(row=1, column=0, sticky=tk.W, padx=5, pady=4)
        self.var_sub_judul = tk.StringVar(value="Pengurus Musholla ....")
        ent_sub = ttk.Entry(f_periode, textvariable=self.var_sub_judul, width=35)
        ent_sub.grid(row=1, column=1, columnspan=3, sticky=tk.W, padx=5, pady=4)

        # SECTION 2: Saldo & Pemasukan
        f_saldo = ttk.LabelFrame(container, text=" 2. Saldo Awal & Saldo Masuk ", padding=10)
        f_saldo.pack(fill=tk.X, pady=(0, 10))

        ttk.Label(f_saldo, text="Sisa Saldo Bulan Kemarin (Rp):").grid(row=0, column=0, sticky=tk.W, padx=5, pady=4)
        self.var_saldo_kemarin = tk.StringVar(value="-166000")
        ent_sk = ttk.Entry(f_saldo, textvariable=self.var_saldo_kemarin, width=18)
        ent_sk.grid(row=0, column=1, sticky=tk.W, padx=5, pady=4)
        ent_sk.bind("<KeyRelease>", lambda e: self.update_hitung())

        ttk.Label(f_saldo, text="Saldo Masuk Bulan Ini (Rp):").grid(row=0, column=2, sticky=tk.W, padx=(20, 5), pady=4)
        self.var_saldo_masuk = tk.StringVar(value="1600000")
        ent_sm = ttk.Entry(f_saldo, textvariable=self.var_saldo_masuk, width=18)
        ent_sm.grid(row=0, column=3, sticky=tk.W, padx=5, pady=4)
        ent_sm.bind("<KeyRelease>", lambda e: self.update_hitung())

        self.lbl_total_masuk = ttk.Label(f_saldo, text="Total Pemasukan: Rp. 1.434.000", font=("Arial", 9, "bold"), foreground="#2563eb")
        self.lbl_total_masuk.grid(row=0, column=4, sticky=tk.W, padx=(25, 5), pady=4)

        # SECTION 3: Daftar Pengeluaran
        f_pengeluaran = ttk.LabelFrame(container, text=" 3. Daftar Pengeluaran ", padding=10)
        f_pengeluaran.pack(fill=tk.BOTH, expand=True, pady=(0, 10))

        # Form tambah baris
        f_add = tk.Frame(f_pengeluaran)
        f_add.pack(fill=tk.X, pady=(0, 6))

        ttk.Label(f_add, text="Nama Pengeluaran:").pack(side=tk.LEFT, padx=(0, 5))
        self.var_new_nama = tk.StringVar()
        self.ent_new_nama = ttk.Entry(f_add, textvariable=self.var_new_nama, width=22)
        self.ent_new_nama.pack(side=tk.LEFT, padx=(0, 15))

        ttk.Label(f_add, text="Nominal (Rp):").pack(side=tk.LEFT, padx=(0, 5))
        self.var_new_nom = tk.StringVar()
        self.ent_new_nom = ttk.Entry(f_add, textvariable=self.var_new_nom, width=16)
        self.ent_new_nom.pack(side=tk.LEFT, padx=(0, 15))
        self.ent_new_nom.bind("<Return>", lambda e: self.tambah_item())

        btn_tambah = ttk.Button(f_add, text="➕ Tambah Baris", command=self.tambah_item)
        btn_tambah.pack(side=tk.LEFT, padx=(0, 10))

        btn_hapus = ttk.Button(f_add, text="🗑️ Hapus Baris Terpilih", command=self.hapus_item)
        btn_hapus.pack(side=tk.LEFT)

        # Treeview tabel
        tree_frame = tk.Frame(f_pengeluaran)
        tree_frame.pack(fill=tk.BOTH, expand=True)

        cols = ("no", "nama", "nominal")
        self.tree = ttk.Treeview(tree_frame, columns=cols, show="headings", height=7)
        self.tree.heading("no", text="No")
        self.tree.heading("nama", text="Nama Pengeluaran")
        self.tree.heading("nominal", text="Nominal (Rp)")
        
        self.tree.column("no", width=40, anchor=tk.CENTER)
        self.tree.column("nama", width=380, anchor=tk.W)
        self.tree.column("nominal", width=180, anchor=tk.E)

        scrollbar = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)

        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        self.render_table()

        # Rekap baris bawah
        f_rekap = tk.Frame(f_pengeluaran, pady=5)
        f_rekap.pack(fill=tk.X)

        self.lbl_total_keluar = ttk.Label(f_rekap, text="Total Pengeluaran: Rp. 2.022.000", font=("Arial", 9, "bold"))
        self.lbl_total_keluar.pack(side=tk.LEFT, padx=(0, 30))

        self.lbl_saldo_akhir = ttk.Label(f_rekap, text="Sisa Saldo Akhir: Rp. -588.000", font=("Arial", 10, "bold"), foreground="#dc2626")
        self.lbl_saldo_akhir.pack(side=tk.LEFT)

        # SECTION 4: Tanda Tangan
        f_ttd = ttk.LabelFrame(container, text=" 4. Penandatangan ", padding=10)
        f_ttd.pack(fill=tk.X, pady=(0, 10))

        ttk.Label(f_ttd, text="Jabatan Kiri:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=2)
        self.var_jabatan_kiri = tk.StringVar(value="ketua")
        ttk.Entry(f_ttd, textvariable=self.var_jabatan_kiri, width=16).grid(row=0, column=1, sticky=tk.W, padx=5, pady=2)

        ttk.Label(f_ttd, text="Nama Kiri:").grid(row=0, column=2, sticky=tk.W, padx=(20, 5), pady=2)
        self.var_nama_kiri = tk.StringVar(value="Syarif Hidayat")
        ttk.Entry(f_ttd, textvariable=self.var_nama_kiri, width=22).grid(row=0, column=3, sticky=tk.W, padx=5, pady=2)

        ttk.Label(f_ttd, text="Jabatan Kanan:").grid(row=1, column=0, sticky=tk.W, padx=5, pady=2)
        self.var_jabatan_kanan = tk.StringVar(value="Bendahara")
        ttk.Entry(f_ttd, textvariable=self.var_jabatan_kanan, width=16).grid(row=1, column=1, sticky=tk.W, padx=5, pady=2)

        ttk.Label(f_ttd, text="Nama Kanan:").grid(row=1, column=2, sticky=tk.W, padx=(20, 5), pady=2)
        self.var_nama_kanan = tk.StringVar(value="Ahmad Rizal")
        ttk.Entry(f_ttd, textvariable=self.var_nama_kanan, width=22).grid(row=1, column=3, sticky=tk.W, padx=5, pady=2)

        # Bottom buttons
        f_btn = tk.Frame(container)
        f_btn.pack(fill=tk.X, pady=(5, 0))

        btn_word = tk.Button(
            f_btn, text="📄 Buat & Buka File Word (.docx)", bg="#2563eb", fg="white",
            font=("Arial", 11, "bold"), padx=15, pady=8, relief=tk.FLAT, cursor="hand2",
            command=self.simpan_word
        )
        btn_word.pack(side=tk.LEFT, padx=(0, 10))

        btn_contoh = ttk.Button(f_btn, text="📋 Muat Contoh Sesuai Gambar", command=self.muat_contoh)
        btn_contoh.pack(side=tk.LEFT, padx=(0, 10))

    def on_bulan_change(self, event=None):
        bulan = self.var_bulan.get()
        if bulan in self.daftar_bulan:
            idx = self.daftar_bulan.index(bulan)
            prev_idx = (idx - 1 + 12) % 12
            self.var_bulan_kemarin.set(self.daftar_bulan[prev_idx].lower())

    def render_table(self):
        for item in self.tree.get_children():
            self.tree.delete(item)
        for i, row in enumerate(self.list_pengeluaran, 1):
            self.tree.insert("", tk.END, values=(i, row["nama"], f"Rp. {format_rupiah(row['nominal'])}"))

    def tambah_item(self):
        nama = self.var_new_nama.get().strip()
        nom_str = self.var_new_nom.get().replace(".", "").replace(",", "").strip()
        if not nama:
            messagebox.showwarning("Peringatan", "Nama pengeluaran tidak boleh kosong.")
            self.ent_new_nama.focus()
            return
        try:
            nom = int(nom_str)
        except ValueError:
            messagebox.showwarning("Peringatan", "Nominal harus berupa angka valid.")
            self.ent_new_nom.focus()
            return

        self.list_pengeluaran.append({"nama": nama, "nominal": nom})
        self.var_new_nama.set("")
        self.var_new_nom.set("")
        self.ent_new_nama.focus()
        self.render_table()
        self.update_hitung()

    def hapus_item(self):
        selected = self.tree.selection()
        if not selected:
            messagebox.showinfo("Info", "Pilih baris yang ingin dihapus pada tabel.")
            return
        idx = self.tree.index(selected[0])
        del self.list_pengeluaran[idx]
        self.render_table()
        self.update_hitung()

    def update_hitung(self):
        try:
            sk = int(self.var_saldo_kemarin.get().replace(".", "").strip() or 0)
        except:
            sk = 0
        try:
            sm = int(self.var_saldo_masuk.get().replace(".", "").strip() or 0)
        except:
            sm = 0
        tot_masuk = sk + sm
        self.lbl_total_masuk.config(text=f"Total Pemasukan: Rp. {format_rupiah(tot_masuk)}")

        tot_keluar = sum(item["nominal"] for item in self.list_pengeluaran)
        self.lbl_total_keluar.config(text=f"Total Pengeluaran: Rp. {format_rupiah(tot_keluar)}")

        saldo_akhir = tot_masuk - tot_keluar
        self.lbl_saldo_akhir.config(
            text=f"Sisa Saldo Akhir: Rp. {format_rupiah(saldo_akhir)}",
            foreground="#dc2626" if saldo_akhir < 0 else "#16a34a"
        )

    def muat_contoh(self):
        self.var_bulan.set("Agustus")
        self.var_tahun.set("2026")
        self.var_sub_judul.set("Pengurus Musholla ....")
        self.var_bulan_kemarin.set("juni")
        self.var_saldo_kemarin.set("-166000")
        self.var_saldo_masuk.set("1600000")
        self.var_jabatan_kiri.set("ketua")
        self.var_nama_kiri.set("Syarif Hidayat")
        self.var_jabatan_kanan.set("Bendahara")
        self.var_nama_kanan.set("Ahmad Rizal")

        self.list_pengeluaran = [
            {"nama": "Aceng", "nominal": 700000},
            {"nama": "Didin", "nominal": 150000},
            {"nama": "Eko", "nominal": 150000},
            {"nama": "Bu Yam", "nominal": 200000},
            {"nama": "10 Dus Tens", "nominal": 270000},
            {"nama": "Bu Hamina", "nominal": 200000},
            {"nama": "Listrik", "nominal": 352000},
        ]
        self.render_table()
        self.update_hitung()
        messagebox.showinfo("Berhasil", "Data contoh sesuai gambar telah berhasil dimuat!")

    def simpan_word(self):
        bulan = self.var_bulan.get()
        tahun = self.var_tahun.get()
        try:
            sk = int(self.var_saldo_kemarin.get().replace(".", "").strip() or 0)
            sm = int(self.var_saldo_masuk.get().replace(".", "").strip() or 0)
        except ValueError:
            messagebox.showerror("Error", "Saldo kemarin atau saldo masuk harus berupa angka!")
            return

        default_name = f"Laporan_Keuangan_{bulan}_{tahun}.docx"
        file_path = filedialog.asksaveasfilename(
            defaultextension=".docx",
            filetypes=[("Microsoft Word Document", "*.docx")],
            initialfile=default_name
        )
        if not file_path:
            return

        data = {
            "bulanSekarang": bulan,
            "tahun": tahun,
            "subJudul": self.var_sub_judul.get(),
            "bulanKemarin": self.var_bulan_kemarin.get(),
            "saldoKemarin": sk,
            "saldoMasuk": sm,
            "pengeluaran": self.list_pengeluaran,
            "jabatanKiri": self.var_jabatan_kiri.get(),
            "namaKiri": self.var_nama_kiri.get(),
            "jabatanKanan": self.var_jabatan_kanan.get(),
            "namaKanan": self.var_nama_kanan.get(),
            "fontDocx": "Arial"
        }

        try:
            generate_report_docx(data, file_path)
            res = messagebox.askyesno(
                "Berhasil!",
                f"File Word berhasil dibuat di:\n{file_path}\n\nApakah ingin langsung membukanya?"
            )
            if res:
                os.startfile(file_path)
        except Exception as e:
            messagebox.showerror("Gagal", f"Terjadi kesalahan: {e}")

if __name__ == "__main__":
    app = LaporanApp()
    app.mainloop()
