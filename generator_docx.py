"""
Generator Dokumen Word (.docx) untuk Laporan Keuangan Bulanan
Disesuaikan persis dengan template laporan keuangan kas bulanan (format akuntansi kas).
"""

import sys
import os
import json
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import parse_xml, OxmlElement
from docx.oxml.ns import nsdecls, qn

def format_rupiah(num):
    if num is None:
        return "0"
    is_neg = num < 0
    abs_val = abs(int(round(num)))
    formatted = f"{abs_val:,}".replace(",", ".")
    return f"-{formatted}" if is_neg else formatted

def set_cell_borders(cell, top=False, bottom=False, left=False, right=False, border_sz="12", border_color="000000"):
    """Mengatur border spesifik pada sel tabel docx."""
    tcPr = cell._tc.get_or_add_tcPr()
    
    # Hapus tcBorders jika ada
    existing_borders = tcPr.find(qn('w:tcBorders'))
    if existing_borders is not None:
        tcPr.remove(existing_borders)
        
    b_top = f'<w:top w:val="single" w:sz="{border_sz}" w:space="0" w:color="{border_color}"/>' if top else '<w:top w:val="none"/>'
    b_bottom = f'<w:bottom w:val="single" w:sz="{border_sz}" w:space="0" w:color="{border_color}"/>' if bottom else '<w:bottom w:val="none"/>'
    b_left = f'<w:left w:val="single" w:sz="{border_sz}" w:space="0" w:color="{border_color}"/>' if left else '<w:left w:val="none"/>'
    b_right = f'<w:right w:val="single" w:sz="{border_sz}" w:space="0" w:color="{border_color}"/>' if right else '<w:right w:val="none"/>'
    
    xml_str = f'<w:tcBorders {nsdecls("w")}>{b_top}{b_left}{b_bottom}{b_right}</w:tcBorders>'
    tcPr.append(parse_xml(xml_str))

def add_table_row(table, col_widths, desc, eq, rp, val, underline=False, bold=False, font_name="Arial", font_size=11):
    row = table.add_row()
    cells = row.cells
    
    # Set lebar
    for i, w in enumerate(col_widths):
        cells[i].width = Inches(w)
        
    # Text runs
    texts = [desc, eq, rp, val]
    aligns = [WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.RIGHT]
    
    for i in range(4):
        p = cells[i].paragraphs[0]
        p.alignment = aligns[i]
        p.paragraph_format.space_before = Pt(1)
        p.paragraph_format.space_after = Pt(1)
        p.paragraph_format.line_spacing = 1.15
        
        run = p.add_run(texts[i])
        run.bold = bold
        run.font.name = font_name
        run.font.size = Pt(font_size)
        
        # Atur border: jika ada underline (garis jumlah), beri border bawah pada kolom perhitungan (=, Rp, dan angka)
        if underline and i >= 1:
            set_cell_borders(cells[i], bottom=True)
        else:
            set_cell_borders(cells[i])
            
    return row

def generate_report_docx(data, output_filepath=None):
    """
    Membuat file docx laporan keuangan berdasarkan dict data.
    """
    bulan_sekarang = data.get("bulanSekarang", "Agustus")
    tahun = str(data.get("tahun", "2026"))
    bulan_kemarin = data.get("bulanKemarin", "juni")
    saldo_kemarin = int(data.get("saldoKemarin", -166000))
    saldo_masuk = int(data.get("saldoMasuk", 1600000))
    pengeluaran = data.get("pengeluaran", [])
    
    jabatan_kiri = data.get("jabatanKiri", "ketua")
    nama_kiri = data.get("namaKiri", "Syarif Hidayat")
    jabatan_kanan = data.get("jabatanKanan", "Bendahara")
    nama_kanan = data.get("namaKanan", "Ahmad Rizal")
    
    font_name = data.get("fontDocx", "Arial")
    font_size = 11

    doc = docx.Document()

    # Set page margin standar 1 inci (2.54 cm)
    for section in doc.sections:
        section.top_margin = Inches(1.0)
        section.bottom_margin = Inches(1.0)
        section.left_margin = Inches(1.0)
        section.right_margin = Inches(1.0)

    # 1. Header Judul
    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_title.paragraph_format.space_after = Pt(20)
    run_title = p_title.add_run(f"{bulan_sekarang} {tahun}")
    run_title.bold = True
    run_title.font.name = font_name
    run_title.font.size = Pt(14)

    # Lebar kolom dalam inci (Total ~ 5.5 inci)
    col_widths = [3.0, 0.4, 0.5, 1.4]

    # SECTION 1: Saldo Awal & Saldo Masuk
    t1 = doc.add_table(rows=0, cols=4)
    t1.alignment = WD_TABLE_ALIGNMENT.LEFT
    
    total_pemasukan = saldo_kemarin + saldo_masuk
    add_table_row(t1, col_widths, f"Sisa Saldo Bulan {bulan_kemarin} {tahun}", "=", "Rp.", format_rupiah(saldo_kemarin), underline=False, bold=False, font_name=font_name)
    add_table_row(t1, col_widths, f"Saldo Masuk Bulan {bulan_sekarang}", "=", "Rp.", format_rupiah(saldo_masuk), underline=True, bold=False, font_name=font_name)
    add_table_row(t1, col_widths, "Total", "=", "Rp.", format_rupiah(total_pemasukan), underline=False, bold=True, font_name=font_name)

    # Spacing antar tabel
    p_sep1 = doc.add_paragraph()
    p_sep1.paragraph_format.space_before = Pt(14)
    p_sep1.paragraph_format.space_after = Pt(0)

    # SECTION 2: Daftar Pengeluaran
    t2 = doc.add_table(rows=0, cols=4)
    t2.alignment = WD_TABLE_ALIGNMENT.LEFT
    
    total_pengeluaran = 0
    num_items = len(pengeluaran)
    for idx, item in enumerate(pengeluaran):
        nom = int(item.get("nominal", 0))
        total_pengeluaran += nom
        is_last = (idx == num_items - 1)
        add_table_row(t2, col_widths, item.get("nama", ""), "=", "Rp.", format_rupiah(nom), underline=is_last, bold=False, font_name=font_name)
        
    add_table_row(t2, col_widths, "", "", "Rp.", format_rupiah(total_pengeluaran), underline=False, bold=True, font_name=font_name)

    # Spacing antar tabel
    p_sep2 = doc.add_paragraph()
    p_sep2.paragraph_format.space_before = Pt(14)
    p_sep2.paragraph_format.space_after = Pt(0)

    # SECTION 3: Rekapitulasi (Pemasukan vs Pengeluaran)
    t3 = doc.add_table(rows=0, cols=4)
    t3.alignment = WD_TABLE_ALIGNMENT.LEFT
    
    saldo_akhir = total_pemasukan - total_pengeluaran
    add_table_row(t3, col_widths, "Pemasukan", "=", "Rp.", format_rupiah(total_pemasukan), underline=False, bold=False, font_name=font_name)
    add_table_row(t3, col_widths, "Pengeluaran", "=", "Rp.", format_rupiah(total_pengeluaran), underline=True, bold=False, font_name=font_name)
    add_table_row(t3, col_widths, "", "", "Rp.", format_rupiah(saldo_akhir), underline=False, bold=True, font_name=font_name)

    # Spacing ke tanda tangan
    p_sep3 = doc.add_paragraph()
    p_sep3.paragraph_format.space_before = Pt(36)
    p_sep3.paragraph_format.space_after = Pt(0)

    # SECTION 4: Tanda Tangan (2 Kolom)
    t_sig = doc.add_table(rows=3, cols=2)
    t_sig.alignment = WD_TABLE_ALIGNMENT.CENTER
    sig_widths = [3.2, 3.2]

    # Baris 0: Jabatan
    for i, w in enumerate(sig_widths):
        t_sig.rows[0].cells[i].width = Inches(w)
        set_cell_borders(t_sig.rows[0].cells[i])
        p = t_sig.rows[0].cells[i].paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run(jabatan_kiri if i == 0 else jabatan_kanan)
        run.bold = True
        run.font.name = font_name
        run.font.size = Pt(font_size)

    # Baris 1: Ruang tanda tangan (kosong setinggi ~60 pt)
    for i, w in enumerate(sig_widths):
        t_sig.rows[1].cells[i].width = Inches(w)
        set_cell_borders(t_sig.rows[1].cells[i])
        p = t_sig.rows[1].cells[i].paragraphs[0]
        p.paragraph_format.space_before = Pt(45)
        p.paragraph_format.space_after = Pt(0)

    # Baris 2: Nama
    for i, w in enumerate(sig_widths):
        t_sig.rows[2].cells[i].width = Inches(w)
        set_cell_borders(t_sig.rows[2].cells[i])
        p = t_sig.rows[2].cells[i].paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run(nama_kiri if i == 0 else nama_kanan)
        run.bold = True
        run.font.name = font_name
        run.font.size = Pt(font_size)

    if output_filepath is None:
        output_filepath = f"Laporan_Keuangan_{bulan_sekarang}_{tahun}.docx"

    doc.save(output_filepath)
    return output_filepath

if __name__ == "__main__":
    contoh_data = {
        "bulanSekarang": "Agustus",
        "tahun": "2026",
        "bulanKemarin": "juni",
        "saldoKemarin": -166000,
        "saldoMasuk": 1600000,
        "pengeluaran": [
            {"nama": "Aceng", "nominal": 700000},
            {"nama": "Didin", "nominal": 150000},
            {"nama": "Eko", "nominal": 150000},
            {"nama": "Bu Yam", "nominal": 200000},
            {"nama": "10 Dus Tens", "nominal": 270000},
            {"nama": "Bu Hamina", "nominal": 200000},
            {"nama": "Listrik", "nominal": 352000},
        ],
        "jabatanKiri": "ketua",
        "namaKiri": "Syarif Hidayat",
        "jabatanKanan": "Bendahara",
        "namaKanan": "Ahmad Rizal",
        "fontDocx": "Arial"
    }

    out = generate_report_docx(contoh_data, "Laporan_Keuangan_Agustus_2026.docx")
    print(f"[OK] Berhasil membuat file Word: {out}")
