import streamlit as st
import pandas as pd
from fpdf import FPDF
from PIL import Image
from datetime import datetime
import os

# ==========================
# Helper: hitung tinggi teks
# ==========================
def get_num_lines(pdf, text, col_width):
    words = str(text).split(" ")
    line = ""
    lines = 1
    for word in words:
        if pdf.get_string_width(line + " " + word) <= col_width - 2:
            line += " " + word
        else:
            lines += 1
            line = word
    return lines

# ==========================
# Buat Voucher Jurnal
# ==========================
def buat_voucher(df, no_voucher, settings):
    pdf = FPDF(orientation="P", unit="mm", format="A4")
    pdf.add_page()
    pdf.set_font("Arial", "B", 12)

    # Header Logo & Perusahaan
    if settings.get("logo"):
        pdf.image(settings["logo"], 10, 8, 20)
    pdf.cell(0, 10, settings.get("nama_perusahaan", ""), ln=1, align="C")
    pdf.set_font("Arial", "", 10)
    pdf.multi_cell(0, 5, settings.get("alamat", ""), align="C")
    pdf.ln(5)

    # Judul
    pdf.set_font("Arial", "B", 12)
    pdf.cell(0, 10, "BUKTI VOUCHER JURNAL", ln=1, align="C")
    pdf.set_font("Arial", "", 10)
    pdf.cell(0, 7, f"No Voucher : {no_voucher}", ln=1, align="C")
    pdf.ln(3)

    # ==========================
    # Layout Tabel
    # ==========================
    page_width = 210
    margin = 10
    usable_width = page_width - 2 * margin
    fixed = [25, 25, 25, 25]  # Tanggal, Akun, Debit, Kredit
    fixed_total = sum(fixed)
    nama_akun_width = 45
    deskripsi_width = usable_width - (fixed_total + nama_akun_width)
    col_widths = [25, 25, nama_akun_width, deskripsi_width, 25, 25]

    headers = ["Tanggal", "Akun", "Nama Akun", "Deskripsi", "Debit", "Kredit"]

    pdf.set_font("Arial", "B", 9)
    for i, h in enumerate(headers):
        pdf.cell(col_widths[i], 8, h, border=1, align="C")
    pdf.ln()

    pdf.set_font("Arial", "", 9)

    total_debit, total_kredit = 0, 0

    for _, row in df.iterrows():
        debit_val = int(row.get("Debet", 0)) if pd.notna(row.get("Debet")) else 0
        kredit_val = int(row.get("Kredit", 0)) if pd.notna(row.get("Kredit")) else 0

        # format tanggal dd/mm/yyyy
        try:
            tgl = pd.to_datetime(row.get("Tanggal")).strftime("%d/%m/%Y")
        except:
            tgl = str(row.get("Tanggal"))

        values = [
            tgl,
            str(row.get("No Akun", "")),
            str(row.get("Akun", "")),   # Nama Akun
            str(row.get("Deskripsi", "")),
            f"{debit_val:,}".replace(",", "."),
            f"{kredit_val:,}".replace(",", ".")
        ]

        # Hitung tinggi baris berdasarkan Nama Akun & Deskripsi
        n_lines_nama = get_num_lines(pdf, values[2], col_widths[2])
        n_lines_desc = get_num_lines(pdf, values[3], col_widths[3])
        row_height = max(n_lines_nama, n_lines_desc) * 6

        # Simpan posisi awal
        x = pdf.get_x()
        y = pdf.get_y()

        # Tanggal & Akun (cell biasa)
        pdf.cell(col_widths[0], row_height, values[0], border=1, align="L")
        pdf.cell(col_widths[1], row_height, values[1], border=1, align="L")

        # Nama Akun (wrap)
        x2, y2 = pdf.get_x(), pdf.get_y()
        pdf.multi_cell(col_widths[2], 6, values[2], border=1, align="L")
        pdf.set_xy(x2 + col_widths[2], y2)

        # Deskripsi (wrap)
        x3, y3 = pdf.get_x(), pdf.get_y()
        pdf.multi_cell(col_widths[3], 6, values[3], border=1, align="L")
        pdf.set_xy(x3 + col_widths[3], y3)

        # Debit & Kredit (cell biasa)
        pdf.cell(col_widths[4], row_height, values[4], border=1, align="R")
        pdf.cell(col_widths[5], row_height, values[5], border=1, align="R")
        pdf.ln(row_height)

        total_debit += debit_val
        total_kredit += kredit_val

    # Total
    pdf.set_font("Arial", "B", 9)
    pdf.cell(col_widths[0] + col_widths[1] + col_widths[2] + col_widths[3], 8, "Total", border=1, align="R")
    pdf.cell(col_widths[4], 8, f"{total_debit:,}".replace(",", "."), border=1, align="R")
    pdf.cell(col_widths[5], 8, f"{total_kredit:,}".replace(",", "."), border=1, align="R")
    pdf.ln()

    # Terbilang
    pdf.set_font("Arial", "I", 8)
    pdf.cell(0, 8, f"Terbilang: {angka_terbilang(total_debit)} rupiah", ln=1)

    # ==========================
    # Tanda tangan format resmi
    # ==========================
    pdf.ln(15)
    pdf.set_font("Arial", "", 10)
    pdf.cell(95, 6, "Direktur,", align="C")
    pdf.cell(95, 6, "Finance,", align="C", ln=1)
    pdf.ln(20)  # spasi tanda tangan
    pdf.cell(95, 6, f"({settings.get('direktur','')})", align="C")
    pdf.cell(95, 6, f"({settings.get('finance','')})", align="C", ln=1)

    filename = f"voucher_{no_voucher.replace('/', '-')}.pdf"
    pdf.output(filename)
    return filename

# ==========================
# Angka ke Terbilang
# ==========================
def angka_terbilang(n):
    from num2words import num2words
    return num2words(n, lang="id").replace("koma nol", "")

# ==========================
# STREAMLIT APP
# ==========================
st.title("📑 Mini Akunting - Voucher Jurnal")

# Sidebar Settings (langsung ambil value)
st.sidebar.header("⚙️ Pengaturan Perusahaan")
nama_perusahaan = st.sidebar.text_input("Nama Perusahaan")
alamat = st.sidebar.text_area("Alamat Perusahaan")
direktur = st.sidebar.text_input("Nama Direktur")
finance = st.sidebar.text_input("Nama Finance")
logo_file = st.sidebar.file_uploader("Upload Logo (PNG/JPG)", type=["png", "jpg", "jpeg"])

# Upload file
uploaded_file = st.file_uploader("Upload Daftar Jurnal (Excel)", type=["xlsx", "xls", "csv"])

if uploaded_file:
    df = pd.read_excel(uploaded_file) if uploaded_file.name.endswith("xlsx") else pd.read_csv(uploaded_file)
    st.dataframe(df.head())

    vouchers = df["Nomor Voucher Jurnal"].unique().tolist()
    pilih_no = st.selectbox("Pilih Nomor Voucher", vouchers)

    if st.button("Cetak Voucher Jurnal"):
        # Ambil settings saat tombol ditekan
        settings = {
            "nama_perusahaan": nama_perusahaan,
            "alamat": alamat,
            "direktur": direktur,
            "finance": finance
        }
        if logo_file:
            img = Image.open(logo_file)
            logo_path = "logo.png"
            img.save(logo_path)
            settings["logo"] = logo_path

        subset = df[df["Nomor Voucher Jurnal"] == pilih_no]
        file_pdf = buat_voucher(subset, pilih_no, settings)
        with open(file_pdf, "rb") as f:
            st.download_button("⬇️ Download Voucher PDF", f, file_name=file_pdf)
