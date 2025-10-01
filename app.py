import streamlit as st
import pandas as pd
from fpdf import FPDF
from num2words import num2words
import base64
from io import BytesIO
from datetime import datetime

# ==============================
# Fungsi format angka
# ==============================
def format_rupiah(x):
    try:
        return f"{int(x):,}".replace(",", ".")
    except:
        return "0"

# ==============================
# Fungsi konversi ke terbilang
# ==============================
def terbilang_rupiah(x):
    try:
        return num2words(int(x), lang='id') + " rupiah"
    except:
        return "nol rupiah"

# ==============================
# Fungsi buat voucher ke PDF
# ==============================
def buat_voucher(df, voucher_no, perusahaan, alamat, direktur, finance, logo=None):
    buffer = BytesIO()
    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Arial", "B", 12)

    # Header dengan logo
    if logo:
        pdf.image(logo, 10, 8, 25)
        pdf.set_xy(40, 10)
    else:
        pdf.set_xy(10, 10)

    pdf.multi_cell(0, 8, perusahaan, align="L")
    pdf.set_font("Arial", "", 10)
    pdf.multi_cell(0, 6, alamat, align="L")
    pdf.ln(4)

    # Judul
    pdf.set_font("Arial", "B", 12)
    pdf.cell(0, 10, "BUKTI VOUCHER JURNAL", align="C", ln=True)
    pdf.set_font("Arial", "", 10)
    pdf.cell(0, 8, f"No Voucher : {voucher_no}", align="C", ln=True)
    pdf.ln(4)

    # Table Header
    pdf.set_font("Arial", "B", 9)
    col_widths = [25, 20, 55, 50, 20, 20]
    headers = ["Tanggal", "Akun", "Nama Akun", "Deskripsi", "Debit", "Kredit"]

    for i, h in enumerate(headers):
        pdf.cell(col_widths[i], 8, h, border=1, align="C")
    pdf.ln()

    # Isi tabel
    pdf.set_font("Arial", "", 9)
    total_debit, total_kredit = 0, 0

    for _, row in df.iterrows():
        tanggal = pd.to_datetime(row["Tanggal"]).strftime("%d/%m/%Y")
        akun = str(row["No Akun"])
        nama = str(row["Akun"])
        desk = str(row["Deskripsi"])
        debit = row["Debet"] if not pd.isna(row["Debet"]) else 0
        kredit = row["Kredit"] if not pd.isna(row["Kredit"]) else 0

        total_debit += debit
        total_kredit += kredit

        # Tulis baris
        pdf.cell(col_widths[0], 8, tanggal, border=1)
        pdf.cell(col_widths[1], 8, akun, border=1)
        pdf.multi_cell(col_widths[2], 8, nama, border=1)
        x = pdf.get_x()
        y = pdf.get_y()
        pdf.set_xy(x + col_widths[2], y - 8)
        pdf.multi_cell(col_widths[3], 8, desk, border=1)
        pdf.set_xy(x + col_widths[2] + col_widths[3], y)
        pdf.cell(col_widths[4], 8, format_rupiah(debit), border=1, align="R")
        pdf.cell(col_widths[5], 8, format_rupiah(kredit), border=1, align="R")
        pdf.ln()

    # Total
    pdf.set_font("Arial", "B", 9)
    pdf.cell(sum(col_widths[:-2]), 8, "Total", border=1)
    pdf.cell(col_widths[4], 8, format_rupiah(total_debit), border=1, align="R")
    pdf.cell(col_widths[5], 8, format_rupiah(total_kredit), border=1, align="R")
    pdf.ln()

    # Terbilang
    pdf.set_font("Arial", "I", 8)
    pdf.multi_cell(0, 6, f"Terbilang: {terbilang_rupiah(total_debit)}", align="L")
    pdf.ln(10)

    # Tanda tangan
    pdf.set_font("Arial", "", 10)
    pdf.cell(90, 6, "Direktur,", align="C")
    pdf.cell(90, 6, "Finance,", align="C", ln=True)
    pdf.ln(15)
    pdf.cell(90, 6, f"({direktur})", align="C")
    pdf.cell(90, 6, f"({finance})", align="C", ln=True)

    pdf.output(buffer)
    buffer.seek(0)
    return buffer

# ==============================
# Fungsi Preview HTML Voucher
# ==============================
def preview_voucher(df, voucher_no, perusahaan, alamat, direktur, finance, logo=None):
    total_debit = df["Debet"].fillna(0).sum()
    total_kredit = df["Kredit"].fillna(0).sum()

    html = f"""
    <div style='text-align:center'>
        {'<img src="data:image/png;base64,'+base64.b64encode(open(logo, "rb").read()).decode()+"\" width='50' />" if logo else ""}
        <h3>{perusahaan}</h3>
        <p>{alamat}</p>
        <h4>BUKTI VOUCHER JURNAL</h4>
        <p><b>No Voucher:</b> {voucher_no}</p>
    </div>
    <table border='1' cellspacing='0' cellpadding='5' width='100%'>
        <tr style='background:#eee; text-align:center'>
            <th>Tanggal</th><th>Akun</th><th>Nama Akun</th><th>Deskripsi</th><th>Debit</th><th>Kredit</th>
        </tr>
    """
    for _, row in df.iterrows():
        tanggal = pd.to_datetime(row["Tanggal"]).strftime("%d/%m/%Y")
        akun = str(row["No Akun"])
        nama = str(row["Akun"])
        desk = str(row["Deskripsi"])
        debit = format_rupiah(row["Debet"]) if not pd.isna(row["Debet"]) else "0"
        kredit = format_rupiah(row["Kredit"]) if not pd.isna(row["Kredit"]) else "0"

        html += f"""
        <tr>
            <td>{tanggal}</td>
            <td>{akun}</td>
            <td style='word-wrap:break-word'>{nama}</td>
            <td style='word-wrap:break-word'>{desk}</td>
            <td align='right'>{debit}</td>
            <td align='right'>{kredit}</td>
        </tr>
        """

    html += f"""
        <tr>
            <td colspan='4' align='right'><b>Total</b></td>
            <td align='right'><b>{format_rupiah(total_debit)}</b></td>
            <td align='right'><b>{format_rupiah(total_kredit)}</b></td>
        </tr>
    </table>
    <p><i>Terbilang: {terbilang_rupiah(total_debit)}</i></p>
    <br><br>
    <table width='100%'>
        <tr>
            <td align='center'>Direktur,<br><br><br>({direktur})</td>
            <td align='center'>Finance,<br><br><br>({finance})</td>
        </tr>
    </table>
    """
    return html

# ==============================
# STREAMLIT APP
# ==============================
st.title("Mini Akunting - Voucher Jurnal")

uploaded_file = st.file_uploader("Upload Jurnal (Excel)", type=["xlsx", "xls"])
if uploaded_file:
    df = pd.read_excel(uploaded_file)

    # Input pengaturan custom
    st.subheader("Pengaturan Perusahaan")
    perusahaan = st.text_input("Nama Perusahaan", "Perusahaan ABC")
    alamat = st.text_area("Alamat Perusahaan", "Jl. Contoh No.123 Jakarta")
    direktur = st.text_input("Nama Direktur", "Direktur Utama")
    finance = st.text_input("Nama Finance", "Manager Finance")
    logo = st.file_uploader("Upload Logo", type=["png", "jpg", "jpeg"])

    voucher_no = st.selectbox("Pilih Nomor Voucher", df["Nomor Voucher Jurnal"].unique())

    if st.button("Preview Voucher"):
        df_voucher = df[df["Nomor Voucher Jurnal"] == voucher_no]
        html_preview = preview_voucher(df_voucher, voucher_no, perusahaan, alamat, direktur, finance, logo.name if logo else None)
        st.markdown(html_preview, unsafe_allow_html=True)

    if st.button("Cetak Voucher Jurnal"):
        df_voucher = df[df["Nomor Voucher Jurnal"] == voucher_no]
        buffer = buat_voucher(df_voucher, voucher_no, perusahaan, alamat, direktur, finance, logo.name if logo else None)
        b64 = base64.b64encode(buffer.read()).decode()
        href = f'<a href="data:application/pdf;base64,{b64}" download="voucher.pdf">Download PDF</a>'
        st.markdown(href, unsafe_allow_html=True)
