import streamlit as st
import pandas as pd
import io, os, zipfile
from fpdf import FPDF
from num2words import num2words
from datetime import datetime

# ===== Fungsi bantu =====
def format_rupiah(x):
    try:
        return "{:,.0f}".format(float(x)).replace(",", ".")
    except:
        return "0"

def terbilang_rupiah(x):
    try:
        return num2words(int(x), lang='id') + " rupiah"
    except:
        return ""

def buat_voucher(df, voucher_no, settings):
    data = df[df["Nomor Voucher Jurnal"] == voucher_no]
    if data.empty:
        return None

    total_debit = data["Debet"].fillna(0).sum()
    total_kredit = data["Kredit"].fillna(0).sum()

    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Arial", "B", 12)

    # Logo
    if settings.get("logo"):
        logo_bytes = settings["logo"].read()
        logo_path = "logo_temp.png"
        with open(logo_path, "wb") as f:
            f.write(logo_bytes)
        pdf.image(logo_path, 10, 8, 25)
        os.remove(logo_path)

    # Header perusahaan
    pdf.cell(0, 5, settings.get("perusahaan", ""), ln=1, align="C")
    pdf.set_font("Arial", "", 10)
    pdf.multi_cell(0, 5, settings.get("alamat", ""), align="C")
    pdf.ln(5)

    # Judul
    pdf.set_font("Arial", "B", 12)
    pdf.cell(0, 7, "BUKTI VOUCHER JURNAL", ln=1, align="C")
    pdf.set_font("Arial", "", 10)
    pdf.cell(0, 6, f"No Voucher : {voucher_no}", ln=1, align="C")
    pdf.ln(5)

    # Tabel header
    pdf.set_font("Arial", "B", 9)
    col_widths = [25, 20, 55, 40, 25, 25]
    headers = ["Tanggal", "Akun", "Nama Akun", "Deskripsi", "Debit", "Kredit"]
    for h, w in zip(headers, col_widths):
        pdf.cell(w, 8, h, border=1, align="C")
    pdf.ln()

    # Isi tabel
    pdf.set_font("Arial", "", 9)
    for _, row in data.iterrows():
        tanggal = pd.to_datetime(row["Tanggal"]).strftime("%d/%m/%Y")
        pdf.cell(col_widths[0], 8, tanggal, border=1)
        pdf.cell(col_widths[1], 8, str(row["No Akun"]), border=1)
        pdf.cell(col_widths[2], 8, str(row["Akun"])[:30], border=1)
        pdf.cell(col_widths[3], 8, str(row["Deskripsi"])[:30], border=1)
        pdf.cell(col_widths[4], 8, format_rupiah(row["Debet"]), border=1, align="R")
        pdf.cell(col_widths[5], 8, format_rupiah(row["Kredit"]), border=1, align="R")
        pdf.ln()

    # Total
    pdf.set_font("Arial", "B", 9)
    pdf.cell(sum(col_widths[:4]), 8, "Total", border=1, align="R")
    pdf.cell(col_widths[4], 8, format_rupiah(total_debit), border=1, align="R")
    pdf.cell(col_widths[5], 8, format_rupiah(total_kredit), border=1, align="R")
    pdf.ln(10)

    # Terbilang
    pdf.set_font("Arial", "I", 9)
    pdf.cell(0, 6, f"Terbilang: {terbilang_rupiah(total_debit)}", ln=1)

    # Tanda tangan
    pdf.ln(15)
    pdf.set_font("Arial", "", 10)
    pdf.cell(95, 6, "Direktur,", align="C")
    pdf.cell(95, 6, "Finance,", align="C")
    pdf.ln(20)
    pdf.set_font("Arial", "B", 10)
    pdf.cell(95, 6, f"({settings.get('direktur','')})", align="C")
    pdf.cell(95, 6, f"({settings.get('finance','')})", align="C")

    # ✅ fix output (selalu bytes)
    pdf_bytes = pdf.output(dest="S")
    if isinstance(pdf_bytes, str):
        pdf_bytes = pdf_bytes.encode("latin-1")
    elif isinstance(pdf_bytes, bytearray):
        pdf_bytes = bytes(pdf_bytes)
    return pdf_bytes

# ===== App Utama =====
st.title("📑 Mini Akunting - Voucher Jurnal")

# Upload file
file = st.file_uploader("Upload Jurnal (Excel)", type=["xlsx", "xls"])
if file:
    df = pd.read_excel(file)

    # Sidebar custom
    st.sidebar.header("⚙️ Pengaturan Perusahaan")
    perusahaan = st.sidebar.text_input("Nama Perusahaan")
    alamat = st.sidebar.text_area("Alamat Perusahaan")
    direktur = st.sidebar.text_input("Nama Direktur")
    finance = st.sidebar.text_input("Nama Finance")
    logo = st.sidebar.file_uploader("Upload Logo (PNG/JPG)", type=["png", "jpg", "jpeg"])

    settings = {
        "perusahaan": perusahaan,
        "alamat": alamat,
        "direktur": direktur,
        "finance": finance,
        "logo": logo,
    }

    # Pilihan mode download
    mode = st.radio("Pilih Mode Download", ["Single", "Multi"], horizontal=True)

    if mode == "Single":
        voucher_no = st.selectbox("Pilih Nomor Voucher", df["Nomor Voucher Jurnal"].unique())
        if st.button("Download Voucher PDF"):
            pdf_bytes = buat_voucher(df, voucher_no, settings)
            if pdf_bytes:
                st.download_button(
                    "⬇️ Download PDF",
                    data=pdf_bytes,
                    file_name=f"Voucher_{voucher_no}.pdf",
                    mime="application/pdf"
                )

    else:  # Multi
        voucher_list = st.multiselect("Pilih Nomor Voucher", df["Nomor Voucher Jurnal"].unique())
        if st.button("Download ZIP Voucher"):
            buffer = io.BytesIO()
            with zipfile.ZipFile(buffer, "w") as zf:
                for v in voucher_list:
                    pdf_bytes = buat_voucher(df, v, settings)
                    if pdf_bytes:
                        zf.writestr(f"Voucher_{v}.pdf", pdf_bytes)
            buffer.seek(0)
            st.download_button(
                "⬇️ Download ZIP",
                data=buffer,
                file_name="Vouchers.zip",
                mime="application/zip"
            )
