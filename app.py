import streamlit as st
import pandas as pd
from fpdf import FPDF
from num2words import num2words
from PIL import Image
import io

# --- Fungsi Terbilang Indonesia ---
def terbilang(nominal):
    try:
        return num2words(nominal, lang='id').replace("koma nol", "")
    except:
        return str(nominal)

# --- Fungsi Generate Voucher ---
def buat_voucher(jurnal_df, no_voucher, settings):
    data = jurnal_df[jurnal_df['nomor voucher jurnal'] == no_voucher]

    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Arial", size=10)

    # Logo
    if settings.get("logo"):
        try:
            pdf.image(settings["logo"], 10, 8, 25, 25)
        except:
            pass

    # Header Perusahaan
    pdf.set_xy(40, 10)
    pdf.set_font("Arial", 'B', 12)
    pdf.cell(100, 6, settings.get("nama_perusahaan", ""), ln=True)
    pdf.set_font("Arial", '', 9)
    pdf.set_x(40)
    pdf.multi_cell(120, 5, settings.get("alamat", ""))

    pdf.ln(5)
    pdf.set_font("Arial", 'B', 14)
    pdf.cell(200, 10, "BUKTI VOUCHER JURNAL", ln=True, align="C")

    pdf.set_font("Arial", '', 11)
    pdf.cell(200, 8, f"No Voucher : {no_voucher}", ln=True, align="C")

    pdf.ln(5)

    # Tabel Jurnal
    pdf.set_font("Arial", 'B', 10)
    pdf.cell(40, 8, "Akun", 1)
    pdf.cell(70, 8, "Deskripsi", 1)
    pdf.cell(40, 8, "Debit", 1, align="R")
    pdf.cell(40, 8, "Kredit", 1, align="R")
    pdf.ln()

    total = 0
    pdf.set_font("Arial", '', 10)
    for _, row in data.iterrows():
        debit_val = row.get("debit", 0) if pd.notna(row.get("debit", 0)) else 0
        kredit_val = row.get("kredit", 0) if pd.notna(row.get("kredit", 0)) else 0

        pdf.cell(40, 8, str(row.get('no akun', '')), 1)
        pdf.cell(70, 8, str(row.get('deskripsi', '')), 1)
        pdf.cell(40, 8, f"{int(debit_val):,}", 1, align="R")
        pdf.cell(40, 8, f"{int(kredit_val):,}", 1, align="R")
        pdf.ln()

        total += debit_val

    # Total
    pdf.set_font("Arial", 'B', 10)
    pdf.cell(110, 8, "Total", 1)
    pdf.cell(40, 8, f"{int(total):,}", 1, align="R")
    pdf.cell(40, 8, "", 1)
    pdf.ln()

    pdf.set_font("Arial", 'I', 9)
    pdf.multi_cell(0, 6, f"Terbilang: {terbilang(int(total))} rupiah")

    pdf.ln(15)
    pdf.set_font("Arial", '', 10)
    pdf.cell(100, 6, "Direktur,", align="C")
    pdf.cell(90, 6, "Finance,", align="C")
    pdf.ln(20)

    pdf.cell(100, 6, settings.get("direktur", ""), align="C")
    pdf.cell(90, 6, settings.get("finance", ""), align="C")

    # Simpan ke memori (bukan file fisik)
    buf = io.BytesIO()
    pdf.output(buf)
    pdf_data = buf.getvalue()
    return pdf_data

# --- STREAMLIT APP ---
st.title("📄 Cetak Voucher Jurnal")

# Setting Perusahaan
st.sidebar.header("⚙️ Pengaturan Perusahaan")
nama_perusahaan = st.sidebar.text_input("Nama Perusahaan")
alamat = st.sidebar.text_area("Alamat Perusahaan")
direktur = st.sidebar.text_input("Nama Direktur")
finance = st.sidebar.text_input("Nama Finance")
logo_file = st.sidebar.file_uploader("Upload Logo (PNG/JPG)", type=["png", "jpg", "jpeg"])

settings = {"nama_perusahaan": nama_perusahaan, "alamat": alamat, "direktur": direktur, "finance": finance}
if logo_file:
    img = Image.open(logo_file)
    logo_path = "logo.png"
    img.save(logo_path)
    settings["logo"] = logo_path

# Upload Jurnal
uploaded_file = st.file_uploader("Upload Daftar Jurnal (Excel/CSV)", type=["xlsx", "csv"])

if uploaded_file:
    if uploaded_file.name.endswith("csv"):
        df = pd.read_csv(uploaded_file)
    else:
        df = pd.read_excel(uploaded_file)

    # Normalisasi kolom
    df.columns = df.columns.str.strip().str.lower()

    # Preview
    st.subheader("Preview Data Jurnal")
    st.dataframe(df.head())

    # Pilih Voucher
    if "nomor voucher jurnal" in df.columns:
        no_list = df["nomor voucher jurnal"].unique()
        pilih_no = st.selectbox("Pilih Nomor Voucher", no_list)
        if st.button("Generate Voucher PDF"):
            pdf_bytes = buat_voucher(df, pilih_no, settings)
            st.download_button(
                "Download Voucher PDF",
                data=pdf_bytes,
                file_name=f"voucher_{pilih_no}.pdf",
                mime="application/pdf"
            )
