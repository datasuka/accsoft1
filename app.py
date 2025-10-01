import streamlit as st
import pandas as pd
from fpdf import FPDF
from num2words import num2words
from PIL import Image

# Fungsi terbilang
def terbilang(nominal):
    try:
        return num2words(nominal, lang='id').replace("koma nol", "")
    except:
        return str(nominal)

# Generate Voucher
def buat_voucher(jurnal_df, no_jurnal, settings):
    data = jurnal_df[jurnal_df['No Jurnal'] == no_jurnal]

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
    pdf.cell(200, 8, f"No Voucher : {no_jurnal}", ln=True, align="C")

    # Subjek
    if "Subjek" in data.columns and data["Subjek"].notna().any():
        subjek = data["Subjek"].iloc[0]
        pdf.cell(200, 6, f"Subjek : {subjek}", ln=True)

    pdf.ln(5)

    # Tabel Jurnal
    pdf.set_font("Arial", 'B', 10)
    pdf.cell(30, 8, "Tanggal", 1)
    pdf.cell(40, 8, "Akun", 1)
    pdf.cell(50, 8, "Deskripsi", 1)
    pdf.cell(30, 8, "Debit", 1, align="R")
    pdf.cell(30, 8, "Kredit", 1, align="R")
    pdf.ln()

    total = 0
    pdf.set_font("Arial", '', 10)
    for _, row in data.iterrows():
        pdf.cell(30, 8, str(row.get('Tanggal', '')), 1)
        pdf.cell(40, 8, str(row.get('Akun', row.get('No Akun', ''))), 1)
        pdf.cell(50, 8, str(row['Deskripsi']), 1)
        pdf.cell(30, 8, f"{row['Debit']:.0f}", 1, align="R")
        pdf.cell(30, 8, f"{row['Kredit']:.0f}", 1, align="R")
        pdf.ln()
        total += row['Debit']

    # Total
    pdf.set_font("Arial", 'B', 10)
    pdf.cell(120, 8, "Total", 1)
    pdf.cell(30, 8, f"{total:.0f}", 1, align="R")
    pdf.cell(30, 8, "", 1)
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

    filename = f"voucher_{no_jurnal}.pdf"
    pdf.output(filename)
    return filename

# --- STREAMLIT APP ---
st.title("📊 Mini Accounting System - Single Entity")

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

    # Normalisasi nama kolom
    df = df.rename(columns=lambda x: str(x).strip())
    if "Debet" in df.columns: df.rename(columns={"Debet": "Debit"}, inplace=True)
    if "No Akun" in df.columns: df.rename(columns={"No Akun": "Akun"}, inplace=True)

    # Tambahkan No Jurnal jika tidak ada
    if "No Jurnal" not in df.columns:
        if "No" in df.columns:
            df.rename(columns={"No": "No Jurnal"}, inplace=True)

    st.subheader("Preview Data Jurnal")
    st.dataframe(df.head())

    # Tombol Eksekusi
    tab1, tab2, tab3 = st.tabs(["📄 Voucher Jurnal", "📊 Laba Rugi", "📊 Neraca"])

    with tab1:
        no_list = df["No Jurnal"].unique()
        pilih_no = st.selectbox("Pilih No Jurnal", no_list)
        if st.button("Generate Voucher PDF"):
            filename = buat_voucher(df, pilih_no, settings)
            with open(filename, "rb") as f:
                st.download_button("Download Voucher PDF", f, file_name=filename)

    with tab2:
        st.subheader("Laporan Laba Rugi")
        laba_rugi = df[df['Akun'].astype(str).str.startswith(("4","5"))] \
                      .groupby("Akun")[["Debit","Kredit"]].sum()
        st.dataframe(laba_rugi)

    with tab3:
        st.subheader("Laporan Neraca")
        neraca = df[df['Akun'].astype(str).str.startswith(("1","2","3"))] \
                    .groupby("Akun")[["Debit","Kredit"]].sum()
        st.dataframe(neraca)
