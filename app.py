import streamlit as st
import pandas as pd
from fpdf import FPDF
from num2words import num2words
from PIL import Image

# --- Fungsi Terbilang Indonesia ---
def terbilang(nominal):
    try:
        return num2words(nominal, lang='id').replace("koma nol", "")
    except:
        return str(nominal)

# --- Fungsi Generate Voucher ---
def buat_voucher(jurnal_df, no_jurnal, settings):
    data = jurnal_df[jurnal_df['no jurnal'] == no_jurnal]

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
    if "subjek" in data.columns and data["subjek"].notna().any():
        subjek = data["subjek"].iloc[0]
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
        debit_val = row.get("debit", 0) if pd.notna(row.get("debit", 0)) else 0
        kredit_val = row.get("kredit", 0) if pd.notna(row.get("kredit", 0)) else 0

        pdf.cell(30, 8, str(row.get('tanggal', '')), 1)
        pdf.cell(40, 8, str(row.get('akun', row.get('no akun', ''))), 1)
        pdf.cell(50, 8, str(row.get('deskripsi', '')), 1)
        pdf.cell(30, 8, f"{int(debit_val):,}", 1, align="R")
        pdf.cell(30, 8, f"{int(kredit_val):,}", 1, align="R")
        pdf.ln()

        total += debit_val

    # Total
    pdf.set_font("Arial", 'B', 10)
    pdf.cell(120, 8, "Total", 1)
    pdf.cell(30, 8, f"{int(total):,}", 1, align="R")
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

    # --- Normalisasi kolom ---
    df.columns = df.columns.str.strip().str.lower()

    # Buat nama kolom unik otomatis (jika ada duplikat)
    from pandas.io.parsers.base_parser import ParserBase
    df.columns = ParserBase({'names': df.columns})._maybe_dedup_names(df.columns)

    # Alias umum
    if "debet" in df.columns and "debit" not in df.columns:
        df.rename(columns={"debet": "debit"}, inplace=True)
    if "no" in df.columns and "no jurnal" not in df.columns:
        df.rename(columns={"no": "no jurnal"}, inplace=True)
    if "no akun" in df.columns and "akun" not in df.columns:
        df.rename(columns={"no akun": "akun"}, inplace=True)

    # Pastikan numeric
    if "debit" in df.columns:
        df["debit"] = pd.to_numeric(df["debit"], errors="coerce").fillna(0)
    if "kredit" in df.columns:
        df["kredit"] = pd.to_numeric(df["kredit"], errors="coerce").fillna(0)

    st.subheader("Preview Data Jurnal")
    st.dataframe(df.head())

    # Tabs eksekusi
    tab1, tab2, tab3 = st.tabs(["📄 Voucher Jurnal", "📊 Laba Rugi", "📊 Neraca"])

    with tab1:
        if "no jurnal" in df.columns:
            no_list = df["no jurnal"].unique()
            pilih_no = st.selectbox("Pilih No Jurnal", no_list)
            if st.button("Generate Voucher PDF"):
                filename = buat_voucher(df, pilih_no, settings)
                with open(filename, "rb") as f:
                    st.download_button("Download Voucher PDF", f, file_name=filename)

    with tab2:
        st.subheader("Laporan Laba Rugi")
        if "akun" in df.columns:
            laba_rugi = (
                df[df['akun'].astype(str).str.startswith(("4", "5"))]
                .groupby("akun")[["debit", "kredit"]]
                .sum()
            )
            st.dataframe(laba_rugi)

    with tab3:
        st.subheader("Laporan Neraca")
        if "akun" in df.columns:
            neraca = (
                df[df['akun'].astype(str).str.startswith(("1", "2", "3"))]
                .groupby("akun")[["debit", "kredit"]]
                .sum()
            )
            st.dataframe(neraca)
