import streamlit as st
import pandas as pd
from fpdf import FPDF
from num2words import num2words
from PIL import Image
import io
import math

# --- Fungsi Terbilang Indonesia ---
def terbilang(nominal):
    try:
        return num2words(nominal, lang='id').replace("koma nol", "")
    except:
        return str(nominal)

# --- Fungsi bersihkan data jurnal ---
def bersihkan_jurnal(df):
    # Hapus kolom Unnamed
    df = df.loc[:, ~df.columns.str.contains("^Unnamed")]

    # Normalisasi nama kolom
    df.columns = df.columns.str.strip().str.lower()

    # Alias nama kolom
    rename_map = {
        "debet": "debit",
        "no": "nomor voucher jurnal",
        "no akun": "no akun"
    }
    df = df.rename(columns={k: v for k, v in rename_map.items() if k in df.columns})

    # Hapus baris kosong / header duplikat
    df = df.dropna(how="all")
    if "akun" in df.columns:
        df = df[df["akun"].str.lower() != "akun"]

    # Format tanggal jadi dd/mm/yyyy
    if "tanggal" in df.columns:
        df["tanggal"] = pd.to_datetime(df["tanggal"], errors="coerce").dt.strftime("%d/%m/%Y")

    # Pastikan No Akun jadi string
    if "no akun" in df.columns:
        df["no akun"] = df["no akun"].astype(str).str.strip()

    # Bersihkan angka debit/kredit
    for col in ["debit", "kredit"]:
        if col in df.columns:
            df[col] = (
                df[col]
                .astype(str)
                .str.replace(".", "", regex=False)   # hapus titik ribuan
                .str.replace(",", ".", regex=False)  # koma → titik
                .replace("", "0")
            )
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0).astype(int)

    return df

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

    # Tabel Header
    pdf.set_font("Arial", 'B', 10)
    col_widths = [25, 25, 50, 60, 25, 25]  # Deskripsi diperlebar jadi 60
    headers = ["Tanggal", "Akun", "Nama Akun", "Deskripsi", "Debit", "Kredit"]
    for i, h in enumerate(headers):
        pdf.cell(col_widths[i], 8, h, 1, align="C")
    pdf.ln()

    total_debit = 0
    total_kredit = 0

    pdf.set_font("Arial", '', 10)
    for _, row in data.iterrows():
        debit_val = int(row.get("debit", 0)) if pd.notna(row.get("debit", 0)) else 0
        kredit_val = int(row.get("kredit", 0)) if pd.notna(row.get("kredit", 0)) else 0

        col_values = [
            str(row.get('tanggal', '')),
            str(row.get('no akun', '')),
            str(row.get('akun', '')),
            str(row.get('deskripsi', '')),
            f"{debit_val:,}".replace(",", "."),
            f"{kredit_val:,}".replace(",", ".")
        ]

        # --- Hitung tinggi baris maksimum ---
        line_counts = []
        for i, val in enumerate(col_values):
            text_width = pdf.get_string_width(val)
            n_lines = max(1, math.ceil(text_width / (col_widths[i] - 2)))
            line_counts.append(n_lines)
        row_height = max(line_counts) * 6

        # --- Render baris dengan row lock ---
        x = pdf.get_x()
        y = pdf.get_y()
        for i, val in enumerate(col_values):
            pdf.multi_cell(col_widths[i], 6, val, border=1, align="L")
            pdf.set_xy(x + col_widths[i], y)  # reset posisi ke kanan
            x += col_widths[i]
        pdf.set_y(y + row_height)

        total_debit += debit_val
        total_kredit += kredit_val

    # Total
    pdf.set_font("Arial", 'B', 10)
    pdf.cell(sum(col_widths[:-2]), 8, "Total", 1)
    pdf.cell(col_widths[-2], 8, f"{total_debit:,}".replace(",", "."), 1, align="R")
    pdf.cell(col_widths[-1], 8, f"{total_kredit:,}".replace(",", "."), 1, align="R")
    pdf.ln()

    # Terbilang
    nilai_terbilang = total_debit if total_debit != 0 else total_kredit
    pdf.set_font("Arial", 'I', 9)
    pdf.multi_cell(0, 6, f"Terbilang: {terbilang(int(nilai_terbilang))} rupiah")

    pdf.ln(15)
    pdf.set_font("Arial", '', 10)
    pdf.cell(100, 6, "Direktur,", align="C")
    pdf.cell(90, 6, "Finance,", align="C")
    pdf.ln(20)

    pdf.cell(100, 6, settings.get("direktur", ""), align="C")
    pdf.cell(90, 6, settings.get("finance", ""), align="C")

    buf = io.BytesIO()
    pdf.output(buf)
    return buf.getvalue()

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
        df = pd.read_csv(uploaded_file, dtype=str)
    else:
        df = pd.read_excel(uploaded_file, dtype=str)

    # Bersihkan data
    df = bersihkan_jurnal(df)

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
