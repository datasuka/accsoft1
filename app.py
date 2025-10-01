import streamlit as st
import pandas as pd
from fpdf import FPDF
from io import BytesIO
from num2words import num2words
import zipfile
import calendar

# --- bersihkan data ---
def bersihkan_jurnal(df):
    df = df.rename(columns=lambda x: str(x).strip().lower())
    mapping = {
        "tanggal": "Tanggal",
        "nomor voucher jurnal": "Nomor Voucher Jurnal",
        "no akun": "No Akun",
        "akun": "Akun",
        "deskripsi": "Deskripsi",
        "debet": "Debet",
        "kredit": "Kredit"
    }
    df = df.rename(columns={k.lower(): v for k,v in mapping.items() if k.lower() in df.columns})
    df["Debet"] = pd.to_numeric(df.get("Debet", 0), errors="coerce").fillna(0)
    df["Kredit"] = pd.to_numeric(df.get("Kredit", 0), errors="coerce").fillna(0)
    return df

# --- generate voucher ---
def buat_voucher(df, no_voucher, settings):
    pdf = FPDF("P", "mm", "A4")
    pdf.set_auto_page_break(auto=True, margin=15)
    pdf.add_page()

    # Header perusahaan + logo
    if settings.get("logo"):
        pdf.image(settings["logo"], 15, 8, settings.get("logo_size", 20))

    pdf.set_font("Arial", "B", 12)
    pdf.cell(210, 6, settings.get("perusahaan",""), ln=1, align="C")
    pdf.set_font("Arial", "", 10)
    pdf.multi_cell(210, 5, settings.get("alamat",""), align="C")
    pdf.ln(5)

    # Judul
    pdf.set_font("Arial", "B", 12)
    pdf.cell(0, 10, "BUKTI VOUCHER JURNAL", ln=1, align="C")
    pdf.set_font("Arial", "", 10)
    pdf.cell(0, 8, f"No Voucher : {no_voucher}", ln=1, align="C")
    pdf.ln(4)

    # ambil data voucher
    data = df[df["Nomor Voucher Jurnal"] == no_voucher]

    # table header
    col_widths = [25, 20, 55, 55, 25, 25]
    headers = ["Tanggal","Akun","Nama Akun","Deskripsi","Debit","Kredit"]

    pdf.set_font("Arial","B",9)
    for h,w in zip(headers,col_widths):
        pdf.cell(w, 8, h, border=1, align="C")
    pdf.ln()

    total_debit, total_kredit = 0,0
    pdf.set_font("Arial","",9)

    for _, row in data.iterrows():
        debit_val = int(row["Debet"]) if pd.notna(row["Debet"]) else 0
        kredit_val = int(row["Kredit"]) if pd.notna(row["Kredit"]) else 0
        try:
            tgl = pd.to_datetime(row["Tanggal"]).strftime("%d/%m/%Y")
        except:
            tgl = str(row["Tanggal"])

        values = [
            tgl,
            str(row["No Akun"]),
            str(row["Akun"]),
            str(row["Deskripsi"]),
            f"{debit_val:,}".replace(",", "."),
            f"{kredit_val:,}".replace(",", ".")
        ]

        # hitung tinggi baris hanya untuk kolom teks
        nama_lines = pdf.multi_cell(col_widths[2], 6, values[2], split_only=True)
        desc_lines = pdf.multi_cell(col_widths[3], 6, values[3], split_only=True)
        max_lines = max(len(nama_lines), len(desc_lines), 1)
        row_height = max_lines * 6

        # gambar baris
        x = pdf.l_margin
        y = pdf.get_y()

        # Tanggal
        pdf.rect(x, y, col_widths[0], row_height)
        pdf.multi_cell(col_widths[0], 6, values[0], border=0)
        x += col_widths[0]
        pdf.set_xy(x, y)

        # Akun
        pdf.rect(x, y, col_widths[1], row_height)
        pdf.multi_cell(col_widths[1], 6, values[1], border=0)
        x += col_widths[1]
        pdf.set_xy(x, y)

        # Nama Akun
        pdf.rect(x, y, col_widths[2], row_height)
        pdf.multi_cell(col_widths[2], 6, values[2], border=0)
        x += col_widths[2]
        pdf.set_xy(x, y)

        # Deskripsi
        pdf.rect(x, y, col_widths[3], row_height)
        pdf.multi_cell(col_widths[3], 6, values[3], border=0)
        x += col_widths[3]
        pdf.set_xy(x, y)

        # Debit (angka single line)
        pdf.rect(x, y, col_widths[4], row_height)
        pdf.cell(col_widths[4], row_height, values[4], border=0, align="R")
        x += col_widths[4]
        pdf.set_xy(x, y)

        # Kredit (angka single line)
        pdf.rect(x, y, col_widths[5], row_height)
        pdf.cell(col_widths[5], row_height, values[5], border=0, align="R")

        pdf.set_y(y + row_height)

        total_debit += debit_val
        total_kredit += kredit_val

    # total row
    pdf.set_font("Arial","B",9)
    pdf.cell(sum(col_widths[:-2]),8,"Total",border=1,align="R")
    pdf.cell(col_widths[4],8,f"{total_debit:,}".replace(",", "."),border=1,align="R")
    pdf.cell(col_widths[5],8,f"{total_kredit:,}".replace(",", "."),border=1,align="R")
    pdf.ln()

    # terbilang
    pdf.set_font("Arial","I",9)
    pdf.multi_cell(0,6,f"Terbilang: {num2words(total_debit, lang='id')} rupiah")

    # tanda tangan
    pdf.ln(15)
    pdf.set_font("Arial","",10)
    pdf.cell(95,6,"Direktur,",align="C")
    pdf.cell(95,6,"Finance,",align="C",ln=1)
    pdf.ln(20)
    pdf.cell(95,6,f"({settings.get('direktur','')})",align="C")
    pdf.cell(95,6,f"({settings.get('finance','')})",align="C",ln=1)

    buffer = BytesIO()
    pdf.output(buffer)
    return buffer

# --- Streamlit ---
st.set_page_config(page_title="Mini Akunting", layout="wide")
st.title("📑 Mini Akunting - Voucher Jurnal")

# Sidebar
st.sidebar.header("⚙️ Pengaturan Perusahaan")
settings = {}
settings["perusahaan"] = st.sidebar.text_input("Nama Perusahaan")
settings["alamat"] = st.sidebar.text_area("Alamat Perusahaan")
settings["direktur"] = st.sidebar.text_input("Nama Direktur")
settings["finance"] = st.sidebar.text_input("Nama Finance")
settings["logo_size"] = st.sidebar.slider("Ukuran Logo (mm)", 10, 50, 25)
logo_file = st.sidebar.file_uploader("Upload Logo (PNG/JPG)", type=["png","jpg","jpeg"])
if logo_file:
    tmp = BytesIO(logo_file.read())
    settings["logo"] = tmp

# Main content
file = st.file_uploader("Upload Jurnal (Excel)", type=["xlsx","xls"])
if file:
    df = pd.read_excel(file)
    df = bersihkan_jurnal(df)
    st.dataframe(df.head())

    mode = st.radio("Pilih Mode Cetak", ["Single Voucher", "Per Bulan"])

    if mode == "Single Voucher":
        no_voucher = st.selectbox("Pilih Nomor Voucher", df["Nomor Voucher Jurnal"].unique())
        if st.button("Cetak Voucher Jurnal"):
            pdf_file = buat_voucher(df, no_voucher, settings)
            st.download_button("⬇️ Download PDF", data=pdf_file, file_name=f"{no_voucher}.pdf")

    else:  # per bulan
        df["Tanggal"] = pd.to_datetime(df["Tanggal"], errors="coerce")
        df = df.dropna(subset=["Tanggal"])
        bulan = st.selectbox("Pilih Bulan", range(1,13), format_func=lambda x: calendar.month_name[x])

        if st.button("Cetak Semua Voucher Bulan Ini"):
            buffer_zip = BytesIO()
            with zipfile.ZipFile(buffer_zip, "w") as zf:
                for v in df[df["Tanggal"].dt.month==bulan]["Nomor Voucher Jurnal"].unique():
                    pdf_file = buat_voucher(df, v, settings)
                    zf.writestr(f"{v}.pdf", pdf_file.getvalue())
            buffer_zip.seek(0)
            st.download_button("⬇️ Download ZIP", data=buffer_zip, file_name=f"voucher_{bulan}.zip", mime="application/zip")
