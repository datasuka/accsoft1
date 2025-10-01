import streamlit as st
import pandas as pd
from fpdf import FPDF
from io import BytesIO
from num2words import num2words

# --- helper untuk hitung line ---
def get_num_lines(pdf, text, col_width):
    result = pdf.multi_cell(col_width, 6, text, split_only=True)
    return len(result)

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
    pdf.add_page()
    pdf.set_font("Arial", "B", 12)

    # Header perusahaan
    if settings.get("logo"):
        pdf.image(settings["logo"], 10, 8, settings.get("logo_size", 25))

    pdf.set_xy(0, 10)
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
    col_widths = [25, 20, 50, 55, 20, 20]
    headers = ["Tanggal","Akun","Nama Akun","Deskripsi","Debit","Kredit"]

    pdf.set_font("Arial","B",9)
    for h,w in zip(headers,col_widths):
        pdf.cell(w, 8, h, border=1, align="C")
    pdf.ln()

    pdf.set_font("Arial","",9)
    total_debit, total_kredit = 0,0

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

        # Sinkronisasi multi-line
        nama_lines = pdf.multi_cell(col_widths[2], 6, values[2], split_only=True)
        desc_lines = pdf.multi_cell(col_widths[3], 6, values[3], split_only=True)
        max_lines = max(len(nama_lines), len(desc_lines))
        row_height = max_lines * 6

        # Posisi awal
        x = pdf.get_x()
        y = pdf.get_y()

        # Tanggal & No Akun
        pdf.cell(col_widths[0], row_height, values[0], border=1, align="L")
        pdf.cell(col_widths[1], row_height, values[1], border=1, align="L")

        # Nama Akun sinkron
        x2, y2 = pdf.get_x(), pdf.get_y()
        for i in range(max_lines):
            line = nama_lines[i] if i < len(nama_lines) else ""
            border = "LR" if i < max_lines-1 else 1
            pdf.multi_cell(col_widths[2], 6, line, border=border, align="L")
        pdf.set_xy(x2+col_widths[2], y2)

        # Deskripsi sinkron
        x3, y3 = pdf.get_x(), pdf.get_y()
        for i in range(max_lines):
            line = desc_lines[i] if i < len(desc_lines) else ""
            border = "LR" if i < max_lines-1 else 1
            pdf.multi_cell(col_widths[3], 6, line, border=border, align="L")
        pdf.set_xy(x3+col_widths[3], y3)

        # Debit & Kredit
        pdf.cell(col_widths[4], row_height, values[4], border=1, align="R")
        pdf.cell(col_widths[5], row_height, values[5], border=1, align="R")
        pdf.ln(row_height)

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

    no_voucher = st.selectbox("Pilih Nomor Voucher", df["Nomor Voucher Jurnal"].unique())

    if st.button("Cetak Voucher Jurnal"):
        pdf_file = buat_voucher(df, no_voucher, settings)
        st.download_button("Download PDF", data=pdf_file, file_name="voucher_jurnal.pdf")
