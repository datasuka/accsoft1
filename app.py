import streamlit as st
import pandas as pd
from fpdf import FPDF
from io import BytesIO
from num2words import num2words
import zipfile

# --- Bersihkan data ---
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
    df = df.rename(columns={k.lower(): v for k, v in mapping.items() if k.lower() in df.columns})
    df["Debet"] = pd.to_numeric(df.get("Debet", 0), errors="coerce").fillna(0)
    df["Kredit"] = pd.to_numeric(df.get("Kredit", 0), errors="coerce").fillna(0)
    df["Tanggal"] = pd.to_datetime(df["Tanggal"], errors="coerce")
    return df

# --- Buat voucher ---
def buat_voucher(df, no_voucher, settings):
    pdf = FPDF("P", "mm", "A4")
    pdf.add_page()

    # Logo perusahaan
    if settings.get("logo"):
        pdf.image(settings["logo"], 10, 8, settings.get("logo_size", 25))

    # Nama & alamat perusahaan
    pdf.set_xy(0, 10)
    pdf.set_font("Arial", "B", 12)
    pdf.cell(210, 6, settings.get("perusahaan",""), ln=1, align="C")
    pdf.set_font("Arial", "", 10)
    pdf.multi_cell(210, 5, settings.get("alamat",""), align="C")
    pdf.ln(5)

    # Judul voucher
    pdf.set_font("Arial", "B", 12)
    pdf.cell(0, 10, "BUKTI VOUCHER JURNAL", ln=1, align="C")
    pdf.set_font("Arial", "", 10)
    pdf.cell(0, 8, f"No Voucher : {no_voucher}", ln=1, align="C")
    pdf.ln(4)

    # Filter data voucher
    data = df[df["Nomor Voucher Jurnal"] == no_voucher]

    # Header tabel
    col_widths = [25, 20, 50, 55, 20, 20]
    headers = ["Tanggal","Akun","Nama Akun","Deskripsi","Debit","Kredit"]

    pdf.set_font("Arial","B",9)
    for h, w in zip(headers, col_widths):
        pdf.cell(w, 8, h, border=1, align="C")
    pdf.ln()

    pdf.set_font("Arial","",9)
    total_debit, total_kredit = 0, 0

    # Baris voucher
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

        # Hitung jumlah baris per kolom
        split_texts = []
        for i, val in enumerate(values):
            split = pdf.multi_cell(col_widths[i], 6, val, split_only=True)
            split_texts.append(split)
        max_lines = max(len(s) for s in split_texts)
        row_height = max_lines * 6

        # Cetak per kolom sinkron
        y_start = pdf.get_y()
        for i, val in enumerate(values):
            x_start = pdf.get_x()
            lines = split_texts[i]
            for j in range(max_lines):
                line = lines[j] if j < len(lines) else ""
                border = "LR" if j < max_lines-1 else 1
                align = "L" if i < 4 else "R"
                pdf.multi_cell(col_widths[i], 6, line, border=border, align=align)
            pdf.set_xy(x_start + col_widths[i], y_start)
        pdf.ln(row_height)

        total_debit += debit_val
        total_kredit += kredit_val

    # Total row
    pdf.set_font("Arial","B",9)
    pdf.cell(sum(col_widths[:-2]),8,"Total",border=1,align="R")
    pdf.cell(col_widths[4],8,f"{total_debit:,}".replace(",", "."),border=1,align="R")
    pdf.cell(col_widths[5],8,f"{total_kredit:,}".replace(",", "."),border=1,align="R")
    pdf.ln()

    # Terbilang
    pdf.set_font("Arial","I",9)
    pdf.multi_cell(0,6,f"Terbilang: {num2words(total_debit, lang='id')} rupiah")

    # Tanda tangan
    pdf.ln(15)
    pdf.set_font("Arial","",10)
    pdf.cell(95,6,"Direktur,",align="C")
    pdf.cell(95,6,"Finance,",align="C",ln=1)
    pdf.ln(20)
    pdf.cell(95,6,f"({settings.get('direktur','')})",align="C")
    pdf.cell(95,6,f"({settings.get('finance','')})",align="C",ln=1)

    buffer = BytesIO()
    pdf.output(buffer)
    return buffer.getvalue()

# --- Streamlit ---
st.set_page_config(page_title="Mini Akunting", layout="wide")
st.title("📑 Mini Akunting - Voucher Jurnal")

# Sidebar: Pengaturan Perusahaan
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

# Main Content
file = st.file_uploader("Upload Jurnal (Excel)", type=["xlsx","xls"])
if file:
    df = pd.read_excel(file)
    df = bersihkan_jurnal(df)
    st.dataframe(df.head())

    pilihan = st.radio("Pilih Mode Download", ["Single Voucher", "Multi Voucher"])

    if pilihan == "Single Voucher":
        no_voucher = st.selectbox("Pilih Nomor Voucher", df["Nomor Voucher Jurnal"].unique())
        if st.button("⬇️ Download Voucher Jurnal"):
            pdf_file = buat_voucher(df, no_voucher, settings)
            st.download_button("Download PDF", data=pdf_file, file_name=f"voucher_{no_voucher}.pdf", mime="application/pdf")

    else:  # Multi Voucher
        df["bulan"] = df["Tanggal"].dt.strftime("%B %Y")  # contoh: Januari 2024
        bulan_opsi = ["Semua Bulan"] + sorted(df["bulan"].dropna().unique().tolist())
        pilih_bulan = st.selectbox("Pilih Bulan Voucher", bulan_opsi)

        if pilih_bulan != "Semua Bulan":
            df_filtered = df[df["bulan"] == pilih_bulan]
        else:
            df_filtered = df

        if st.button("⬇️ Download Semua Voucher (ZIP)"):
            buffer = BytesIO()
            with zipfile.ZipFile(buffer, "w") as zipf:
                for no_voucher in df_filtered["Nomor Voucher Jurnal"].unique():
                    pdf_bytes = buat_voucher(df, no_voucher, settings)
                    zipf.writestr(f"voucher_{no_voucher}.pdf", pdf_bytes)
            buffer.seek(0)
            nama_file = "all_vouchers.zip" if pilih_bulan == "Semua Bulan" else f"vouchers_{pilih_bulan.replace(' ','_')}.zip"
            st.download_button("Download ZIP", data=buffer, file_name=nama_file, mime="application/zip")
