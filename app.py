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
    pdf.set_left_margin(15)
    pdf.set_right_margin(15)
    pdf.add_page()
    pdf.set_font("Arial", "B", 12)

    # Header perusahaan
    if settings.get("logo"):
        pdf.image(settings["logo"], 15, 8, settings.get("logo_size", 20))

    pdf.set_xy(0, 10)
    pdf.set_font("Arial", "B", 12)
    pdf.cell(0, 6, settings.get("perusahaan",""), ln=1, align="C")

    pdf.set_font("Arial", "", 10)
    pdf.multi_cell(0, 5, settings.get("alamat",""), align="C")
    pdf.ln(3)

    # Judul
    pdf.set_font("Arial", "B", 12)
    pdf.cell(0, 8, "BUKTI VOUCHER JURNAL", ln=1, align="C")
    pdf.set_font("Arial", "", 10)
    pdf.cell(0, 6, f"No Voucher : {no_voucher}", ln=1, align="C")
    pdf.ln(4)

    # ambil data voucher
    data = df[df["Nomor Voucher Jurnal"] == no_voucher]

    # table header (tanpa deskripsi)
    col_widths = [35, 25, 70, 25, 25]
    headers = ["Tanggal","Akun","Nama Akun","Debit","Kredit"]

    pdf.set_font("Arial","B",9)
    for h,w in zip(headers,col_widths):
        pdf.cell(w, 8, h, border=1, align="C")
    pdf.ln()

    total_debit, total_kredit = 0,0
    pdf.set_font("Arial","",9)

    desc_all = []

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
            f"{debit_val:,}".replace(",", "."),
            f"{kredit_val:,}".replace(",", ".")
        ]

        # tinggi baris untuk teks wrap
        line_counts = []
        for i, (val, w) in enumerate(zip(values, col_widths)):
            if i in [3,4]:  # angka
                line_counts.append(1)
            else:
                lines = pdf.multi_cell(w, 6, val, split_only=True)
                line_counts.append(len(lines))
        max_lines = max(line_counts)
        row_height = max_lines * 6

        # isi tabel
        x_start = pdf.get_x()
        y_start = pdf.get_y()
        for i, (val, w) in enumerate(zip(values, col_widths)):
            pdf.rect(x_start, y_start, w, row_height)
            pdf.set_xy(x_start, y_start)
            if i in [3,4]:  
                pdf.cell(w, row_height, val, align="R")
            else:
                pdf.multi_cell(w, 6, val, align="L")
            x_start += w
        pdf.set_y(y_start + row_height)

        total_debit += debit_val
        total_kredit += kredit_val
        desc_all.append(str(row["Deskripsi"]))

    # total row
    pdf.set_font("Arial","B",9)
    pdf.cell(sum(col_widths[:-2]),8,"Total",border=1,align="R")
    pdf.cell(col_widths[3],8,f"{total_debit:,}".replace(",", "."),border=1,align="R")
    pdf.cell(col_widths[4],8,f"{total_kredit:,}".replace(",", "."),border=1,align="R")
    pdf.ln()

    # deskripsi gabungan
    pdf.set_font("Arial","",9)
    full_desc = "; ".join([d for d in desc_all if d and d.strip() != ""])
    pdf.multi_cell(0,6,f"Deskripsi: {full_desc}", border=1)

    # terbilang
    pdf.set_font("Arial","I",9)
    pdf.multi_cell(0,6,f"Terbilang: {num2words(total_debit, lang='id')} rupiah")

    # tanda tangan
    pdf.ln(15)
    pdf.set_font("Arial","",10)
    pdf.cell(90,6,"Direktur,",align="C")
    pdf.cell(90,6,"Finance,",align="C",ln=1)
    pdf.ln(20)
    pdf.cell(90,6,f"({settings.get('direktur','')})",align="C")
    pdf.cell(90,6,f"({settings.get('finance','')})",align="C",ln=1)

    buffer = BytesIO()
    pdf.output(buffer)
    return buffer
