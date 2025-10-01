import streamlit as st
import pandas as pd
from fpdf import FPDF
from io import BytesIO
from datetime import datetime

# ==============================
# Helper Function
# ==============================
def format_rupiah(angka):
    try:
        return f"{int(angka):,}".replace(",", ".")
    except:
        return "0"

def terbilang_rupiah(n):
    satuan = ["", "satu", "dua", "tiga", "empat", "lima",
              "enam", "tujuh", "delapan", "sembilan"]
    belasan = ["sepuluh", "sebelas", "dua belas", "tiga belas",
               "empat belas", "lima belas", "enam belas",
               "tujuh belas", "delapan belas", "sembilan belas"]
    puluhan = ["", "", "dua puluh", "tiga puluh", "empat puluh",
               "lima puluh", "enam puluh", "tujuh puluh",
               "delapan puluh", "sembilan puluh"]
    ribuan = ["", "ribu", "juta", "miliar", "triliun"]

    def terbilang_bilangan(n):
        if n < 10:
            return satuan[n]
        elif n < 20:
            return belasan[n-10]
        elif n < 100:
            return puluhan[n//10] + (" " + satuan[n%10] if n % 10 != 0 else "")
        elif n < 200:
            return "seratus " + terbilang_bilangan(n-100)
        elif n < 1000:
            return satuan[n//100] + " ratus " + terbilang_bilangan(n%100)
        elif n < 2000:
            return "seribu " + terbilang_bilangan(n-1000)
        else:
            for i, r in enumerate(ribuan[1:], 1):
                if n < 1000**(i+1):
                    return terbilang_bilangan(n//(1000**i)) + " " + r + " " + terbilang_bilangan(n % (1000**i))

    return terbilang_bilangan(int(n)).strip() + " rupiah"

# ==============================
# Voucher Preview (HTML)
# ==============================
def preview_voucher(df, voucher_no, perusahaan, alamat, direktur, finance, logo=None):
    total_debit = df["Debet"].fillna(0).sum()
    total_kredit = df["Kredit"].fillna(0).sum()

    # Header
    html = f"""
    <div style='text-align:center'>
        <h3>{perusahaan}</h3>
        <p>{alamat}</p>
        <h4>BUKTI VOUCHER JURNAL</h4>
        <p><b>No Voucher:</b> {voucher_no}</p>
    </div>
    <br>
    <table border='1' cellspacing='0' cellpadding='5' width='100%' style='border-collapse:collapse; font-size:12px'>
        <thead style='background:#f0f0f0; text-align:center'>
            <tr>
                <th>Tanggal</th>
                <th>Akun</th>
                <th>Nama Akun</th>
                <th>Deskripsi</th>
                <th>Debit</th>
                <th>Kredit</th>
            </tr>
        </thead>
        <tbody>
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
        </tbody>
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
# Streamlit App
# ==============================
st.set_page_config("Mini Akunting - Voucher Jurnal", layout="wide")

st.title("📑 Mini Akunting - Voucher Jurnal")

uploaded = st.file_uploader("Upload Jurnal (Excel)", type=["xlsx", "xls"])

if uploaded:
    df = pd.read_excel(uploaded)

    # pastikan kolom konsisten
    df.columns = [c.strip().title() for c in df.columns]

    st.write("### Data Jurnal")
    st.dataframe(df)

    # ambil nomor voucher unik
    voucher_list = df["Nomor Voucher Jurnal"].dropna().unique()
    voucher_no = st.selectbox("Pilih Nomor Voucher", voucher_list)

    if voucher_no:
        data_voucher = df[df["Nomor Voucher Jurnal"] == voucher_no]

        st.subheader("🖨 Print Preview Voucher")

        # form custom data perusahaan
        perusahaan = st.text_input("Nama Perusahaan", "Perusahaan ABC")
        alamat = st.text_area("Alamat Perusahaan", "Jl. Contoh No. 123 Jakarta")
        direktur = st.text_input("Nama Direktur", "Direktur Utama")
        finance = st.text_input("Nama Finance", "Manager Finance")

        if st.button("Preview Voucher"):
            html = preview_voucher(data_voucher, voucher_no, perusahaan, alamat, direktur, finance)
            st.markdown(html, unsafe_allow_html=True)
