import streamlit as st
import pandas as pd
from io import BytesIO
from fpdf import FPDF
from num2words import num2words

def bersihkan_jurnal(df):
    df.columns = [c.strip() for c in df.columns]
    df["Tanggal"] = pd.to_datetime(df["Tanggal"], errors="coerce").dt.strftime("%d/%m/%Y")
    df["Debet"] = pd.to_numeric(df["Debet"], errors="coerce").fillna(0)
    df["Kredit"] = pd.to_numeric(df["Kredit"], errors="coerce").fillna(0)
    return df

# --- Fungsi PDF tetap sama (sudah ada) ---

st.title("🧾 Mini Akunting - Voucher Jurnal")

file = st.file_uploader("Upload Jurnal (Excel)", type=["xlsx", "xls"])

with st.form("settings_form"):
    st.subheader("⚙️ Pengaturan Perusahaan")
    perusahaan = st.text_input("Nama Perusahaan")
    alamat = st.text_area("Alamat Perusahaan")
    direktur = st.text_input("Nama Direktur")
    finance = st.text_input("Nama Finance")
    logo = st.file_uploader("Upload Logo", type=["png", "jpg", "jpeg"])
    submitted = st.form_submit_button("Simpan Pengaturan")

settings = {
    "perusahaan": perusahaan,
    "alamat": alamat,
    "direktur": direktur,
    "finance": finance,
    "logo": logo
}

if file:
    df = pd.read_excel(file)
    df = bersihkan_jurnal(df)

    no_voucher = st.selectbox("Pilih Nomor Voucher", df["Nomor Voucher Jurnal"].unique())

    if no_voucher:
        group = df[df["Nomor Voucher Jurnal"] == no_voucher]
        total_debit = group["Debet"].sum()
        total_kredit = group["Kredit"].sum()

        # --- Print Preview mirip PDF ---
        st.subheader("🖨️ Print Preview Voucher")

        html = f"""
        <div style="text-align:center;">
            <h3>{settings.get('perusahaan','')}</h3>
            <p>{settings.get('alamat','')}</p>
            <h4 style="margin-top:20px;">BUKTI VOUCHER JURNAL</h4>
            <p>No Voucher : {no_voucher}</p>
        </div>
        <br>
        <table border="1" cellspacing="0" cellpadding="4" style="border-collapse:collapse; width:100%; font-size:13px;">
            <tr style="background-color:#f0f0f0; font-weight:bold; text-align:center;">
                <td>Tanggal</td>
                <td>Akun</td>
                <td>Nama Akun</td>
                <td>Deskripsi</td>
                <td>Debit</td>
                <td>Kredit</td>
            </tr>
        """

        for _, row in group.iterrows():
            html += f"""
            <tr>
                <td>{row['Tanggal']}</td>
                <td>{row['No Akun']}</td>
                <td>{row['Akun']}</td>
                <td>{row['Deskripsi']}</td>
                <td style="text-align:right;">{row['Debet']:,.0f}</td>
                <td style="text-align:right;">{row['Kredit']:,.0f}</td>
            </tr>
            """

        html += f"""
            <tr style="font-weight:bold;">
                <td colspan="4" style="text-align:right;">Total</td>
                <td style="text-align:right;">{total_debit:,.0f}</td>
                <td style="text-align:right;">{total_kredit:,.0f}</td>
            </tr>
        </table>
        <br>
        <p><i>Terbilang: {num2words(int(total_debit), lang='id')} rupiah</i></p>
        <br><br><br>
        <table style="width:100%; text-align:center;">
            <tr>
                <td>Direktur,<br><br><br>({settings.get('direktur','')})</td>
                <td>Finance,<br><br><br>({settings.get('finance','')})</td>
            </tr>
        </table>
        """

        st.markdown(html, unsafe_allow_html=True)

        # --- Tombol cetak PDF ---
        if st.button("Cetak Voucher Jurnal"):
            pdf_file = buat_voucher(df, no_voucher, settings)
            st.download_button("Download PDF", data=pdf_file,
                               file_name=f"voucher_{no_voucher}.pdf")
