import streamlit as st
import pandas as pd
from io import BytesIO
from fpdf import FPDF
from num2words import num2words

# =========================
# Fungsi bersih-bersih data
# =========================
def bersihkan_jurnal(df):
    df.columns = [c.strip() for c in df.columns]
    df["Tanggal"] = pd.to_datetime(df["Tanggal"], errors="coerce").dt.strftime("%d/%m/%Y")
    df["Debet"] = pd.to_numeric(df["Debet"], errors="coerce").fillna(0)
    df["Kredit"] = pd.to_numeric(df["Kredit"], errors="coerce").fillna(0)
    return df

# =========================
# Fungsi cetak PDF voucher
# =========================
def buat_voucher(df, no_voucher, settings):
    pdf = FPDF("P", "mm", "A4")
    pdf.add_page()
    pdf.set_font("Arial", "B", 12)

    # Logo
    if settings.get("logo") is not None:
        pdf.image(settings["logo"], 10, 8, 25)

    # Nama perusahaan & alamat
    pdf.cell(200, 10, settings.get("perusahaan",""), ln=1, align="C")
    pdf.set_font("Arial", "", 10)
    pdf.multi_cell(200, 5, settings.get("alamat",""), align="C")
    pdf.ln(5)

    pdf.set_font("Arial", "B", 12)
    pdf.cell(200, 10, "BUKTI VOUCHER JURNAL", ln=1, align="C")
    pdf.set_font("Arial", "", 10)
    pdf.cell(200, 7, f"No Voucher : {no_voucher}", ln=1, align="C")
    pdf.ln(3)

    group = df[df["Nomor Voucher Jurnal"] == no_voucher]

    # Header tabel
    pdf.set_font("Arial", "B", 9)
    col_widths = [25, 20, 55, 50, 20, 20]
    headers = ["Tanggal", "Akun", "Nama Akun", "Deskripsi", "Debet", "Kredit"]
    for i, h in enumerate(headers):
        pdf.cell(col_widths[i], 8, h, 1, 0, "C")
    pdf.ln()

    # Isi tabel
    pdf.set_font("Arial", "", 9)
    for _, row in group.iterrows():
        pdf.cell(col_widths[0], 8, str(row["Tanggal"]), 1)
        pdf.cell(col_widths[1], 8, str(row["No Akun"]), 1)
        pdf.cell(col_widths[2], 8, str(row["Akun"])[:30], 1)  # wrap manual
        pdf.cell(col_widths[3], 8, str(row["Deskripsi"])[:30], 1)
        pdf.cell(col_widths[4], 8, f"{row['Debet']:,.0f}", 1, 0, "R")
        pdf.cell(col_widths[5], 8, f"{row['Kredit']:,.0f}", 1, 0, "R")
        pdf.ln()

    # Total
    total_debit = group["Debet"].sum()
    total_kredit = group["Kredit"].sum()
    pdf.set_font("Arial", "B", 9)
    pdf.cell(sum(col_widths[:-2]), 8, "Total", 1)
    pdf.cell(col_widths[4], 8, f"{total_debit:,.0f}", 1, 0, "R")
    pdf.cell(col_widths[5], 8, f"{total_kredit:,.0f}", 1, 0, "R")
    pdf.ln(12)

    # Terbilang
    pdf.set_font("Arial", "I", 9)
    pdf.cell(200, 5, f"Terbilang: {num2words(int(total_debit), lang='id')} rupiah", ln=1)

    # Tanda tangan
    pdf.ln(20)
    pdf.set_font("Arial", "", 10)
    pdf.cell(100, 5, "Direktur,", 0, 0, "C")
    pdf.cell(100, 5, "Finance,", 0, 1, "C")
    pdf.ln(20)
    pdf.cell(100, 5, f"({settings.get('direktur','')})", 0, 0, "C")
    pdf.cell(100, 5, f"({settings.get('finance','')})", 0, 1, "C")

    out = BytesIO()
    pdf.output(out)
    return out.getvalue()

# =========================
# Streamlit UI
# =========================
st.title("🧾 Mini Akunting - Voucher Jurnal")

file = st.file_uploader("Upload Jurnal (Excel)", type=["xlsx", "xls"])

# Pengaturan perusahaan (langsung di main form)
with st.form("settings_form"):
    st.subheader("⚙️ Pengaturan Perusahaan")
    perusahaan = st.text_input("Nama Perusahaan")
    alamat = st.text_area("Alamat Perusahaan")
    direktur = st.text_input("Nama Direktur")
    finance = st.text_input("Nama Finance")
    logo = st.file_uploader("Upload Logo (PNG/JPG)", type=["png", "jpg", "jpeg"])
    submitted = st.form_submit_button("Simpan Pengaturan")

settings = {
    "perusahaan": perusahaan,
    "alamat": alamat,
    "direktur": direktur,
    "finance": finance,
    "logo": logo if logo else None
}

if file:
    df = pd.read_excel(file)
    df = bersihkan_jurnal(df)

    no_voucher = st.selectbox("Pilih Nomor Voucher", df["Nomor Voucher Jurnal"].unique())

    if no_voucher:
        group = df[df["Nomor Voucher Jurnal"] == no_voucher]

        # =========================
        # PRINT PREVIEW VOUCHER
        # =========================
        st.subheader("🖨️ Print Preview Voucher")

        html = f"""
        <div style="text-align:center;">
            <h3>{settings.get('perusahaan','')}</h3>
            <p>{settings.get('alamat','')}</p>
            <h4>BUKTI VOUCHER JURNAL</h4>
            <p>No Voucher : {no_voucher}</p>
        </div>
        <table style="width:100%; border-collapse: collapse;" border="1">
            <tr>
                <th>Tanggal</th>
                <th>Akun</th>
                <th>Nama Akun</th>
                <th>Deskripsi</th>
                <th>Debet</th>
                <th>Kredit</th>
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
        total_debit = group["Debet"].sum()
        total_kredit = group["Kredit"].sum()
        html += f"""
            <tr>
                <td colspan="4" style="text-align:right;"><b>Total</b></td>
                <td style="text-align:right;"><b>{total_debit:,.0f}</b></td>
                <td style="text-align:right;"><b>{total_kredit:,.0f}</b></td>
            </tr>
        </table>
        <p><i>Terbilang: {num2words(int(total_debit), lang='id')} rupiah</i></p>
        <br><br>
        <table style="width:100%; text-align:center; border:0;">
            <tr>
                <td>Direktur,</td>
                <td>Finance,</td>
            </tr>
            <tr><td height="60px"></td><td></td></tr>
            <tr>
                <td>({settings.get('direktur','')})</td>
                <td>({settings.get('finance','')})</td>
            </tr>
        </table>
        """
        st.markdown(html, unsafe_allow_html=True)

        # =========================
        # Tombol Cetak PDF
        # =========================
        if st.button("📥 Cetak Voucher Jurnal"):
            pdf_file = buat_voucher(df, no_voucher, settings)
            st.download_button("Download PDF", data=pdf_file,
                               file_name=f"voucher_{no_voucher}.pdf")
