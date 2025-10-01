def buat_voucher(df, no_voucher, settings):
    pdf = FPDF("P", "mm", "A4")
    pdf.set_left_margin(15)
    pdf.set_right_margin(15)
    pdf.add_page()
    pdf.set_font("Arial", "B", 12)

    # Header perusahaan + logo
    if settings.get("logo"):
        pdf.image(settings["logo"], 15, 8, settings.get("logo_size", 20))

    pdf.set_xy(0, 10)
    pdf.set_font("Arial", "B", 12)
    pdf.cell(0, 6, settings.get("perusahaan", ""), ln=1, align="C")

    pdf.set_font("Arial", "", 10)
    pdf.multi_cell(0, 5, settings.get("alamat", ""), align="C")
    pdf.ln(3)

    # Judul
    pdf.set_font("Arial", "B", 12)
    pdf.cell(0, 8, "BUKTI VOUCHER JURNAL", ln=1, align="C")
    pdf.set_font("Arial", "", 10)
    pdf.cell(0, 6, f"No Voucher : {no_voucher}", ln=1, align="C")
    pdf.ln(4)

    # Data voucher
    data = df[df["Nomor Voucher Jurnal"] == no_voucher]

    # Lebar kolom (pakai margin 15mm kiri kanan)
    page_width = 210 - 30  # A4 (210mm) - left+right margin (15+15)
    col_widths = [28, 20, 60, 40, 31, 31]
    headers = ["Tanggal", "Akun", "Nama Akun", "Deskripsi", "Debit", "Kredit"]

    # Header tabel
    pdf.set_font("Arial", "B", 9)
    for h, w in zip(headers, col_widths):
        pdf.cell(w, 8, h, border=1, align="C")
    pdf.ln()

    total_debit, total_kredit = 0, 0
    pdf.set_font("Arial", "", 9)

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

        # Hitung tinggi baris
        line_counts = []
        for i, (val, w) in enumerate(zip(values, col_widths)):
            if i in [4, 5]:  # Debit, Kredit → single line
                line_counts.append(1)
            else:
                lines = pdf.multi_cell(w, 6, val, split_only=True)
                line_counts.append(len(lines))
        row_height = max(line_counts) * 6

        # Gambar baris tabel
        x_start = pdf.get_x()
        y_start = pdf.get_y()
        for i, (val, w) in enumerate(zip(values, col_widths)):
            pdf.rect(x_start, y_start, w, row_height)  # kotak
            pdf.set_xy(x_start, y_start)

            if i in [4, 5]:
                pdf.cell(w, row_height, val, align="R")
            else:
                pdf.multi_cell(w, 6, val, align="L")
            x_start += w
        pdf.set_y(y_start + row_height)

        total_debit += debit_val
        total_kredit += kredit_val

    # Total row
    pdf.set_font("Arial", "B", 9)
    pdf.cell(sum(col_widths[:-2]), 8, "Total", border=1, align="L")
    pdf.cell(col_widths[4], 8, f"{total_debit:,}".replace(",", "."), border=1, align="R")
    pdf.cell(col_widths[5], 8, f"{total_kredit:,}".replace(",", "."), border=1, align="R")
    pdf.ln()

    # Ambil deskripsi pertama
    first_desc = str(data.iloc[0]["Deskripsi"]) if not data.empty else ""

    # Terbilang & Deskripsi sebagai baris tabel dengan border
    pdf.set_font("Arial", "I", 9)
    pdf.multi_cell(page_width, 8, f"Terbilang : {num2words(total_debit, lang='id')} rupiah", border=1)

    pdf.set_font("Arial", "", 9)
    pdf.multi_cell(page_width, 8, f"Deskripsi : {first_desc}", border=1)

    # Tanda tangan
    pdf.ln(12)
    pdf.set_font("Arial", "", 10)
    pdf.cell(page_width/2, 6, "Direktur,", align="C")
    pdf.cell(page_width/2, 6, "Finance,", align="C", ln=1)
    pdf.ln(20)
    pdf.cell(page_width/2, 6, f"({settings.get('direktur','')})", align="C")
    pdf.cell(page_width/2, 6, f"({settings.get('finance','')})", align="C", ln=1)

    buffer = BytesIO()
    pdf.output(buffer)
    return buffer
