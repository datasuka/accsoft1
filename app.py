def buat_voucher(df, no_voucher, settings):
    pdf = FPDF("P", "mm", "A4")
    pdf.add_page()
    pdf.set_font("Arial", "B", 12)

    # Header perusahaan dengan logo
    if settings.get("logo"):
        logo_w = settings.get("logo_size", 25)
        pdf.image(settings["logo"], 10, 8, logo_w)

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

    # table header (tanpa deskripsi)
    col_widths = [30, 25, 80, 30, 30]
    headers = ["Tanggal","Akun","Nama Akun","Debit","Kredit"]

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
            f"{debit_val:,}".replace(",", "."),
            f"{kredit_val:,}".replace(",", ".")
        ]

        # hitung tinggi baris
        line_counts = []
        for val, w in zip(values, col_widths):
            lines = pdf.multi_cell(w, 6, val, split_only=True)
            line_counts.append(len(lines))
        max_lines = max(line_counts)
        row_height = max_lines * 6

        # gambar kotak tiap kolom
        x_start = pdf.get_x()
        y_start = pdf.get_y()
        for i, (val, w) in enumerate(zip(values, col_widths)):
            pdf.rect(x_start, y_start, w, row_height)   # kotak
            pdf.set_xy(x_start, y_start)
            align = "R" if i in [3,4] else "L"
            pdf.multi_cell(w, 6, val, align=align)      # isi text
            x_start += w
        pdf.set_y(y_start + row_height)

        total_debit += debit_val
        total_kredit += kredit_val

    # total row
    pdf.set_font("Arial","B",9)
    pdf.cell(sum(col_widths[:-2]),8,"Total",border=1,align="L")  # rata kiri
    pdf.cell(col_widths[3],8,f"{total_debit:,}".replace(",", "."),border=1,align="R")
    pdf.cell(col_widths[4],8,f"{total_kredit:,}".replace(",", "."),border=1,align="R")
    pdf.ln()

    # deskripsi ambil dari baris pertama saja
    first_desc = str(data.iloc[0]["Deskripsi"]) if not data.empty else ""
    if first_desc.strip() != "":
        pdf.set_font("Arial","",9)
        pdf.multi_cell(0,6,f"Deskripsi: {first_desc}", border=1)

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
