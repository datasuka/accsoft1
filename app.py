    # total row
    pdf.set_font("Arial","B",9)
    pdf.cell(sum(col_widths[:-2]),8,"Total",border=1,align="R")
    pdf.cell(col_widths[3],8,f"{total_debit:,}".replace(",", "."),border=1,align="R")
    pdf.cell(col_widths[4],8,f"{total_kredit:,}".replace(",", "."),border=1,align="R")
    pdf.ln()

    # ambil deskripsi dari baris pertama saja
    first_desc = str(data.iloc[0]["Deskripsi"]) if not data.empty else ""
    if first_desc.strip() != "":
        pdf.set_font("Arial","",9)
        pdf.multi_cell(0,6,f"Deskripsi: {first_desc}", border=1)
