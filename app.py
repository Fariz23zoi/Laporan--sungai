# ======================
# EXPORT EXCEL A4 PRO
# ======================
if st.button("Export Excel A4"):

    from openpyxl import Workbook
    from openpyxl.drawing.image import Image as XLImage
    from openpyxl.styles import Alignment, Font, Border, Side
    import tempfile
    import math
    import requests

    wb = Workbook()
    ws = wb.active

    # ======================
    # SET A4
    # ======================
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.orientation = ws.ORIENTATION_PORTRAIT

    # ======================
    # STYLE
    # ======================
    center = Alignment(horizontal='center', vertical='center', wrap_text=True)
    left = Alignment(horizontal='left', vertical='top', wrap_text=True)
    bold = Font(bold=True)

    thin = Side(style='thin')
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    # ======================
    # KOLOM
    # ======================
    ws.column_dimensions['A'].width = 5
    ws.column_dimensions['B'].width = 25
    ws.column_dimensions['C'].width = 20
    ws.column_dimensions['D'].width = 25
    ws.column_dimensions['E'].width = 35

    # ======================
    # JUDUL
    # ======================
    ws.merge_cells('A1:E1')
    ws.merge_cells('A2:E2')
    ws.merge_cells('A3:E3')

    ws['A1'] = "LAPORAN MONITORING SUNGAI"
    ws['A2'] = "OPSDA"
    ws['A3'] = datetime.now().strftime("%B %Y")

    for i in range(1,4):
        ws[f"A{i}"].alignment = center
        ws[f"A{i}"].font = bold

    # ======================
    # HEADER
    # ======================
    headers = ["NO","URAIAN","LOKASI","KETERANGAN","FOTO"]

    for col, val in enumerate(headers,1):
        c = ws.cell(row=5, column=col)
        c.value = val
        c.font = bold
        c.alignment = center
        c.border = border

    # ======================
    # DATA
    # ======================
    row = 6

    for i, d in enumerate(data, start=1):

        ws.row_dimensions[row].height = 140

        ws.cell(row=row, column=1, value=i)
        ws.cell(row=row, column=2, value=d["uraian"])
        ws.cell(row=row, column=3, value=d["lokasi"])
        ws.cell(row=row, column=4, value=d["keterangan"])

        for col in [1,2,3,4]:
            ws.cell(row=row, column=col).alignment = left
            ws.cell(row=row, column=col).border = border

        # ======================
        # AMBIL FOTO DARI DB
        # ======================
        fotos = supabase.table("foto").select("*").eq("laporan_id", d["id"]).execute().data

        if fotos:
            try:
                img_url = fotos[0]["url_foto"]
                img_bytes = requests.get(img_url).content

                with tempfile.NamedTemporaryFile(delete=False) as tmp:
                    tmp.write(img_bytes)
                    img = XLImage(tmp.name)

                # KOMPRES + SIZE
                img.width = 240
                img.height = 130

                img.anchor = f"E{row}"
                ws.add_image(img)

            except:
                pass

        ws.cell(row=row, column=5).border = border

        row += 1

    # ======================
    # SAVE
    # ======================
    from io import BytesIO
    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)

    st.download_button(
        "⬇️ Download Excel A4",
        buf,
        file_name="laporan_sungai_A4.xlsx"
    )
