from openpyxl import Workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.styles import Alignment, Font, Border, Side
import tempfile

if st.button("📊 Export Excel"):
    wb = Workbook()
    ws = wb.active

    center = Alignment(horizontal='center', vertical='center', wrap_text=True)
    bold = Font(bold=True)
    border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )

    # HEADER
    ws.merge_cells('A1:F1')
    ws['A1'] = "LAPORAN PENELUSURAN SUNGAI CIPETUNGAN"
    ws['A1'].font = Font(size=14, bold=True)
    ws['A1'].alignment = center

    ws.merge_cells('A2:F2')
    ws['A2'] = "KOTA/KAB. CIAMIS"
    ws['A2'].alignment = center

    ws.merge_cells('A3:F3')
    ws['A3'] = "OPERASI DAN PEMELIHARAAN SDA III"
    ws['A3'].alignment = center

    ws.merge_cells('A4:F4')
    ws['A4'] = f"BULAN {bulan.upper()} TAHUN {tahun}"
    ws['A4'].alignment = center

    headers = ["NO","SUNGAI","LOKASI","KOORDINAT","FOTO","KETERANGAN"]
    ws.append(headers)

    for col in range(1, 7):
        cell = ws.cell(row=5, column=col)
        cell.font = bold
        cell.alignment = center
        cell.border = border

    ws.column_dimensions['A'].width = 5
    ws.column_dimensions['B'].width = 20
    ws.column_dimensions['C'].width = 35
    ws.column_dimensions['D'].width = 20
    ws.column_dimensions['E'].width = 35
    ws.column_dimensions['F'].width = 30

    row = 6

    for i, d in enumerate(st.session_state.data, start=1):
        ws.cell(row=row, column=1, value=i)
        ws.cell(row=row, column=2, value=d["sungai"])
        ws.cell(row=row, column=3, value=d["lokasi"])
        ws.cell(row=row, column=4, value=d["koordinat"])
        ws.cell(row=row, column=6, value=d["keterangan"])

        for col in [1,2,3,4,6]:
            c = ws.cell(row=row, column=col)
            c.alignment = center
            c.border = border

        # FOTO
        with tempfile.NamedTemporaryFile(delete=False) as tmp:
            tmp.write(d["foto"].read())
            img = XLImage(tmp.name)

        img.width = 300
        img.height = 180
        ws.add_image(img, f'E{row}')

        ws.cell(row=row, column=5).border = border
        ws.row_dimensions[row].height = 140

        row += 1

    filename = f"Laporan_Sungai_{bulan}_{tahun}.xlsx"
    wb.save(filename)

    with open(filename, "rb") as f:
        st.download_button("⬇️ Download Excel", f, file_name=filename)
