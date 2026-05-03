import streamlit as st
from openpyxl import Workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.styles import Alignment, Font, Border, Side
from datetime import datetime
from io import BytesIO
from PIL import Image as PILImage
import tempfile

st.title("📋 Laporan Penelusuran Sungai")

# ======================
# INPUT
# ======================
bulan = st.selectbox("Bulan", [
    "Januari","Februari","Maret","April","Mei","Juni",
    "Juli","Agustus","September","Oktober","November","Desember"
])

tahun = st.number_input("Tahun", value=2026)

if "data" not in st.session_state:
    st.session_state.data = []

st.subheader("Input Data")

uraian = st.text_input("Uraian")
lokasi = st.text_input("Lokasi")
koordinat = st.text_input("Koordinat")
keterangan = st.text_input("Keterangan")
foto = st.file_uploader("Upload Foto", type=["jpg","png","jpeg"])

if st.button("Tambah Data"):
    if uraian and lokasi and foto:
        st.session_state.data.append({
            "uraian": uraian,
            "lokasi": lokasi,
            "koordinat": koordinat,
            "keterangan": keterangan,
            "foto": foto
        })
        st.success("Data ditambahkan")

# ======================
# EXPORT EXCEL
# ======================
if st.button("Export Excel"):

    wb = Workbook()
    ws = wb.active

    # STYLE
    center = Alignment(horizontal='center', vertical='center', wrap_text=True)
    bold = Font(bold=True)
    border = Border(
        left=Side(style='medium'),
        right=Side(style='medium'),
        top=Side(style='medium'),
        bottom=Side(style='medium')
    )

    # KOLOM
    ws.column_dimensions['A'].width = 5
    ws.column_dimensions['B'].width = 25
    ws.column_dimensions['C'].width = 20
    ws.column_dimensions['D'].width = 25
    ws.column_dimensions['E'].width = 45
    ws.column_dimensions['F'].width = 30

    # ======================
    # HEADER
    # ======================
    ws.merge_cells('A1:F1')
    ws['A1'] = "LAPORAN SUNGAI CIMUNTUR"
    ws['A1'].alignment = center
    ws['A1'].font = Font(size=16, bold=True)

    ws.merge_cells('A2:F2')
    ws['A2'] = "KOTA/KAB. CIAMIS"
    ws['A2'].alignment = center
    ws['A2'].font = Font(size=14)

    ws.merge_cells('A3:F3')
    ws['A3'] = "OPERASI DAN PEMELIHARAAN SDA III"
    ws['A3'].alignment = center
    ws['A3'].font = Font(size=14)

    ws.merge_cells('A4:F4')
    ws['A4'] = f"BULAN {bulan.upper()} TAHUN {tahun}"
    ws['A4'].alignment = center
    ws['A4'].font = Font(size=14, bold=True)

    ws.merge_cells('A5:F5')
    ws['A5'] = f"Nomor: 1/1/OPSDA-03/V/{tahun}"
    ws['A5'].alignment = center

    # ======================
    # HEADER TABEL
    # ======================
    headers = ["NO","URAIAN","LOKASI","KOORDINAT","FOTO","KETERANGAN"]

    for col, val in enumerate(headers, 1):
        cell = ws.cell(row=7, column=col, value=val)
        cell.alignment = center
        cell.font = Font(bold=True)
        cell.border = border

    # ======================
    # DATA
    # ======================
    row = 8

    for i, d in enumerate(st.session_state.data, start=1):

        ws.row_dimensions[row].height = 210

        ws.cell(row=row, column=1, value=i).alignment = center
        ws.cell(row=row, column=2, value=d["uraian"]).alignment = center
        ws.cell(row=row, column=3, value=d["lokasi"]).alignment = center
        ws.cell(row=row, column=4, value=d["koordinat"]).alignment = center
        ws.cell(row=row, column=6, value=d["keterangan"]).alignment = center

        for col in [1,2,3,4,5,6]:
            ws.cell(row=row, column=col).border = border

        # FOTO
        img_file = d["foto"]
        img = PILImage.open(img_file)

        temp = tempfile.NamedTemporaryFile(delete=False, suffix=".jpg")
        img.save(temp.name)

        xl_img = XLImage(temp.name)
        xl_img.width = 300
        xl_img.height = 170
        xl_img.anchor = f"E{row}"

        ws.add_image(xl_img)

        row += 1

    # ======================
    # TTD
    # ======================
    ttd_row = row + 1
    today = datetime.now()

    ws.merge_cells(start_row=ttd_row, start_column=5, end_row=ttd_row, end_column=6)
    ws.cell(row=ttd_row, column=5, value=f"Ciamis, {today.day} {bulan} {today.year}").alignment = center

    ws.merge_cells(start_row=ttd_row+1, start_column=5, end_row=ttd_row+1, end_column=6)
    ws.cell(row=ttd_row+1, column=5, value="Juru Sungai OPSDA 03").alignment = center

    ws.merge_cells(start_row=ttd_row+4, start_column=5, end_row=ttd_row+4, end_column=6)
    ws.cell(row=ttd_row+4, column=5, value="Fariz Rionaldi").alignment = center

    # SAVE
    output = BytesIO()
    wb.save(output)
    output.seek(0)

    st.download_button(
        "Download Excel",
        data=output,
        file_name="Laporan_Sungai.xlsx"
    )
