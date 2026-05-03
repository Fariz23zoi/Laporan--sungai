import streamlit as st
from openpyxl import Workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.styles import Alignment, Font, Border, Side
from datetime import datetime
import tempfile
import math
import re

st.title("📋 Laporan Penelusuran Sungai")

# ======================
# INPUT
# ======================
bulan_list = [
    "Januari","Februari","Maret","April","Mei","Juni",
    "Juli","Agustus","September","Oktober","November","Desember"
]

bulan = st.selectbox("Bulan", bulan_list)
tahun = st.number_input("Tahun", value=2026)

bulan_romawi = {
    "Januari": "I","Februari": "II","Maret": "III","April": "IV",
    "Mei": "V","Juni": "VI","Juli": "VII","Agustus": "VIII",
    "September": "IX","Oktober": "X","November": "XI","Desember": "XII"
}

# ======================
# DATA
# ======================
if "data" not in st.session_state:
    st.session_state.data = []

st.subheader("Input Data")

uraian = st.text_input("Uraian (contoh: Sungai Cimuntur)")
lokasi = st.text_area("Lokasi")
koordinat = st.text_input("Koordinat")
keterangan = st.text_area("Keterangan")
foto = st.file_uploader("Upload Foto", type=["jpg","png","jpeg"])

if st.button("➕ Tambah Data"):
    if foto:
        st.session_state.data.append({
            "uraian": uraian,
            "lokasi": lokasi,
            "koordinat": koordinat,
            "keterangan": keterangan,
            "foto": foto
        })
        st.success("Data ditambahkan!")
    else:
        st.warning("Upload foto dulu!")

st.write(f"Jumlah data: {len(st.session_state.data)}")

# ======================
# EXPORT
# ======================
if st.button("📊 Export Excel"):
    wb = Workbook()
    ws = wb.active

    # ======================
    # SET A4 PORTRAIT
    # ======================
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.orientation = ws.ORIENTATION_PORTRAIT

    # ======================
    # STYLE
    # ======================
    center = Alignment(horizontal='center', vertical='center', wrap_text=True)
    left_top = Alignment(horizontal='left', vertical='top', wrap_text=True)
    bold = Font(bold=True)

    thin = Side(style='thin')
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    # ======================
    # KOLOM
    # ======================
    ws.column_dimensions['A'].width = 5
    ws.column_dimensions['B'].width = 24
    ws.column_dimensions['C'].width = 20
    ws.column_dimensions['D'].width = 22
    ws.column_dimensions['E'].width = 30
    ws.column_dimensions['F'].width = 32

    # ======================
    # HITUNG HALAMAN
    # ======================
    total_data = len(st.session_state.data)
    data_per_page = 5
    total_halaman = max(1, math.ceil(total_data / data_per_page))

    nomor_laporan = f"{total_halaman}/OPSDA-03/{bulan_romawi[bulan]}/{tahun}"

    # ======================
    # AMBIL NAMA SUNGAI DARI URAIAN
    # ======================
    uraian_text = st.session_state.data[0]["uraian"] if total_data > 0 else ""
    match = re.search(r"sungai\s+(.*)", uraian_text.lower())
    nama_sungai = match.group(1).upper() if match else uraian_text.upper()

    # ======================
    # JUDUL
    # ======================
    ws.merge_cells('A1:F1')
    ws.merge_cells('A2:F2')
    ws.merge_cells('A3:F3')
    ws.merge_cells('A4:F4')
    ws.merge_cells('A5:F5')

    ws['A1'] = f"LAPORAN PENELUSURAN SUNGAI {nama_sungai}"
    ws['A2'] = "KOTA/KAB. CIAMIS"
    ws['A3'] = "OPERASI DAN PEMELIHARAAN SDA III"
    ws['A4'] = f"BULAN {bulan.upper()} TAHUN {tahun}"
    ws['A5'] = f"Nomor: {nomor_laporan}"

    for i in range(1, 6):
        ws[f"A{i}"].alignment = center
        ws[f"A{i}"].font = bold

    # ======================
    # HEADER
    # ======================
    headers = ["NO","URAIAN","LOKASI","KOORDINAT","FOTO","KETERANGAN"]

    for col, val in enumerate(headers, 1):
        c = ws.cell(row=6, column=col)
        c.value = val
        c.font = bold
        c.alignment = center
        c.border = border

    # ======================
    # DATA
    # ======================
    row = 8

    for i, d in enumerate(st.session_state.data, start=1):

        ws.row_dimensions[row].height = 130

        ws.cell(row=row, column=1, value=i)
        ws.cell(row=row, column=2, value=d["uraian"])
        ws.cell(row=row, column=3, value=d["lokasi"])
        ws.cell(row=row, column=4, value=d["koordinat"])
        ws.cell(row=row, column=6, value=d["keterangan"])

        for col in [1,2,3,4,6]:
            ws.cell(row=row, column=col).alignment = left_top
            ws.cell(row=row, column=col).border = border

        # FOTO MASUK DALAM CELL
        with tempfile.NamedTemporaryFile(delete=False) as tmp:
            tmp.write(d["foto"].read())
            img = XLImage(tmp.name)

        img.width = 220
        img.height = 120
        img.anchor = f"E{row}"
        ws.add_image(img)

        ws.cell(row=row, column=5).border = border

        row += 1

    # ======================
    # TANGGAL + TTD
    # ======================
    ttd_row = row + 2

    today = datetime.now()
    tanggal = f"Ciamis, {today.day} {bulan_list[today.month-1]} {today.year}"

    ws.merge_cells(start_row=ttd_row-1, start_column=5, end_row=ttd_row-1, end_column=6)
    ws.cell(row=ttd_row-1, column=5).value = tanggal
    ws.cell(row=ttd_row-1, column=5).alignment = center

    ws.merge_cells(start_row=ttd_row, start_column=5, end_row=ttd_row, end_column=6)
    ws.cell(row=ttd_row, column=5).value = "Juru Sungai OPSDA 03"
    ws.cell(row=ttd_row, column=5).alignment = center

    try:
        ttd_img = XLImage("ttd.png")
        ttd_img.width = 160
        ttd_img.height = 80
        ws.add_image(ttd_img, f"E{ttd_row+1}")
    except:
        pass

    ws.merge_cells(start_row=ttd_row+3, start_column=5, end_row=ttd_row+3, end_column=6)
    ws.cell(row=ttd_row+3, column=5).value = "Fariz Rionaldi"
    ws.cell(row=ttd_row+3, column=5).alignment = center

    # ======================
    # SAVE
    # ======================
    filename = f"Laporan_Sungai_{bulan}_{tahun}.xlsx"
    wb.save(filename)

    with open(filename, "rb") as f:
        st.download_button("⬇️ Download Excel", f, file_name=filename)
