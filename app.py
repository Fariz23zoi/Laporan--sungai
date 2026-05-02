import streamlit as st
from openpyxl import Workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.styles import Alignment, Font, Border, Side
from datetime import datetime
import tempfile

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

# ======================
# ROMAWI
# ======================
bulan_romawi = {
    "Januari": "I","Februari": "II","Maret": "III","April": "IV",
    "Mei": "V","Juni": "VI","Juli": "VII","Agustus": "VIII",
    "September": "IX","Oktober": "X","November": "XI","Desember": "XII"
}

nomor_laporan = f"600/OPSDA-03/{bulan_romawi[bulan]}/{tahun}"

# ======================
# DATA
# ======================
if "data" not in st.session_state:
    st.session_state.data = []

st.subheader("Input Data")

uraian = st.text_input("Uraian")
lokasi = st.text_area("Lokasi")
koordinat = st.text_input("Koordinat")
keterangan = st.text_area("Keterangan", "Kondisi tebing masih dalam keadaan baik")
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
# EXPORT EXCEL
# ======================
if st.button("📊 Export Excel"):
    wb = Workbook()
    ws = wb.active

    # STYLE
    center = Alignment(horizontal='center', vertical='center', wrap_text=True)
    left_top = Alignment(horizontal='left', vertical='top', wrap_text=True)
    bold = Font(bold=True)

    thin = Side(style='thin')
    border_all = Border(left=thin, right=thin, top=thin, bottom=thin)

    # ======================
    # UKURAN KOLOM
    # ======================
    ws.column_dimensions['A'].width = 5
    ws.column_dimensions['B'].width = 24
    ws.column_dimensions['C'].width = 20
    ws.column_dimensions['D'].width = 22
    ws.column_dimensions['E'].width = 30
    ws.column_dimensions['F'].width = 32

    # ======================
    # TINGGI BARIS
    # ======================
    ws.row_dimensions[1].height = 30
    ws.row_dimensions[2].height = 20
    ws.row_dimensions[3].height = 20
    ws.row_dimensions[4].height = 20
    ws.row_dimensions[5].height = 20
    ws.row_dimensions[6].height = 30

    # ======================
    # JUDUL
    # ======================
    ws.merge_cells('A1:F1')
    ws.merge_cells('A2:F2')
    ws.merge_cells('A3:F3')
    ws.merge_cells('A4:F4')
    ws.merge_cells('A5:F5')

    ws['A1'] = "LAPORAN PENELUSURAN SUNGAI CIWADORI"
    ws['A2'] = "KOTA/KAB. CIAMIS"
    ws['A3'] = "OPERASI DAN PEMELIHARAAN SDA III"
    ws['A4'] = f"BULAN {bulan.upper()} TAHUN {tahun}"
    ws['A5'] = f"Nomor: {nomor_laporan}"

    for i in range(1, 6):
        ws[f"A{i}"].alignment = center
        if i == 1:
            ws[f"A{i}"].font = Font(bold=True, size=14)

    # ======================
    # HEADER
    # ======================
    headers = ["NO","URAIAN","LOKASI","KOORDINAT","FOTO","KETERANGAN"]

    for col, val in enumerate(headers, 1):
        c = ws.cell(row=6, column=col)
        c.value = val
        c.font = bold
        c.alignment = center
        c.border = border_all

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

        ws.cell(row=row, column=1).alignment = center
        ws.cell(row=row, column=2).alignment = left_top
        ws.cell(row=row, column=3).alignment = left_top
        ws.cell(row=row, column=4).alignment = left_top
        ws.cell(row=row, column=6).alignment = left_top

        for col in range(1, 7):
            ws.cell(row=row, column=col).border = border_all

        # FOTO
        with tempfile.NamedTemporaryFile(delete=False) as tmp:
            tmp.write(d["foto"].read())
            img = XLImage(tmp.name)

        img.width = 260
        img.height = 150
        ws.add_image(img, f'E{row}')

        row += 1

    # ======================
    # TANDA TANGAN + TANGGAL
    # ======================
    ttd_row = row + 2

    bulan_indo = bulan_list
    today = datetime.now()
    tanggal = f"Ciamis, {today.day} {bulan_indo[today.month-1]} {today.year}"

    # Tanggal
    ws.merge_cells(start_row=ttd_row-1, start_column=5, end_row=ttd_row-1, end_column=6)
    ws.cell(row=ttd_row-1, column=5).value = tanggal
    ws.cell(row=ttd_row-1, column=5).alignment = center

    # Jabatan
    ws.merge_cells(start_row=ttd_row, start_column=5, end_row=ttd_row, end_column=6)
    ws.cell(row=ttd_row, column=5).value = "Juru Sungai OPSDA 03"
    ws.cell(row=ttd_row, column=5).alignment = center

    # Gambar tanda tangan (opsional)
    try:
        ttd_img = XLImage("ttd.png")
        ttd_img.width = 160
        ttd_img.height = 80
        ws.add_image(ttd_img, f"E{ttd_row+1}")
    except:
        pass

    # Nama
    ws.merge_cells(start_row=ttd_row+3, start_column=5, end_row=ttd_row+3, end_column=6)
    ws.cell(row=ttd_row+3, column=5).value = "Faniz Rionaldi"
    ws.cell(row=ttd_row+3, column=5).alignment = center

    # ======================
    # SAVE
    # ======================
    filename = f"Laporan_Sungai_{bulan}_{tahun}.xlsx"
    wb.save(filename)

    with open(filename, "rb") as f:
        st.download_button("⬇️ Download Excel", f, file_name=filename)
