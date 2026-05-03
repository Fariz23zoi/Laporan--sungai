import streamlit as st
from openpyxl import Workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.styles import Alignment, Font, Border, Side
from openpyxl.worksheet.pagebreak import Break
from datetime import datetime
from io import BytesIO
from PIL import Image as PILImage
from concurrent.futures import ThreadPoolExecutor
import math
import re

# ======================
# CONFIG
# ======================
st.set_page_config(page_title="Laporan Sungai FINAL", layout="wide")
st.title("📋 Laporan Penelusuran Sungai FINAL STABLE")

# ======================
# BULAN
# ======================
bulan_list = [
    "Januari","Februari","Maret","April","Mei","Juni",
    "Juli","Agustus","September","Oktober","November","Desember"
]

bulan_romawi = {
    "Januari": "I","Februari": "II","Maret": "III","April": "IV",
    "Mei": "V","Juni": "VI","Juli": "VII","Agustus": "VIII",
    "September": "IX","Oktober": "X","November": "XI","Desember": "XII"
}

bulan = st.selectbox("Bulan", bulan_list)
tahun = st.number_input("Tahun", value=2026)

# ======================
# SESSION DATA
# ======================
if "data" not in st.session_state:
    st.session_state.data = []

st.subheader("Input Data")

uraian = st.text_input("Uraian")
lokasi = st.text_area("Lokasi")
koordinat = st.text_input("Koordinat")
foto = st.file_uploader("Upload Foto", type=["jpg","png","jpeg"])

col1, col2 = st.columns(2)

with col1:
    if st.button("➕ Tambah Data"):
        if not uraian or not lokasi:
            st.warning("Uraian & Lokasi wajib diisi!")
        elif not foto:
            st.warning("Upload foto dulu!")
        else:
            st.session_state.data.append({
                "uraian": uraian,
                "lokasi": lokasi,
                "koordinat": koordinat,
                "foto": foto
            })
            st.success("Data berhasil ditambahkan!")

with col2:
    if st.button("🗑 Reset"):
        st.session_state.data = []

st.write(f"Total data: {len(st.session_state.data)}")

if st.session_state.data:
    st.dataframe([{k:v for k,v in d.items() if k!="foto"} for d in st.session_state.data])

# ======================
# IMAGE COMPRESS
# ======================
def compress_image(file_bytes):
    img = PILImage.open(BytesIO(file_bytes))

    if img.mode in ("RGBA", "P"):
        img = img.convert("RGB")

    img.thumbnail((1280, 720))

    buffer = BytesIO()
    img.save(buffer, format="JPEG", quality=65, optimize=True)

    return buffer.getvalue()

# ======================
# EXPORT
# ======================
if st.button("📊 EXPORT EXCEL FINAL 🚀"):

    data = st.session_state.data
    total = len(data)

    if total == 0:
        st.warning("Belum ada data!")
        st.stop()

    progress = st.progress(0)

    def process(d):
        return compress_image(d["foto"].getvalue())

    with ThreadPoolExecutor(max_workers=4) as executor:
        images = list(executor.map(process, data))

    progress.progress(30)

    wb = Workbook()
    ws = wb.active
    ws.title = "Laporan"

    center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    left = Alignment(horizontal="left", vertical="top", wrap_text=True)
    bold = Font(bold=True)

    thin = Side(style="thin")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    ws.column_dimensions["A"].width = 5
    ws.column_dimensions["B"].width = 24
    ws.column_dimensions["C"].width = 22
    ws.column_dimensions["D"].width = 22
    ws.column_dimensions["E"].width = 40
    ws.column_dimensions["F"].width = 30

    ws.page_setup.fitToWidth = 1

    match = re.search(r"sungai\s+(.+)", data[0]["uraian"].lower())
    nama_sungai = match.group(1).upper() if match else data[0]["uraian"].upper()

    per_page = 5
    pages = math.ceil(total / per_page)

    row_global = 1

    for p in range(pages):

        if p > 0:
            ws.row_breaks.append(Break(id=row_global))

        # HEADER (AMAN TANPA merge_cells ERROR)
        def merge(row):
            ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=6)

        merge(row_global)
        ws.cell(row_global,1,f"LAPORAN SUNGAI {nama_sungai}").alignment = center
        ws.cell(row_global,1).font = bold

        merge(row_global+1)
        ws.cell(row_global+1,1,"KOTA/KAB. CIAMIS").alignment = center

        merge(row_global+2)
        ws.cell(row_global+2,1,"OPSDA 03").alignment = center

        merge(row_global+3)
        ws.cell(row_global+3,1,f"BULAN {bulan.upper()} {tahun}").alignment = center

        merge(row_global+4)
        nomor = f"{p+1}/{pages}/OPSDA-03/{bulan_romawi[bulan]}/{tahun}"
        ws.cell(row_global+4,1,f"Nomor: {nomor}").alignment = center

        headers = ["NO","URAIAN","LOKASI","KOORDINAT","FOTO","KETERANGAN"]

        for i,h in enumerate(headers,1):
            c = ws.cell(row_global+5,i,h)
            c.font = bold
            c.alignment = center
            c.border = border

        start = p * per_page
        end = start + per_page
        chunk = data[start:end]

        row = row_global + 7

        for i, d in enumerate(chunk,1):

            ws.row_dimensions[row].height = 160

            ws.cell(row=row, column=1, value=i)
            ws.cell(row=row, column=2, value=d["uraian"])
            ws.cell(row=row, column=3, value=d["lokasi"])
            ws.cell(row=row, column=4, value=d["koordinat"])
            ws.cell(row=row, column=6, value="Kondisi tebing masih dalam keadaan baik")

            for col in [1,2,3,4,6]:
                ws.cell(row=row, column=col).alignment = left
                ws.cell(row=row, column=col).border = border

            img = XLImage(BytesIO(images[start+i-1]))
            img.width = 260
            img.height = 150
            img.anchor = f"E{row}"
            ws.add_image(img)

            row += 1

        ttd = row + 2
        now = datetime.now()

        def merge_sig(r):
            ws.merge_cells(start_row=r, start_column=5, end_row=r, end_column=6)

        merge_sig(ttd)
        ws.cell(ttd,5,f"Ciamis, {now.day} {bulan_list[now.month-1]} {now.year}").alignment = center

        merge_sig(ttd+1)
        ws.cell(ttd+1,5,"Juru Sungai OPSDA 03").alignment = center

        merge_sig(ttd+4)
        ws.cell(ttd+4,5,"Fariz Rionaldi").alignment = center

        row_global = ttd + 6

        progress.progress(int(((p+1)/pages)*100))

    output = BytesIO()
    wb.save(output)
    output.seek(0)

    progress.progress(100)

    st.success("Export berhasil 🚀")

    st.download_button(
        "⬇️ Download Excel",
        data=output,
        file_name=f"Laporan_Sungai_{bulan}_{tahun}.xlsx"
    )
