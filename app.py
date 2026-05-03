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
st.title("📋 Laporan Penelusuran Sungai FINAL")

bulan_list = [
    "Januari","Februari","Maret","April","Mei","Juni",
    "Juli","Agustus","September","Oktober","November","Desember"
]

bulan = st.selectbox("Bulan", bulan_list)
tahun = st.number_input("Tahun", value=2026)

# ======================
# SESSION
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
        if not uraian or not lokasi or not foto:
            st.warning("Lengkapi semua data!")
        else:
            st.session_state.data.append({
                "uraian": uraian,
                "lokasi": lokasi,
                "koordinat": koordinat,
                "foto": foto
            })
            st.success("Data ditambahkan!")

with col2:
    if st.button("🗑 Reset"):
        st.session_state.data = []

st.write(f"Total data: {len(st.session_state.data)}")

# ======================
# COMPRESS IMAGE
# ======================
def compress_image(file_bytes):
    img = PILImage.open(BytesIO(file_bytes))

    if img.mode in ("RGBA", "P"):
        img = img.convert("RGB")

    img.thumbnail((1280, 720))

    buffer = BytesIO()
    img.save(buffer, format="JPEG", quality=70, optimize=True)
    return buffer.getvalue()

# ======================
# EXPORT FINAL
# ======================
if st.button("📊 EXPORT EXCEL FINAL 🚀"):

    data = st.session_state.data
    if not data:
        st.warning("Belum ada data!")
        st.stop()

    progress = st.progress(0)

    # compress paralel
    with ThreadPoolExecutor(max_workers=4) as executor:
        images = list(executor.map(lambda d: compress_image(d["foto"].getvalue()), data))

    wb = Workbook()
    ws = wb.active
    ws.title = "Laporan"

    # STYLE
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    left = Alignment(horizontal="left", vertical="center", wrap_text=True)
    bold = Font(bold=True)

    thin = Side(style="thin")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    # COLUMN WIDTH
    ws.column_dimensions["A"].width = 5
    ws.column_dimensions["B"].width = 25
    ws.column_dimensions["C"].width = 25
    ws.column_dimensions["D"].width = 22
    ws.column_dimensions["E"].width = 42
    ws.column_dimensions["F"].width = 30

    ws.page_setup.fitToWidth = 1

    # ambil nama sungai
    match = re.search(r"sungai\s+(.+)", data[0]["uraian"].lower())
    nama_sungai = match.group(1).upper() if match else data[0]["uraian"].upper()

    per_page = 5
    pages = math.ceil(len(data) / per_page)

    row_global = 1
    nomor_global = 1

    for p in range(pages):

        if p > 0:
            ws.row_breaks.append(Break(id=row_global))

        def merge(r):
            ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=6)

        # ================= HEADER =================
        merge(row_global)
        ws.cell(row_global,1,f"LAPORAN SUNGAI {nama_sungai}").alignment = center
        ws.cell(row_global,1).font = Font(bold=True, size=14)

        merge(row_global+1)
        ws.cell(row_global+1,1,"KOTA/KAB. CIAMIS").alignment = center
        ws.cell(row_global+1,1).font = bold

        merge(row_global+2)
        ws.cell(row_global+2,1,"OPERASI DAN PEMELIHARAAN SDA III").alignment = center
        ws.cell(row_global+2,1).font = bold

        merge(row_global+3)
        ws.cell(row_global+3,1,f"BULAN {bulan.upper()} TAHUN {tahun}").alignment = center
        ws.cell(row_global+3,1).font = bold

        # ================= TABLE HEADER (ROW 7) =================
        headers = ["NO","URAIAN","LOKASI","KOORDINAT","FOTO","KETERANGAN"]

        for i,h in enumerate(headers,1):
            c = ws.cell(row_global+6,i,h)
            c.font = bold
            c.alignment = center
            c.border = border

        start = p * per_page
        chunk = data[start:start+per_page]

        row = row_global + 7

        for d in chunk:

            ws.row_dimensions[row].height = 165

            ws.cell(row=row, column=1, value=nomor_global)
            ws.cell(row=row, column=2, value=d["uraian"])
            ws.cell(row=row, column=3, value=d["lokasi"])
            ws.cell(row=row, column=4, value=d["koordinat"])
            ws.cell(row=row, column=6, value="Kondisi tebing masih dalam keadaan baik")

            # alignment
            ws.cell(row=row, column=1).alignment = center
            ws.cell(row=row, column=4).alignment = center

            for col in [1,2,3,4,6]:
                ws.cell(row=row, column=col).border = border

            # FOTO
            img = XLImage(BytesIO(images[nomor_global-1]))
            img.width = 320
            img.height = 180
            ws.add_image(img, f"E{row}")

            nomor_global += 1
            row += 1

        # ================= TTD =================
        ttd = row + 1
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

    st.success("Export berhasil 🚀")

    st.download_button(
        "⬇️ Download Excel",
        data=output,
        file_name=f"Laporan_Sungai_{bulan}_{tahun}.xlsx"
    )
