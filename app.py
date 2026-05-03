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

    # compress parallel
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
    ws
