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

# PDF
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Image, Spacer, Paragraph
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.units import cm

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
# EXPORT EXCEL
# ======================
if st.button("📊 EXPORT EXCEL FINAL"):

    data = st.session_state.data
    if not data:
        st.warning("Belum ada data!")
        st.stop()

    with ThreadPoolExecutor(max_workers=4) as executor:
        images = list(executor.map(lambda d: compress_image(d["foto"].getvalue()), data))

    wb = Workbook()
    ws = wb.active
    ws.title = "Laporan"

    center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    left = Alignment(horizontal="left", vertical="center", wrap_text=True)
    bold = Font(bold=True)

    thin = Side(style="thin")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    ws.column_dimensions["A"].width = 5
    ws.column_dimensions["B"].width = 25
    ws.column_dimensions["C"].width = 25
    ws.column_dimensions["D"].width = 22
    ws.column_dimensions["E"].width = 42
    ws.column_dimensions["F"].width = 30

    match = re.search(r"sungai\s+(.+)", data[0]["uraian"].lower())
    nama_sungai = match.group(1).upper() if match else data[0]["uraian"].upper()

    row_global = 1
    nomor = 1

    def merge(r):
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=6)

    # HEADER
    merge(1)
    ws.cell(1,1,f"LAPORAN SUNGAI {nama_sungai}").alignment = center
    ws.cell(1,1).font = Font(bold=True, size=14)

    merge(2)
    ws.cell(2,1,"KOTA/KAB. CIAMIS").alignment = center

    merge(3)
    ws.cell(3,1,"OPERASI DAN PEMELIHARAAN SDA III").alignment = center

    merge(4)
    ws.cell(4,1,f"BULAN {bulan.upper()} TAHUN {tahun}").alignment = center

    headers = ["NO","URAIAN","LOKASI","KOORDINAT","FOTO","KETERANGAN"]

    for i,h in enumerate(headers,1):
        c = ws.cell(6,i,h)
        c.font = bold
        c.alignment = center
        c.border = border

    row = 7

    for d in data:
        ws.row_dimensions[row].height = 165

        ws.cell(row=row, column=1, value=nomor)
        ws.cell(row=row, column=2, value=d["uraian"])
        ws.cell(row=row, column=3, value=d["lokasi"])
        ws.cell(row=row, column=4, value=d["koordinat"])
        ws.cell(row=row, column=6, value="Kondisi tebing masih dalam keadaan baik")

        ws.cell(row=row, column=1).alignment = center
        ws.cell(row=row, column=4).alignment = center

        for col in [1,2,3,4,6]:
            ws.cell(row=row, column=col).border = border

        img = XLImage(BytesIO(images[nomor-1]))
        img.width = 320
        img.height = 180
        ws.add_image(img, f"E{row}")

        nomor += 1
        row += 1

    # TTD
    ttd = row + 1
    now = datetime.now()

    ws.merge_cells(start_row=ttd, start_column=5, end_row=ttd, end_column=6)
    ws.cell(ttd,5,f"Ciamis, {now.day} {bulan_list[now.month-1]} {now.year}").alignment = center

    ws.merge_cells(start_row=ttd+1, start_column=5, end_row=ttd+1, end_column=6)
    ws.cell(ttd+1,5,"Juru Sungai OPSDA 03").alignment = center

    ws.merge_cells(start_row=ttd+4, start_column=5, end_row=ttd+4, end_column=6)
    ws.cell(ttd+4,5,"Fariz Rionaldi").alignment = center

    output = BytesIO()
    wb.save(output)
    output.seek(0)

    st.download_button(
        "⬇️ Download Excel",
        data=output,
        file_name=f"Laporan_Sungai_{bulan}_{tahun}.xlsx"
    )

# ======================
# EXPORT PDF
# ======================
if st.button("📄 EXPORT PDF A4"):

    data = st.session_state.data
    if not data:
        st.warning("Belum ada data!")
        st.stop()

    buffer = BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4)

    styles = getSampleStyleSheet()
    elements = []

    match = re.search(r"sungai\s+(.+)", data[0]["uraian"].lower())
    nama_sungai = match.group(1).upper() if match else data[0]["uraian"].upper()

    elements.append(Paragraph(f"<b>LAPORAN SUNGAI {nama_sungai}</b>", styles["Title"]))
    elements.append(Paragraph("<b>KOTA/KAB. CIAMIS</b>", styles["Normal"]))
    elements.append(Paragraph("<b>OPERASI DAN PEMELIHARAAN SDA III</b>", styles["Normal"]))
    elements.append(Paragraph(f"<b>BULAN {bulan.upper()} TAHUN {tahun}</b>", styles["Normal"]))
    elements.append(Spacer(1, 12))

    table_data = [["NO","URAIAN","LOKASI","KOORDINAT","FOTO","KETERANGAN"]]

    for i, d in enumerate(data, 1):
        img = PILImage.open(d["foto"])
        img.thumbnail((400, 250))

        img_buffer = BytesIO()
        img.save(img_buffer, format="JPEG")
        img_buffer.seek(0)

        rl_img = Image(img_buffer, width=6*cm, height=4*cm)

        table_data.append([
            str(i),
            d["uraian"],
            d["lokasi"],
            d["koordinat"],
            rl_img,
            "Kondisi tebing masih dalam keadaan baik"
        ])

    table = Table(table_data, colWidths=[1*cm,4*cm,4*cm,4*cm,6*cm,5*cm])

    table.setStyle(TableStyle([
        ("GRID",(0,0),(-1,-1),1,colors.black),
        ("ALIGN",(0,0),(-1,0),"CENTER"),
        ("VALIGN",(0,0),(-1,-1),"MIDDLE"),
    ]))

    elements.append(table)
    elements.append(Spacer(1, 25))

    now = datetime.now()

    elements.append(Paragraph(
        f"Ciamis, {now.day} {bulan_list[now.month-1]} {now.year}",
        styles["Normal"]
    ))
    elements.append(Paragraph("Juru Sungai OPSDA 03", styles["Normal"]))
    elements.append(Spacer(1, 40))
    elements.append(Paragraph("<b>Fariz Rionaldi</b>", styles["Normal"]))

    doc.build(elements)
    buffer.seek(0)

    st.download_button(
        "⬇️ Download PDF",
        data=buffer,
        file_name=f"Laporan_Sungai_{bulan}_{tahun}.pdf",
        mime="application/pdf"
    )
