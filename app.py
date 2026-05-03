import streamlit as st
from openpyxl import Workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.styles import Alignment, Font, Border, Side
from openpyxl.worksheet.pagebreak import Break
from datetime import datetime
from io import BytesIO
from PIL import Image as PILImage
from concurrent.futures import ThreadPoolExecutor
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

# ======================
# IMAGE COMPRESS
# ======================
def compress_image(file_bytes):
    img = PILImage.open(BytesIO(file_bytes))
    if img.mode in ("RGBA", "P"):
        img = img.convert("RGB")
    img.thumbnail((1280, 720))
    buf = BytesIO()
    img.save(buf, format="JPEG", quality=70)
    return buf.getvalue()

# ======================
# EXPORT EXCEL FINAL
# ======================
if st.button("📊 EXPORT EXCEL FINAL"):

    data = st.session_state.data
    if not data:
        st.warning("Belum ada data!")
        st.stop()

    images = [compress_image(d["foto"].getvalue()) for d in data]

    wb = Workbook()
    ws = wb.active
    ws.title = "Laporan"

    # POTRET A4
    ws.page_setup.orientation = ws.ORIENTATION_PORTRAIT
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.fitToWidth = 1

    # STYLE
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    left = Alignment(horizontal="left", vertical="center", wrap_text=True)
    bold = Font(bold=True)

    thin = Side(style="thin")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    # KOLOM (PAS POTRET)
    ws.column_dimensions["A"].width = 5
    ws.column_dimensions["B"].width = 19
    ws.column_dimensions["C"].width = 19
    ws.column_dimensions["D"].width = 17
    ws.column_dimensions["E"].width = 38
    ws.column_dimensions["F"].width = 22

    # nama sungai
    match = re.search(r"sungai\s+(.+)", data[0]["uraian"].lower())
    nama_sungai = match.group(1).upper() if match else data[0]["uraian"].upper()

    def header(start_row):
        ws.merge_cells(start_row=start_row, start_column=1, end_row=start_row, end_column=6)
        ws.cell(start_row,1,f"LAPORAN SUNGAI {nama_sungai}").alignment = center
        ws.cell(start_row,1).font = Font(bold=True, size=14)

        ws.merge_cells(start_row=start_row+1, start_column=1, end_row=start_row+1, end_column=6)
        ws.cell(start_row+1,1,"KOTA/KAB. CIAMIS").alignment = center

        ws.merge_cells(start_row=start_row+2, start_column=1, end_row=start_row+2, end_column=6)
        ws.cell(start_row+2,1,"OPERASI DAN PEMELIHARAAN SDA III").alignment = center

        ws.merge_cells(start_row=start_row+3, start_column=1, end_row=start_row+3, end_column=6)
        ws.cell(start_row+3,1,f"BULAN {bulan.upper()} TAHUN {tahun}").alignment = center

        headers = ["NO","URAIAN","LOKASI","KOORDINAT","FOTO","KETERANGAN"]
        for i,h in enumerate(headers,1):
            c = ws.cell(start_row+5,i,h)
            c.font = bold
            c.alignment = center
            c.border = border

        return start_row + 6

    row = header(1)
    nomor = 1
    row_limit = 20
    counter = 0

    for d in data:

        if counter >= row_limit:
            ws.row_breaks.append(Break(id=row))
            row = header(row)
            counter = 0

        ws.row_dimensions[row].height = 165

        ws.cell(row=row, column=1, value=nomor).alignment = center
        ws.cell(row=row, column=2, value=d["uraian"])
        ws.cell(row=row, column=3, value=d["lokasi"])
        ws.cell(row=row, column=4, value=d["koordinat"]).alignment = center
        ws.cell(row=row, column=6, value="Kondisi tebing masih dalam keadaan baik")

        for col in [1,2,3,4,6]:
            ws.cell(row=row, column=col).border = border

        img = XLImage(BytesIO(images[nomor-1]))
        img.width = 320
        img.height = 180
        ws.add_image(img, f"E{row}")

        nomor += 1
        row += 1
        counter += 1

    # TTD
    ttd = row + 1
    now = datetime.now()

    ws.merge_cells(start_row=ttd, start_column=5, end_row=ttd, end_column=6)
    ws.cell(ttd,5,f"Ciamis, {now.day} {bulan_list[now.month-1]} {now.year}").alignment = center

    ws.merge_cells(start_row=ttd+1, start_column=5, end_row=ttd+1, end_column=6)
    ws.cell(ttd+1,5,"Juru Sungai OPSDA 03").alignment = center

    ws.merge_cells(start_row=ttd+4, start_column=5, end_row=ttd+4, end_column=6)
    ws.cell(ttd+4,5,"Fariz Rionaldi").alignment = center

    out = BytesIO()
    wb.save(out)
    out.seek(0)

    st.download_button("⬇️ Download Excel", data=out, file_name="laporan.xlsx")

# ======================
# EXPORT PDF FINAL
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
    elements.append(Spacer(1, 10))

    table_data = [["NO","URAIAN","LOKASI","KOORDINAT","FOTO","KETERANGAN"]]

    for i, d in enumerate(data, 1):
        img = PILImage.open(d["foto"])
        img.thumbnail((400, 250))

        buf = BytesIO()
        img.save(buf, format="JPEG")
        buf.seek(0)

        rl_img = Image(buf, width=6*cm, height=4*cm)

        table_data.append([
            str(i),
            d["uraian"],
            d["lokasi"],
            d["koordinat"],
            rl_img,
            "Kondisi tebing masih dalam keadaan baik"
        ])

    table = Table(table_data, repeatRows=1)

    table.setStyle(TableStyle([
        ("GRID",(0,0),(-1,-1),1,colors.black),
        ("ALIGN",(0,0),(-1,0),"CENTER"),
        ("VALIGN",(0,0),(-1,-1),"MIDDLE"),
    ]))

    elements.append(table)
    elements.append(Spacer(1, 20))

    now = datetime.now()

    elements.append(Paragraph(f"Ciamis, {now.day} {bulan_list[now.month-1]} {now.year}", styles["Normal"]))
    elements.append(Paragraph("Juru Sungai OPSDA 03", styles["Normal"]))
    elements.append(Spacer(1, 40))
    elements.append(Paragraph("<b>Fariz Rionaldi</b>", styles["Normal"]))

    doc.build(elements)
    buffer.seek(0)

    st.download_button("⬇️ Download PDF", data=buffer, file_name="laporan.pdf")
