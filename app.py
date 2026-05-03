import streamlit as st
from openpyxl import Workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.styles import Alignment, Font, Border, Side
from openpyxl.worksheet.pagebreak import Break
from datetime import datetime
from io import BytesIO
from PIL import Image as PILImage
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
# PREVIEW (REVIEW)
# ======================
if st.session_state.data:
    st.subheader("🔍 Preview Data")
    for i, d in enumerate(st.session_state.data, 1):
        col1, col2 = st.columns([1,2])
        with col1:
            st.image(d["foto"], use_container_width=True)
        with col2:
            st.write(f"**No:** {i}")
            st.write(f"**Uraian:** {d['uraian']}")
            st.write(f"**Lokasi:** {d['lokasi']}")
            st.write(f"**Koordinat:** {d['koordinat']}")
            st.write("**Keterangan:** Kondisi tebing masih dalam keadaan baik")

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
# EXPORT EXCEL
# ======================
if st.button("📊 EXPORT EXCEL FINAL"):

    data = st.session_state.data
    if not data:
        st.warning("Belum ada data!")
        st.stop()

    images = [compress_image(d["foto"].getvalue()) for d in data]

    wb = Workbook()
    ws = wb.active

    # A4 POTRET
    ws.page_setup.orientation = ws.ORIENTATION_PORTRAIT
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.fitToWidth = 1

    center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    bold = Font(bold=True)

    thin = Side(style="thin")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    # KOLOM
    ws.column_dimensions["A"].width = 5
    ws.column_dimensions["B"].width = 19
    ws.column_dimensions["C"].width = 19
    ws.column_dimensions["D"].width = 17
    ws.column_dimensions["E"].width = 38
    ws.column_dimensions["F"].width = 22

    match = re.search(r"sungai\s+(.+)", data[0]["uraian"].lower())
    nama_sungai = match.group(1).upper() if match else data[0]["uraian"].upper()

    def header(r):
        for i, text in enumerate([
            f"LAPORAN SUNGAI {nama_sungai}",
            "KOTA/KAB. CIAMIS",
            "OPERASI DAN PEMELIHARAAN SDA III",
            f"BULAN {bulan.upper()} TAHUN {tahun}"
        ]):
            ws.merge_cells(start_row=r+i, start_column=1, end_row=r+i, end_column=6)
            ws.cell(r+i,1,text).alignment = center
            ws.cell(r+i,1).font = Font(bold=True, size=16)

        headers = ["NO","URAIAN","LOKASI","KOORDINAT","FOTO","KETERANGAN"]
        for i,h in enumerate(headers,1):
            c = ws.cell(r+5,i,h)
            c.font = bold
            c.alignment = center
            c.border = border

        return r + 6

    row = header(1)
    nomor = 1
    limit = 20
    count = 0

    for d in data:

        if count >= limit:
            ws.row_breaks.append(Break(id=row))
            row = header(row)
            count = 0

        ws.row_dimensions[row].height = 150

        ws.cell(row,1,nomor).alignment = center
        ws.cell(row,2,d["uraian"]).alignment = center
        ws.cell(row,3,d["lokasi"]).alignment = center
        ws.cell(row,4,d["koordinat"]).alignment = center
        ws.cell(row,6,"Kondisi tebing masih dalam keadaan baik").alignment = center

        for col in [1,2,3,4,6]:
            ws.cell(row,col).border = border

        img = XLImage(BytesIO(images[nomor-1]))
        img.width = 300
        img.height = 170
        img.anchor = f"E{row}"
        ws.add_image(img)

        nomor += 1
        row += 1
        count += 1

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
    elements.append(Paragraph("<b>KOTA/KAB. CIAMIS</b>", styles["Title"]))
    elements.append(Paragraph("<b>OPERASI DAN PEMELIHARAAN SDA III</b>", styles["Title"]))
    elements.append(Paragraph(f"<b>BULAN {bulan.upper()} TAHUN {tahun}</b>", styles["Title"]))
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
        ("ALIGN",(0,0),(-1,-1),"CENTER"),
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
