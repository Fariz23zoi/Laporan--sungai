import streamlit as st
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Image as RLImage, Paragraph, Spacer
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import cm
from reportlab.lib.enums import TA_CENTER
from PIL import Image
import io

st.title("📋 Laporan Penelusuran Sungai")

bulan = st.selectbox("Bulan", [
    "Januari","Februari","Maret","April","Mei","Juni",
    "Juli","Agustus","September","Oktober","November","Desember"
])
tahun = st.number_input("Tahun", value=2026)

if "data" not in st.session_state:
    st.session_state.data = []

st.subheader("Input Data")

sungai = st.text_input("Nama Sungai", "SUNGAI CIMANTAJA")
lokasi = st.text_area("Lokasi")
koordinat = st.text_input("Koordinat")
keterangan = st.text_area("Keterangan", "Kondisi tebing masih dalam keadaan baik")
foto = st.file_uploader("Upload Foto", type=["jpg","png","jpeg"])

if st.button("➕ Tambah"):
    if foto:
        st.session_state.data.append({
            "sungai": sungai,
            "lokasi": lokasi,
            "koordinat": koordinat,
            "keterangan": keterangan,
            "foto": foto
        })
        st.success("Data ditambahkan!")

st.write(f"Jumlah data: {len(st.session_state.data)}")

# EXPORT PDF
if st.button("📄 Export PDF Resmi"):
    filename = f"Laporan_Sungai_{bulan}_{tahun}.pdf"

    doc = SimpleDocTemplate(filename, pagesize=A4)
    elements = []
    styles = getSampleStyleSheet()

    center_title = ParagraphStyle(name='t', alignment=TA_CENTER, fontSize=14)
    center = ParagraphStyle(name='c', alignment=TA_CENTER, fontSize=11)

    elements.append(Paragraph("LAPORAN PENELUSURAN SUNGAI CIPETUNGAN", center_title))
    elements.append(Paragraph("KOTA/KABUPATEN CIAMIS", center))
    elements.append(Paragraph("OPERASI DAN PEMELIHARAAN SDA III", center))
    elements.append(Paragraph(f"BULAN {bulan.upper()} TAHUN {tahun}", center))
    elements.append(Spacer(1, 12))

    data = [["NO","SUNGAI","LOKASI","KOORDINAT","FOTO","KETERANGAN"]]

    for i, d in enumerate(st.session_state.data, start=1):
        img = Image.open(d["foto"])
        img_io = io.BytesIO()
        img.save(img_io, format='JPEG')
        img_io.seek(0)

        rl_img = RLImage(img_io, width=5*cm, height=3*cm)

        data.append([
            str(i),
            d["sungai"],
            d["lokasi"],
            d["koordinat"],
            rl_img,
            d["keterangan"]
        ])

    table = Table(data, repeatRows=1)

    table.setStyle(TableStyle([
        ('GRID',(0,0),(-1,-1),1,colors.black),
        ('BACKGROUND',(0,0),(-1,0),colors.lightgrey),
        ('ALIGN',(0,0),(-1,-1),'CENTER'),
    ]))

    elements.append(table)
    doc.build(elements)

    with open(filename, "rb") as f:
        st.download_button("⬇️ Download PDF", f, file_name=filename)
