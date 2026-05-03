import streamlit as st
from supabase import create_client
import requests
from datetime import datetime
from io import BytesIO
from PIL import Image as PILImage, ImageDraw, ImageFont
from openpyxl import Workbook
import folium
from streamlit_folium import st_folium
from geopy.distance import geodesic
import textwrap
import time

# ======================
# KONFIG (HARDCODE DULU BIAR AMAN)
# ======================
SUPABASE_URL = "https://gkadevvilhbcwyyojfm.supabase.co"
SUPABASE_KEY = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9..."  # PASTE FULL KEY DI SINI

supabase = create_client(SUPABASE_URL, SUPABASE_KEY)

st.set_page_config(layout="wide")
st.title("📋 Monitoring Sungai")

# ======================
# KETERANGAN LOCK
# ======================
KETERANGAN_DEFAULT = "Kondisi tebing masih dalam keadaan baik"
st.info(f"Keterangan otomatis: {KETERANGAN_DEFAULT}")

# ======================
# INPUT
# ======================
st.subheader("Input Data")

uraian = st.text_input("Uraian")
lokasi = st.text_area("Lokasi")

# Koordinat manual (AMAN)
lat = st.text_input("Latitude (opsional)")
lon = st.text_input("Longitude (opsional)")

foto_list = st.file_uploader("Multi Foto", accept_multiple_files=True)

# ======================
# CUACA
# ======================
def get_weather():
    try:
        return "-"
    except:
        return "-"

cuaca_auto = get_weather()

# ======================
# OVERLAY FOTO
# ======================
def overlay_ts(file_bytes, lokasi, koordinat, uraian, ket, cuaca, waktu):
    img = PILImage.open(BytesIO(file_bytes)).convert("RGB")
    w, h = img.size

    overlay_h = int(h * 0.3)
    overlay = PILImage.new("RGBA", (w, overlay_h), (0, 0, 0, 160))
    img.paste(overlay, (0, h - overlay_h), overlay)

    draw = ImageDraw.Draw(img)

    try:
        font = ImageFont.truetype("arial.ttf", 22)
    except:
        font = ImageFont.load_default()

    y = h - overlay_h + 10
    texts = [waktu, cuaca, lokasi, koordinat, uraian, ket]

    for t in texts:
        for line in textwrap.wrap(str(t), 40):
            draw.text((10, y), line, fill="white", font=font)
            y += 25

    buf = BytesIO()
    img.save(buf, format="JPEG", quality=65)
    return buf.getvalue()

# ======================
# SIMPAN
# ======================
def simpan(data, fotos):
    try:
        res = supabase.table("laporan").insert(data).execute()
        laporan_id = res.data[0]["id"]

        for f in fotos:
            raw = f.read()

            img_bytes = overlay_ts(
                raw,
                data["lokasi"],
                data["koordinat"],
                data["uraian"],
                KETERANGAN_DEFAULT,
                data["cuaca"],
                data["waktu"]
            )

            filename = f"{laporan_id}_{time.time()}.jpg"

            supabase.storage.from_("foto").upload(filename, img_bytes)
            url = supabase.storage.from_("foto").get_public_url(filename)

            supabase.table("foto").insert({
                "laporan_id": laporan_id,
                "url_foto": url
            }).execute()

    except Exception as e:
        st.error(f"Gagal simpan: {e}")

# ======================
# BUTTON SIMPAN
# ======================
if st.button("Simpan"):
    koordinat = f"{lat},{lon}" if lat and lon else "-"

    data = {
        "uraian": uraian,
        "lokasi": lokasi,
        "koordinat": koordinat,
        "keterangan": KETERANGAN_DEFAULT,
        "cuaca": cuaca_auto,
        "waktu": datetime.now().strftime("%d %B %Y %H:%M")
    }

    simpan(data, foto_list)
    st.success("Data masuk!")

# ======================
# AMBIL DATA
# ======================
data = []
try:
    res = supabase.table("laporan").select("*").execute()
    if res.data:
        data = res.data
except:
    st.warning("Database belum terbaca")

# ======================
# DASHBOARD
# ======================
st.subheader("📊 Dashboard")
st.metric("Total Data", len(data))

# ======================
# MAP
# ======================
coords = []
for d in data:
    try:
        la, lo = map(float, d["koordinat"].split(","))
        coords.append((la, lo))
    except:
        pass

if coords:
    m = folium.Map(location=coords[0], zoom_start=14)

    for (la, lo) in coords:
        folium.Marker([la, lo]).add_to(m)

    folium.PolyLine(coords).add_to(m)
    st_folium(m, width=700, height=500)

    total = 0
    for i in range(len(coords) - 1):
        total += geodesic(coords[i], coords[i + 1]).meters

    st.metric("Total Jarak (m)", int(total))

# ======================
# EXPORT EXCEL
# ======================
if st.button("Export Excel"):
    wb = Workbook()
    ws = wb.active

    ws.append(["NO", "URAIAN", "LOKASI", "KETERANGAN"])

    for i, d in enumerate(data, start=1):
        ws.append([
            i,
            d["uraian"],
            d["lokasi"],
            KETERANGAN_DEFAULT
        ])

    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)

    st.download_button("Download Excel", buf, "laporan.xlsx")
