import streamlit as st
from supabase import create_client
from streamlit_geolocation import streamlit_geolocation
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
# KONFIG
# ======================
SUPABASE_URL = st.secrets.get("SUPABASE_URL", "")
SUPABASE_KEY = st.secrets.get("SUPABASE_KEY", "")

supabase = None
if SUPABASE_URL and SUPABASE_KEY:
    try:
        supabase = create_client(SUPABASE_URL, SUPABASE_KEY)
    except:
        supabase = None

st.set_page_config(layout="wide")
st.title("📋 Monitoring Sungai")

# ======================
# INFO KETERANGAN (LOCK)
# ======================
KETERANGAN_DEFAULT = "Kondisi tebing masih dalam keadaan baik"
st.info(f"Keterangan otomatis: {KETERANGAN_DEFAULT}")

# ======================
# GPS
# ======================
loc = streamlit_geolocation()
lat, lon = None, None
if loc and loc.get("latitude"):
    lat = loc["latitude"]
    lon = loc["longitude"]

# ======================
# CUACA
# ======================
def get_weather(lat, lon):
    try:
        url = f"https://api.open-meteo.com/v1/forecast?latitude={lat}&longitude={lon}&current_weather=true"
        res = requests.get(url, timeout=5).json()
        w = res.get("current_weather", {})
        suhu = w.get("temperature")
        return f"{suhu}°C"
    except:
        return "-"

cuaca_auto = get_weather(lat, lon) if lat else "-"

# ======================
# OVERLAY FOTO
# ======================
def overlay_ts(file_bytes, lokasi, koordinat, uraian, ket, cuaca, waktu):
    img = PILImage.open(BytesIO(file_bytes)).convert("RGB")
    w, h = img.size

    overlay_h = int(h * 0.32)
    overlay = PILImage.new("RGBA", (w, overlay_h), (0, 0, 0, 170))
    img.paste(overlay, (0, h - overlay_h), overlay)

    draw = ImageDraw.Draw(img)

    try:
        f_big = ImageFont.truetype("arial.ttf", 32)
        f_small = ImageFont.truetype("arial.ttf", 22)
    except:
        f_big = f_small = ImageFont.load_default()

    x, y = 30, h - overlay_h + 20
    draw.text((x, y), waktu, fill="white", font=f_big); y += 40
    draw.text((x, y), cuaca, fill="white", font=f_small); y += 30

    for txt in [lokasi, koordinat, uraian, ket]:
        for line in textwrap.wrap(txt or "-", 40):
            draw.text((x, y), line, fill="white", font=f_small)
            y += 28

    buf = BytesIO()
    img.save(buf, format="JPEG", quality=65)
    return buf.getvalue()

# ======================
# INPUT
# ======================
st.subheader("Input Data")

uraian = st.text_input("Uraian")
lokasi = st.text_area("Lokasi")

mode_1klik = st.toggle("⚡ Mode 1 Klik Survey")
foto_list = st.file_uploader("Multi Foto", accept_multiple_files=True)

# ======================
# SIMPAN
# ======================
def simpan(data, fotos):
    if not supabase:
        st.error("❌ Supabase belum terhubung")
        return

    try:
        res = supabase.table("laporan").insert(data).execute()
        laporan_id = res.data[0]["id"]

        for f in fotos:
            if f is None:
                continue

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
# MODE 1 KLIK
# ======================
if mode_1klik:
    cam = st.camera_input("Ambil Foto")

    if cam and lat and lon:
        data = {
            "uraian": uraian or "Penelusuran",
            "lokasi": lokasi or "GPS",
            "koordinat": f"{lat},{lon}",
            "keterangan": KETERANGAN_DEFAULT,
            "cuaca": cuaca_auto,
            "waktu": datetime.now().strftime("%d %B %Y %H:%M")
        }

        simpan(data, [cam])
        st.success("Tersimpan!")
        st.rerun()

# ======================
# SIMPAN MANUAL
# ======================
if st.button("Simpan"):
    if lat and lon:
        data = {
            "uraian": uraian,
            "lokasi": lokasi,
            "koordinat": f"{lat},{lon}",
            "keterangan": KETERANGAN_DEFAULT,
            "cuaca": cuaca_auto,
            "waktu": datetime.now().strftime("%d %B %Y %H:%M")
        }

        simpan(data, foto_list)
        st.success("Data masuk!")
    else:
        st.warning("GPS belum aktif")

# ======================
# AMBIL DATA (ANTI ERROR)
# ======================
data = []
if supabase:
    for _ in range(2):
        try:
            res = supabase.table("laporan").select("*").execute()
            if res and res.data:
                data = res.data
            break
        except:
            time.sleep(1)

# ======================
# DASHBOARD
# ======================
st.subheader("📊 Dashboard")
st.metric("Total Data", len(data))

# ======================
# MAP
# ======================
st.subheader("🗺️ Map Tracking")

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
st.subheader("📊 Export")

if st.button("Export Excel A4"):
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

    st.download_button(
        "⬇️ Download Excel",
        buf,
        file_name="laporan_sungai.xlsx"
    )
