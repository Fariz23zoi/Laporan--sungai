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

# ======================
# KONFIG SUPABASE (STREAMLIT CLOUD)
# ======================
SUPABASE_URL = st.secrets["SUPABASE_URL"]
SUPABASE_KEY = st.secrets["SUPABASE_KEY"]

supabase = create_client(SUPABASE_URL, SUPABASE_KEY)

st.set_page_config(layout="wide")
st.title("📋 Sistem Monitoring Sungai")

# ======================
# LOGIN
# ======================
if "user" not in st.session_state:
    st.subheader("🔐 Login")

    email = st.text_input("Email")
    password = st.text_input("Password", type="password")

    if st.button("Login"):
        try:
            res = supabase.auth.sign_in_with_password({
                "email": email,
                "password": password
            })
            st.session_state.user = res.user
            st.success("Login berhasil")
            st.rerun()
        except:
            st.error("Login gagal")

    st.stop()

# ======================
# GPS
# ======================
loc = streamlit_geolocation()

lat, lon = None, None
if loc and loc["latitude"]:
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
        code = w.get("weathercode")

        kondisi_map = {
            0:"Cerah",1:"Cerah Berawan",2:"Berawan",3:"Mendung",
            61:"Hujan Ringan",63:"Hujan",65:"Hujan Lebat"
        }
        kondisi = kondisi_map.get(code,"-")

        return f"{kondisi} {suhu}°C"
    except:
        return "-"

cuaca_auto = get_weather(lat, lon) if lat else "-"

# ======================
# OVERLAY FOTO
# ======================
def overlay_ts(file_bytes, lokasi, koordinat, uraian, ket, cuaca, waktu):
    img = PILImage.open(BytesIO(file_bytes)).convert("RGB")
    w, h = img.size

    overlay_h = int(h*0.32)
    overlay = PILImage.new("RGBA",(w,overlay_h),(0,0,0,170))
    img.paste(overlay,(0,h-overlay_h),overlay)

    draw = ImageDraw.Draw(img)

    try:
        f_big = ImageFont.truetype("arial.ttf",32)
        f_small = ImageFont.truetype("arial.ttf",24)
    except:
        f_big = f_small = ImageFont.load_default()

    draw.rectangle([(10,h-overlay_h),(15,h)],fill=(255,200,0))

    x,y = 30, h-overlay_h+20
    draw.text((x,y), waktu, fill="white", font=f_big); y+=40
    draw.text((x,y), cuaca, fill="white", font=f_small); y+=30

    def wrap(txt,y):
        for line in textwrap.wrap(txt,40):
            draw.text((x,y), line, fill="white", font=f_small)
            y+=28
        return y

    y=wrap(lokasi,y)
    draw.text((x,y), koordinat, fill="white", font=f_small); y+=30
    y=wrap(uraian,y)
    y=wrap(ket,y)

    buf = BytesIO()
    img.save(buf, format="JPEG", quality=65)
    return buf.getvalue()

# ======================
# INPUT
# ======================
st.subheader("Input Data")

uraian = st.text_input("Uraian")
lokasi = st.text_area("Lokasi")
keterangan = st.text_area("Keterangan")

mode_1klik = st.toggle("⚡ Mode 1 Klik Survey")

foto_list = st.file_uploader("Multi Foto", accept_multiple_files=True)

# ======================
# SIMPAN DATA
# ======================
def simpan(data, fotos):
    try:
        res = supabase.table("laporan").insert({
            "user_id": st.session_state.user.id,
            "uraian": data["uraian"],
            "lokasi": data["lokasi"],
            "koordinat": data["koordinat"],
            "keterangan": data["keterangan"],
            "cuaca": data["cuaca"],
            "waktu": data["waktu"]
        }).execute()

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
                data["keterangan"],
                data["cuaca"],
                data["waktu"]
            )

            filename = f"{laporan_id}_{datetime.now().timestamp()}.jpg"

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
            "keterangan": keterangan or "-",
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
            "keterangan": keterangan,
            "cuaca": cuaca_auto,
            "waktu": datetime.now().strftime("%d %B %Y %H:%M")
        }

        simpan(data, foto_list)
        st.success("Data masuk!")
    else:
        st.warning("GPS belum aktif")

# ======================
# AMBIL DATA
# ======================
data = supabase.table("laporan").select("*").execute().data or []

# ======================
# DASHBOARD
# ======================
st.subheader("📊 Dashboard")
st.metric("Total Data", len(data))

# ======================
# MAP
# ======================
st.subheader("🗺️ Map Tracking")

coords=[]
for d in data:
    try:
        lat,lon=map(float,d["koordinat"].split(","))
        coords.append((lat,lon))
    except:
        pass

if coords:
    m = folium.Map(location=coords[0], zoom_start=14)

    for i,(lat,lon) in enumerate(coords):
        folium.Marker([lat,lon]).add_to(m)

    folium.PolyLine(coords).add_to(m)
    st_folium(m, width=700, height=500)

    total=0
    for i in range(len(coords)-1):
        total+=geodesic(coords[i],coords[i+1]).meters

    st.metric("Total Jarak (m)", int(total))

# ======================
# EXPORT EXCEL
# ======================
if st.button("Export Excel"):

    wb=Workbook()
    ws=wb.active

    ws.append(["Uraian","Lokasi","Koordinat","Cuaca","Waktu"])

    for d in data:
        ws.append([
            d["uraian"],
            d["lokasi"],
            d["koordinat"],
            d["cuaca"],
            d["waktu"]
        ])

    buf=BytesIO()
    wb.save(buf)
    buf.seek(0)

    st.download_button("Download Excel", buf, file_name="laporan.xlsx")
