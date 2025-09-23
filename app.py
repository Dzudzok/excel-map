import io
import re
import time
import pandas as pd
import streamlit as st
import folium
from folium.plugins import MarkerCluster
from geopy.geocoders import Nominatim
from geopy.extra.rate_limiter import RateLimiter
from streamlit_folium import st_folium


st.set_page_config(page_title="Mapa z Excela", layout="wide")
st.title("📍 Mapa klientów z pliku Excel/CSV (online, free)")

st.markdown(
"""
Wgraj plik **.xlsx / .csv** lub podaj **link online** (np. Google Sheets → *File > Share > Publish to the web* → CSV).  
Wymagane kolumny: **Adres**, **Miasto**, **PSC**.  
Opcjonalne: **Nazwa odbiorcy**, **Obrót w czk**, **email**.  
Jeśli masz gotowe współrzędne (**lat**, **lon**), geokodowanie będzie pominięte.
"""
)

# ---------- Helpers ----------
REQ_ADDR_COLS = ["Adres", "Miasto", "PSC"]

def norm_col(s: pd.Series) -> pd.Series:
    return (
        s.fillna("")
         .astype(str)
         .str.replace(r"^\s*0\s*$", "", regex=True)
         .str.replace(r"\s+", " ", regex=True)
         .str.strip()
    )

def build_full_address(df: pd.DataFrame) -> pd.Series:
    a = norm_col(df["Adres"])
    m = norm_col(df["Miasto"])
    p = norm_col(df["PSC"]).str.replace(" ", "")
    return (a + ", " + m + " " + p).str.strip(", ").str.strip()

def load_from_url(url: str) -> pd.DataFrame:
    # Google Sheets "share" URL → CSV export
    gs = re.match(r"https://docs\.google\.com/spreadsheets/d/([a-zA-Z0-9-_]+)", url)
    if gs:
        sheet_csv = f"https://docs.google.com/spreadsheets/d/{gs.group(1)}/export?format=csv"
        return pd.read_csv(sheet_csv)

    # inne: spróbuj CSV, a jeśli nie, to xlsx
    try:
        if url.lower().endswith(".csv"):
            return pd.read_csv(url)
        else:
            # pandas umie czytać xlsx z URL, ale czasem lepiej pobrać bajty
            import requests
            r = requests.get(url, timeout=30)
            r.raise_for_status()
            return pd.read_excel(io.BytesIO(r.content))
    except Exception as e:
        st.error(f"Nie udało się wczytać danych z URL: {e}")
        st.stop()

@st.cache_data(show_spinner=False)
def geocode_address(addr: str):
    # pojedynczy adres → (lat, lon) lub None
    geolocator = Nominatim(user_agent="excel_map_online")
    geocode = RateLimiter(geolocator.geocode, min_delay_seconds=1, swallow_exceptions=True)
    loc = geocode(addr)
    if loc:
        return (loc.latitude, loc.longitude)
    return None

def to_float(x):
    try:
        return float(x)
    except:
        return None

# ---------- UI: input ----------
left, right = st.columns([1,1])

with left:
    uploaded = st.file_uploader("Wgraj Excel/CSV", type=["xlsx", "csv"])

with right:
    url = st.text_input("…albo wklej link do pliku online (Google Sheets/CSV/OneDrive)")

if not uploaded and not url:
    st.info("Wgraj plik lub podaj link, aby kontynuować.")
    st.stop()

# ---------- Load data ----------
if uploaded:
    if uploaded.name.lower().endswith(".csv"):
        df = pd.read_csv(uploaded)
    else:
        df = pd.read_excel(uploaded)
else:
    df = load_from_url(url)

if df.empty:
    st.error("Plik nie zawiera danych.")
    st.stop()

st.subheader("Podgląd danych")
st.dataframe(df.head(50), use_container_width=True)

# ---------- Column checks / normalize ----------
missing = [c for c in REQ_ADDR_COLS if c not in df.columns]
if missing and not ({"lat","lon"} <= set(df.columns)):
    st.error(f"Brakuje wymaganych kolumn adresowych: {', '.join(missing)} "
             f"(chyba że podasz kolumny lat i lon — wtedy adres nie jest potrzebny).")
    st.stop()

# sprzątanie
if "Adres" in df.columns:  df["Adres"]  = norm_col(df["Adres"])
if "Miasto" in df.columns: df["Miasto"] = norm_col(df["Miasto"])
if "PSC" in df.columns:    df["PSC"]    = norm_col(df["PSC"]).str.replace(" ", "")

# FullAddress jeśli nie ma lat/lon
if not ({"lat","lon"} <= set(df.columns)):
    df["FullAddress"] = build_full_address(df)
    df = df[df["FullAddress"].str.len() > 0].copy()

# ---------- Geocoding / coordinates ----------
with st.spinner("Geokoduję adresy (OpenStreetMap/Nominatim)…"):
    if {"lat","lon"} <= set(df.columns):
        df["lat"] = df["lat"].apply(to_float)
        df["lon"] = df["lon"].apply(to_float)
        geo_df = df.dropna(subset=["lat","lon"]).copy()
        skipped = len(df) - len(geo_df)
    else:
        coords = []
        for addr in df["FullAddress"]:
            loc = geocode_address(addr)
            coords.append(loc)
        df["lat"] = [c[0] if c else None for c in coords]
        df["lon"] = [c[1] if c else None for c in coords]
        geo_df = df.dropna(subset=["lat","lon"]).copy()
        skipped = df["lat"].isna().sum()

if geo_df.empty:
    st.error("Nie udało się uzyskać żadnych współrzędnych. Sprawdź adresy lub dodaj kolumny lat/lon.")
    st.stop()

st.success(f"Gotowe! Pinezek: {len(geo_df)}.  Pominętych wierszy: {skipped}.")

# ---------- Map ----------
m = folium.Map(location=[geo_df["lat"].mean(), geo_df["lon"].mean()], zoom_start=8)
cluster = MarkerCluster().add_to(m)

def val(col, row, default=""):
    return row[col] if col in geo_df.columns and pd.notna(row[col]) else default

for _, r in geo_df.iterrows():
    popup_html = f"""
    <div style="font-size:14px">
      <b>{val('Nazwa odbiorcy', r)}</b><br>
      {('Obrót: {:,.2f} CZK'.format(val('Obrót w czk', r)) if pd.notna(val('Obrót w czk', r)) else '')}<br>
      {('Email: ' + val('email', r)) if val('email', r) else ''}<br>
      {('Adres: ' + val('FullAddress', r)) if 'FullAddress' in geo_df.columns else ''}
    </div>
    """
    folium.Marker(
        [r["lat"], r["lon"]],
        tooltip=val('Nazwa odbiorcy', r) or "Klient",
        popup=folium.Popup(popup_html, max_width=350)
    ).add_to(cluster)

st_folium(m, height=700)

# ---------- Export ----------
with st.expander("💾 Eksport"):
    st.download_button(
        "Pobierz CSV z geokodowanymi współrzędnymi",
        data=geo_df.to_csv(index=False).encode("utf-8"),
        file_name="geokodowane_dane.csv",
        mime="text/csv"
    )
    html = m.get_root().render()
    st.download_button(
        "Pobierz mapę jako plik HTML",
        data=html.encode("utf-8"),
        file_name="mapa.html",
        mime="text/html"
    )
