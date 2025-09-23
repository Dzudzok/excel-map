import io
import re
import pandas as pd
import streamlit as st
import folium
from folium.plugins import MarkerCluster
from streamlit_folium import st_folium

st.set_page_config(page_title="Mapa z Excela (Google Sheets)", layout="wide")
st.title("📍 Mapa klientów z Excela / Google Sheets (online, free)")

# === STAŁY PUBLICZNY CSV Z GOOGLE SHEETS ===
GOOGLE_SHEETS_CSV = "https://docs.google.com/spreadsheets/d/e/2PACX-1vSijSBg7JqZkg4T8aY56FEhox0pqw5huE7oWRmSbaB25LJj9nFyo76JLPKSXHZecd4nZEyu92jesaor/pub?gid=0&single=true&output=csv"

st.markdown(
"""
Wymagane kolumny: **Adres**, **Miasto**, **PSC**.  
Opcjonalne: **Nazwa odbiorcy**, **Obrót w czk**, **email**.  
Jeśli podasz **lat** i **lon** w danych, geokodowanie jest pomijane.
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
    p = norm_col(df["PSC"]).str.replace(" ", "")  # PSC bez spacji
    return (a + ", " + m + " " + p).str.strip(", ").str.strip()

@st.cache_data(show_spinner=False)
def load_google_csv(url: str) -> pd.DataFrame:
    import requests
    headers = {"User-Agent": "Mozilla/5.0"}
    r = requests.get(url, timeout=30, headers=headers)
    r.raise_for_status()
    # próba CSV
    try:
        return pd.read_csv(io.StringIO(r.text))
    except Exception:
        # awaryjnie spróbuj xlsx (gdyby ktoś kiedyś zmienił output)
        return pd.read_excel(io.BytesIO(r.content))

def to_float(x):
    try:
        return float(x)
    except:
        return None

# ---------- UI ----------
col1, col2 = st.columns([1,1])

with col1:
    uploaded = st.file_uploader("Wgraj własny plik (Excel/CSV)", type=["xlsx", "csv"])

with col2:
    if st.button("⬇️ Pobierz z Google (stały link)"):
        st.session_state["_use_google"] = True

use_google = st.session_state.get("_use_google", False)

# ---------- Load data ----------
df = None

if uploaded is not None and not use_google:
    if uploaded.name.lower().endswith(".csv"):
        df = pd.read_csv(uploaded)
    else:
        df = pd.read_excel(uploaded)
elif use_google:
    try:
        df = load_google_csv(GOOGLE_SHEETS_CSV)
    except Exception as e:
        st.error(f"Nie udało się pobrać danych z Google Sheets: {e}")
        st.stop()

if df is None:
    st.info("Wgraj plik lub kliknij „Pobierz z Google (stały link)”.")
    st.stop()

if df.empty:
    st.error("Plik nie zawiera danych.")
    st.stop()

st.subheader("Podgląd danych")
st.dataframe(df.head(50), use_container_width=True)

# ---------- Column checks / normalize ----------
has_coords = {"lat","lon"} <= set(df.columns)

if not has_coords:
    missing = [c for c in REQ_ADDR_COLS if c not in df.columns]
    if missing:
        st.error(f"Brakuje wymaganych kolumn adresowych: {', '.join(missing)} "
                 f"(albo dodaj lat/lon, wtedy adres nie jest potrzebny).")
        st.stop()

# sprzątanie kolumn
if "Adres" in df.columns:  df["Adres"]  = norm_col(df["Adres"])
if "Miasto" in df.columns: df["Miasto"] = norm_col(df["Miasto"])
if "PSC" in df.columns:    df["PSC"]    = norm_col(df["PSC"]).str.replace(" ", "")

# FullAddress jeśli nie ma lat/lon
if not has_coords:
    df["FullAddress"] = build_full_address(df)
    df = df[df["FullAddress"].str.len() > 0].copy()

# ---------- Geocoding / coordinates ----------
with st.spinner("Przygotowuję współrzędne… (jeśli brak lat/lon, geokodowanie może potrwać)"):
    if has_coords:
        df["lat"] = df["lat"].apply(to_float)
        df["lon"] = df["lon"].apply(to_float)
        geo_df = df.dropna(subset=["lat","lon"]).copy()
        skipped = len(df) - len(geo_df)
    else:
        # Użyj OpenCage/OSM geocodingu tylko jeśli naprawdę potrzebujesz.
        # Na darmowym hostingu Nominatim jest wolny i limitowany — lepiej raz dodać lat/lon w pliku.
        # Tu dla bezpieczeństwa NIE geokodujemy online (żeby nie blokować appki i nie wpaść w limity).
        st.warning("Brak kolumn lat/lon — mapę mogę narysować, ale potrzebuję współrzędnych. "
                   "Dodaj lat/lon do pliku (albo daj znać, to włączę wolne, darmowe geokodowanie OSM).")
        geo_df = df.copy()
        geo_df["lat"] = None
        geo_df["lon"] = None
        skipped = len(geo_df)

if geo_df.empty:
    st.error("Brak wierszy z poprawnymi danymi.")
    st.stop()

# Jeśli nie mamy żadnych współrzędnych, to narysujemy mapę na bazowym centrum
if pd.isna(geo_df["lat"]).all() or pd.isna(geo_df["lon"]).all():
    # środek na Czechy/Śląsk jako domyślna perspektywa
    m = folium.Map(location=[49.8, 18.2], zoom_start=7)
    st_folium(m, height=650)
    st.info("Dodaj kolumny 'lat' i 'lon' do danych, żeby zobaczyć pinezki. "
            "Jeśli wolisz, mogę włączyć darmowe geokodowanie OSM (wolne, ale bez klucza).")
    st.stop()

st.success(f"Pinezek z koordynatami: {geo_df[['lat','lon']].dropna().shape[0]}. Pominętych: {skipped}.")

# ---------- Map ----------
m = folium.Map(location=[geo_df['lat'].dropna().mean(), geo_df['lon'].dropna().mean()], zoom_start=8)
cluster = MarkerCluster().add_to(m)

def val(col, row, default=""):
    return row[col] if col in geo_df.columns and pd.notna(row[col]) else default

for _, r in geo_df.dropna(subset=["lat","lon"]).iterrows():
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
        "Pobierz CSV z danymi (w tym lat/lon jeśli były w źródle)",
        data=geo_df.to_csv(index=False).encode("utf-8"),
        file_name="dane_z_koordynatami.csv",
        mime="text/csv"
    )
    html = m.get_root().render()
    st.download_button(
        "Pobierz mapę jako plik HTML",
        data=html.encode("utf-8"),
        file_name="mapa.html",
        mime="text/html"
    )
