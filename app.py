import io
import time
import math
import hashlib
import pandas as pd
import streamlit as st
import folium
from folium.plugins import MarkerCluster
from streamlit_folium import st_folium
from geopy.geocoders import Nominatim
from geopy.extra.rate_limiter import RateLimiter
import requests

# -------------------- KONFIG --------------------
GOOGLE_SHEETS_CSV = "https://docs.google.com/spreadsheets/d/e/2PACX-1vSijSBg7JqZkg4T8aY56FEhox0pqw5huE7oWRmSbaB25LJj9nFyo76JLPKSXHZecd4nZEyu92jesaor/pub?gid=0&single=true&output=csv"
SPREADSHEET_ID = st.secrets.get("SPREADSHEET_ID", "")        # .../spreadsheets/d/<ID>/edit
WORKSHEET_NAME = st.secrets.get("WORKSHEET_NAME", "Arkusz1")

st.set_page_config(page_title="Mapa klientów z Excela / Google Sheets", layout="wide")
st.title("📍 Mapa klientów z Excela / Google Sheets (online, free)")
st.caption("Wymagane: **Adres**, **Miasto**, **PSC**. Opcjonalne: **Nazwa odbiorcy**, **Obrót w czk**, **email**, **lat**, **lon**.")

REQ_ADDR_COLS = ["Adres", "Miasto", "PSC"]

# -------------------- HELPERY --------------------
def to_float_or_none(x):
    try:
        if x is None or (isinstance(x, float) and math.isnan(x)):
            return None
        s = str(x).strip().replace("\xa0", " ").replace(" ", "")
        s = s.replace(",", ".")
        return float(s)
    except Exception:
        return None

def fmt_czk(x):
    v = to_float_or_none(x)
    if v is None:
        return ""
    txt = f"{v:,.2f}".replace(",", " ").replace(".", ",")
    return f"Obrót: {txt} CZK"

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

@st.cache_data(show_spinner=False)
def load_google_csv(url: str) -> pd.DataFrame:
    headers = {"User-Agent": "Mozilla/5.0"}
    r = requests.get(url, timeout=30, headers=headers)
    r.raise_for_status()
    content = r.content
    for enc in ("utf-8-sig", "utf-8", "cp1250", "iso-8859-2"):
        try:
            return pd.read_csv(io.StringIO(content.decode(enc)))
        except Exception:
            continue
    return pd.read_excel(io.BytesIO(content))

@st.cache_data(show_spinner=False)
def geocode_one(address: str):
    geolocator = Nominatim(user_agent="mroauto-excel-map (contact: info@mroauto.cz)")
    geocode = RateLimiter(geolocator.geocode, min_delay_seconds=1, swallow_exceptions=True)
    loc = geocode(address)
    if loc:
        return (loc.latitude, loc.longitude)
    return None

def save_to_google_sheet(df_to_save: pd.DataFrame):
    if not SPREADSHEET_ID:
        st.error("Brakuje SPREADSHEET_ID w Secrets.")
        return False
    try:
        import gspread
        from gspread_dataframe import set_with_dataframe
        from google.oauth2.service_account import Credentials
        scopes = ["https://www.googleapis.com/auth/spreadsheets"]
        creds = Credentials.from_service_account_info(st.secrets["gcp_service_account"], scopes=scopes)
        gc = gspread.authorize(creds)
        sh = gc.open_by_key(SPREADSHEET_ID)
        try:
            ws = sh.worksheet(WORKSHEET_NAME)
        except gspread.WorksheetNotFound:
            ws = sh.add_worksheet(title=WORKSHEET_NAME, rows="1000", cols="26")
        set_with_dataframe(ws, df_to_save, include_index=False, include_column_header=True)
        return True
    except Exception as e:
        st.error(f"Nie udało się zapisać do Google Sheets: {e}")
        return False

def df_hash(df: pd.DataFrame) -> str:
    # stabilny hash treści danych do kontroli rerenderu
    return hashlib.md5(pd.util.hash_pandas_object(df, index=True).values).hexdigest()

# -------------------- UI --------------------
c1, c2, c3 = st.columns([1,1,1])
with c1:
    uploaded = st.file_uploader("Wgraj plik (Excel/CSV)", type=["xlsx", "csv"])
with c2:
    go_btn = st.button("⬇️ Pobierz z Google (stały link)")
with c3:
    auto_geocode = st.checkbox("Auto-geokoduj, jeśli brak lat/lon", value=False)

# Guzik Reset (czyści stan, pomaga przy testach)
reset = st.button("♻️ Reset (wyczyść wyniki)")
if reset:
    for k in ["_use_google", "geo_df", "map_hash", "data_hash"]:
        st.session_state.pop(k, None)
    st.experimental_rerun()

if go_btn:
    st.session_state["_use_google"] = True
    # nowy import = wyczyść poprzednie
    for k in ["geo_df", "map_hash", "data_hash"]:
        st.session_state.pop(k, None)

use_google = st.session_state.get("_use_google", False)

# -------------------- DANE --------------------
df = None
if uploaded is not None and not use_google:
    if uploaded.name.lower().endswith(".csv"):
        df = pd.read_csv(uploaded)
    else:
        df = pd.read_excel(uploaded)
elif use_google:
    df = load_google_csv(GOOGLE_SHEETS_CSV)

if df is None or df.empty:
    st.info("Wgraj plik lub kliknij „Pobierz z Google (stały link)”.")
    st.stop()

st.subheader("Podgląd danych")
st.dataframe(df.head(50), width="stretch")

# -------------------- NORMALIZACJA --------------------
has_coords = {"lat", "lon"} <= set(df.columns)
if "Adres" in df.columns:  df["Adres"]  = norm_col(df["Adres"])
if "Miasto" in df.columns: df["Miasto"] = norm_col(df["Miasto"])
if "PSC" in df.columns:    df["PSC"]    = norm_col(df["PSC"]).str.replace(" ", "")

if not has_coords:
    missing = [c for c in REQ_ADDR_COLS if c not in df.columns]
    if missing:
        st.error(f"Brakuje kolumn: {', '.join(missing)} (lub dodaj lat/lon).")
        st.stop()
    if "FullAddress" not in df.columns:
        df["FullAddress"] = build_full_address(df)
    df = df[df["FullAddress"].str.len() > 0].copy()

# slot na mapę
map_slot = st.empty()

# -------------------- GEO_DF Z SESJI --------------------
geo_df = None
if "geo_df" in st.session_state:
    try:
        geo_df = pd.DataFrame(st.session_state["geo_df"])
    except Exception:
        geo_df = None

# Jeśli wejściowe dane mają lat/lon a nie mamy jeszcze geo_df
if has_coords and geo_df is None:
    df["lat"] = df["lat"].apply(to_float_or_none)
    df["lon"] = df["lon"].apply(to_float_or_none)
    geo_df = df.dropna(subset=["lat", "lon"]).copy()
    st.session_state["geo_df"] = geo_df.to_dict(orient="records")

# -------------------- GEOKODOWANIE --------------------
if not has_coords and geo_df is None:
    st.warning("Brak kolumn lat/lon — mogę policzyć współrzędne (OSM/Nominatim, ~1 zapytanie/s).")
    max_rows = 300
    if len(df) > max_rows:
        st.info(f"Adresów: {len(df)}. Dla bezpieczeństwa geokoduję pierwsze {max_rows}.")
    to_geo = df["FullAddress"].head(max_rows).tolist()

    trigger = auto_geocode or st.button("📍 Geokoduj adresy (OSM)", key="btn_geocode")
    if trigger:
        results = []
        prog = st.progress(0.0)
        for i, addr in enumerate(to_geo, start=1):
            coords = geocode_one(addr)
            results.append((addr, coords))
            prog.progress(i/len(to_geo))
            time.sleep(0.05)

        mapping = {a: c for a, c in results}
        df["lat"] = df["FullAddress"].map(lambda a: mapping.get(a, (None, None))[0] if mapping.get(a) else None)
        df["lon"] = df["FullAddress"].map(lambda a: mapping.get(a, (None, None))[1] if mapping.get(a) else None)
        geo_df = df.dropna(subset=["lat", "lon"]).copy()

        st.session_state["geo_df"] = geo_df.to_dict(orient="records")
        st.success(f"Znaleziono współrzędne dla {geo_df.shape[0]} / {len(to_geo)} adresów.")
        st.experimental_rerun()

# Jeśli dalej nie mamy współrzędnych – pokaż ramkę z instrukcją, ale NIE rysuj neutralnej mapy, by nie przykrywać
if geo_df is None or geo_df.empty:
    st.info("Brak gotowych współrzędnych. Użyj geokodowania albo dodaj kolumny 'lat' i 'lon' w arkuszu.")
    st.stop()

# -------------------- MAPA (render tylko gdy zmieniły się dane) --------------------
current_hash = df_hash(geo_df)
if st.session_state.get("map_hash") != current_hash:
    m = folium.Map(location=[geo_df["lat"].mean(), geo_df["lon"].mean()], zoom_start=8)
    cluster = MarkerCluster().add_to(m)

    def val(col, row, default=""):
        return row[col] if col in geo_df.columns and pd.notna(row[col]) else default

    for _, r in geo_df.iterrows():
        amount_text = fmt_czk(val('Obrót w czk', r))
        popup_html = f"""
        <div style="font-size:14px">
          <b>{val('Nazwa odbiorcy', r)}</b><br>
          {amount_text}<br>
          {('Email: ' + val('email', r)) if val('email', r) else ''}<br>
          {('Adres: ' + val('FullAddress', r)) if 'FullAddress' in geo_df.columns else ''}
        </div>
        """
        folium.Marker(
            [r["lat"], r["lon"]],
            tooltip=val('Nazwa odbiorcy', r) or "Klient",
            popup=folium.Popup(popup_html, max_width=350)
        ).add_to(cluster)

    with map_slot:
        st_folium(m, height=700, key="mapview")

    st.session_state["map_hash"] = current_hash
else:
    # nic się nie zmieniło – nie przebudowuj mapy
    with map_slot:
        st_folium(folium.Map(), height=1, key="mapview")  # utrzymaj widget, ale nie renderuj ciężkiej mapy ponownie
        st.write("")  # no-op

# -------------------- EKSPORT / ZAPIS --------------------
with st.expander("💾 Eksport / Zapis"):
    st.download_button(
        "Pobierz CSV z lat/lon",
        data=geo_df.to_csv(index=False).encode("utf-8"),
        file_name="geokodowane_dane.csv",
        mime="text/csv"
    )
    html = folium.Map().get_root().render()  # placeholder by uniknąć rerunów od renderu m.get_root()
    # jeśli chcesz realny eksport HTML ostatniej mapy, trzymaj m w session_state i tu go użyj

    if st.button("📤 Zapisz do Google Sheets (nadpisze zakładkę)"):
        ok = save_to_google_sheet(geo_df)
        if ok:
            st.success("Zapisano do Google Sheets ✅")
