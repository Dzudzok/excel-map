import math
import time
import hashlib
from typing import Tuple, Optional

import pandas as pd
import streamlit as st
import folium
from folium.plugins import MarkerCluster
from geopy.geocoders import Nominatim
from geopy.extra.rate_limiter import RateLimiter

# === USTAWIENIA ===
st.set_page_config(page_title="Mapa klientów z Excela / Google Sheets", layout="wide")
if "saved_latlon_keys" not in st.session_state:
    st.session_state["saved_latlon_keys"] = set()


# 1) ŹRÓDŁO CSV — użyj secrets albo stałej
CSV_URL = st.secrets.get(
    "GOOGLE_SHEETS_CSV",
    "https://docs.google.com/spreadsheets/d/e/2PACX-1vSijSBg7JqZkg4T8aY56FEhox0pqw5huE7oWRmSbaB25LJj9nFyo76JLPKSXHZecd4nZEyu92jesaor/pub?gid=0&single=true&output=csv"
)

# 2) Opcjonalne zapisywanie lat/lon do arkusza (wymaga sekretnych danych serwisowego konta)
ENABLE_WRITE_BACK = all(
    k in st.secrets
    for k in ("SPREADSHEET_ID", "WORKSHEET_NAME", "gcp_service_account")
)

# === GEOCODER ===
_geolocator = Nominatim(user_agent="mroauto-excel-map")
_geocode = RateLimiter(_geolocator.geocode, min_delay_seconds=1, max_retries=2, swallow_exceptions=True)

@st.cache_data(show_spinner=False)
def load_csv(url: str) -> pd.DataFrame:
    df = pd.read_csv(url, dtype=str).fillna("")
    # Zapewnij obecność kolumn
    for col in ["lp.", "Nazwa odbiorcy", "Obrót w czk", "email", "Adres", "Miasto", "PSC", "lat", "lon"]:
        if col not in df.columns:
            df[col] = ""
    # Współrzędne jako numery (jeśli są)
    df["lat"] = pd.to_numeric(df["lat"], errors="coerce")
    df["lon"] = pd.to_numeric(df["lon"], errors="coerce")
    return df

def build_full_address(row: pd.Series) -> str:
    parts = [str(row.get("Adres", "")).strip(),
             str(row.get("Miasto", "")).strip(),
             str(row.get("PSC", "")).strip(),
             "Czechy"]  # możesz zmienić/dostawiać kraj, jeśli masz też PL itd.
    return ", ".join([p for p in parts if p])



def row_key(row: pd.Series) -> str:
    # Klucz oparty o pełny adres + nazwę odbiorcy
    base = (build_full_address(row) + "|" + str(row.get("Nazwa odbiorcy",""))).strip()
    return hashlib.sha1(base.encode("utf-8")).hexdigest()


@st.cache_data(show_spinner=False)
def geocode_one(address: str) -> Optional[Tuple[float, float]]:
    """Geokoduje adres i zwraca (lat, lon) albo None."""
    if not address:
        return None
    try:
        loc = _geocode(address)
        if loc:
            return (float(loc.latitude), float(loc.longitude))
    except Exception:
        pass
    return None

def geocode_missing(df: pd.DataFrame) -> pd.DataFrame:
    """Dla wierszy bez lat/lon — wylicza współrzędne. Zwraca kopię df."""
    out = df.copy()
    need = out["lat"].isna() | out["lon"].isna()
    idxs = out.index[need].tolist()

    if not idxs:
        return out

    progress = st.progress(0)
    status = st.empty()

    for i, idx in enumerate(idxs, start=1):
        address = build_full_address(out.loc[idx])
        status.text(f"Geokoduję: {address}")
        coords = geocode_one(address)
        if coords:
            out.at[idx, "lat"] = coords[0]
            out.at[idx, "lon"] = coords[1]
        progress.progress(i / len(idxs))
    status.empty()
    progress.empty()
    return out

def write_back_latlon(original: pd.DataFrame, updated: pd.DataFrame) -> int:
    """
    Dopisuje do Google Sheets tylko te wiersze, które miały NaN w lat/lon,
    a teraz mają wartości. Zwraca liczbę zaktualizowanych wierszy.
    Wymaga: st.secrets["SPREADSHEET_ID"], ["WORKSHEET_NAME"], ["gcp_service_account"]
    """
    if not ENABLE_WRITE_BACK:
        return 0

    import gspread
    from google.oauth2.service_account import Credentials

    # Przygotuj maskę „co nowego dopisać”
    orig_nan = original["lat"].isna() | original["lon"].isna()
    now_ok = updated["lat"].notna() & updated["lon"].notna()
    changed_mask = orig_nan & now_ok
    if not changed_mask.any():
        return 0

    # Autoryzacja
    scope = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive"
    ]
    creds = Credentials.from_service_account_info(
        st.secrets["gcp_service_account"], scopes=scope
    )
    gc = gspread.authorize(creds)
    sh = gc.open_by_key(st.secrets["SPREADSHEET_ID"])
    ws = sh.worksheet(st.secrets["WORKSHEET_NAME"])

    # Pobierz pierwszy wiersz (nagłówki), aby znać indeksy kolumn
    headers = ws.row_values(1)
    def col_index(name: str) -> int:
        try:
            return headers.index(name) + 1  # 1-based index
        except ValueError:
            # Jeśli brak kolumny — dołóż ją na końcu
            ws.update_cell(1, len(headers) + 1, name)
            headers.append(name)
            return len(headers)

    col_lat = col_index("lat")
    col_lon = col_index("lon")

    # GSpread ma indeks wierszy 1-based (1 = nagłówki). Zakładamy, że kolejność w CSV == kolejność w arkuszu
    updated_rows = 0
    for i, (idx, row) in enumerate(updated.iterrows(), start=2):  # 2 = pierwszy wiersz danych
        if not changed_mask.loc[idx]:
            continue
        key = row_key(row)
        if key in st.session_state["saved_latlon_keys"]:
            continue  # już zapisane w tej sesji
        lat_val = float(row["lat"])
        lon_val = float(row["lon"])
        ws.update_cell(i, col_lat, lat_val)
        ws.update_cell(i, col_lon, lon_val)
        st.session_state["saved_latlon_keys"].add(key)
        updated_rows += 1
        time.sleep(0.2)

    return updated_rows

def make_map(df: pd.DataFrame) -> folium.Map:
    # Filtrowanie tylko z kompletem współrzędnych
    dd = df.dropna(subset=["lat", "lon"]).copy()

    if dd.empty:
        center = (49.820, 15.470)  # środek CZ
        m = folium.Map(location=center, zoom_start=7, control_scale=True)
        return m

    # Środek mapy
    center = (dd["lat"].mean(), dd["lon"].mean())
    m = folium.Map(location=center, zoom_start=7, control_scale=True)

    cluster = MarkerCluster().add_to(m)

    def fmt_popup(r: pd.Series) -> str:
        lines = []
        lines.append(f"<b>{r.get('Nazwa odbiorcy','').strip()}</b>")
        obr = str(r.get("Obrót w czk", "")).strip()
        if obr:
            lines.append(f"Obrót: {obr} CZK")
        email = str(r.get("email", "")).strip()
        if email:
            lines.append(f"Email: {email}")
        addr = build_full_address(r)
        lines.append(addr)
        return "<br>".join(lines)

    for _, r in dd.iterrows():
        popup_html = folium.Popup(fmt_popup(r), max_width=320)
        folium.Marker(
            location=(float(r["lat"]), float(r["lon"])),
            popup=popup_html,
            tooltip=r.get("Nazwa odbiorcy", "")
        ).add_to(cluster)
    return m

# === UI ===
st.title("🗺️ Mapa klientów z Google Sheets (CSV)")

with st.sidebar:
    st.subheader("Źródło danych")
    st.code(CSV_URL, language="text")
    st.caption("Aby edytować dane, modyfikuj plik Google Sheets. Aplikacja wczyta je z linku CSV.")

df_orig = load_csv(CSV_URL)

st.markdown("#### Podgląd danych (pierwsze 20 wierszy)")
st.dataframe(df_orig.head(20), use_container_width=True)

# Geokodowanie braków
st.markdown("### Geokodowanie brakujących współrzędnych")
df_geo = geocode_missing(df_orig)

# Opcjonalny zapis do arkusza
updated_rows = 0
if ENABLE_WRITE_BACK:
    st.markdown("### Zapis do Google Sheets")
    if st.button("💾 Zapisz nowe współrzędne do arkusza"):
        with st.spinner("Zapisuję nowe współrzędne do Google Sheets…"):
            try:
                updated_rows = write_back_latlon(df_orig, df_geo)
                if updated_rows == 0:
                    st.success("Brak nowych współrzędnych do zapisania.")
                else:
                    st.success(f"Zaktualizowano w arkuszu: {updated_rows} wierszy.")
            except Exception as e:
                st.warning(f"Nie udało się zapisać do arkusza: {e}")
    else:
        st.caption("Współrzędne zapiszą się dopiero po kliknięciu przycisku.")


# Mapa
st.markdown("### Mapa")
m = make_map(df_geo)
from streamlit_folium import st_folium  # wymaga: pip install streamlit-folium
st_folium(m, width=None, height=700)

# Podsumowanie
missing_after = df_geo["lat"].isna() | df_geo["lon"].isna()
st.info(
    f"Znaleziono współrzędne dla {(~missing_after).sum()} rekordów. "
    f"Brakujących: {missing_after.sum()}. "
    + (f"Zaktualizowano w arkuszu: {updated_rows} wierszy." if ENABLE_WRITE_BACK else "")
)

