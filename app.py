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
import streamlit.components.v1 as components


# -------------------- KONFIG --------------------
GOOGLE_SHEETS_CSV = "https://docs.google.com/spreadsheets/d/e/2PACX-1vSijSBg7JqZkg4T8aY56FEhox0pqw5huE7oWRmSbaB25LJj9nFyo76JLPKSXHZecd4nZEyu92jesaor/pub?gid=0&single=true&output=csv"

SPREADSHEET_ID  = st.secrets.get("SPREADSHEET_ID", "")        # .../spreadsheets/d/<ID>/edit
WORKSHEET_NAME  = st.secrets.get("WORKSHEET_NAME", "Arkusz1")  # zakładka do nadpisania

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

@st.cache_data(show_spinner=False, ttl=60)
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
    """Zwraca (lat, lon) albo None. Ograniczone do CZ/PL i odrzuca wyniki poza bbox."""
    geolocator = Nominatim(user_agent="mroauto-excel-map (contact: info@mroauto.cz)")
    geocode = RateLimiter(geolocator.geocode, min_delay_seconds=1, swallow_exceptions=True)

    # preferencje: język czeski/polski, tylko CZ/PL, zawęź do bbox (Czechy+Polska)
    params = dict(
        addressdetails=False,
        language="cs",
        country_codes="cz,pl",
        viewbox=((12.0, 51.6), (24.3, 48.0)),  # (west, north) , (east, south)
        bounded=True,
        exactly_one=True
    )
    loc = geocode(address, **params)
    if not loc:
        # druga próba po polsku (czasem pomaga)
        params["language"] = "pl"
        loc = geocode(address, **params)
    if not loc:
        return None

    lat, lon = float(loc.latitude), float(loc.longitude)

    # twarda walidacja — tylko CZ/PL
    if not (48.0 <= lat <= 55.0 and 12.0 <= lon <= 25.0):
        return None
    return (lat, lon)


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
    return hashlib.md5(pd.util.hash_pandas_object(df, index=True).values).hexdigest()

# -------------------- UI --------------------
c1, c2, c3, c4 = st.columns([1,1,1,1])
with c1:
    uploaded = st.file_uploader("Wgraj plik (Excel/CSV)", type=["xlsx", "csv"])
with c2:
    go_btn = st.button("⬇️ Pobierz z Google (stały link)")
with c3:
    auto_geocode = st.checkbox("Auto-geokoduj, jeśli brak lat/lon", value=True)
with c4:
    reset = st.button("♻️ Reset (wyczyść)")

if reset:
    st.cache_data.clear() 
    for k in ["_use_google", "geo_df", "map_hash", "data_hash", "geocoded_done"]:
        st.session_state.pop(k, None)
    st.rerun()   # zamiast st.experimental_rerun()


if go_btn:
    st.cache_data.clear()
    st.session_state["_use_google"] = True
    for k in ["geo_df", "map_hash", "data_hash", "geocoded_done"]:
        st.session_state.pop(k, None)
    st.rerun()   # zamiast st.experimental_rerun()





use_google = st.session_state.get("_use_google", True)  # domyślnie czytaj z Google

# -------------------- DANE --------------------
df = None
if uploaded is not None and not use_google:
    if uploaded.name.lower().endswith(".csv"):
        df = pd.read_csv(uploaded)
    else:
        df = pd.read_excel(uploaded)
else:
    df = load_google_csv(GOOGLE_SHEETS_CSV)

if df is None or df.empty:
    st.info("Brak danych – wrzuć plik albo użyj przycisku Google.")
    st.stop()

st.subheader("Podgląd danych")
st.dataframe(df.head(50), width="stretch")


# -------------------- NORMALIZACJA --------------------
# 1) wykryj kolumny współrzędnych i zamień na numeric
has_coord_cols = {"lat", "lon"}.issubset(df.columns)
if has_coord_cols:
    df["lat"] = pd.to_numeric(df["lat"], errors="coerce")
    df["lon"] = pd.to_numeric(df["lon"], errors="coerce")

# 2) sprzątanie adresu
if "Adres" in df.columns:  df["Adres"]  = norm_col(df["Adres"])
if "Miasto" in df.columns: df["Miasto"] = norm_col(df["Miasto"])
if "PSC" in df.columns:    df["PSC"]    = norm_col(df["PSC"]).str.replace(" ", "")

# 3) FullAddress (potrzebny, gdy będziemy cokolwiek geokodować)
need_address_cols = [c for c in REQ_ADDR_COLS if c not in df.columns]
if need_address_cols:
    st.error(f"Brakuje kolumn: {', '.join(need_address_cols)} (lub dodaj lat/lon).")
    st.stop()
if "FullAddress" not in df.columns:
    df["FullAddress"] = build_full_address(df)
def valid_coord(lat, lon):
    try:
        lat = float(lat); lon = float(lon)
    except Exception:
        return False
    return (48.0 <= lat <= 55.0) and (12.0 <= lon <= 25.0)

# zamień przecinki w istniejących lat/lon (jeśli ktoś wpisał „49,83”)
if "lat" in df.columns: df["lat"] = df["lat"].astype(str).str.replace(",", ".")
if "lon" in df.columns: df["lon"] = df["lon"].astype(str).str.replace(",", ".")
df["lat"] = pd.to_numeric(df.get("lat"), errors="coerce")
df["lon"] = pd.to_numeric(df.get("lon"), errors="coerce")

# wiersze do naprawy = brak lub nieprawidłowe współrzędne (np. 34/70)
to_fix_idx = df.index[(df["lat"].isna() | df["lon"].isna()) |
                      (~df[["lat","lon"]].apply(lambda r: valid_coord(r["lat"], r["lon"]), axis=1))]

# 4) status współrzędnych
any_coords = has_coord_cols and not df[["lat", "lon"]].dropna().empty
missing_mask = (df["lat"].isna() | df["lon"].isna()) if has_coord_cols else (df["FullAddress"].str.len() > 0)

# --- SLOT NA MAPĘ ---
map_slot = st.empty()

# --- GEO_DF Z SESJI ---
geo_df = None
if "geo_df" in st.session_state:
    try:
        geo_df = pd.DataFrame(st.session_state["geo_df"])
    except Exception:
        geo_df = None

# --- przypadek 1: mamy jakieś koordynaty -> pokaż je, ale pozwól dogeokodować brakujące ---
if any_coords:
    geo_df = df.dropna(subset=["lat", "lon"]).copy()
    st.session_state["geo_df"] = geo_df.to_dict(orient="records")

    # czy są brakujące?
    need_fill = (df["lat"].isna() | df["lon"].isna()).any()
    if need_fill:
        st.info("W arkuszu są brakujące współrzędne. Mogę dogeokodować tylko brakujące rekordy.")
        trigger_fill = auto_geocode or st.button("📍 Geokoduj BRAKUJĄCE (OSM)")
        if trigger_fill and not st.session_state.get("geocoded_done", False):
            to_geo = df.loc[(df["lat"].isna() | df["lon"].isna()), "FullAddress"].tolist()
            results = []
            prog = st.progress(0.0)
            for i, addr in enumerate(to_geo, start=1):
                coords = geocode_one(addr)
                results.append((addr, coords))
                prog.progress(i/len(to_geo))
                time.sleep(0.05)
            mapping = {a: c for a, c in results}
            # uzupełnij TYLKO brakujące
            idx_missing = df.index[(df["lat"].isna() | df["lon"].isna())]
            df.loc[idx_missing, "lat"] = df.loc[idx_missing, "FullAddress"].map(lambda a: mapping.get(a, (None, None))[0] if mapping.get(a) else None)
            df.loc[idx_missing, "lon"] = df.loc[idx_missing, "FullAddress"].map(lambda a: mapping.get(a, (None, None))[1] if mapping.get(a) else None)

            geo_df = df.dropna(subset=["lat", "lon"]).copy()
            st.session_state["geo_df"] = geo_df.to_dict(orient="records")
            st.session_state["geocoded_done"] = True

            saved_ok = save_to_google_sheet(df)
            st.success("Uzupełniono brakujące współrzędne i zapisano do Google Sheets." if saved_ok else "Uzupełniono lokalnie (zapis do Sheets nieudany).")
            st.rerun()

# --- przypadek 2: nie mamy żadnych koordynat – geokoduj WSZYSTKIE ---
# --- GEOKODOWANIE ---
need_any_coords = df[["lat","lon"]].dropna().shape[0] > 0
need_fix = len(to_fix_idx) > 0

if need_fix and geo_df is None:
    info = "Brak części współrzędnych – uzupełnię brakujące i błędne." if need_any_coords \
           else "Brak współrzędnych – policzę wszystkie."
    st.warning(info + " (OSM/Nominatim, ~1 zapytanie/s).")

    to_geo = df.loc[to_fix_idx, "FullAddress"].tolist()
    trigger = auto_geocode or st.button("📍 Geokoduj (uzupełnij brakujące/błędne)")

    if trigger and not st.session_state.get("geocoded_done", False):
        results = []
        prog = st.progress(0.0)
        n = len(to_geo)
        for i, addr in enumerate(to_geo, start=1):
            coords = geocode_one(addr)
            results.append((addr, coords))
            prog.progress(i/n)
            time.sleep(0.05)

        mapping = {a: c for a, c in results}

        # uzupełnij WYŁĄCZNIE wiersze z to_fix_idx
        df.loc[to_fix_idx, "lat"] = df.loc[to_fix_idx, "FullAddress"].map(lambda a: mapping.get(a, (None, None))[0] if mapping.get(a) else None)
        df.loc[to_fix_idx, "lon"] = df.loc[to_fix_idx, "FullAddress"].map(lambda a: mapping.get(a, (None, None))[1] if mapping.get(a) else None)

        # ponowna walidacja – odfiltruj ewentualne śmieci, których nie udało się poprawić
        df.loc[~df.apply(lambda r: valid_coord(r["lat"], r["lon"]), axis=1), ["lat","lon"]] = pd.NA

        geo_df = df.dropna(subset=["lat","lon"]).copy()
        st.session_state["geo_df"] = geo_df.to_dict(orient="records")
        st.session_state["geocoded_done"] = True

        # zapis do Sheets (nadpisuje zakładkę)
        saved_ok = save_to_google_sheet(df)
        if saved_ok:
            st.success(f"Zapisano poprawione współrzędne do Google Sheets (zakładka: {WORKSHEET_NAME}).")
        else:
            st.info("Uzupełniono lokalnie (zapis do Sheets nieudany).")
        st.rerun()

# jeśli po powyższym nadal nie ma poprawnych współrzędnych:
geo_df = df.dropna(subset=["lat","lon"]).copy()
if geo_df.empty:
    st.info("Brak poprawnych współrzędnych po walidacji. Sprawdź adresy/kody pocztowe i spróbuj ponownie.")
    st.stop()



# -------------------- MAPA (stabilny render przez HTML) --------------------
# budujemy mapę ZA KAŻDYM przebiegiem (lekko wolniej, ale bez migania)
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

# render stabilny – bez „migania” streamlit_folium
html_map = m.get_root().render()
components.html(html_map, height=700, scrolling=False)


# -------------------- EKSPORT / ZAPIS --------------------
with st.expander("💾 Eksport / Zapis"):
    st.download_button(
        "Pobierz CSV z lat/lon (to, co rysuje mapa)",
        data=geo_df.to_csv(index=False).encode("utf-8"),
        file_name="geokodowane_dane.csv",
        mime="text/csv"
    )
    if st.button("📤 Zapisz do Google Sheets teraz (nadpisze zakładkę)"):
        ok = save_to_google_sheet(geo_df)
        if ok:
            st.success("Zapisano do Google Sheets ✅")
            st.rerun()   # (opcjonalnie) odśwież po ręcznym zapisie
