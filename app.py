import io, time, math, hashlib, requests, numpy as np
import pandas as pd
import streamlit as st
import streamlit.components.v1 as components
import folium
from folium.plugins import MarkerCluster
from geopy.geocoders import Nominatim
from geopy.extra.rate_limiter import RateLimiter

# ===================== KONFIG =====================
GOOGLE_SHEETS_CSV = "https://docs.google.com/spreadsheets/d/e/2PACX-1vSijSBg7JqZkg4T8aY56FEhox0pqw5huE7oWRmSbaB25LJj9nFyo76JLPKSXHZecd4nZEyu92jesaor/pub?gid=0&single=true&output=csv"
SPREADSHEET_ID  = st.secrets.get("SPREADSHEET_ID", "")
WORKSHEET_NAME  = st.secrets.get("WORKSHEET_NAME", "Arkusz1")

st.set_page_config(page_title="Mapa klientów z Excela / Google Sheets", layout="wide")
st.title("📍 Mapa klientów z Excela / Google Sheets (online, free)")
st.caption("Wymagane kolumny: **Adres**, **Miasto**, **PSC**. Opcjonalne: **Nazwa odbiorcy**, **Obrót w czk**, **email**, **lat**, **lon**.")

REQ_ADDR_COLS = ["Adres", "Miasto", "PSC"]
OPT_COLS = ["Nazwa odbiorcy", "Obrót w czk", "email", "lat", "lon", "FullAddress"]

# ===================== HELPERY =====================
def to_float_or_none(x):
    try:
        if x is None or (isinstance(x, float) and math.isnan(x)):
            return None
        s = str(x).strip().replace("\xa0"," ").replace(" ","").replace(",", ".")
        return float(s)
    except Exception:
        return None


def fmt_czk(x):
    v = to_float_or_none(x)
    if v is None:
        return ""
    return f"Obrót: {v:,.2f}".replace(",", " ").replace(".", ",") + " CZK"


def norm_col(s: pd.Series) -> pd.Series:
    return (
        s.fillna("").astype(str)
        .str.replace(r"^\s*0\s*$","", regex=True)
        .str.replace(r"\s+"," ", regex=True)
        .str.strip()
    )


def build_full_address(df: pd.DataFrame) -> pd.Series:
    a = norm_col(df["Adres"]).fillna("")
    m = norm_col(df["Miasto"]).fillna("")
    p = norm_col(df["PSC"]).str.replace(" ","").fillna("")
    return (a + ", " + m + " " + p).str.strip(", ").str.strip()


# ===================== ODCZYT DANYCH =====================
@st.cache_data(show_spinner=False, ttl=10)
def load_google(url_csv: str, spreadsheet_id: str, worksheet_name: str) -> pd.DataFrame:
    """Czytaj z API (świeże) – fallback na publikowany CSV."""
    try:
        if spreadsheet_id and "gcp_service_account" in st.secrets:
            import gspread
            from google.oauth2.service_account import Credentials

            creds = Credentials.from_service_account_info(
                st.secrets["gcp_service_account"],
                scopes=["https://www.googleapis.com/auth/spreadsheets.readonly"],
            )
            sh = gspread.authorize(creds).open_by_key(spreadsheet_id)
            ws = sh.worksheet(worksheet_name)
            df = pd.DataFrame(ws.get_all_records(numericise_ignore=['all']))
            if not df.empty:
                return df
    except Exception as e:
        st.warning(f"Odczyt API nieudany ({e}). Próbuję przez publikowany CSV…")

    r = requests.get(url_csv, timeout=30, headers={"User-Agent": "Mozilla/5.0"})
    r.raise_for_status()
    b = r.content
    for enc in ("utf-8-sig", "utf-8", "cp1250", "iso-8859-2"):
        try:
            return pd.read_csv(io.StringIO(b.decode(enc)))
        except Exception:
            pass
    return pd.read_excel(io.BytesIO(b))


def save_to_google_sheet(df_to_save: pd.DataFrame) -> bool:
    if not SPREADSHEET_ID:
        st.error("Brakuje SPREADSHEET_ID w Secrets.")
        return False
    try:
        import gspread
        from gspread_dataframe import set_with_dataframe
        from google.oauth2.service_account import Credentials

        creds = Credentials.from_service_account_info(
            st.secrets["gcp_service_account"], scopes=["https://www.googleapis.com/auth/spreadsheets"]
        )
        gc = gspread.authorize(creds)
        sh = gc.open_by_key(SPREADSHEET_ID)
        try:
            ws = sh.worksheet(WORKSHEET_NAME)
        except gspread.WorksheetNotFound:
            ws = sh.add_worksheet(title=WORKSHEET_NAME, rows="2000", cols="50")
        from gspread.exceptions import APIError
        for attempt in range(5):
            try:
                set_with_dataframe(ws, df_to_save, include_index=False, include_column_header=True)
                return True
            except APIError as e:
                msg = str(e)
                code = getattr(getattr(e, 'response', None), 'status_code', None)
                if code == 429 or ('Quota exceeded' in msg):
                    wait = min(32, 2 ** attempt)
                    st.warning(f"Przekroczono limit Google Sheets (429). Ponawiam za {wait}s… [próba {attempt+1}/5]")
                    time.sleep(wait)
                    continue
                raise
        st.error("Przekroczono limit Google Sheets – nie udało się zapisać po 5 próbach.")
        return False
    except Exception as e:
        st.error(f"Nie udało się zapisać do Google Sheets: {e}")
        return False


# ===================== GEOKODER =====================
@st.cache_data(show_spinner=False)
def geocode_one(address: str):
    """Zwraca (lat, lon) dla adresów CZ/PL albo None – z timeoutami i retry."""
    geolocator = Nominatim(
        user_agent="mroauto-excel-map (contact: info@mroauto.cz)", timeout=10
    )
    geocode = RateLimiter(
        geolocator.geocode,
        min_delay_seconds=1.1,
        error_wait_seconds=3.0,
        max_retries=2,
        swallow_exceptions=True,
    )
    params = dict(
        addressdetails=False,
        language="cs",
        country_codes="cz,pl",
        viewbox=((12.0, 51.6), (24.3, 48.0)),
        bounded=True,
        exactly_one=True,
    )
    loc = geocode(address, **params) or geocode(address, **{**params, "language": "pl"})
    if not loc:
        return None
    lat, lon = float(loc.latitude), float(loc.longitude)
    if not (48.0 <= lat <= 55.0 and 12.0 <= lon <= 25.0):
        return None
    return (lat, lon)


def valid_coord(lat, lon) -> bool:
    try:
        lat = float(lat)
        lon = float(lon)
    except Exception:
        return False
    return (48.0 <= lat <= 55.0) and (12.0 <= lon <= 25.0)


# ===================== UI: WYBÓR ŹRÓDŁA =====================
left, right = st.columns([3, 2])
with left:
    source = st.radio(
        "Źródło danych",
        ["Google Sheet", "Plik (upload)"],
        horizontal=True,
        index=0,
        help="Wybierz skąd czytać dane wejściowe.",
    )
with right:
    auto_geocode = st.checkbox("📍 Auto-geokoduj brakujące/błędne", value=True)

uploaded = None
if source == "Plik (upload)":
    uploaded = st.file_uploader("Wgraj plik (Excel/CSV)", type=["xlsx", "csv"])

st.divider()

# ===================== WCZYTANIE DANYCH =====================
if source == "Google Sheet":
    df = load_google(GOOGLE_SHEETS_CSV, SPREADSHEET_ID, WORKSHEET_NAME)
else:
    if uploaded is None:
        st.info("Wgraj plik, aby kontynuować.")
        st.stop()
    df = (
        pd.read_csv(uploaded)
        if uploaded.name.lower().endswith(".csv")
        else pd.read_excel(uploaded)
    )

if df is None or df.empty:
    st.info("Brak danych – uzupełnij arkusz lub wgraj plik.")
    st.stop()

# Upewnij się, że kolumny istnieją
for c in REQ_ADDR_COLS:
    if c not in df.columns:
        st.error(f"Brakuje kolumny: {c}")
        st.stop()

# Zadbaj o istnienie kolumn lat/lon/FullAddress
if "FullAddress" not in df.columns:
    df["FullAddress"] = build_full_address(df)

if "lat" not in df.columns:
    df["lat"] = np.nan
if "lon" not in df.columns:
    df["lon"] = np.nan

# Normalizacja lat/lon
for col in ["lat", "lon"]:
    df[col] = (
        pd.to_numeric(df[col].astype(str).str.replace(",", "."), errors="coerce")
        .astype("float64")
    )

# Wiersze wymagające geokodowania
needs_geo_mask = (
    df["lat"].isna()
    | df["lon"].isna()
    | (~df[["lat", "lon"]].apply(lambda r: valid_coord(r["lat"], r["lon"]), axis=1))
)
needs_geo_idx = df.index[needs_geo_mask]

# ===================== PODGLĄD =====================
st.subheader("Podgląd danych")
st.dataframe(df.head(50), width="stretch")

# ===================== GEOKODOWANIE =====================
if len(needs_geo_idx) > 0:
    st.warning(
        f"Do uzupełnienia/błędne współrzędne: {len(needs_geo_idx)}. (OSM/Nominatim: ~1 zapytanie/s)"
    )
    if auto_geocode or st.button("Geokoduj teraz"):
        addrs = df.loc[needs_geo_idx, "FullAddress"].tolist()
        results = []
        prog = st.progress(0.0)
        for i, addr in enumerate(addrs, start=1):
            results.append((addr, geocode_one(addr)))
            prog.progress(i / len(addrs))
            time.sleep(0.05)

        mapping = {a: c for a, c in results}

        def pick_lat(a):
            c = mapping.get(a)
            return c[0] if c else np.nan

        def pick_lon(a):
            c = mapping.get(a)
            return c[1] if c else np.nan

        df.loc[needs_geo_idx, "lat"] = (
            df.loc[needs_geo_idx, "FullAddress"].map(pick_lat).astype("float64")
        )
        df.loc[needs_geo_idx, "lon"] = (
            df.loc[needs_geo_idx, "FullAddress"].map(pick_lon).astype("float64")
        )

        # Ostateczna walidacja – jeśli dalej błędne, zerujemy
        bad_mask = ~df.apply(lambda r: valid_coord(r["lat"], r["lon"]), axis=1)
        df.loc[bad_mask, ["lat", "lon"]] = np.nan

        # Zapis do Google (jeśli źródło to Google Sheet)
        if source == "Google Sheet":
            if save_to_google_sheet(df):
                st.success(
                    f"Zapisano uzupełnione współrzędne do Google Sheets (zakładka: {WORKSHEET_NAME})."
                )
            else:
                st.info("Uzupełniono lokalnie (zapis do Sheets nieudany).")
        st.rerun()

# Dane z poprawnymi współrzędnymi do rysowania
geo_df = df.dropna(subset=["lat", "lon"]).copy()
if geo_df.empty:
    st.info(
        "Brak poprawnych współrzędnych. Użyj geokodowania albo uzupełnij lat/lon w arkuszu."
    )
    st.stop()

# ===================== MAPA =====================
# Środek i bounds
m = folium.Map(location=[geo_df["lat"].mean(), geo_df["lon"].mean()], zoom_start=7)
cluster = MarkerCluster().add_to(m)

# kalkulacja bounds do dopasowania
try:
    sw = [geo_df["lat"].min(), geo_df["lon"].min()]
    ne = [geo_df["lat"].max(), geo_df["lon"].max()]
    m.fit_bounds([sw, ne])
except Exception:
    pass


def val(col, row, default=""):
    return row[col] if col in geo_df.columns and pd.notna(row[col]) else default

for _, r in geo_df.iterrows():
    popup_html = f"""
    <div style=\"font-size:14px\">
      <b>{val('Nazwa odbiorcy', r)}</b><br>
      {fmt_czk(val('Obrót w czk', r))}<br>
      {('Email: ' + val('email', r)) if val('email', r) else ''}<br>
      {('Adres: ' + val('FullAddress', r)) if 'FullAddress' in geo_df.columns else ''}
    </div>
    """
    folium.Marker(
        [r["lat"], r["lon"]],
        tooltip=val("Nazwa odbiorcy", r) or "Klient",
        popup=folium.Popup(popup_html, max_width=350),
    ).add_to(cluster)

components.html(m.get_root().render(), height=700, scrolling=False)

# ===================== EKSPORT / ZAPIS =====================
with st.expander("💾 Eksport / Zapis"):
    c1, c2 = st.columns(2)
    with c1:
        st.download_button(
            "Pobierz CSV z lat/lon (to, co rysuje mapa)",
            data=geo_df.to_csv(index=False).encode("utf-8"),
            file_name="geokodowane_dane.csv",
            mime="text/csv",
        )
    with c2:
        if source == "Google Sheet":
            if st.button("📤 Zapisz cały arkusz do Google Sheets (nadpisze zakładkę)"):
                if save_to_google_sheet(df):
                    st.success("Zapisano do Google Sheets ✅")
                    st.rerun()
        else:
            st.caption("Źródło = plik. Aby zapisać w Google, przełącz się na źródło 'Google Sheet'.")

# ===================== STATUS =====================
ok_cnt = len(geo_df)
bad_cnt = int((~df.index.isin(geo_df.index)).sum())
st.info(f"✅ Na mapie: {ok_cnt} punktów. ❗ Bez współrzędnych (nie pokazane): {bad_cnt}.")
