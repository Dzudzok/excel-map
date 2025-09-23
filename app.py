import io, time, math, hashlib, requests
import pandas as pd
import streamlit as st
import streamlit.components.v1 as components
import folium
from datetime import datetime, timedelta
from folium.plugins import MarkerCluster
from geopy.geocoders import Nominatim
from geopy.extra.rate_limiter import RateLimiter

# ===================== KONFIG =====================
GOOGLE_SHEETS_CSV = "https://docs.google.com/spreadsheets/d/e/2PACX-1vSijSBg7JqZkg4T8aY56FEhox0pqw5huE7oWRmSbaB25LJj9nFyo76JLPKSXHZecd4nZEyu92jesaor/pub?gid=0&single=true&output=csv"
SPREADSHEET_ID  = st.secrets.get("SPREADSHEET_ID", "")
WORKSHEET_NAME  = st.secrets.get("WORKSHEET_NAME", "Arkusz1")

st.set_page_config(page_title="Mapa klientów z Excela / Google Sheets", layout="wide")
st.title("📍 Mapa klientów z Excela / Google Sheets (online, free)")
st.caption("Wymagane kolumny: **Adres**, **Miasto**, **PSC**. Opcjonalne: **Nazwa odbiorcy**, **Obrót w czk**, **email**.")

REQ_ADDR_COLS = ["Adres", "Miasto", "PSC"]
OPT_COLS = ["Nazwa odbiorcy", "Obrót w czk", "email", "FullAddress"]

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
@st.cache_data(show_spinner=False, ttl=60)
def load_google(url_csv: str, spreadsheet_id: str, worksheet_name: str) -> pd.DataFrame:
    # fallback: CSV (może mieć opóźnienie)
    r = requests.get(url_csv, timeout=30, headers={"User-Agent": "Mozilla/5.0"})
    r.raise_for_status()
    b = r.content
    for enc in ("utf-8-sig", "utf-8", "cp1250", "iso-8859-2"):
        try:
            return pd.read_csv(io.StringIO(b.decode(enc)))
        except Exception:
            pass
    return pd.read_excel(io.BytesIO(b))

# ===================== GEOKODER =====================
@st.cache_data(show_spinner=False)
def geocode_one(address: str):
    geolocator = Nominatim(user_agent="mroauto-excel-map (contact: info@mroauto.cz)", timeout=10)
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

# FullAddress – zawsze budujemy/uzupełniamy z Adres + Miasto + PSC
built_full = build_full_address(df)
if "FullAddress" not in df.columns:
    df["FullAddress"] = built_full
else:
    # jeśli w arkuszu było puste/"None"/"nan" – nadpisz tymczasowo zbudowaną wartością
    mask_blank = df["FullAddress"].astype(str).str.strip().isin(["", "None", "nan", "NaN", "NONE"]) | df["FullAddress"].isna()
    df.loc[mask_blank, "FullAddress"] = built_full[mask_blank]

# Wszystkie rekordy wymagają geokodowania
st.subheader("Podgląd danych")
# Pokaż podgląd bez kolumn lat/lon (nawet jeśli są w arkuszu)
preview_cols_order = ["lp", "Nazwa odbiorcy", "Obrót w czk", "email", "Adres", "Miasto", "PSC", "FullAddress"]
show_cols = [c for c in preview_cols_order if c in df.columns]
if show_cols:
    st.dataframe(df[show_cols].head(50), width="stretch")
else:
    st.dataframe(df.head(50), width="stretch")

st.warning("Geokodowanie wszystkich adresów (OSM/Nominatim, ~1 zapytanie/s)")
if auto_geocode or st.button("Geokoduj teraz"):
    addrs = df["FullAddress"].tolist()
    results = []
    prog = st.progress(0.0)
    for i, addr in enumerate(addrs, start=1):
        results.append((addr, geocode_one(addr)))
        prog.progress(i / len(addrs))
        time.sleep(0.05)

    mapping = {a: c for a, c in results}
    df["lat"] = df["FullAddress"].map(lambda a: mapping.get(a)[0] if mapping.get(a) else None)
    df["lon"] = df["FullAddress"].map(lambda a: mapping.get(a)[1] if mapping.get(a) else None)

    df = df.dropna(subset=["lat","lon"]).copy()
    st.session_state["geo_df"] = df.to_dict(orient="records")
    st.success("Geokodowanie zakończone – dane gotowe do mapy.")
st.rerun()

# Dane z poprawnymi współrzędnymi
geo_df = pd.DataFrame(st.session_state.get("geo_df", []))
if geo_df.empty:
    st.info("Brak poprawnych współrzędnych – uruchom geokodowanie.")
    st.stop()

# ===================== MAPA =====================
m = folium.Map(location=[geo_df["lat"].mean(), geo_df["lon"].mean()], zoom_start=7)
cluster = MarkerCluster().add_to(m)

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
    st.download_button(
        "Pobierz CSV z geokodowanymi danymi",
        data=geo_df.to_csv(index=False).encode("utf-8"),
        file_name="geokodowane_dane.csv",
        mime="text/csv",
    )

# ===================== STATUS =====================
st.info(f"✅ Na mapie: {len(geo_df)} punktów.")
