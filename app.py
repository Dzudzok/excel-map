import io, time, math, hashlib, requests, numpy as np
import pandas as pd
import streamlit as st
import streamlit.components.v1 as components
import folium
from folium.plugins import MarkerCluster
from geopy.geocoders import Nominatim
from geopy.extra.rate_limiter import RateLimiter

# -------------------- KONFIG --------------------
GOOGLE_SHEETS_CSV = "https://docs.google.com/spreadsheets/d/e/2PACX-1vSijSBg7JqZkg4T8aY56FEhox0pqw5huE7oWRmSbaB25LJj9nFyo76JLPKSXHZecd4nZEyu92jesaor/pub?gid=0&single=true&output=csv"
SPREADSHEET_ID  = st.secrets.get("SPREADSHEET_ID", "")
WORKSHEET_NAME  = st.secrets.get("WORKSHEET_NAME", "Arkusz1")

st.set_page_config(page_title="Mapa klientów z Excela / Google Sheets", layout="wide")
st.title("📍 Mapa klientów z Excela / Google Sheets (online, free)")
st.caption("Wymagane: **Adres**, **Miasto**, **PSC**. Opcjonalne: **Nazwa odbiorcy**, **Obrót w czk**, **email**, **lat**, **lon**.")

REQ_ADDR_COLS = ["Adres", "Miasto", "PSC"]

# -------------------- HELPERY --------------------
def to_float_or_none(x):
    try:
        if x is None or (isinstance(x, float) and math.isnan(x)): return None
        s = str(x).strip().replace("\xa0"," ").replace(" ","").replace(",", ".")
        return float(s)
    except Exception:
        return None

def fmt_czk(x):
    v = to_float_or_none(x)
    if v is None: return ""
    return f"Obrót: {v:,.2f}".replace(",", " ").replace(".", ",") + " CZK"

def norm_col(s: pd.Series) -> pd.Series:
    return (s.fillna("").astype(str)
             .str.replace(r"^\s*0\s*$","", regex=True)
             .str.replace(r"\s+"," ", regex=True)
             .str.strip())

def build_full_address(df: pd.DataFrame) -> pd.Series:
    a = norm_col(df["Adres"]); m = norm_col(df["Miasto"])
    p = norm_col(df["PSC"]).str.replace(" ","")
    return (a + ", " + m + " " + p).str.strip(", ").str.strip()

# ----------- ODCZYT: preferuj API (świeże), fallback: published CSV -----------
@st.cache_data(show_spinner=False, ttl=10)
def load_google(url_csv: str, spreadsheet_id: str, worksheet_name: str) -> pd.DataFrame:
    try:
        if spreadsheet_id and "gcp_service_account" in st.secrets:
            import gspread
            from google.oauth2.service_account import Credentials
            creds = Credentials.from_service_account_info(
                st.secrets["gcp_service_account"],
                scopes=["https://www.googleapis.com/auth/spreadsheets.readonly"]
            )
            sh = gspread.authorize(creds).open_by_key(spreadsheet_id)
            ws = sh.worksheet(worksheet_name)
            df = pd.DataFrame(ws.get_all_records(numericise_ignore=['all']))
            if not df.empty:
                return df
    except Exception as e:
        st.warning(f"Odczyt API nieudany ({e}). Próbuję przez publikowany CSV…")

    # fallback: CSV (może mieć opóźnienie)
    r = requests.get(url_csv, timeout=30, headers={"User-Agent":"Mozilla/5.0"})
    r.raise_for_status()
    b = r.content
    for enc in ("utf-8-sig","utf-8","cp1250","iso-8859-2"):
        try: return pd.read_csv(io.StringIO(b.decode(enc)))
        except Exception: pass
    return pd.read_excel(io.BytesIO(b))

# ----------- ZAPIS: API do tej samej zakładki -----------
def save_to_google_sheet(df_to_save: pd.DataFrame):
    if not SPREADSHEET_ID:
        st.error("Brakuje SPREADSHEET_ID w Secrets."); return False
    try:
        import gspread
        from gspread_dataframe import set_with_dataframe
        from google.oauth2.service_account import Credentials
        creds = Credentials.from_service_account_info(
            st.secrets["gcp_service_account"],
            scopes=["https://www.googleapis.com/auth/spreadsheets"]
        )
        gc = gspread.authorize(creds)
        sh = gc.open_by_key(SPREADSHEET_ID)
        try: ws = sh.worksheet(WORKSHEET_NAME)
        except gspread.WorksheetNotFound:
            ws = sh.add_worksheet(title=WORKSHEET_NAME, rows="2000", cols="50")
        set_with_dataframe(ws, df_to_save, include_index=False, include_column_header=True)
        return True
    except Exception as e:
        st.error(f"Nie udało się zapisać do Google Sheets: {e}")
        return False

# ----------- Geokoder: timeout + retry + CZ/PL -----------
@st.cache_data(show_spinner=False)
def geocode_one(address: str):
    """(lat, lon) w CZ/PL albo None – solidne timeouty i retry."""
    geolocator = Nominatim(user_agent="mroauto-excel-map (contact: info@mroauto.cz)", timeout=10)
    geocode = RateLimiter(
        geolocator.geocode,
        min_delay_seconds=1.1,
        error_wait_seconds=3.0,
        max_retries=2,
        swallow_exceptions=True,
    )
    params = dict(
        addressdetails=False, language="cs", country_codes="cz,pl",
        viewbox=((12.0, 51.6), (24.3, 48.0)), bounded=True, exactly_one=True
    )
    loc = geocode(address, **params) or geocode(address, **{**params, "language":"pl"})
    if not loc: return None
    lat, lon = float(loc.latitude), float(loc.longitude)
    if not (48.0 <= lat <= 55.0 and 12.0 <= lon <= 25.0): return None
    return (lat, lon)

# -------------------- UI --------------------
c1, c2, c3, c4 = st.columns([1,1,1,1])
with c1: uploaded = st.file_uploader("Wgraj plik (Excel/CSV)", type=["xlsx","csv"])
with c2: go_btn = st.button("⬇️ Pobierz z Google (stały link)")
with c3: auto_geocode = st.checkbox("Auto-geokoduj brakujące/błędne", value=True)
with c4: reset = st.button("♻️ Reset (wyczyść)")

if reset:
    st.cache_data.clear()
    for k in ["_use_google","geo_df","geocoded_done"]:
        st.session_state.pop(k, None)
    st.rerun()

if go_btn:
    st.cache_data.clear()
    st.session_state["_use_google"] = True
    for k in ["geo_df","geocoded_done"]:
        st.session_state.pop(k, None)
    st.rerun()

use_google = st.session_state.get("_use_google", True)

# -------------------- DANE --------------------
if uploaded is not None and not use_google:
    df = pd.read_csv(uploaded) if uploaded.name.lower().endswith(".csv") else pd.read_excel(uploaded)
else:
    df = load_google(GOOGLE_SHEETS_CSV, SPREADSHEET_ID, WORKSHEET_NAME)

if df is None or df.empty:
    st.info("Brak danych – wrzuć plik albo użyj przycisku Google."); st.stop()

st.subheader("Podgląd danych")
st.dataframe(df.head(50), width="stretch")

# -------------------- NORMALIZACJA --------------------
# wymuś obecność kolumn adresowych
missing_req = [c for c in REQ_ADDR_COLS if c not in df.columns]
if missing_req:
    st.error(f"Brakuje kolumn: {', '.join(missing_req)} (lub dodaj lat/lon)."); st.stop()

# FullAddress
if "FullAddress" not in df.columns:
    df["FullAddress"] = build_full_address(df)

# lat/lon → kropki, numeric, dtype float64
if "lat" not in df.columns: df["lat"] = np.nan
if "lon" not in df.columns: df["lon"] = np.nan
df["lat"] = pd.to_numeric(df["lat"].astype(str).str.replace(",", "."), errors="coerce").astype("float64")
df["lon"] = pd.to_numeric(df["lon"].astype(str).str.replace(",", "."), errors="coerce").astype("float64")

def valid_coord(lat, lon):
    try: lat=float(lat); lon=float(lon)
    except Exception: return False
    return (48.0 <= lat <= 55.0) and (12.0 <= lon <= 25.0)

# wiersze do poprawy (brak lub poza CZ/PL)
to_fix_idx = df.index[(df["lat"].isna() | df["lon"].isna()) |
                      (~df[["lat","lon"]].apply(lambda r: valid_coord(r["lat"], r["lon"]), axis=1))]

# -------------------- JEDEN BLOK GEOKODOWANIA --------------------
if len(to_fix_idx) > 0 and not st.session_state.get("geocoded_done", False):
    st.warning(f"Do uzupełnienia/błędne współrzędne: {len(to_fix_idx)}. (OSM/Nominatim, ~1 zapytanie/s)")
    trigger = auto_geocode or st.button("📍 Geokoduj teraz")
    if trigger:
        addrs = df.loc[to_fix_idx, "FullAddress"].tolist()
        results, prog = [], st.progress(0.0)
        for i, addr in enumerate(addrs, start=1):
            results.append((addr, geocode_one(addr)))
            prog.progress(i/len(addrs)); time.sleep(0.05)

        mapping = {a:c for a,c in results}

        def pick_lat(a): 
            c = mapping.get(a); return c[0] if c else np.nan
        def pick_lon(a): 
            c = mapping.get(a); return c[1] if c else np.nan

        df.loc[to_fix_idx, "lat"] = df.loc[to_fix_idx,"FullAddress"].map(pick_lat).astype("float64")
        df.loc[to_fix_idx, "lon"] = df.loc[to_fix_idx,"FullAddress"].map(pick_lon).astype("float64")

        # finalna walidacja
        bad = ~df.apply(lambda r: valid_coord(r["lat"], r["lon"]), axis=1)
        df.loc[bad, ["lat","lon"]] = np.nan

        # zapamiętaj do sesji i zapisz do Sheets
        st.session_state["geocoded_done"] = True
        st.session_state["geo_df"] = df.dropna(subset=["lat","lon"]).to_dict(orient="records")

        if save_to_google_sheet(df):
            st.success(f"Zapisano współrzędne do Google Sheets (zakładka: {WORKSHEET_NAME}).")
        else:
            st.info("Uzupełniono lokalnie (zapis do Sheets nieudany).")
        st.rerun()

# jeśli mamy gotowe punkty – rysujemy; jak nie, prosimy o geokodowanie
geo_df = pd.DataFrame(st.session_state.get("geo_df")) if "geo_df" in st.session_state else df.dropna(subset=["lat","lon"]).copy()
if geo_df.empty:
    st.info("Brak poprawnych współrzędnych. Użyj geokodowania (checkbox/przycisk) albo uzupełnij lat/lon w arkuszu.")
    st.stop()

# -------------------- MAPA (stabilny render przez HTML) --------------------
m = folium.Map(location=[geo_df["lat"].mean(), geo_df["lon"].mean()], zoom_start=8)
cluster = MarkerCluster().add_to(m)

def val(col, row, default=""):
    return row[col] if col in geo_df.columns and pd.notna(row[col]) else default

for _, r in geo_df.iterrows():
    popup_html = f"""
    <div style="font-size:14px">
      <b>{val('Nazwa odbiorcy', r)}</b><br>
      {fmt_czk(val('Obrót w czk', r))}<br>
      {('Email: ' + val('email', r)) if val('email', r) else ''}<br>
      {('Adres: ' + val('FullAddress', r)) if 'FullAddress' in geo_df.columns else ''}
    </div>
    """
    folium.Marker([r["lat"], r["lon"]],
                  tooltip=val('Nazwa odbiorcy', r) or "Klient",
                  popup=folium.Popup(popup_html, max_width=350)).add_to(cluster)

components.html(m.get_root().render(), height=700, scrolling=False)

# -------------------- EKSPORT / ZAPIS --------------------
with st.expander("💾 Eksport / Zapis"):
    st.download_button(
        "Pobierz CSV z lat/lon (to, co rysuje mapa)",
        data=geo_df.to_csv(index=False).encode("utf-8"),
        file_name="geokodowane_dane.csv",
        mime="text/csv"
    )
    if st.button("📤 Zapisz do Google Sheets teraz (nadpisze zakładkę)"):
        if save_to_google_sheet(df):
            st.success("Zapisano do Google Sheets ✅")
            st.rerun()
