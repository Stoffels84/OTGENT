# app.py
import re
import datetime as dt
from pathlib import Path

import pandas as pd
import streamlit as st
import openpyxl


APP_DIR = Path(__file__).parent
XLSM_PATH = APP_DIR / "schade met macro.xlsm"
LOGO_PATH = APP_DIR / "logo.png"
SHEET_NAME = "BRON"

REQUIRED_COLS = [
    "personeelsnr",
    "volledige naam",
    "Datum",
    "Link",
    "Locatie",
    "voertuig",
    "bus/tram",
    "type",
]


def norm(s) -> str:
    return str(s).strip().lower()


def parse_year(v) -> int | None:
    if v is None:
        return None
    if isinstance(v, (dt.date, dt.datetime)):
        return v.year

    s = str(v).strip()
    if not s:
        return None

    m = re.match(r"^(\d{1,2})[/-](\d{1,2})[/-](\d{4})", s)
    if m:
        return int(m.group(3))

    m2 = re.match(r"^(\d{4})[/-](\d{1,2})[/-](\d{1,2})", s)
    if m2:
        return int(m2.group(1))

    try:
        return dt.datetime.fromisoformat(s).year
    except Exception:
        return None


@st.cache_data(show_spinner=False)
def load_bron_df() -> pd.DataFrame:
    if not XLSM_PATH.exists():
        raise FileNotFoundError(
            f"Bestand niet gevonden: {XLSM_PATH.name} (zet dit naast app.py)"
        )

    wb = openpyxl.load_workbook(XLSM_PATH, data_only=True, keep_vba=True)
    if SHEET_NAME not in wb.sheetnames:
        raise ValueError(f"Tabblad '{SHEET_NAME}' niet gevonden in {XLSM_PATH.name}")

    ws = wb[SHEET_NAME]

    header = [c.value for c in ws[1]]
    header_map = {norm(h): idx for idx, h in enumerate(header)}

    def find_idx(col: str) -> int | None:
        key = norm(col)
        if key in header_map:
            return header_map[key]
        if col == "bus/tram":
            for alt in ["bus/ tram", "bus / tram", "bus - tram"]:
                if alt in header_map:
                    return header_map[alt]
        if col == "volledige naam":
            for alt in ["naam", "volledige naam.", "volledige naam "]:
                if alt in header_map:
                    return header_map[alt]
        return None

    idx_map = {c: find_idx(c) for c in REQUIRED_COLS}

    rows = []
    for r in ws.iter_rows(min_row=2, values_only=True):
        obj = {}
        any_val = False
        for col in REQUIRED_COLS:
            i = idx_map.get(col)
            val = r[i] if (i is not None and i < len(r)) else None

            if val is not None and str(val).strip() != "":
                any_val = True

            if col == "Datum" and isinstance(val, (dt.date, dt.datetime)):
                val = val.isoformat()

            obj[col] = val

        if any_val:
            rows.append(obj)

    df = pd.DataFrame(rows)
    for c in REQUIRED_COLS:
        if c not in df.columns:
            df[c] = None

    df["_jaar"] = df["Datum"].apply(parse_year)
    df["_search"] = (
        df["personeelsnr"].fillna("").astype(str) + " " +
        df["volledige naam"].fillna("").astype(str) + " " +
        df["voertuig"].fillna("").astype(str)
    ).str.lower()

    return df


# ----------------------------
# Streamlit page setup
# ----------------------------
st.set_page_config(page_title="Analyse en rapportering OT Gent", layout="wide")

# Make a "top sidebar" (topbar) via CSS and a container
st.markdown(
    """
    <style>
      /* overall dark look */
      .stApp {
        background: radial-gradient(1200px 700px at 15% 5%, rgba(74,163,255,.10), transparent 60%),
                    radial-gradient(900px 600px at 90% 20%, rgba(120,80,255,.10), transparent 55%),
                    #0b0f14;
      }

      /* hide default sidebar if you don't use it */
      section[data-testid="stSidebar"] { display: none; }

      /* Topbar container styling */
      .ot-topbar {
        position: sticky;
        top: 0;
        z-index: 999;
        background: rgba(15, 22, 33, .86);
        backdrop-filter: blur(10px);
        border-bottom: 1px solid rgba(255,255,255,.08);
        padding: 10px 14px;
        border-radius: 14px;
        margin-bottom: 14px;
      }

      .ot-brand {
        display: flex;
        align-items: center;
        gap: 10px;
      }
      .ot-logo {
        width: 38px;
        height: 38px;
        object-fit: contain;
        border-radius: 10px;
        border: 1px solid rgba(255,255,255,.08);
        background: rgba(255,255,255,.03);
        padding: 6px;
      }
      .ot-title { font-size: 16px; font-weight: 700; color: #e6edf3; line-height: 1.1; }
      .ot-sub   { font-size: 12px; color: #9aa4b2; margin-top: 2px; }

      /* reduce padding around main block */
      .block-container { padding-top: 0.5rem; }

      /* make radio look like pills */
      div[role="radiogroup"] > label {
        background: rgba(255,255,255,.02);
        border: 1px solid rgba(255,255,255,.08);
        padding: 8px 12px !important;
        border-radius: 999px !important;
        margin-right: 8px !important;
      }
      div[role="radiogroup"] > label:hover {
        background: rgba(255,255,255,.05);
      }

      /* remove extra label spacing on inputs */
      .stRadio > label, .stSelectbox > label, .stTextInput > label { font-weight: 600; }
    </style>
    """,
    unsafe_allow_html=True,
)

# ----------------------------
# Load data
# ----------------------------
try:
    df = load_bron_df()
except Exception as e:
    st.error(f"Kan data niet laden: {e}")
    st.stop()

years = sorted([y for y in df["_jaar"].dropna().unique().tolist() if y is not None], reverse=True)

# ----------------------------
# Topbar UI (acts as "sidebar bovenaan")
# ----------------------------
st.markdown('<div class="ot-topbar">', unsafe_allow_html=True)

c1, c2, c3 = st.columns([2.4, 1.2, 3.2], vertical_alignment="center")

with c1:
    logo_html = ""
    if LOGO_PATH.exists():
        logo_html = f'<img class="ot-logo" src="logo.png" alt="Logo" />'
    st.markdown(
        f"""
        <div class="ot-brand">
          {logo_html}
          <div>
            <div class="ot-title">Analyse en rapportering OT Gent</div>
            <div class="ot-sub">schade</div>
          </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

with c2:
    year_choice = st.selectbox("Jaar", ["Alle"] + [str(y) for y in years], index=0)

with c3:
    page = st.radio(
        "Menu",
        ["Dashboard", "Chauffeur", "Voertuig", "Locatie", "Coaching", "Analyse"],
        horizontal=True,
        label_visibility="collapsed",
    )

st.markdown("</div>", unsafe_allow_html=True)

# Apply year filter
if year_choice != "Alle":
    df_view = df[df["_jaar"] == int(year_choice)].copy()
else:
    df_view = df.copy()

# ----------------------------
# Pages
# ----------------------------
if page == "Dashboard":
    st.subheader("Dashboard")

    q = st.text_input(
        "Zoek op personeelsnr, volledige naam of voertuig",
        placeholder="Typ om te zoeken…",
    ).strip().lower()

    if q:
        hits = df_view[df_view["_search"].str.contains(re.escape(q), na=False)].copy()
    else:
        hits = df_view.copy()

    st.caption(f"Records: {len(hits)} (jaarfilter: {year_choice})")

    show = hits[REQUIRED_COLS].head(500).copy()
    st.data_editor(
        show,
        use_container_width=True,
        hide_index=True,
        column_config={"Link": st.column_config.LinkColumn("Link")},
        disabled=True,
    )

elif page == "Chauffeur":
    st.subheader("Chauffeur")
    st.info("Later uitwerken (filters/aggregaties op BRON).")

elif page == "Voertuig":
    st.subheader("Voertuig")
    st.info("Later uitwerken (top voertuigen, trends, …).")

elif page == "Locatie":
    st.subheader("Locatie")
    st.info("Later uitwerken (top locaties, …).")

elif page == "Coaching":
    st.subheader("Coaching")
    st.info("Later koppeling met Coachingslijst.xlsx / gesprekken.")

elif page == "Analyse":
    st.subheader("Analyse")
    st.info("Later: grafieken per maand, schade per type, …")
