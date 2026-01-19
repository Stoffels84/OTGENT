# app.py
import re
import datetime as dt
from pathlib import Path

import pandas as pd
import streamlit as st
import openpyxl


APP_DIR = Path(__file__).parent
XLSM_PATH = APP_DIR / "schade met macro.xlsm"
GESPREKKEN_XLSX_PATH = APP_DIR / "overzicht gesprekken (aangepast).xlsx"
LOGO_PATH = APP_DIR / "logo.png"

SCHADESHEET = "BRON"

SCHADE_COLS = [
    "personeelsnr",
    "volledige naam",
    "Datum",
    "Link",
    "Locatie",
    "voertuig",
    "bus/tram",
    "type",
]

# Voor gesprekken: we lezen alle kolommen, maar we moeten minstens deze 2 kunnen vinden
GESPREK_KEY_COLS = ["personeelsnr", "volledige naam"]


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


def _find_col_case_insensitive(df: pd.DataFrame, wanted: str) -> str | None:
    w = norm(wanted)
    for c in df.columns:
        if norm(c) == w:
            return c
    # tolerant voor varianten
    if w == "personeelsnr":
        for alt in ["personeelsnr.", "personeelsnummer", "persnr", "persnr."]:
            for c in df.columns:
                if norm(c) == alt:
                    return c
    if w == "volledige naam":
        for alt in ["naam", "volledige naam.", "volledige_naam", "volledige naam "]:
            for c in df.columns:
                if norm(c) == alt:
                    return c
    return None


@st.cache_data(show_spinner=False)
def load_schade_df() -> pd.DataFrame:
    if not XLSM_PATH.exists():
        raise FileNotFoundError(f"Bestand niet gevonden: {XLSM_PATH.name} (zet dit naast app.py)")

    wb = openpyxl.load_workbook(XLSM_PATH, data_only=True, keep_vba=True)
    if SCHADESHEET not in wb.sheetnames:
        raise ValueError(f"Tabblad '{SCHADESHEET}' niet gevonden in {XLSM_PATH.name}")

    ws = wb[SCHADESHEET]

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

    idx_map = {c: find_idx(c) for c in SCHADE_COLS}

    rows = []
    for r in ws.iter_rows(min_row=2, values_only=True):
        obj = {}
        any_val = False
        for col in SCHADE_COLS:
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
    for c in SCHADE_COLS:
        if c not in df.columns:
            df[c] = None

    df["_jaar"] = df["Datum"].apply(parse_year)
    df["_search"] = (
        df["personeelsnr"].fillna("").astype(str) + " " +
        df["volledige naam"].fillna("").astype(str) + " " +
        df["voertuig"].fillna("").astype(str)
    ).str.lower()

    return df


@st.cache_data(show_spinner=False)
def load_gesprekken_df() -> pd.DataFrame:
    if not GESPREKKEN_XLSX_PATH.exists():
        raise FileNotFoundError(f"Bestand niet gevonden: {GESPREKKEN_XLSX_PATH.name} (zet dit naast app.py)")

    # Lees eerste sheet standaard
    xls = pd.ExcelFile(GESPREKKEN_XLSX_PATH)
    sheet = xls.sheet_names[0]
    df = pd.read_excel(GESPREKKEN_XLSX_PATH, sheet_name=sheet)

    # Vind de juiste kolommen (case-insensitive / tolerant)
    pn_col = _find_col_case_insensitive(df, "personeelsnr")
    nm_col = _find_col_case_insensitive(df, "volledige naam")

    if pn_col is None and nm_col is None:
        raise ValueError(
            "In 'overzicht gesprekken (aangepast).xlsx' vind ik geen kolommen voor personeelsnr/naam. "
            "Controleer de headernamen."
        )

    # Normaliseer naar vaste kolomnamen (zodat de rest van de code simpel blijft)
    if pn_col is not None and pn_col != "personeelsnr":
        df = df.rename(columns={pn_col: "personeelsnr"})
    if nm_col is not None and nm_col != "volledige naam":
        df = df.rename(columns={nm_col: "volledige naam"})

    if "personeelsnr" not in df.columns:
        df["personeelsnr"] = None
    if "volledige naam" not in df.columns:
        df["volledige naam"] = None

    df["_search"] = (
        df["personeelsnr"].fillna("").astype(str) + " " +
        df["volledige naam"].fillna("").astype(str)
    ).str.lower()

    return df


# ----------------------------
# Streamlit page setup
# ----------------------------
st.set_page_config(page_title="Analyse en rapportering OT Gent", layout="wide")

st.markdown(
    """
    <style>
      .stApp {
        background: radial-gradient(1200px 700px at 15% 5%, rgba(74,163,255,.10), transparent 60%),
                    radial-gradient(900px 600px at 90% 20%, rgba(120,80,255,.10), transparent 55%),
                    #0b0f14;
      }
      section[data-testid="stSidebar"] { display: none; }

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
      .ot-brand { display: flex; align-items: center; gap: 10px; }
      .ot-logo {
        width: 38px; height: 38px; object-fit: contain;
        border-radius: 10px; border: 1px solid rgba(255,255,255,.08);
        background: rgba(255,255,255,.03); padding: 6px;
      }
      .ot-title { font-size: 16px; font-weight: 700; color: #e6edf3; line-height: 1.1; }
      .ot-sub   { font-size: 12px; color: #9aa4b2; margin-top: 2px; }
      .block-container { padding-top: 0.5rem; }

      div[role="radiogroup"] > label {
        background: rgba(255,255,255,.02);
        border: 1px solid rgba(255,255,255,.08);
        padding: 8px 12px !important;
        border-radius: 999px !important;
        margin-right: 8px !important;
      }
      div[role="radiogroup"] > label:hover { background: rgba(255,255,255,.05); }
    </style>
    """,
    unsafe_allow_html=True,
)

# ----------------------------
# Load data
# ----------------------------
try:
    df_schade = load_schade_df()
except Exception as e:
    st.error(f"Kan schade-data niet laden: {e}")
    st.stop()

try:
    df_gesprekken = load_gesprekken_df()
except Exception as e:
    st.warning(f"Gesprekkenbestand niet geladen: {e}")
    df_gesprekken = pd.DataFrame(columns=["personeelsnr", "volledige naam", "_search"])

years = sorted([y for y in df_schade["_jaar"].dropna().unique().tolist() if y is not None], reverse=True)

# ----------------------------
# Topbar ("sidebar bovenaan")
# ----------------------------
st.markdown('<div class="ot-topbar">', unsafe_allow_html=True)

c1, c2, c3 = st.columns([2.4, 1.2, 3.2], vertical_alignment="center")

with c1:
    logo_html = f'<img class="ot-logo" src="logo.png" alt="Logo" />' if LOGO_PATH.exists() else ""
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

# Apply year filter on schade (gesprekken hebben geen Datum-filter tenzij jij dat later wil)
df_schade_view = df_schade[df_schade["_jaar"] == int(year_choice)].copy() if year_choice != "Alle" else df_schade.copy()

# ----------------------------
# Pages
# ----------------------------
if page == "Dashboard":
    st.subheader("Dashboard")

    q = st.text_input(
        "Zoek op personeelsnr of volledige naam (en voertuig in schade)",
        placeholder="Typ om te zoeken…",
    ).strip().lower()

    if not q:
        st.caption("Typ iets in het zoekveld om resultaten te zien.")
        st.stop()

    # 1) Zoek in schade (personeelsnr/naam/voertuig)
    schade_hits = df_schade_view[df_schade_view["_search"].str.contains(re.escape(q), na=False)].copy()

    # 2) Zoek in gesprekken (personeelsnr/naam)
    gesprekken_hits = df_gesprekken[df_gesprekken["_search"].str.contains(re.escape(q), na=False)].copy()

    # Toon: gesprekken eerst (compact), dan schade (details)
    st.markdown("#### Overzicht gesprekken")
    if len(gesprekken_hits) == 0:
        st.caption("Geen gesprekken gevonden voor deze zoekterm.")
    else:
        # toon alle originele kolommen behalve interne _search
        cols = [c for c in gesprekken_hits.columns if c != "_search"]
        st.dataframe(
            gesprekken_hits[cols].head(200),
            use_container_width=True,
            hide_index=True,
        )

    st.markdown("#### Schade (BRON)")
    if len(schade_hits) == 0:
        st.caption("Geen schadegevallen gevonden voor deze zoekterm.")
    else:
        show = schade_hits[SCHADE_COLS].head(500).copy()
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
