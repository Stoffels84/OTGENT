# app.py
import re
import base64
import html
import datetime as dt
from pathlib import Path

import pandas as pd
import streamlit as st
import openpyxl


APP_DIR = Path(__file__).parent
XLSM_PATH = APP_DIR / "schade met macro.xlsm"
GESPREKKEN_XLSX_PATH = APP_DIR / "Overzicht gesprekken (aangepast).xlsx"
GESPREKKEN_SHEET_NAME = "gesprekken per thema"
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


def norm(s) -> str:
    return str(s).strip().lower()


def clean_id(v) -> str:
    if v is None:
        return ""
    s = str(v).strip()
    if not s:
        return ""
    if re.fullmatch(r"\d+\.0", s):
        s = s[:-2]
    return s.strip()


def clean_text(v) -> str:
    return "" if v is None else str(v).strip()


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


def format_ddmmyyyy(v) -> str:
    """Toon altijd dd-mm-jjjj; tijd/uurnotatie verdwijnt."""
    if v is None:
        return ""
    s = str(v).strip()
    if not s:
        return ""
    try:
        ts = pd.to_datetime(s, dayfirst=True, errors="coerce")
        if pd.isna(ts):
            return s
        return ts.strftime("%d-%m-%Y")
    except Exception:
        return s


def img_to_data_uri(path: Path) -> str:
    b = path.read_bytes()
    ext = path.suffix.lower().lstrip(".")
    mime = "png" if ext == "png" else ext
    return f"data:image/{mime};base64,{base64.b64encode(b).decode('utf-8')}"


def _find_col(df: pd.DataFrame, wanted: str) -> str | None:
    w = norm(wanted)
    for c in df.columns:
        if norm(c) == w:
            return c

    if w == "nummer":
        for alt in ["nr", "id", "persnr", "personeelsnr", "personeelsnummer"]:
            for c in df.columns:
                if norm(c) == alt:
                    return c

    if w == "datum":
        for alt in ["date", "datum gesprek", "gespreksdatum"]:
            for c in df.columns:
                if norm(c) == alt:
                    return c

    if w == "info":
        for alt in ["informatie", "opmerking", "opmerkingen", "beschrijving", "details"]:
            for c in df.columns:
                if norm(c) == alt:
                    return c

    if w in ["volledige naam", "chauffeurnaam"]:
        for alt in ["chauffeurnaam", "chauffeur naam", "naam", "medewerker", "werknemer", "chauffeur"]:
            for c in df.columns:
                if norm(c) == alt:
                    return c

    return None


def render_html_table(
    df: pd.DataFrame,
    col_order: list[str],
    col_widths: dict[str, str],
    max_height_px: int = 520,
) -> None:
    """
    Render een HTML-tabel met echte tekstterugloop + automatische rijhoogte.
    Dit omzeilt de Streamlit 'virtualized grid' beperking.
    """
    # Alleen gevraagde kolommen, en veilig casten naar string
    view = df[col_order].copy()
    for c in col_order:
        view[c] = view[c].fillna("").astype(str)

    # Header
    ths = []
    for c in col_order:
        w = col_widths.get(c, "auto")
        ths.append(f'<th style="width:{w}">{html.escape(c)}</th>')
    thead = "<tr>" + "".join(ths) + "</tr>"

    # Body
    trs = []
    for _, row in view.iterrows():
        tds = []
        for c in col_order:
            cell = row[c]
            # Convert linebreaks netjes
            safe = html.escape(cell).replace("\n", "<br/>")
            tds.append(f"<td>{safe}</td>")
        trs.append("<tr>" + "".join(tds) + "</tr>")
    tbody = "".join(trs)

    table_html = f"""
    <div class="ot-table-wrap" style="max-height:{max_height_px}px;">
      <table class="ot-table">
        <thead>{thead}</thead>
        <tbody>{tbody}</tbody>
      </table>
    </div>
    """

    st.markdown(table_html, unsafe_allow_html=True)


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

    df["personeelsnr"] = df["personeelsnr"].apply(clean_id)
    df["volledige naam"] = df["volledige naam"].apply(clean_text)
    df["voertuig"] = df["voertuig"].apply(clean_text)

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

    xls = pd.ExcelFile(GESPREKKEN_XLSX_PATH)
    if GESPREKKEN_SHEET_NAME not in xls.sheet_names:
        raise ValueError(
            f"Tabblad '{GESPREKKEN_SHEET_NAME}' niet gevonden in {GESPREKKEN_XLSX_PATH.name}. "
            f"Gevonden tabs: {xls.sheet_names}"
        )

    df = pd.read_excel(GESPREKKEN_XLSX_PATH, sheet_name=GESPREKKEN_SHEET_NAME, dtype=str)

    num_col = _find_col(df, "nummer")
    date_col = _find_col(df, "Datum")
    info_col = _find_col(df, "Info")
    name_col = _find_col(df, "Chauffeurnaam")

    if num_col is None:
        raise ValueError("Kolom 'nummer' (personeelsnr) niet gevonden in 'gesprekken per thema'.")

    if num_col != "nummer":
        df = df.rename(columns={num_col: "nummer"})
    if date_col and date_col != "Datum":
        df = df.rename(columns={date_col: "Datum"})
    if info_col and info_col != "Info":
        df = df.rename(columns={info_col: "Info"})
    if name_col and name_col != "Chauffeurnaam":
        df = df.rename(columns={name_col: "Chauffeurnaam"})

    if "Datum" not in df.columns:
        df["Datum"] = ""
    if "Info" not in df.columns:
        df["Info"] = ""
    if "Chauffeurnaam" not in df.columns:
        df["Chauffeurnaam"] = ""

    df["nummer"] = df["nummer"].apply(clean_id)
    df["Datum"] = df["Datum"].apply(clean_text)
    df["Info"] = df["Info"].apply(clean_text)
    df["Chauffeurnaam"] = df["Chauffeurnaam"].apply(clean_text)

    df["_search"] = (
        df["nummer"].fillna("").astype(str) + " " +
        df["Chauffeurnaam"].fillna("").astype(str) + " " +
        df["Info"].fillna("").astype(str)
    ).str.lower()

    df["_jaar"] = df["Datum"].apply(parse_year)

    return df


# ----------------------------
# Streamlit setup
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

      /* ---- HTML tabel (gesprekken) ---- */
      .ot-table-wrap{
        overflow: auto;
        border: 1px solid rgba(255,255,255,.08);
        border-radius: 12px;
        background: rgba(255,255,255,.02);
      }
      table.ot-table{
        width: 100%;
        border-collapse: collapse;
        table-layout: fixed; /* belangrijk voor vaste kolombreedtes */
      }
      table.ot-table thead th{
        position: sticky;
        top: 0;
        background: rgba(15, 22, 33, .92);
        color: #cbd5e1;
        text-align: left;
        font-weight: 600;
        font-size: 13px;
        padding: 10px 10px;
        border-bottom: 1px solid rgba(255,255,255,.08);
      }
      table.ot-table td{
        color: #e6edf3;
        font-size: 13px;
        padding: 10px 10px;
        vertical-align: top;
        border-bottom: 1px solid rgba(255,255,255,.06);
        white-space: normal;         /* wrap */
        overflow-wrap: anywhere;     /* wrap */
        word-break: break-word;      /* wrap */
      }
      table.ot-table tr:last-child td{
        border-bottom: none;
      }
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
    df_gesprekken = pd.DataFrame(columns=["nummer", "Chauffeurnaam", "Datum", "Info", "_search", "_jaar"])

years_schade = df_schade["_jaar"].dropna().unique().tolist() if "_jaar" in df_schade.columns else []
years_gespr = df_gesprekken["_jaar"].dropna().unique().tolist() if "_jaar" in df_gesprekken.columns else []
years = sorted({int(y) for y in (years_schade + years_gespr) if y is not None}, reverse=True)

# ----------------------------
# Topbar
# ----------------------------
st.markdown('<div class="ot-topbar">', unsafe_allow_html=True)

c1, c2, c3 = st.columns([2.4, 1.2, 3.2])

with c1:
    logo_html = (
        f'<img class="ot-logo" src="{img_to_data_uri(LOGO_PATH)}" alt="Logo" />'
        if LOGO_PATH.exists()
        else ""
    )
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

df_schade_view = (
    df_schade[df_schade["_jaar"] == int(year_choice)].copy()
    if year_choice != "Alle"
    else df_schade.copy()
)

df_gesprekken_view = (
    df_gesprekken[df_gesprekken["_jaar"] == int(year_choice)].copy()
    if year_choice != "Alle"
    else df_gesprekken.copy()
)

# ----------------------------
# Pages
# ----------------------------
if page == "Dashboard":
    st.subheader("Dashboard")

    q = st.text_input(
        "Zoek op personeelsnr of naam (en voertuig in schade). In gesprekken zoekt hij op nummer/chauffeurnaam/info.",
        placeholder="Typ om te zoeken…",
    ).strip().lower()

    if not q:
        st.caption("Typ iets in het zoekveld om resultaten te zien.")
        st.stop()

    schade_hits = df_schade_view[df_schade_view["_search"].str.contains(re.escape(q), na=False)].copy()
    gesprekken_hits = df_gesprekken_view[df_gesprekken_view["_search"].str.contains(re.escape(q), na=False)].copy()

    st.markdown("#### Overzicht gesprekken (gesprekken per thema)")
    if len(gesprekken_hits) == 0:
        st.caption("Geen gesprekken gevonden voor deze zoekterm.")
    else:
        cols = ["nummer", "Chauffeurnaam", "Datum", "Info"]
        display_gesprekken = gesprekken_hits[cols].copy()
        display_gesprekken["Datum"] = display_gesprekken["Datum"].apply(format_ddmmyyyy)

        # ✅ HTML-tabel: eerste 3 kolommen smal, Info krijgt de rest + wrap + volledige rijhoogte
        render_html_table(
            display_gesprekken.head(300),
            col_order=["nummer", "Chauffeurnaam", "Datum", "Info"],
            col_widths={
                "nummer": "90px",
                "Chauffeurnaam": "180px",
                "Datum": "120px",
                "Info": "auto",
            },
            max_height_px=520,
        )

    st.markdown("#### Schade (BRON)")
    if len(schade_hits) == 0:
        st.caption("Geen schadegevallen gevonden voor deze zoekterm.")
    else:
        show = schade_hits[SCHADE_COLS].head(500).copy()
        show["Datum"] = show["Datum"].apply(format_ddmmyyyy)

        # Schade kan gerust via dataframe (meestal minder lange tekst in 1 kolom)
        st.dataframe(
            show,
            use_container_width=True,
            hide_index=True,
            column_config={
                "personeelsnr": st.column_config.TextColumn("personeelsnr", width="small"),
                "volledige naam": st.column_config.TextColumn("volledige naam", width="medium"),
                "Datum": st.column_config.TextColumn("Datum", width="small"),
                "Link": st.column_config.LinkColumn("Link", width="small"),
                "Locatie": st.column_config.TextColumn("Locatie", width="medium"),
                "voertuig": st.column_config.TextColumn("voertuig", width="medium"),
                "bus/tram": st.column_config.TextColumn("bus/tram", width="small"),
                "type": st.column_config.TextColumn("type", width="small"),
            },
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
