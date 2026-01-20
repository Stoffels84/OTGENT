# app.py
import re
import json
import base64
import html
import datetime as dt
from pathlib import Path

import pandas as pd
import streamlit as st
import openpyxl


# ----------------------------
# Paths / Config
# ----------------------------
APP_DIR = Path(__file__).parent

XLSM_PATH = APP_DIR / "schade met macro.xlsm"
SCHADESHEET = "BRON"

GESPREKKEN_XLSX_PATH = APP_DIR / "Overzicht gesprekken (aangepast).xlsx"
GESPREKKEN_SHEET_NAME = "gesprekken per thema"

COACHINGS_XLSX_PATH = APP_DIR / "Coachingslijst.xlsx"
COACHINGS_SHEET_VOLTOOID = "Voltooide coachings"
COACHINGS_SHEET_COACHING = "Coaching"  # ✅ nieuw tabblad

PERSONEEL_JSON_PATH = APP_DIR / "personeelsficheGB.json"
LOGO_PATH = APP_DIR / "logo.png"

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

PAGES = [
    ("dashboard", "Dashboard"),
    ("chauffeur", "Chauffeur"),
    ("voertuig", "Voertuig"),
    ("locatie", "Locatie"),
    ("coaching", "Coaching"),
    ("analyse", "Analyse"),
]


# ----------------------------
# Helpers
# ----------------------------
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

    if w in ["nummer", "personeelsnr", "personeelsnummer", "p-nr", "p_nr", "p nr", "p-nr."]:
        for alt in [
            "nr", "id", "persnr", "personeelsnr", "personeelsnummer",
            "nummer", "employeeid", "employee_id",
            "p-nr", "p nr", "p_nr", "p-nr."
        ]:
            for c in df.columns:
                if norm(c) == norm(alt):
                    return c

    if w == "datum":
        for alt in ["date", "datum gesprek", "gespreksdatum", "datum coaching", "coachingsdatum"]:
            for c in df.columns:
                if norm(c) == norm(alt):
                    return c

    if w == "info":
        for alt in [
            "informatie", "opmerking", "opmerkingen", "beschrijving", "details",
            "thema", "onderwerp", "samenvatting", "actiepunten", "resultaat", "notities", "commentaar",
            "opmerkingen (coach)", "opmerkingen chauffeur"
        ]:
            for c in df.columns:
                if norm(c) == norm(alt):
                    return c

    if w in ["volledige naam", "chauffeurnaam", "naam"]:
        for alt in [
            "chauffeurnaam", "chauffeur naam", "naam", "medewerker", "werknemer", "chauffeur",
            "volledige naam", "full name", "fullname", "displayname", "display_name"
        ]:
            for c in df.columns:
                if norm(c) == norm(alt):
                    return c

    return None


def _flatten_json_to_records(data):
    if data is None:
        return []
    if isinstance(data, list):
        return [x for x in data if isinstance(x, dict)]
    if isinstance(data, dict):
        for k in ["data", "items", "results", "records"]:
            if k in data:
                return _flatten_json_to_records(data[k])
        if data and all(isinstance(v, dict) for v in data.values()):
            out = []
            for key, val in data.items():
                rec = dict(val)
                rec["_key"] = str(key)
                out.append(rec)
            return out
        return [data]
    return []


def render_html_table(
    df: pd.DataFrame,
    col_order: list[str],
    col_widths: dict[str, str],
    max_height_px: int = 520,
) -> None:
    view = df[col_order].copy()
    for c in col_order:
        view[c] = view[c].fillna("").astype(str)

    ths = []
    for c in col_order:
        w = col_widths.get(c, "auto")
        ths.append(f'<th style="width:{w}">{html.escape(c)}</th>')
    thead = "<tr>" + "".join(ths) + "</tr>"

    trs = []
    for _, row in view.iterrows():
        tds = []
        for c in col_order:
            cell = row[c]
            safe = html.escape(cell).replace("\n", "<br/>")
            tds.append(f"<td>{safe}</td>")
        trs.append("<tr>" + "".join(tds) + "</tr>")
    tbody = "".join(trs)

    st.markdown(
        f"""
        <div class="ot-table-wrap" style="max-height:{max_height_px}px;">
          <table class="ot-table">
            <thead>{thead}</thead>
            <tbody>{tbody}</tbody>
          </table>
        </div>
        """,
        unsafe_allow_html=True,
    )


# ----------------------------
# Navigation state
# ----------------------------
def get_page(default="dashboard") -> str:
    try:
        v = st.query_params.get("page", default)
        if isinstance(v, list):
            v = v[0] if v else default
        v = str(v).strip().lower()
    except Exception:
        v = default
    valid = {pid for pid, _ in PAGES}
    return v if v in valid else default


def set_page(page_id: str) -> None:
    st.query_params["page"] = page_id
    st.rerun()


# ----------------------------
# Loaders
# ----------------------------
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

    for c in ["Datum", "Info", "Chauffeurnaam"]:
        if c not in df.columns:
            df[c] = ""

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


@st.cache_data(show_spinner=False)
def load_coaching_voltooid_df() -> pd.DataFrame:
    """Voltooide coachings tab"""
    if not COACHINGS_XLSX_PATH.exists():
        return pd.DataFrame(columns=["nummer", "Chauffeurnaam", "Datum", "Info", "_search", "_jaar"])

    xls = pd.ExcelFile(COACHINGS_XLSX_PATH)
    if COACHINGS_SHEET_VOLTOOID not in xls.sheet_names:
        raise ValueError(
            f"Tabblad '{COACHINGS_SHEET_VOLTOOID}' niet gevonden in {COACHINGS_XLSX_PATH.name}. "
            f"Gevonden tabs: {xls.sheet_names}"
        )

    df = pd.read_excel(COACHINGS_XLSX_PATH, sheet_name=COACHINGS_SHEET_VOLTOOID, dtype=str)

    num_col = _find_col(df, "nummer") or _find_col(df, "personeelsnr")
    name_col = _find_col(df, "Chauffeurnaam") or _find_col(df, "naam") or _find_col(df, "volledige naam")
    date_col = _find_col(df, "Datum")
    info_col = _find_col(df, "Info")

    if num_col is None:
        df["nummer"] = ""
        num_col = "nummer"
    if num_col != "nummer":
        df = df.rename(columns={num_col: "nummer"})

    if name_col is None:
        df["Chauffeurnaam"] = ""
    elif name_col != "Chauffeurnaam":
        df = df.rename(columns={name_col: "Chauffeurnaam"})

    if date_col is None:
        df["Datum"] = ""
    elif date_col != "Datum":
        df = df.rename(columns={date_col: "Datum"})

    if info_col is None:
        candidates = []
        for c in df.columns:
            if norm(c) in [
                "thema", "onderwerp", "opmerking", "opmerkingen", "samenvatting",
                "notities", "commentaar", "actiepunten", "resultaat"
            ]:
                candidates.append(c)
        if candidates:
            df["Info"] = df[candidates].fillna("").astype(str).agg(" | ".join, axis=1)
        else:
            df["Info"] = ""
    elif info_col != "Info":
        df = df.rename(columns={info_col: "Info"})

    df["nummer"] = df["nummer"].apply(clean_id)
    df["Chauffeurnaam"] = df["Chauffeurnaam"].apply(clean_text)
    df["Datum"] = df["Datum"].apply(clean_text)
    df["Info"] = df["Info"].apply(clean_text)

    df["_search"] = (
        df["nummer"].fillna("").astype(str) + " " +
        df["Chauffeurnaam"].fillna("").astype(str) + " " +
        df["Info"].fillna("").astype(str)
    ).str.lower()
    df["_jaar"] = df["Datum"].apply(parse_year)
    return df


@st.cache_data(show_spinner=False)
def load_coaching_tab_df() -> pd.DataFrame:
    """
    ✅ Nieuw: Coachingslijst.xlsx -> tab 'Coaching'
    Kolommen: P-nr, Volledige naam, Opmerkingen
    Normalisatie naar: nummer, Chauffeurnaam, Info (+ _search)
    """
    if not COACHINGS_XLSX_PATH.exists():
        return pd.DataFrame(columns=["nummer", "Chauffeurnaam", "Info", "_search"])

    xls = pd.ExcelFile(COACHINGS_XLSX_PATH)
    if COACHINGS_SHEET_COACHING not in xls.sheet_names:
        raise ValueError(
            f"Tabblad '{COACHINGS_SHEET_COACHING}' niet gevonden in {COACHINGS_XLSX_PATH.name}. "
            f"Gevonden tabs: {xls.sheet_names}"
        )

    df = pd.read_excel(COACHINGS_XLSX_PATH, sheet_name=COACHINGS_SHEET_COACHING, dtype=str)

    pnr_col = _find_col(df, "P-nr") or _find_col(df, "nummer") or _find_col(df, "personeelsnr")
    name_col = _find_col(df, "Volledige naam") or _find_col(df, "naam") or _find_col(df, "chauffeurnaam")
    opm_col = _find_col(df, "Opmerkingen") or _find_col(df, "Info")

    if pnr_col is None:
        df["nummer"] = ""
    else:
        if pnr_col != "nummer":
            df = df.rename(columns={pnr_col: "nummer"})

    if name_col is None:
        df["Chauffeurnaam"] = ""
    else:
        if name_col != "Chauffeurnaam":
            df = df.rename(columns={name_col: "Chauffeurnaam"})

    if opm_col is None:
        df["Info"] = ""
    else:
        if opm_col != "Info":
            df = df.rename(columns={opm_col: "Info"})

    df["nummer"] = df["nummer"].apply(clean_id)
    df["Chauffeurnaam"] = df["Chauffeurnaam"].apply(clean_text)
    df["Info"] = df["Info"].apply(clean_text)

    df["_search"] = (
        df["nummer"].fillna("").astype(str) + " " +
        df["Chauffeurnaam"].fillna("").astype(str) + " " +
        df["Info"].fillna("").astype(str)
    ).str.lower()

    return df


@st.cache_data(show_spinner=False)
def load_personeelsfiche_df() -> pd.DataFrame:
    if not PERSONEEL_JSON_PATH.exists():
        return pd.DataFrame(columns=["_search"])

    try:
        data = json.loads(PERSONEEL_JSON_PATH.read_text(encoding="utf-8"))
    except Exception:
        data = json.loads(PERSONEEL_JSON_PATH.read_text())

    records = _flatten_json_to_records(data)
    if not records:
        return pd.DataFrame(columns=["_search"])

    df = pd.DataFrame(records)

    id_col = _find_col(df, "personeelsnr") or _find_col(df, "nummer") or _find_col(df, "personeelsnummer")
    name_col = _find_col(df, "volledige naam") or _find_col(df, "naam") or _find_col(df, "chauffeurnaam")

    if id_col is None and "_key" in df.columns:
        id_col = "_key"

    if id_col and id_col != "personeelsnr":
        df = df.rename(columns={id_col: "personeelsnr"})
        id_col = "personeelsnr"
    if name_col and name_col != "naam":
        df = df.rename(columns={name_col: "naam"})
        name_col = "naam"

    if id_col is None:
        df["personeelsnr"] = ""
        id_col = "personeelsnr"
    if name_col is None:
        df["naam"] = ""
        name_col = "naam"

    df[id_col] = df[id_col].apply(clean_id)
    df[name_col] = df[name_col].apply(clean_text)

    extra_cols = []
    for c in df.columns:
        if c in ["_search", id_col, name_col]:
            continue
        if norm(c) in ["dienst", "afdeling", "team", "functie", "rol", "standplaats", "locatie"]:
            extra_cols.append(c)

    parts = [df[id_col].fillna("").astype(str), df[name_col].fillna("").astype(str)]
    for c in extra_cols[:6]:
        parts.append(df[c].fillna("").astype(str))

    df["_search"] = parts[0]
    for s in parts[1:]:
        df["_search"] = df["_search"].astype(str) + " " + s.astype(str)
    df["_search"] = df["_search"].str.lower()
    return df


# ----------------------------
# Streamlit setup / CSS
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
        padding: 14px 16px;
        border-radius: 16px;
        margin-bottom: 14px;
      }
      .ot-brand { display: flex; align-items: center; gap: 12px; }
      .ot-logo {
        width: 52px; height: 52px; object-fit: contain;
        border-radius: 12px; border: 1px solid rgba(255,255,255,.08);
        background: rgba(255,255,255,.03); padding: 8px;
      }
      .ot-title { font-size: 22px; font-weight: 800; color: #e6edf3; line-height: 1.15; }
      .ot-sub   { font-size: 14px; color: #9aa4b2; margin-top: 4px; }

      /* tab buttons (Streamlit) */
      div[data-testid="stHorizontalBlock"] .ot-tab-btn button{
        border-radius: 999px !important;
        padding: 10px 14px !important;
        border: 1px solid rgba(255,255,255,.10) !important;
        background: rgba(255,255,255,.02) !important;
        color: #cbd5e1 !important;
        font-weight: 700 !important;
        transition: all .15s ease !important;
        white-space: nowrap !important;
      }
      div[data-testid="stHorizontalBlock"] .ot-tab-btn button:hover{
        background: rgba(255,255,255,.05) !important;
        border-color: rgba(255,255,255,.18) !important;
        transform: translateY(-1px) !important;
      }
      div[data-testid="stHorizontalBlock"] .ot-tab-btn.active button{
        color: #e6edf3 !important;
        background: rgba(74,163,255,.14) !important;
        border-color: rgba(74,163,255,.35) !important;
        box-shadow: 0 0 0 1px rgba(74,163,255,.18) inset,
                    0 10px 30px rgba(74,163,255,.10);
      }

      /* ---- HTML table (wrap + sticky header) ---- */
      .ot-table-wrap{
        overflow: auto;
        border: 1px solid rgba(255,255,255,.08);
        border-radius: 12px;
        background: rgba(255,255,255,.02);
      }
      table.ot-table{
        width: 100%;
        border-collapse: collapse;
        table-layout: fixed;
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
        white-space: normal;
        overflow-wrap: anywhere;
        word-break: break-word;
      }
      table.ot-table tr:last-child td{ border-bottom: none; }
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

try:
    df_coach_voltooid = load_coaching_voltooid_df()
except Exception as e:
    st.warning(f"Voltooide coachings niet geladen: {e}")
    df_coach_voltooid = pd.DataFrame(columns=["nummer", "Chauffeurnaam", "Datum", "Info", "_search", "_jaar"])

try:
    df_coach_tab = load_coaching_tab_df()
except Exception as e:
    st.warning(f"Tabblad 'Coaching' niet geladen: {e}")
    df_coach_tab = pd.DataFrame(columns=["nummer", "Chauffeurnaam", "Info", "_search"])

df_personeel = load_personeelsfiche_df()

years_schade = df_schade["_jaar"].dropna().unique().tolist() if "_jaar" in df_schade.columns else []
years_gespr = df_gesprekken["_jaar"].dropna().unique().tolist() if "_jaar" in df_gesprekken.columns else []
years_volt = df_coach_voltooid["_jaar"].dropna().unique().tolist() if "_jaar" in df_coach_voltooid.columns else []
years = sorted({int(y) for y in (years_schade + years_gespr + years_volt) if y is not None}, reverse=True)

current_page = get_page("dashboard")


# ----------------------------
# Topbar
# ----------------------------
st.markdown('<div class="ot-topbar">', unsafe_allow_html=True)

c1, c2, c3 = st.columns([2.3, 1.2, 3.5], vertical_alignment="center")

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
    tab_cols = st.columns([1, 1, 1, 1, 1, 1, 0.95], gap="small")

    for (pid, label), col in zip(PAGES, tab_cols[:6]):
        with col:
            active = (pid == current_page)
            st.markdown(
                f'<div class="ot-tab-btn {"active" if active else ""}">',
                unsafe_allow_html=True,
            )
            if st.button(label, key=f"tab_{pid}", use_container_width=True):
                set_page(pid)
            st.markdown("</div>", unsafe_allow_html=True)

    with tab_cols[6]:
        st.markdown('<div class="ot-tab-btn">', unsafe_allow_html=True)
        if st.button("↻ Herladen", key="reload_btn", use_container_width=True):
            st.cache_data.clear()
            try:
                st.cache_resource.clear()
            except Exception:
                pass
            st.rerun()
        st.markdown("</div>", unsafe_allow_html=True)

st.markdown("</div>", unsafe_allow_html=True)

# ----------------------------
# Year filter views
# ----------------------------
df_schade_view = df_schade[df_schade["_jaar"] == int(year_choice)].copy() if year_choice != "Alle" else df_schade.copy()
df_gesprekken_view = df_gesprekken[df_gesprekken["_jaar"] == int(year_choice)].copy() if year_choice != "Alle" else df_gesprekken.copy()
df_coach_voltooid_view = df_coach_voltooid[df_coach_voltooid["_jaar"] == int(year_choice)].copy() if year_choice != "Alle" else df_coach_voltooid.copy()
# coaching-tab heeft geen datum/year filter (alleen nr/naam/opmerkingen)


# ----------------------------
# Pages
# ----------------------------
if current_page == "dashboard":
    st.subheader("Dashboard")

    q = st.text_input(
        "Zoek op personeelsnr of naam. (Schade: nr/naam/voertuig) (Gesprekken: nr/naam/info) "
        "(Voltooide coachings: nr/naam/info) (Coaching: P-nr/naam/opmerkingen) (Personeelsfiche: nr/naam)",
        placeholder="Typ om te zoeken…",
    ).strip().lower()

    if not q:
        st.caption("Typ iets in het zoekveld om resultaten te zien.")
        st.stop()

    schade_hits = df_schade_view[df_schade_view["_search"].str.contains(re.escape(q), na=False)].copy()
    gesprekken_hits = df_gesprekken_view[df_gesprekken_view["_search"].str.contains(re.escape(q), na=False)].copy()
    coach_volt_hits = df_coach_voltooid_view[df_coach_voltooid_view["_search"].str.contains(re.escape(q), na=False)].copy()

    coach_tab_hits = pd.DataFrame()
    if "_search" in df_coach_tab.columns and len(df_coach_tab) > 0:
        coach_tab_hits = df_coach_tab[df_coach_tab["_search"].str.contains(re.escape(q), na=False)].copy()

    personeels_hits = pd.DataFrame()
    if "_search" in df_personeel.columns and len(df_personeel) > 0:
        personeels_hits = df_personeel[df_personeel["_search"].str.contains(re.escape(q), na=False)].copy()

    # --- Personeelsfiche ---
    st.markdown("#### Personeelsfiche (personeelsficheGB.json)")
    if len(personeels_hits) == 0:
        st.caption("Geen personeelsfiche gevonden voor deze zoekterm.")
    else:
        summary_cols = [c for c in ["personeelsnr", "naam"] if c in personeels_hits.columns]
        if summary_cols:
            st.dataframe(personeels_hits[summary_cols].head(20), use_container_width=True, hide_index=True)

        max_show = 10
        for i, (_, row) in enumerate(personeels_hits.head(max_show).iterrows(), start=1):
            pid = row.get("personeelsnr", "")
            nm = row.get("naam", "")
            title = f"{i}. {pid} — {nm}".strip(" —")
            with st.expander(title, expanded=(i == 1)):
                rec = row.drop(labels=["_search"], errors="ignore").to_dict()
                st.json(rec)

        if len(personeels_hits) > max_show:
            st.caption(f"… en nog {len(personeels_hits) - max_show} extra matches.")

    # --- Coaching (tabblad Coaching) ---
    st.markdown("#### Coaching (Coachingslijst.xlsx → tabblad 'Coaching')")
    if len(coach_tab_hits) == 0:
        st.caption("Geen coaching-info gevonden voor deze zoekterm.")
    else:
        # ✅ gevraagde kolommen: P-nr, Volledige naam, Opmerkingen -> nu genormaliseerd:
        # nummer, Chauffeurnaam, Info
        display_ct = coach_tab_hits[["nummer", "Chauffeurnaam", "Info"]].copy()

        render_html_table(
            display_ct.head(300),
            col_order=["nummer", "Chauffeurnaam", "Info"],
            col_widths={
                "nummer": "90px",
                "Chauffeurnaam": "220px",
                "Info": "auto",
            },
            max_height_px=520,
        )

    # --- Voltooide coachings ---
    st.markdown("#### Voltooide coachings (Coachingslijst.xlsx)")
    if len(coach_volt_hits) == 0:
        st.caption("Geen voltooide coachings gevonden voor deze zoekterm.")
    else:
        display_v = coach_volt_hits[["nummer", "Chauffeurnaam", "Datum", "Info"]].copy()
        display_v["Datum"] = display_v["Datum"].apply(format_ddmmyyyy)

        render_html_table(
            display_v.head(300),
            col_order=["nummer", "Chauffeurnaam", "Datum", "Info"],
            col_widths={
                "nummer": "90px",
                "Chauffeurnaam": "180px",
                "Datum": "120px",
                "Info": "auto",
            },
            max_height_px=520,
        )

    # --- Gesprekken ---
    st.markdown("#### Overzicht gesprekken (gesprekken per thema)")
    if len(gesprekken_hits) == 0:
        st.caption("Geen gesprekken gevonden voor deze zoekterm.")
    else:
        display_g = gesprekken_hits[["nummer", "Chauffeurnaam", "Datum", "Info"]].copy()
        display_g["Datum"] = display_g["Datum"].apply(format_ddmmyyyy)

        render_html_table(
            display_g.head(300),
            col_order=["nummer", "Chauffeurnaam", "Datum", "Info"],
            col_widths={
                "nummer": "90px",
                "Chauffeurnaam": "180px",
                "Datum": "120px",
                "Info": "auto",
            },
            max_height_px=520,
        )

    # --- Schade ---
    st.markdown("#### Schade (BRON)")
    if len(schade_hits) == 0:
        st.caption("Geen schadegevallen gevonden voor deze zoekterm.")
    else:
        show = schade_hits[SCHADE_COLS].head(500).copy()
        show["Datum"] = show["Datum"].apply(format_ddmmyyyy)

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

elif current_page == "chauffeur":
    st.subheader("Chauffeur")
    st.info("Later uitwerken (filters/aggregaties op BRON).")

elif current_page == "voertuig":
    st.subheader("Voertuig")
    st.info("Later uitwerken (top voertuigen, trends, …).")

elif current_page == "locatie":
    st.subheader("Locatie")
    st.info("Later uitwerken (top locaties, …).")

elif current_page == "coaching":
    st.subheader("Coaching")
    st.info("Later uitwerken (filters/aggregaties op coaching + koppeling).")

elif current_page == "analyse":
    st.subheader("Analyse")
    st.info("Later: grafieken per maand, schade per type, …")
