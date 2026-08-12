"""
Sindh High Court — Cause List Analytics Dashboard
A Streamlit application for exploring daily cause-list data:
case volumes, judge workloads, court sections, advocates, and case search.
"""

import math
import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st
from datetime import datetime
from pathlib import Path

APP_BUILD = "2026-08-12-path-fix-v1"

st.set_page_config(
    page_title="Sindh High Court — Cause List Analytics",
    page_icon="⚖️",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ═══════════════════════════════════════════════════════════════
# THEME
# ═══════════════════════════════════════════════════════════════
COLORS = {
    "dark_green": "#0F2E22",
    "green_mid": "#16523C",
    "green_acc": "#1FA463",
    "green_light": "#4ED191",
    "mint": "#B7F0CE",
    "gold": "#2FBF71",
    "navy": "#123524",
    "text_dark": "#0E241B",
    "text_mid": "#3F6656",
    "text_light": "#8FAFA0",
    "teal": "#17A589",
    "red": "#C0392B",
    "orange": "#E67E22",
    "purple": "#7B68EE",
}
PALETTE = [
    COLORS["dark_green"], COLORS["green_acc"], COLORS["teal"], COLORS["green_light"],
    COLORS["gold"], COLORS["orange"], COLORS["purple"], "#0B8457",
]
DONUT_PALETTE = [
    COLORS["dark_green"], COLORS["gold"], COLORS["teal"], COLORS["red"],
    COLORS["orange"], COLORS["purple"], COLORS["text_light"],
]

FONT = dict(family="Inter", color=COLORS["text_dark"], size=11)
GRID_STYLE = dict(gridcolor="#EAF7F0", linecolor="#D5EDDF")

# Fixed path definition so it looks in the correct folder structure
CAUSE_LIST_FOLDER = Path(__file__).parent / "cause_lists"


# ═══════════════════════════════════════════════════════════════
# ICONS
# ═══════════════════════════════════════════════════════════════
_ICON_PATHS = {
    "cases": '<path d="M9 2h6a2 2 0 0 1 2 2v16a2 2 0 0 1-2 2H9a2 2 0 0 1-2-2V4a2 2 0 0 1 2-2Z"/><path d="M9 8h6M9 12h6M9 16h4"/>',
    "judge": '<path d="M12 3 3 7v2h18V7z"/><path d="M4 21h16M6 10v9M10 10v9M14 10v9M18 10v9"/>',
    "section": '<path d="M3 7h4l2-2h6l2 2h4v12H3z"/>',
    "load": '<path d="M13 2 3 14h7l-1 8 11-14h-7z"/>',
    "category": '<rect x="3" y="3" width="7" height="7" rx="1"/><rect x="14" y="3" width="7" height="7" rx="1"/><rect x="3" y="14" width="7" height="7" rx="1"/><rect x="14" y="14" width="7" height="7" rx="1"/>',
    "lawyer": '<circle cx="12" cy="7" r="4"/><path d="M5.5 21a6.5 6.5 0 0 1 13 0"/>',
    "calendar": '<rect x="3" y="4" width="18" height="18" rx="2"/><path d="M16 2v4M8 2v4M3 10h18"/>',
    "folder": '<path d="M3 7a2 2 0 0 1 2-2h4l2 2h8a2 2 0 0 1 2 2v8a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2Z"/>',
}


def svg_icon(name: str, color: str = "#1FA463", size: int = 18) -> str:
    path = _ICON_PATHS.get(name, "")
    return (
        f'<svg xmlns="http://www.w3.org/2000/svg" width="{size}" height="{size}" '
        f'viewBox="0 0 24 24" fill="none" stroke="{color}" stroke-width="2" '
        f'stroke-linecap="round" stroke-linejoin="round">{path}</svg>'
    )


# ═══════════════════════════════════════════════════════════════
# STYLES
# ═══════════════════════════════════════════════════════════════
def inject_css() -> None:
    st.markdown(
        """
        <style>
        @import url('https://fonts.googleapis.com/css2?family=Playfair+Display:wght@500;600;700&family=Inter:wght@300;400;500;600;700&display=swap');

        html, body, [class*="css"] {
            font-family: 'Inter', sans-serif !important;
            background: linear-gradient(160deg, #EEF1F0 0%, #E4E9E6 100%) !important;
            color: #0E241B !important;
        }
        .main { background: linear-gradient(160deg, #EEF1F0 0%, #E4E9E6 100%) !important; }
        .main .block-container { padding: 0 !important; max-width: 100% !important; }

        [data-testid="stSidebar"] {
            background: transparent !important;
            border-right: none !important;
            padding-top: 0 !important;
        }
        [data-testid="stSidebar"] > div:first-child {
            height: 100vh;
            overflow-y: auto;
            overflow-x: hidden;
            background: linear-gradient(180deg, #E9F2ED 0%, #DCEBE3 100%);
            border-right: 2px solid #1FA463;
            border-radius: 0 24px 24px 0;
            padding-top: 0 !important;
        }
        section[data-testid="stSidebar"] div[data-testid="stVerticalBlock"] { gap: 0 !important; }
        [data-testid="stSidebarContent"] { padding-left: 0 !important; padding-right: 0 !important; }
        [data-testid="stSidebarUserContent"] > div { padding-left: 1.1rem !important; padding-right: 1.1rem !important; }
        [data-testid="stSidebarUserContent"] > div:has(.sb-logo) { padding-left: 0 !important; padding-right: 0 !important; }
        [data-testid="stSidebarHeader"] {
            position: absolute !important;
            top: 8px !important;
            right: 12px !important;
            height: auto !important;
            min-height: 0 !important;
            width: auto !important;
            margin: 0 !important;
            padding: 0 !important;
            background: transparent !important;
            z-index: 1000 !important;
        }
        [data-testid="stSidebarCollapseButton"] button {
            background: rgba(255,255,255,0.14) !important;
            border-radius: 8px !important;
            padding: 4px !important;
        }
        [data-testid="stSidebarCollapseButton"] svg { fill: #FFFFFF !important; }

        [data-testid="stSidebarUserContent"] { padding-top: 0 !important; }
        [data-testid="stSidebarUserContent"] > div:first-child {
            margin-top: 0 !important;
            padding-top: 0 !important;
        }
        [data-testid="stSidebar"] * { color: #0E241B !important; }
        [data-testid="stSidebar"] .stSelectbox label,
        [data-testid="stSidebar"] .stMultiSelect label,
        [data-testid="stSidebar"] .stDateInput label {
            font-size: 0.72rem !important; font-weight: 600 !important;
            letter-spacing: 0.05em !important; text-transform: uppercase !important;
            color: #3F6656 !important; margin-bottom: 0.2rem !important;
        }
        [data-testid="stSidebar"] .stSelectbox > div > div,
        [data-testid="stSidebar"] .stMultiSelect > div > div {
            background: #FFFFFF !important; border: 1px solid #D9F0E2 !important;
            border-radius: 12px !important; font-size: 0.82rem !important;
            transition: border-color 0.15s ease;
        }
        [data-testid="stSidebar"] .stSelectbox > div > div:hover,
        [data-testid="stSidebar"] .stMultiSelect > div > div:hover { border-color: #1FA463 !important; }
        [data-testid="stSidebar"] .stButton > button {
            background: linear-gradient(135deg, #16523C, #1FA463) !important;
            color: #FFFFFF !important; border: none !important; border-radius: 12px !important;
            font-size: 0.8rem !important; font-weight: 600 !important; padding: 0.5rem 1rem !important;
            box-shadow: 0 4px 14px rgba(31,164,99,0.28) !important;
            transition: all 0.18s ease;
        }
        [data-testid="stSidebar"] .stButton > button:hover {
            background: linear-gradient(135deg, #1FA463, #4ED191) !important;
            box-shadow: 0 6px 18px rgba(31,164,99,0.38) !important;
            transform: translateY(-1px);
        }
        [data-testid="stSidebar"] .stButton > button p { color: #FFFFFF !important; }

        .top-banner {
            background: linear-gradient(120deg, #0F2E22 0%, #16523C 45%, #1FA463 100%);
            padding: 1.6rem 2.2rem; display: flex; align-items: center; justify-content: space-between;
            margin: 0.8rem 1.2rem 0 1.2rem; border-radius: 20px; position: relative;
            box-shadow: 0 10px 30px rgba(15,46,34,0.25);
        }
        .top-banner::after {
            content: ''; position: absolute; left: 20px; right: 20px; bottom: 0; height: 3px;
            background: linear-gradient(90deg, #4ED191 0%, #B7F0CE 50%, #4ED191 100%);
        }
        .banner-left { display: flex; align-items: center; gap: 1.1rem; }
        .banner-emblem {
            width: 56px; height: 56px;
            background: linear-gradient(135deg, rgba(78,209,145,0.35), rgba(183,240,206,0.15));
            border: 2px solid rgba(183,240,206,0.55); border-radius: 50%;
            display: flex; align-items: center; justify-content: center; font-size: 1.6rem;
            box-shadow: 0 0 0 4px rgba(255,255,255,0.05), inset 0 0 12px rgba(78,209,145,0.2);
        }
        .banner-title {
            font-family: 'Playfair Display', serif; font-size: 1.7rem; font-weight: 700;
            color: #FFFFFF; line-height: 1.15; margin: 0; letter-spacing: 0.01em;
        }
        .banner-subtitle {
            font-size: 0.72rem; font-weight: 600; color: rgba(183,240,206,0.95);
            letter-spacing: 0.14em; text-transform: uppercase; margin-top: 0.35rem;
        }
        .banner-right { text-align: right; }
        .banner-meta-label {
            font-size: 0.64rem; color: rgba(255,255,255,0.6);
            text-transform: uppercase; letter-spacing: 0.1em; font-weight: 600;
        }
        .banner-meta-value { font-size: 0.95rem; font-weight: 700; color: #FFFFFF; margin-top: 0.15rem; }
        .banner-records {
            font-size: 1.6rem; font-weight: 700; color: #FFAA00; margin-top: 0.15rem;
            text-shadow: 0 1px 12px rgba(183,240,206,0.3);
        }

        .kpi-wrap { background: transparent; padding: 1.1rem 2rem 1.2rem 2rem; }
        .kpi-row { display: grid; grid-template-columns: repeat(6, 1fr); gap: 1rem; }
        .kpi-card {
            background: linear-gradient(155deg, #FFFFFF 0%, #F1FBF5 100%);
            padding: 1.2rem 1.3rem 1.1rem 1.3rem; border: 1px solid #E1F5E9; border-radius: 18px;
            position: relative; box-shadow: 0 6px 18px rgba(15,70,45,0.07);
            transition: box-shadow 0.18s ease, transform 0.18s ease; overflow: hidden;
        }
        .kpi-card:hover { box-shadow: 0 12px 28px rgba(15,70,45,0.14); transform: translateY(-3px); }
        .kpi-card::before { content: ''; position: absolute; top: 0; left: 0; right: 0; height: 4px; border-radius: 18px 18px 0 0; }
        .kpi-card.yellow::before { background: linear-gradient(90deg,#0F2E22,#16523C); }
        .kpi-card.blue::before   { background: linear-gradient(90deg,#16523C,#1FA463); }
        .kpi-card.navy::before   { background: linear-gradient(90deg,#1FA463,#4ED191); }
        .kpi-card.red::before    { background: linear-gradient(90deg,#17A589,#5CD1B8); }
        .kpi-card.green::before  { background: linear-gradient(90deg,#2FBF71,#7EE2A8); }
        .kpi-card.purple::before { background: linear-gradient(90deg,#4ED191,#B7F0CE); }

        .kpi-icon-badge {
            width: 34px; height: 34px; border-radius: 11px;
            display: flex; align-items: center; justify-content: center;
            font-size: 0.95rem; margin-bottom: 0.65rem;
        }
        .kpi-card.yellow .kpi-icon-badge { background: linear-gradient(135deg, rgba(15,46,34,0.14), rgba(22,82,60,0.08)); }
        .kpi-card.blue   .kpi-icon-badge { background: linear-gradient(135deg, rgba(22,82,60,0.16), rgba(31,164,99,0.08)); }
        .kpi-card.navy   .kpi-icon-badge { background: linear-gradient(135deg, rgba(31,164,99,0.18), rgba(78,209,145,0.10)); }
        .kpi-card.red    .kpi-icon-badge { background: linear-gradient(135deg, rgba(23,165,137,0.18), rgba(92,209,184,0.10)); }
        .kpi-card.green  .kpi-icon-badge { background: linear-gradient(135deg, rgba(47,191,113,0.18), rgba(126,226,168,0.10)); }
        .kpi-card.purple .kpi-icon-badge { background: linear-gradient(135deg, rgba(78,209,145,0.18), rgba(183,240,206,0.14)); }

        .kpi-label { font-size: 0.66rem; font-weight: 700; letter-spacing: 0.07em; text-transform: uppercase; color: #6E9784; margin-bottom: 0.35rem; }
        .kpi-value { font-family: 'Playfair Display', serif; font-size: 1.9rem; font-weight: 700; color: #0E241B; line-height: 1.05; }
        .kpi-sub { font-size: 0.72rem; color: #7FA795; margin-top: 0.25rem; }

        .content-area { padding: 0rem 1.2rem 2.2rem 1.2rem; background: transparent; }

        .sec-title {
            font-size: 0.95rem; font-weight: 700; color: #0E241B;
            margin: 1.5rem 0 1rem 0; padding: 0 0 0.5rem 0.7rem;
            border-left: 3px solid #1FA463; border-bottom: 1px solid #E1F5E9; line-height: 1.6;
        }

        [data-testid="stDataFrame"] thead tr th {
            background: linear-gradient(135deg, #0F2E22, #1FA463) !important;
            color: white !important; font-size: 0.75rem !important; font-weight: 600 !important;
        }
        [data-testid="stDataFrame"] tbody tr:nth-child(even) td { background-color: #F3FBF6 !important; }
        [data-testid="stDataFrame"] tbody tr:hover td { background-color: #E7F7ED !important; }

        .res-card {
            background: linear-gradient(155deg, #FFFFFF 0%, #F3FBF6 100%);
            border: 1px solid #E1F5E9; border-left: 4px solid #1FA463; border-radius: 14px;
            padding: 0.95rem 1.15rem; margin-bottom: 0.6rem;
            box-shadow: 0 4px 14px rgba(15,70,45,0.05);
            transition: box-shadow 0.15s, transform 0.15s;
        }
        .res-card:hover { box-shadow: 0 10px 24px rgba(15,70,45,0.12); transform: translateY(-2px); }
        .res-case { font-size: 0.72rem; font-weight: 700; color: #17A589; letter-spacing: 0.05em; }
        .res-title { font-size: 0.9rem; font-weight: 600; color: #0E241B; margin: 0.15rem 0; }
        .res-meta { font-size: 0.74rem; color: #7FA795; }
        .badge {
            display: inline-block; background: rgba(15,70,45,0.08); color: #0F2E22;
            border: 1px solid rgba(15,70,45,0.2); border-radius: 6px; padding: 0.07rem 0.5rem;
            font-size: 0.65rem; font-weight: 700; text-transform: uppercase; letter-spacing: 0.05em; margin-right: 0.3rem;
        }
        .rbadge {
            display: inline-block; background: rgba(192,57,43,0.08); color: #C0392B;
            border: 1px solid rgba(192,57,43,0.2); border-radius: 6px; padding: 0.07rem 0.5rem;
            font-size: 0.65rem; font-weight: 600; margin-right: 0.3rem;
        }
        div[data-testid="stTabs"] {
            background: linear-gradient(155deg, #FFFFFF 0%, #DCF2E4 100%) !important;
            box-shadow: 0 8px 22px rgba(20,70,50,0.1) !important;
        }
        div[data-testid="stTabs"] [role="tablist"] {
            display: flex !important; width: 100% !important; gap: 0.25rem !important;
            overflow-x: auto !important; background: transparent !important;
        }
        div[data-testid="stTab"] {
            flex: 1 1 0 !important; display: flex !important; align-items: center !important;
            justify-content: center !important; padding: 0.6rem 0.3rem !important;
            font-size: 0.95rem !important; font-weight: 600 !important; color: #16523C !important;
            letter-spacing: 0.02em !important; white-space: nowrap !important;
            background-color: transparent !important; border: 1px solid #D9F0E2 !important;
            border-radius: 12px !important;
            transition: background 0.18s ease, color 0.18s ease, box-shadow 0.18s ease, transform 0.18s ease !important;
        }
        div[data-testid="stTab"] p {
            color: inherit !important; font-size: 1rem !important; font-weight: 600 !important;
            letter-spacing: 0.02em !important; margin: 0 !important;
        }
        div[data-testid="stTab"]:hover {
            background-color: #EAF7F0 !important; box-shadow: 0 4px 12px rgba(15,70,45,0.10) !important;
            transform: translateY(-1px);
        }
        div[data-testid="stTab"][aria-selected="true"] {
            color: #FFFFFF !important; background: linear-gradient(135deg, #16523C, #1FA463) !important;
            box-shadow: 0 6px 16px rgba(31,164,99,0.30) !important;
        }
        div[data-testid="stTab"][aria-selected="true"] p { color: #FFFFFF !important; }

        .empty { text-align:center; padding:3rem; color:#8FAFA0; }
        .empty .ei { font-size:2.5rem; }
        .empty .et { font-size:1rem; font-weight:600; color:#0E241B; margin-top:0.5rem; }
        .empty .es { font-size:0.82rem; margin-top:0.2rem; }

        .sb-logo {
            background: linear-gradient(135deg, #0F2E22, #1FA463);
            padding: 1.3rem 1.1rem;
            margin: 0;
            width: 100%;
            border-radius: 0 24px 0 0;
            position: sticky;
            top: 0;
            z-index: 999;
            box-shadow: 0 6px 16px rgba(0,0,0,0.18);
        }
        .sb-logo::after {
            content: ''; position: absolute; left: 16px; right: 16px; bottom: 0; height: 2px;
            background: linear-gradient(90deg, #4ED191, #B7F0CE, #4ED191);
        }
        .sb-logo-title { font-family: 'Playfair Display', serif; font-size: 0.95rem; font-weight: 700; color: #FFFFFF !important; line-height: 1.2; }
        .sb-logo-sub { font-size: 0.62rem; color: rgba(183,240,206,0.95) !important; text-transform: uppercase; letter-spacing: 0.1em; margin-top: 0.2rem; }
        .sb-section {
            font-size: 0.64rem; font-weight: 700; letter-spacing: 0.1em; text-transform: uppercase;
            color: #8FAFA0 !important; padding: 1rem 0.2rem 0.4rem 0.2rem; border-bottom: 1px solid #E1F5E9; margin-bottom: 0.6rem;
        }

        .info-box {
            background: linear-gradient(135deg, #EAFBF1, #DFF7E9); border: 1px solid #A8E6C1;
            border-left: 4px solid #1FA463; border-radius: 12px; padding: 0.75rem 1.1rem;
            font-size: 0.8rem; color: #0F5C34; margin-bottom: 1rem; box-shadow: 0 4px 12px rgba(15,70,45,0.06);
        }

        ::-webkit-scrollbar { width: 5px; height: 5px; }
        ::-webkit-scrollbar-track { background: #F0FAF4; }
        ::-webkit-scrollbar-thumb { background: linear-gradient(#1FA463,#4ED191); border-radius: 3px; }

        .stDownloadButton > button {
            background: linear-gradient(135deg, #16523C, #1FA463) !important; color: white !important;
            border: none !important; border-radius: 12px !important; font-size: 0.8rem !important;
            font-weight: 600 !important; padding: 0.5rem 1.2rem !important;
            box-shadow: 0 4px 14px rgba(31,164,99,0.25) !important;
        }
        .stDownloadButton > button:hover {
            background: linear-gradient(135deg, #1FA463, #4ED191) !important;
            box-shadow: 0 6px 18px rgba(31,164,99,0.35) !important;
        }

        #MainMenu, footer { visibility: hidden !important; }
        header { background: transparent !important; box-shadow: none !important; }
        header [data-testid="stToolbar"] { visibility: hidden !important; }

        [data-testid="stExpandSidebarButton"] {
            visibility: visible !important;
            display: flex !important;
            position: fixed !important;
            top: 12px !important;
            left: 12px !important;
            z-index: 999999 !important;
            background: #1FA463 !important;
            border-radius: 8px !important;
            padding: 6px !important;
            box-shadow: 0 4px 12px rgba(0,0,0,0.25) !important;
        }
        [data-testid="stExpandSidebarButton"] svg {
            visibility: visible !important;
            fill: #FFFFFF !important;
            stroke: #FFFFFF !important;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )


# ═══════════════════════════════════════════════════════════════
# DATA LOADING & CLEANING
# ═══════════════════════════════════════════════════════════════
import re as _re

_ADVOCATE_OVERFLOW = _re.compile(
    r"(?:(?<=\s)|^)"
    r"(?:Cr\.\w[\w\.]*\s*(?:Appln|Appeal|Rev|Bail|Acq|Tran|Acctt[^,]*)?\s+|"
    r"Const\.\s*P\.\s+|Spl\.\w+\s+)"
    r"[\w\-\.]+/\d{4}",
    _re.IGNORECASE,
)

_CAT_MAP: dict[str, str] = {
    "AGAINST THE ORDER":             "AGAINST ORDER",
    "AGAINST THE JUDGEMENT":         "AGAINST JUDGEMENT",
    "SALES TAX.":                    "SALES TAX",
    "QUASHEMENT OF F.I.R.":          "QUASHMENT OF F.I.R.",
    "QUASHMENT OF F.I.R / FREE-WILL": "QUASHMENT OF F.I.R.",
    "Land Matters":                  "LAND MATTERS",
    "W.W.F":                         "WWF",
}

_CRIMINAL_CASE_PREFIX = _re.compile(r"^(Cr\.|Spl\.Cr\.|Criminal)", _re.IGNORECASE)

_VALID_SECTIONS = {
    "FOR ANNOUNCEMENT OF JUDGMENT/ORDER",
    "FOR APPLICATION IN DISPOSED OF CASES",
    "FOR BAIL AFTER ARREST APPLICATIONS AND APPLICATIONS UNDER SECTION 426 CR.P.C",
    "FOR BAIL BEFORE ARREST APPLICATIONS",
    "FOR DIRECTIONS",
    "FOR FRESH CASES",
    "FOR HEARING OF CASES",
    "FOR HEARING OF CASES(PRIORITY)",
    "FOR ORDERS AS TO NON-PROSECUTION",
}

def _clean_section(df: pd.DataFrame) -> pd.DataFrame:
    if "Section" not in df.columns:
        df["Section"] = ""
    df["Section"] = df["Section"].apply(
        lambda s: s if str(s).upper() in _VALID_SECTIONS else ""
    )
    df["Section"] = df["Section"].replace("", pd.NA)
    if "Bench" in df.columns:
        df["Section"] = (
            df.groupby("Bench", sort=False)["Section"]
            .transform(lambda s: s.ffill().bfill())
        )
    df["Section"] = df["Section"].fillna("UNKNOWN").astype(str)
    return df


def _clean_advocate(val: str) -> str:
    val = str(val).strip()
    if not val:
        return val
    m = _ADVOCATE_OVERFLOW.search(val)
    if not m:
        return val
    cleaned = val[: m.start()].strip(" ,;-")
    return cleaned if cleaned else ""


def _clean_case_category(row: pd.Series) -> str:
    cat = _CAT_MAP.get(str(row.get("Case_Category", "")), str(row.get("Case_Category", ""))).strip()
    if cat:
        return cat
    if _CRIMINAL_CASE_PREFIX.match(str(row.get("Case_No", "")).strip()):
        return "CRIMINAL MATTER"
    return "UNCATEGORIZED"


def folder_signature(folder: Path) -> tuple:
    if not folder.exists():
        return ()
    files = sorted(folder.glob("Sindh_Cause_List_*.xlsx"))
    return tuple((f.name, f.stat().st_size, f.stat().st_mtime) for f in files)


def _parse_date(row: pd.Series):
    for fmt in ("%d %B %Y", "%d %b %Y"):
        try:
            return datetime.strptime(f"{row['Day']} {row['Month']} {row['Year']}", fmt)
        except (ValueError, TypeError, KeyError):
            continue
    return None


@st.cache_data(ttl=300)
def load_data(folder: Path, _signature: tuple) -> pd.DataFrame:
    if not folder.exists():
        return pd.DataFrame()
    files = sorted(folder.glob("Sindh_Cause_List_*.xlsx"))
    if not files:
        return pd.DataFrame()

    frames = []
    for f in files:
        try:
            frames.append(pd.read_excel(f, dtype=str))
        except Exception:
            continue
    if not frames:
        return pd.DataFrame()

    df = pd.concat(frames, ignore_index=True)
    df.fillna("", inplace=True)
    df.drop_duplicates(inplace=True)

    df = _clean_section(df)
    if "Respondent_Advocate" in df.columns:
        df["Respondent_Advocate"] = df["Respondent_Advocate"].apply(_clean_advocate)
    if "Case_Category" in df.columns:
        df["Case_Category"] = df.apply(_clean_case_category, axis=1)

    df["Date"] = df.apply(_parse_date, axis=1)
    df["Date_Str"] = df["Date"].apply(lambda d: d.strftime("%d %B %Y") if pd.notnull(d) else "")
    if "Bench" in df.columns:
        df["Judge_Short"] = (
            df["Bench"].str.replace("MR. JUSTICE ", "", regex=False)
            .str.replace("MRS. JUSTICE ", "", regex=False)
            .str.strip()
        )
    else:
        df["Judge_Short"] = "UNKNOWN"
        df["Bench"] = "UNKNOWN"

    if "Section" in df.columns:
        df["Section_Clean"] = df["Section"].str.replace("FOR ", "", regex=False).str.title().str.strip()
    else:
        df["Section_Clean"] = "Unknown"
        
    return df


# ═══════════════════════════════════════════════════════════════
# MAIN APP ENTRYWAY
# ═══════════════════════════════════════════════════════════════
def main():
    inject_css()
    
    # Check folder availability
    sig = folder_signature(CAUSE_LIST_FOLDER)
    df = load_data(CAUSE_LIST_FOLDER, sig)

    # Sidebar setup
    with st.sidebar:
        st.markdown(
            """
            <div class="sb-logo">
                <div class="sb-logo-title">Sindh High Court</div>
                <div class="sb-logo-sub">Cause List Analytics</div>
            </div>
            """,
            unsafe_allow_html=True,
        )
        st.markdown("<div style='height: 1rem;'></div>", unsafe_allow_html=True)
        st.markdown('<div class="sb-section">Dashboard Filters</div>', unsafe_allow_html=True)

    if df.empty:
        st.markdown(
            f"""
            <div class="empty">
                <div class="ei">📂</div>
                <div class="et">No data found in cause_lists folder.</div>
                <div class="es">Please make sure the Excel files (.xlsx) matching 'Sindh_Cause_List_*.xlsx' are placed under: <code>{CAUSE_LIST_FOLDER}</code></div>
            </div>
            """,
            unsafe_allow_html=True,
        )
        return

    # Basic app display execution if files are present
    st.markdown(
        f"""
        <div class="top-banner">
            <div class="banner-left">
                <div class="banner-emblem">⚖️</div>
                <div>
                    <div class="banner-title">Sindh High Court Analytics</div>
                    <div class="banner-subtitle">Cause List Intelligence Platform</div>
                </div>
            </div>
            <div class="banner-right">
                <div class="banner-meta-label">Total Records Loaded</div>
                <div class="banner-records">{len(df):,}</div>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

if __name__ == "__main__":
    main()
