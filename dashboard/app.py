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

APP_BUILD = "2026-07-22-axis-fix-v2"

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

CAUSE_LIST_FOLDER = Path(__file__).parent.parent / "cause_lists"

# New multi-bench master data (Karachi, Hyderabad, Sukkur, Larkana, Mirpurkhas
# combined into one file with a City column).
MASTER_DATA_FOLDER = Path(r"C:\Users\Hafiz Ahmed\Desktop\Sindh\Sindh\sindh_causelist_master")
MASTER_DATA_FILE = MASTER_DATA_FOLDER / "Sindh_Cause_List_Master_Combined.xlsx"


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
    """Return an inline SVG icon (replaces emoji for a cleaner look)."""
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

        /* ── SIDEBAR SHELL ──────────────────────────────────────
           Outer sidebar is transparent; the inner wrapper carries
           the background, rounded corners and its own scrollbar,
           so the banner can stick to the top of that scroll area. */
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
        /* Take Streamlit's built-in sidebar header out of the flow entirely,
           so it can no longer push the banner down or leave a gap above it.
           Its collapse icon is repositioned as a simple button sitting on
           top of the banner's top-right corner. */
        /* Remove Streamlit's built-in left/right padding on the whole
           sidebar content area (this was the source of the side gaps
           around the banner), then restore that padding only for the
           elements below the banner (filters, buttons, info text). */
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

        /* ── SIDEBAR BANNER ──────────────────────────────────────
           Sticky to the top of the sidebar's own scroll area, fills
           the full width/top corners, and never scrolls away. */
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

        /* Sidebar re-open button — must stay visible even though the
           header/toolbar around it is hidden (Streamlit 1.59+) */
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
# DATA LOADING
# ═══════════════════════════════════════════════════════════════
import re as _re

# ── Cleaning constants ───────────────────────────────────────────

# Advocate: overflow pattern — a case reference token that signals
# the start of bundled companion-case data.
# Matches patterns like "Cr.Bail 1457/2023", "Cr.Misc. Appln 1148/2025",
# "Const. P. 3612/2024" embedded inside an advocate cell, whether they
# appear mid-string (after a space) or at the very start of the value.
_ADVOCATE_OVERFLOW = _re.compile(
    r"(?:(?<=\s)|^)"
    r"(?:Cr\.\w[\w\.]*\s*(?:Appln|Appeal|Rev|Bail|Acq|Tran|Acctt[^,]*)?\s+|"
    r"Const\.\s*P\.\s+|Spl\.\w+\s+)"
    r"[\w\-\.]+/\d{4}",
    _re.IGNORECASE,
)

# Case_Category: normalise typos and near-duplicate labels.
_CAT_MAP: dict[str, str] = {
    "AGAINST THE ORDER":              "AGAINST ORDER",
    "AGAINST THE JUDGEMENT":          "AGAINST JUDGEMENT",
    "SALES TAX.":                     "SALES TAX",
    "QUASHEMENT OF F.I.R.":           "QUASHMENT OF F.I.R.",
    "QUASHMENT OF F.I.R / FREE-WILL": "QUASHMENT OF F.I.R.",
    "Land Matters":                   "LAND MATTERS",
    "W.W.F":                          "WWF",
}

# Case_No prefix pattern that identifies criminal-matter cases
# (used to fill empty Case_Category rows).
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
    df["Section"] = df["Section"].apply(
        lambda s: s if s.upper() in _VALID_SECTIONS else ""
    )

    print("=== SECTION DEBUG ===")
    print(df["Section"].value_counts(dropna=False).head(20))
    candidates = df["Section"].replace("", pd.NA).dropna().str.upper().str.strip()
    near_misses = set(df["Section"].str.upper().str.strip().unique()) - _VALID_SECTIONS - {""}
    print("NOT IN WHITELIST:", near_misses)

    df["Section"] = df["Section"].replace("", pd.NA)
    df["Section"] = (
        df.groupby("Bench", sort=False)["Section"]
        .transform(lambda s: s.ffill().bfill())
    )
    df["Section"] = df["Section"].fillna("UNKNOWN").astype(str)
    return df


def _clean_advocate(val: str) -> str:
    """Strip companion-case overflow text from an advocate cell.
    Overflow always follows a whitespace boundary (the regex requires it),
    so after stripping we check if anything useful remains before the match.
    """
    val = val.strip()
    if not val:
        return val
    m = _ADVOCATE_OVERFLOW.search(val)
    if not m:
        return val
    cleaned = val[: m.start()].strip(" ,;-")
    return cleaned if cleaned else ""


def _clean_case_category(row: pd.Series) -> str:
    """Return a normalised Case_Category:
    • Apply the explicit deduplication map.
    • For still-empty cells, infer from the Case_No prefix.
    """
    cat = _CAT_MAP.get(row["Case_Category"], row["Case_Category"]).strip()
    if cat:
        return cat
    if _CRIMINAL_CASE_PREFIX.match(row["Case_No"].strip()):
        return "CRIMINAL MATTER"
    return "UNCATEGORIZED"


def folder_signature(folder: Path) -> tuple:
    """Fingerprint of the folder's contents (name, size, mtime) so that
    st.cache_data invalidates automatically whenever a file changes,
    without waiting for the TTL to expire."""
    files = sorted(folder.glob("Sindh_Cause_List_*.xlsx"))
    return tuple((f.name, f.stat().st_size, f.stat().st_mtime) for f in files)


def master_file_signature(file_path: Path) -> tuple:
    """Fingerprint of the single master combined file (name, size, mtime)
    so that st.cache_data invalidates automatically whenever it changes."""
    if not file_path.exists():
        return ()
    stat = file_path.stat()
    return ((file_path.name, stat.st_size, stat.st_mtime),)


def _parse_date(row: pd.Series):
    for fmt in ("%d %B %Y", "%d %b %Y"):
        try:
            return datetime.strptime(f"{row['Day']} {row['Month']} {row['Year']}", fmt)
        except (ValueError, TypeError):
            continue
    return None


@st.cache_data(ttl=300)
def load_data(folder: Path, _signature: tuple) -> pd.DataFrame:
    # Prefer the new multi-bench master file (has a City column covering
    # Karachi, Hyderabad, Sukkur, Larkana, Mirpurkhas). Fall back to any
    # other "Sindh_Cause_List_*.xlsx" files sitting in the same folder.
    if MASTER_DATA_FILE.exists():
        files = [MASTER_DATA_FILE]
    else:
        files = sorted(MASTER_DATA_FOLDER.glob("Sindh_Cause_List_*.xlsx"))
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
    if "City" not in df.columns:
        df["City"] = "Karachi"
    df["City"] = df["City"].astype(str).str.strip()
    # in load_data, right after df.fillna("", inplace=True)
    garbage_mask = df["Section"].str.contains(r"(?i)^for\s+", regex=True)
    print(df.loc[garbage_mask, "Section"].value_counts().head(20))
    print("---RAW UNIQUE SECTIONS---")
    print(df["Section"].unique()[:40])
    df.drop_duplicates(inplace=True)

    # ── Data cleaning ────────────────────────────────────────────
    # 1. Section: remove garbage strings, forward/back-fill within bench
    df = _clean_section(df)

    # 2. Advocate overflow: strip embedded companion-case text
    df["Respondent_Advocate"] = df["Respondent_Advocate"].apply(_clean_advocate)

    # 3. Case_Category: normalise typos + infer from Case_No where blank
    df["Case_Category"] = df.apply(_clean_case_category, axis=1)
    # ── End cleaning ─────────────────────────────────────────────

    df["Date"] = df.apply(_parse_date, axis=1)
    df["Date_Str"] = df["Date"].apply(lambda d: d.strftime("%d %B %Y") if pd.notnull(d) else "")
    df["Judge_Short"] = (
        df["Bench"].str.replace("MR. JUSTICE ", "", regex=False)
        .str.replace("MRS. JUSTICE ", "", regex=False)
        .str.strip()
    )
    df["Section_Clean"] = df["Section"].str.replace("FOR ", "", regex=False).str.title().str.strip()
    return df


# ═══════════════════════════════════════════════════════════════
# CHART HELPERS
# ═══════════════════════════════════════════════════════════════
def base_layout(height: int = 320, **overrides) -> dict:
    layout = dict(
        paper_bgcolor="#FFFFFF",
        plot_bgcolor="#FFFFFF",
        font=FONT,
        margin=dict(l=10, r=10, t=10, b=10),
        height=height,
    )
    layout.update(overrides)
    return layout


def axis_x(rotate: int = 0, **overrides) -> dict:
    axis = dict(tickfont=dict(size=10, color=COLORS["text_mid"]), tickangle=rotate, **GRID_STYLE)
    axis.update(overrides)
    return axis


def axis_y(**overrides) -> dict:
    axis = dict(tickfont=dict(size=10, color=COLORS["text_mid"]), **GRID_STYLE)
    axis.update(overrides)
    return axis


def truncate(text: str, length: int) -> str:
    return text if len(text) <= length else text[:length] + "…"


def collapse_top_n(counts: pd.DataFrame, label_col: str, value_col: str, top_n: int, other_label: str = "Others") -> pd.DataFrame:
    """Keep the top N rows (by value_col, already sorted descending) and fold the rest into one 'Others' row."""
    if len(counts) <= top_n:
        return counts.copy()
    top = counts.head(top_n).copy()
    other_sum = counts[value_col].iloc[top_n:].sum()
    other_row = pd.DataFrame([{label_col: other_label, value_col: other_sum}])
    return pd.concat([top, other_row], ignore_index=True)


def navy_to_steel_gradient(n: int) -> list:
    """Return n colors fading from dark navy (index 0) to light steel-blue (index n-1)."""
    start, end = np.array([27, 42, 63]), np.array([190, 205, 225])
    colors = []
    for i in range(n):
        t = i / max(n - 1, 1)
        rgb = (start + (end - start) * t).astype(int)
        colors.append(f"rgb({rgb[0]},{rgb[1]},{rgb[2]})")
    return colors


def nice_dtick(max_val: float, target_ticks: int = 5) -> float:
    """Return a 'round' tick spacing (1/2/2.5/5 x a power of 10) so that a numeric
    axis gets ~target_ticks evenly, cleanly spaced labels instead of letting
    Plotly's auto tick-picker choose uneven values that render squeezed together."""
    if max_val <= 0:
        return 1
    raw_step = max_val / target_ticks
    magnitude = 10 ** math.floor(math.log10(raw_step))
    for m in (1, 2, 2.5, 5, 10):
        step = m * magnitude
        if raw_step <= step:
            return step
    return 10 * magnitude


def horizontal_bar_ranked(
    df_counts: pd.DataFrame,
    label_col: str,
    value_col: str,
    height: int,
    label_len: int = 30,
    single_color: str | None = None,
    gradient_desc: bool = False,
    value_labels: pd.Series | None = None,
    right_margin: int = 60,
    x_headroom: float = 1.18,
) -> go.Figure:
    """A horizontal ranked bar chart (ascending order, highest value on top),
    with outside value labels — the pattern reused across every 'Top N' panel."""
    data = df_counts.copy()
    data["_disp"] = data[label_col].apply(lambda x: truncate(x, label_len))
    data = data.sort_values(value_col, ascending=True)
    n = len(data)

    if gradient_desc:
        colors = navy_to_steel_gradient(n)
        colors = [colors[n - 1 - i] for i in range(n)]
    else:
        colors = single_color or COLORS["dark_green"]

    text = value_labels if value_labels is not None else data[value_col]
    max_val = data[value_col].max() if n else 10

    fig = go.Figure(
        go.Bar(
            x=data[value_col],
            y=data["_disp"],
            orientation="h",
            marker=dict(color=colors, line=dict(width=0)),
            text=text,
            textposition="outside",
            textfont=dict(size=10, color=COLORS["text_dark"]),
            cliponaxis=False,
            hovertemplate="<b>%{y}</b><br>Cases: %{x:,}<extra></extra>",
        )
    )
    fig.update_layout(**base_layout(height=height, margin=dict(l=10, r=right_margin, t=10, b=40), bargap=0.3, showlegend=False, uniformtext_minsize=7, uniformtext_mode="hide",))
    axis_max = max_val * x_headroom
    fig.update_xaxes(**axis_x(), title=dict(text="Number of Cases", font=dict(size=10, color=COLORS["text_mid"])),
                      range=[0, axis_max], tick0=0, dtick=nice_dtick(max_val, target_ticks=4), automargin=True)
    fig.update_yaxes(**axis_y(tickfont=dict(size=10, color=COLORS["text_dark"])), automargin=True)
    return fig


def donut_chart(
    df_counts: pd.DataFrame,
    label_col: str,
    value_col: str,
    height: int,
    label_len: int = 26,
    colors: list | None = None,
    total_label: str = "Total Cases",
    hole: float = 0.55,
) -> go.Figure:
    """A donut chart with a centered total annotation — reused across every category/section breakdown."""
    data = df_counts.copy()
    data["_disp"] = data[label_col].apply(lambda x: truncate(x, label_len))
    total = data[value_col].sum()
    palette = colors or DONUT_PALETTE

    fig = go.Figure(
        go.Pie(
            labels=data["_disp"],
            values=data[value_col],
            hole=hole,
            marker=dict(colors=palette[: len(data)], line=dict(color="#FFFFFF", width=2)),
            textposition="inside",
            texttemplate="%{value:,}<br>%{percent}",
            insidetextorientation="horizontal",
            textfont=dict(size=10, color="#FFFFFF"),
            sort=False,
            hovertemplate="<b>%{label}</b><br>Cases: %{value:,} (%{percent})<extra></extra>",
        )
    )
    fig.update_layout(
        **base_layout(
            height=height,
            margin=dict(l=10, r=10, t=20, b=10),
            showlegend=True,
            legend=dict(orientation="h", x=0.5, y=-0.12, xanchor="center", yanchor="top", font=dict(size=10), bgcolor="rgba(0,0,0,0)"),
            annotations=[
                dict(
                    text=f"<b>{total:,}</b><br><span style='font-size:10px;color:{COLORS['text_mid']}'>{total_label}</span>",
                    x=0.5, y=0.5, font=dict(size=20, color=COLORS["text_dark"]), showarrow=False,
                )
            ],
        )
    )
    return fig


def value_counts_df(series: pd.Series, label_col: str, value_col: str = "Cases", blank_label: str | None = None) -> pd.DataFrame:
    s = series.replace("", blank_label) if blank_label else series
    out = s.value_counts().reset_index()
    out.columns = [label_col, value_col]
    return out


def kpi_card_html(label: str, value: str, sub: str, tone: str, icon_name: str, icon_color: str) -> str:
    return f"""
    <div class="kpi-card {tone}" style="margin-bottom:0.9rem;">
        <div class="kpi-icon-badge">{svg_icon(icon_name, icon_color)}</div>
        <div class="kpi-label">{label}</div>
        <div class="kpi-value" style="font-size:1.5rem;">{value}</div>
        <div class="kpi-sub">{sub}</div>
    </div>"""


# ═══════════════════════════════════════════════════════════════
# SIDEBAR
# ═══════════════════════════════════════════════════════════════
def render_sidebar() -> pd.DataFrame:
    with st.sidebar:
        st.markdown(
            """
            <div class="sb-logo">
                <div style="display:flex;align-items:center;gap:0.7rem;">
                    <div style="font-size:1.6rem;">⚖️</div>
                    <div>
                        <div class="sb-logo-title">SHC Analytics</div>
                        <div class="sb-logo-sub">Gallup Pakistan</div>
                    </div>
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )
        st.caption(f"Build: {APP_BUILD}")

        st.markdown('<div class="sb-section" style="padding-top:0.9rem;">Data</div>', unsafe_allow_html=True)
        if st.button("🔄  Refresh Data", use_container_width=True):
            load_data.clear()
            st.rerun()

        df_all = load_data(MASTER_DATA_FOLDER, master_file_signature(MASTER_DATA_FILE))
        if df_all.empty:
            st.error("No data found. Check the master file path/name.")
            st.stop()

        st.markdown('<div class="sb-section">Filters</div>', unsafe_allow_html=True)

        # ── COURT LOCATION (City) — drives every other filter below ────
        all_cities = sorted(c for c in df_all["City"].unique() if c)
        sel_cities = st.multiselect("COURT LOCATION", all_cities, default=[], placeholder="All Locations")
        df_city = df_all if not sel_cities else df_all[df_all["City"].isin(sel_cities)]
        sel_city = "All Locations" if not sel_cities else (sel_cities[0] if len(sel_cities) == 1 else ", ".join(sel_cities))

        avail_dates = sorted(df_city["Date_Str"].unique(), reverse=True)
        sel_dates = st.multiselect("HEARING DATE", avail_dates, default=[], placeholder="All Dates")

        all_judges = sorted(df_city["Judge_Short"].unique())
        sel_judges = st.multiselect("JUDGE / BENCH", all_judges, default=[], placeholder="All Judges")

        all_sections = sorted(df_city["Section_Clean"].unique())
        sel_sections = st.multiselect("CAUSE LIST TYPE", all_sections, default=[], placeholder="All Types")

        all_cats = sorted(df_city["Case_Category"].replace("", "Uncategorized").unique())
        sel_cats = st.multiselect("CASE CATEGORY", all_cats, default=[], placeholder="All Categories")

        st.markdown("<hr style='border-color:#D5EDDF;margin:1rem 0;'>", unsafe_allow_html=True)
        st.markdown(
            f"""
            <div class="sb-section">Data Info</div>
            <div style="font-size:0.78rem;color:#3F6656;line-height:1.9;">
                📍 <b>{sel_city}</b><br>
                📁 <b>{len(avail_dates)}</b> day(s) loaded<br>
                ⚖️ <b>{df_city['Case_No'].nunique():,}</b> unique cases<br>
                👨‍⚖️ <b>{df_city['Judge_Short'].nunique()}</b> judges<br>
                📅 Latest: <b>{avail_dates[0] if avail_dates else '—'}</b>
            </div>
            """,
            unsafe_allow_html=True,
        )

    filtered = df_city.copy()
    if sel_dates:
        filtered = filtered[filtered["Date_Str"].isin(sel_dates)]
    if sel_judges:
        filtered = filtered[filtered["Judge_Short"].isin(sel_judges)]
    if sel_sections:
        filtered = filtered[filtered["Section_Clean"].isin(sel_sections)]
    if sel_cats:
        cat_f = filtered["Case_Category"].replace("", "Uncategorized")
        filtered = filtered[cat_f.isin(sel_cats)]

    st.session_state["_avail_dates"] = avail_dates
    st.session_state["_sel_city"] = sel_city
    return filtered


# ═══════════════════════════════════════════════════════════════
# TOP BANNER + KPIs
# ═══════════════════════════════════════════════════════════════
def render_banner(df: pd.DataFrame) -> None:
    avail_dates = st.session_state.get("_avail_dates", [])
    last_updated = avail_dates[0] if avail_dates else "—"
    sel_city = st.session_state.get("_sel_city", "All Locations")
    if sel_city == "All Locations":
        scope_label = "All Benches — Karachi, Hyderabad, Sukkur, Larkana & Mirpurkhas"
    elif "," in sel_city:
        scope_label = f"{sel_city} Benches"
    else:
        scope_label = f"{sel_city} Bench"

    st.markdown(
        f"""
        <div class="top-banner">
            <div class="banner-left">
                <div class="banner-emblem">⚖️</div>
                <div>
                    <div class="banner-title">Sindh High Court<br>Analytics Dashboard</div>
                    <div class="banner-subtitle">Judicial Intelligence Platform — {scope_label}</div>
                </div>
            </div>
            <div class="banner-right">
                <div class="banner-meta-label">Last Updated</div>
                <div class="banner-meta-value">{last_updated}</div>
                <div class="banner-meta-label" style="margin-top:0.5rem;">Filtered Records</div>
                <div class="banner-records">{len(df):,}</div>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_kpis(df: pd.DataFrame) -> None:
    total_cases = len(df)
    total_judges = df["Judge_Short"].nunique()
    total_sections = df["Section_Clean"].nunique()
    avg_daily = int(df.groupby("Date_Str").size().mean()) if df["Date_Str"].nunique() > 0 else 0
    total_cats = df["Case_Category"].nunique()
    total_advs = pd.concat(
        [
            df.loc[df["Petitioner_Advocate"] != "", "Petitioner_Advocate"],
            df.loc[df["Respondent_Advocate"] != "", "Respondent_Advocate"],
        ]
    ).nunique()

    cards = [
        ("yellow", "cases", "#0F2E22", "Total Cases", f"{total_cases:,}", "Case listings"),
        ("blue", "judge", "#16523C", "Judges", f"{total_judges}", "Active benches"),
        ("navy", "folder", "#1FA463", "Sections", f"{total_sections}", "Hearing types"),
        ("red", "load", "#17A589", "Avg Daily Load", f"{avg_daily:,}", "Cases per day"),
        ("green", "category", "#2FBF71", "Categories", f"{total_cats}", "Case types"),
        ("purple", "lawyer", "#4ED191", "Lawyers", f"{total_advs:,}", "Active advocates"),
    ]
    cards_html = "".join(
        f"""
        <div class="kpi-card {tone}">
            <div class="kpi-icon-badge">{svg_icon(icon, color)}</div>
            <div class="kpi-label">{label}</div>
            <div class="kpi-value">{value}</div>
            <div class="kpi-sub">{sub}</div>
        </div>"""
        for tone, icon, color, label, value, sub in cards
    )
    st.markdown(f'<div class="kpi-wrap"><div class="kpi-row">{cards_html}</div></div>', unsafe_allow_html=True)


# ═══════════════════════════════════════════════════════════════
# TAB 1 — OVERVIEW
# ═══════════════════════════════════════════════════════════════
def tab_overview(df: pd.DataFrame) -> None:
    st.markdown('<div class="content-area">', unsafe_allow_html=True)

    st.markdown('<div class="sec-title">📈 Case Volume Over Time</div>', unsafe_allow_html=True)
    daily = df.groupby("Date").size().reset_index(name="Cases").sort_values("Date")
    daily["MA3"] = daily["Cases"].rolling(3, min_periods=1).mean().round(0)
    # Use a categorical (string) axis for the dates instead of a continuous date
    # axis — with only a handful of days loaded, Plotly's automatic date-tick
    # generator was producing duplicate/overlapping labels (each date rendered
    # twice, see reported bug). A category axis guarantees exactly one tick
    # per date, eliminating the overlap.
    daily["Date_Label"] = daily["Date"].dt.strftime("%d %b")

    fig_trend = go.Figure()
    fig_trend.add_trace(
        go.Bar(
            x=daily["Date_Label"], y=daily["Cases"], name="Daily Cases",
            marker_color=COLORS["navy"], marker_line_width=0, opacity=0.8,
            hovertemplate="<b>%{x}</b><br>Cases: %{y:,}<extra></extra>",
        )
    )
    fig_trend.add_trace(
        go.Scatter(
            x=daily["Date_Label"], y=daily["MA3"], name="3-Day Avg",
            line=dict(color=COLORS["gold"], width=2.5), mode="lines",
            hovertemplate="<b>%{x}</b><br>3-Day Avg: %{y:,.0f}<extra></extra>",
        )
    )
    fig_trend.update_layout(
        **base_layout(
            height=300, margin=dict(l=50, r=30, t=20, b=50), bargap=0.3,
            legend=dict(orientation="h", x=0, y=1.12, font=dict(size=11), bgcolor="rgba(0,0,0,0)", bordercolor="#D5EDDF", borderwidth=1),
        )
    )
    fig_trend.update_xaxes(
        **axis_x(-30 if len(daily) > 10 else 0, tickfont=dict(size=11, color=COLORS["text_mid"])),
        type="category", categoryorder="array", categoryarray=daily["Date_Label"],
    )
    fig_trend.update_yaxes(**axis_y(tickfont=dict(size=11, color=COLORS["text_mid"]), title=dict(text="Number of Cases", font=dict(size=11, color=COLORS["text_mid"]))))
    st.plotly_chart(fig_trend, use_container_width=True, key="chart_899_fig_trend")
    st.markdown("<div style='height:8px'></div>", unsafe_allow_html=True)

    col1, col2 = st.columns(2, gap="medium")
    with col1:
        st.markdown('<div class="sec-title">📂 Case Category Mix</div>', unsafe_allow_html=True)
        cc_full = value_counts_df(df["Case_Category"], "Category", blank_label="Uncategorized")
        cc = collapse_top_n(cc_full, "Category", "Cases", top_n=6)
        cc["Category"] = cc["Category"].apply(lambda x: truncate(x, 32))
        cc["Pct"] = (cc["Cases"] / cc["Cases"].sum() * 100).round(1)
        cc["Label"] = cc.apply(lambda r: f"{r['Cases']:,}  ({r['Pct']}%)", axis=1)
        fig_cat = horizontal_bar_ranked(
            cc, "Category", "Cases", height=360, label_len=32,
            single_color=PALETTE[: len(cc)], value_labels=cc.sort_values("Cases")["Label"],
            right_margin=195, x_headroom=1.3,
        )
        st.plotly_chart(fig_cat, use_container_width=True, key="chart_915_fig_cat")

    with col2:
        st.markdown('<div class="sec-title">📋 Cause List Distribution</div>', unsafe_allow_html=True)
        sc_full = value_counts_df(df["Section_Clean"], "Section")
        sc = collapse_top_n(sc_full, "Section", "Cases", top_n=6)
        colors = [COLORS["dark_green"], COLORS["gold"], COLORS["teal"], COLORS["red"], COLORS["orange"], COLORS["purple"], COLORS["text_light"]]
        fig_sec = donut_chart(sc, "Section", "Cases", height=360, label_len=26, colors=colors)
        st.plotly_chart(fig_sec, use_container_width=True, key="chart_923_fig_sec")

    st.markdown("<div style='height:8px'></div>", unsafe_allow_html=True)

    col3, col4 = st.columns(2, gap="medium")
    with col3:
        st.markdown('<div class="sec-title">👨‍⚖️ Top 10 Judges by Caseload</div>', unsafe_allow_html=True)
        jc = value_counts_df(df["Judge_Short"], "Judge").head(10)
        fig_jc = horizontal_bar_ranked(jc, "Judge", "Cases", height=380, label_len=28, single_color=COLORS["dark_green"], right_margin=80)
        st.plotly_chart(fig_jc, use_container_width=True, key="chart_932_fig_jc")

    with col4:
        st.markdown('<div class="sec-title">🧑‍💼 Top 10 Petitioner Advocates</div>', unsafe_allow_html=True)
        pa = value_counts_df(df.loc[df["Petitioner_Advocate"] != "", "Petitioner_Advocate"], "Advocate").head(10)
        fig_pa = horizontal_bar_ranked(pa, "Advocate", "Cases", height=380, label_len=28, single_color=COLORS["teal"], right_margin=80)
        st.plotly_chart(fig_pa, use_container_width=True, key="chart_938_fig_pa")

    st.markdown("<div style='height:8px'></div>", unsafe_allow_html=True)

    col5, col6 = st.columns(2, gap="medium")
    with col5:
        st.markdown('<div class="sec-title">📅 Cases by Day of Week</div>', unsafe_allow_html=True)
        dow_order = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]
        dow_counts = df["Date"].dt.day_name().value_counts().reindex(dow_order).dropna().reset_index()
        dow_counts.columns = ["Day", "Cases"]
        fig_dow = px.bar(dow_counts, x="Day", y="Cases", color_discrete_sequence=[COLORS["navy"]], text="Cases")
        fig_dow.update_traces(marker_line_width=0, textposition="outside", textfont=dict(size=11, color=COLORS["text_dark"]),
                               hovertemplate="<b>%{x}</b><br>Cases: %{y:,}<extra></extra>")
        fig_dow.update_layout(**base_layout(height=320, margin=dict(l=50, r=30, t=20, b=60), bargap=0.35))
        fig_dow.update_xaxes(**axis_x(0, tickfont=dict(size=11, color=COLORS["text_mid"])))
        fig_dow.update_yaxes(**axis_y(tickfont=dict(size=11, color=COLORS["text_mid"]), title=dict(text="Number of Cases", font=dict(size=10, color=COLORS["text_mid"]))))
        st.plotly_chart(fig_dow, use_container_width=True, key="chart_954_fig_dow")

    with col6:
        st.markdown('<div class="sec-title">🗓️ Cases by Month</div>', unsafe_allow_html=True)
        month_order = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"]
        mon = df["Month"].value_counts().reindex(month_order).dropna().reset_index()
        mon.columns = ["Month", "Cases"]
        fig_mon = px.bar(mon, x="Month", y="Cases", color_discrete_sequence=[COLORS["gold"]], text="Cases")
        fig_mon.update_traces(marker_line_width=0, textposition="outside", textfont=dict(size=11, color=COLORS["text_dark"]),
                              hovertemplate="<b>%{x}</b><br>Cases: %{y:,}<extra></extra>")
        fig_mon.update_layout(**base_layout(height=320, margin=dict(l=50, r=30, t=20, b=60), bargap=0.35))
        fig_mon.update_xaxes(**axis_x(-30, tickfont=dict(size=11, color=COLORS["text_mid"])))
        fig_mon.update_yaxes(**axis_y(tickfont=dict(size=11, color=COLORS["text_mid"]), title=dict(text="Number of Cases", font=dict(size=10, color=COLORS["text_mid"]))))
        st.plotly_chart(fig_mon, use_container_width=True, key="chart_967_fig_mon")

    st.markdown("<div style='height:8px'></div>", unsafe_allow_html=True)

    col7, col8 = st.columns(2, gap="medium")
    with col7:
        st.markdown('<div class="sec-title">📊 Cumulative Case Growth</div>', unsafe_allow_html=True)
        cum = daily.copy()
        cum["Cumulative"] = cum["Cases"].cumsum()
        fig_cum = go.Figure(go.Scatter(
            x=cum["Date_Label"], y=cum["Cumulative"], mode="lines+markers", fill="tozeroy",
            line=dict(color=COLORS["dark_green"], width=2.5), marker=dict(size=6, color=COLORS["dark_green"]),
            fillcolor="rgba(31,164,99,0.15)",
            hovertemplate="<b>%{x}</b><br>Cumulative Cases: %{y:,}<extra></extra>",
        ))
        fig_cum.update_layout(**base_layout(height=320, margin=dict(l=50, r=30, t=20, b=50)))
        fig_cum.update_xaxes(**axis_x(-30 if len(cum) > 10 else 0, tickfont=dict(size=11, color=COLORS["text_mid"])),
                              type="category", categoryorder="array", categoryarray=cum["Date_Label"])
        fig_cum.update_yaxes(**axis_y(tickfont=dict(size=11, color=COLORS["text_mid"])), title=dict(text="Cumulative Cases", font=dict(size=10, color=COLORS["text_mid"])))
        st.plotly_chart(fig_cum, use_container_width=True, key="chart_986_fig_cum")

    with col8:
        st.markdown('<div class="sec-title">🧑‍💼 Top 10 Respondent Advocates</div>', unsafe_allow_html=True)
        ra = value_counts_df(df.loc[df["Respondent_Advocate"] != "", "Respondent_Advocate"], "Advocate").head(10)
        fig_ra = horizontal_bar_ranked(ra, "Advocate", "Cases", height=320, label_len=28, single_color=COLORS["orange"], right_margin=80)
        st.plotly_chart(fig_ra, use_container_width=True, key="chart_992_fig_ra")

    st.markdown("<div style='height:8px'></div>", unsafe_allow_html=True)

    st.markdown('<div class="sec-title">🗂️ Top Case Categories — Daily Trend</div>', unsafe_allow_html=True)
    cat_trend_df = df.dropna(subset=["Date"]).copy()
    if cat_trend_df.empty:
        st.info("No dated records available for the current filter.")
    else:
        cat_trend_df["Case_Category"] = cat_trend_df["Case_Category"].replace("", "Uncategorized")
        cat_trend_df["Date_Label"] = cat_trend_df["Date"].dt.strftime("%d %b")
        top_cats_trend = cat_trend_df["Case_Category"].value_counts().head(5).index.tolist()
        order_lbls = daily["Date_Label"].tolist()
        ct_line = (
            cat_trend_df[cat_trend_df["Case_Category"].isin(top_cats_trend)]
            .groupby(["Date", "Date_Label", "Case_Category"]).size()
            .reset_index(name="Cases").sort_values("Date")
        )
        line_colors2 = [COLORS["dark_green"], COLORS["gold"], COLORS["teal"], COLORS["red"], COLORS["purple"]]
        fig_ct = go.Figure()
        for i, cat in enumerate(top_cats_trend):
            cdata = ct_line[ct_line["Case_Category"] == cat]
            label = truncate(cat, 26)
            fig_ct.add_trace(go.Scatter(
                x=cdata["Date_Label"], y=cdata["Cases"], mode="lines+markers", name=label,
                line=dict(width=2.5, color=line_colors2[i % len(line_colors2)]), marker=dict(size=6),
                hovertemplate=f"<b>{label}</b><br>%{{x}}<br>Cases: %{{y}}<extra></extra>",
            ))
        fig_ct.update_layout(**base_layout(
            height=340, margin=dict(l=50, r=20, t=40, b=50),
            legend=dict(orientation="h", x=0.5, y=1.15, xanchor="center", yanchor="bottom", font=dict(size=10), bgcolor="rgba(0,0,0,0)"),
        ))
        fig_ct.update_xaxes(**axis_x(-30 if len(order_lbls) > 10 else 0, tickfont=dict(size=10, color=COLORS["text_mid"])),
                             type="category", categoryorder="array", categoryarray=order_lbls)
        fig_ct.update_yaxes(**axis_y(tickfont=dict(size=10, color=COLORS["text_mid"])), title=dict(text="Number of Cases", font=dict(size=10, color=COLORS["text_mid"])))
        st.plotly_chart(fig_ct, use_container_width=True, key="chart_1027_fig_ct")

    st.markdown("</div>", unsafe_allow_html=True)


# ═══════════════════════════════════════════════════════════════
# TAB 2 — DAILY CAUSE LIST
# ═══════════════════════════════════════════════════════════════
def tab_daily_cause_list(df: pd.DataFrame) -> None:
    st.markdown('<div class="content-area">', unsafe_allow_html=True)
    st.markdown('<div class="sec-title">📅 Daily Cause List</div>', unsafe_allow_html=True)

    col1, col2, col3, col4 = st.columns(4, gap="medium")
    with col1:
        sel_d = st.selectbox("📅 Date", ["All Dates"] + sorted(df["Date_Str"].unique(), reverse=True), key="dcl_date")
    with col2:
        sel_j = st.selectbox("👨‍⚖️ Judge", ["All Judges"] + sorted(df["Judge_Short"].unique()), key="dcl_judge")
    with col3:
        sel_s = st.selectbox("📁 Section", ["All Sections"] + sorted(df["Section_Clean"].unique()), key="dcl_sec")
    with col4:
        sel_c = st.selectbox("🗂 Category", ["All Categories"] + sorted(df["Case_Category"].replace("", "Uncategorized").unique()), key="dcl_cat")

    filt = df.copy()
    if sel_d != "All Dates":
        filt = filt[filt["Date_Str"] == sel_d]
    if sel_j != "All Judges":
        filt = filt[filt["Judge_Short"] == sel_j]
    if sel_s != "All Sections":
        filt = filt[filt["Section_Clean"] == sel_s]
    if sel_c != "All Categories":
        filt = filt[filt["Case_Category"].replace("", "Uncategorized") == sel_c]

    st.markdown(
        f"""
        <div class="info-box">
            📋 <b>{len(filt):,}</b> cases &nbsp;|&nbsp;
            👨‍⚖️ <b>{filt['Judge_Short'].nunique()}</b> judges active &nbsp;|&nbsp;
            📁 <b>{filt['Section_Clean'].nunique()}</b> sections &nbsp;|&nbsp;
            📅 <b>{filt['Date_Str'].nunique()}</b> date(s)
        </div>
        """,
        unsafe_allow_html=True,
    )
    st.markdown("<div style='height:4px'></div>", unsafe_allow_html=True)

    col_a, col_b = st.columns([1.3, 1], gap="medium")
    with col_a:
        st.markdown('<div class="sec-title">👨‍⚖️ Cases per Judge</div>', unsafe_allow_html=True)
        jc_filt = value_counts_df(filt["Judge_Short"], "Judge")
        fig_jf = horizontal_bar_ranked(
            jc_filt, "Judge", "Cases", height=380, label_len=34,
            gradient_desc=True, right_margin=60
        )
        st.plotly_chart(fig_jf, use_container_width=True, key="chart_1080_fig_jf")

    with col_b:
        st.markdown('<div class="sec-title">📋 Cases per Section</div>', unsafe_allow_html=True)
        sc_don_full = value_counts_df(filt["Section_Clean"], "Section").sort_values("Cases", ascending=False)
        sc_don = collapse_top_n(sc_don_full, "Section", "Cases", top_n=6, other_label="Other")
        colors = [COLORS["dark_green"], COLORS["gold"], COLORS["teal"], COLORS["green_acc"], COLORS["red"], COLORS["orange"], COLORS["text_light"]]
        fig_donut = donut_chart(sc_don, "Section", "Cases", height=380, label_len=28, colors=colors, hole=0.62)
        st.plotly_chart(fig_donut, use_container_width=True, key="chart_1088_fig_donut")

    st.markdown("<div style='height:8px'></div>", unsafe_allow_html=True)

    col_c, col_d = st.columns(2, gap="medium")
    with col_c:
        st.markdown('<div class="sec-title">🗂️ Cases per Category</div>', unsafe_allow_html=True)
        cc_filt = value_counts_df(filt["Case_Category"], "Category", blank_label="Uncategorized").head(10)
        fig_ccf = horizontal_bar_ranked(cc_filt, "Category", "Cases", height=340, label_len=30, single_color=COLORS["teal"], right_margin=60)
        st.plotly_chart(fig_ccf, use_container_width=True, key="chart_1097_fig_ccf")

    with col_d:
        st.markdown('<div class="sec-title">📅 Cases by Day of Week</div>', unsafe_allow_html=True)
        dow_order = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]
        dow_f = filt["Date"].dt.day_name().value_counts().reindex(dow_order).dropna().reset_index()
        dow_f.columns = ["Day", "Cases"]
        fig_dowf = px.bar(dow_f, x="Day", y="Cases", color_discrete_sequence=[COLORS["navy"]], text="Cases")
        fig_dowf.update_traces(marker_line_width=0, textposition="outside", textfont=dict(size=11, color=COLORS["text_dark"]),
                                hovertemplate="<b>%{x}</b><br>Cases: %{y:,}<extra></extra>")
        fig_dowf.update_layout(**base_layout(height=340, margin=dict(l=50, r=30, t=20, b=60), bargap=0.35))
        fig_dowf.update_xaxes(**axis_x(0, tickfont=dict(size=11, color=COLORS["text_mid"])))
        fig_dowf.update_yaxes(**axis_y(tickfont=dict(size=11, color=COLORS["text_mid"])), title=dict(text="Number of Cases", font=dict(size=10, color=COLORS["text_mid"])))
        st.plotly_chart(fig_dowf, use_container_width=True, key="chart_1110_fig_dowf")

    st.markdown("<div style='height:8px'></div>", unsafe_allow_html=True)

    col_e, col_f = st.columns(2, gap="medium")
    with col_e:
        st.markdown('<div class="sec-title">🧑‍💼 Top 10 Petitioner Advocates</div>', unsafe_allow_html=True)
        pa_f = value_counts_df(filt.loc[filt["Petitioner_Advocate"] != "", "Petitioner_Advocate"], "Advocate").head(10)
        fig_paf = horizontal_bar_ranked(pa_f, "Advocate", "Cases", height=340, label_len=28, single_color=COLORS["dark_green"], right_margin=60)
        st.plotly_chart(fig_paf, use_container_width=True, key="chart_1119_fig_paf")

    with col_f:
        st.markdown('<div class="sec-title">🧑‍💼 Top 10 Respondent Advocates</div>', unsafe_allow_html=True)
        ra_f = value_counts_df(filt.loc[filt["Respondent_Advocate"] != "", "Respondent_Advocate"], "Advocate").head(10)
        fig_raf = horizontal_bar_ranked(ra_f, "Advocate", "Cases", height=340, label_len=28, single_color=COLORS["orange"], right_margin=60)
        st.plotly_chart(fig_raf, use_container_width=True, key="chart_1125_fig_raf")

    st.markdown("<div style='height:8px'></div>", unsafe_allow_html=True)

    col_g, col_h = st.columns(2, gap="medium")
    with col_g:
        st.markdown('<div class="sec-title">📈 Cases Trend (Filtered)</div>', unsafe_allow_html=True)
        daily_f = filt.dropna(subset=["Date"]).groupby("Date").size().reset_index(name="Cases").sort_values("Date")
        if len(daily_f) < 2:
            st.info("Select more than one date to see a trend line.")
        else:
            daily_f["Date_Label"] = daily_f["Date"].dt.strftime("%d %b")
            fig_trendf = go.Figure(go.Bar(
                x=daily_f["Date_Label"], y=daily_f["Cases"], marker_color=COLORS["dark_green"], marker_line_width=0,
                hovertemplate="<b>%{x}</b><br>Cases: %{y:,}<extra></extra>",
            ))
            fig_trendf.update_layout(**base_layout(height=320, margin=dict(l=50, r=30, t=20, b=50), bargap=0.3))
            fig_trendf.update_xaxes(**axis_x(-30 if len(daily_f) > 10 else 0, tickfont=dict(size=11, color=COLORS["text_mid"])),
                                     type="category", categoryorder="array", categoryarray=daily_f["Date_Label"])
            fig_trendf.update_yaxes(**axis_y(tickfont=dict(size=11, color=COLORS["text_mid"])), title=dict(text="Number of Cases", font=dict(size=10, color=COLORS["text_mid"])))
            st.plotly_chart(fig_trendf, use_container_width=True, key="chart_1145_fig_trendf")

    with col_h:
        st.markdown('<div class="sec-title">👨‍⚖️ Active Judges (Filtered)</div>', unsafe_allow_html=True)
        jud_f = value_counts_df(filt["Judge_Short"], "Judge").head(8)
        colors_j = [COLORS["dark_green"], COLORS["gold"], COLORS["teal"], COLORS["green_acc"], COLORS["red"], COLORS["orange"], COLORS["purple"], COLORS["text_light"]]
        fig_jdon = donut_chart(jud_f, "Judge", "Cases", height=320, label_len=24, colors=colors_j, total_label="Cases", hole=0.6)
        st.plotly_chart(fig_jdon, use_container_width=True, key="chart_1152_fig_jdon")

    st.markdown("<div style='height:8px'></div>", unsafe_allow_html=True)

    col_i, col_j = st.columns(2, gap="medium")
    with col_i:
        st.markdown('<div class="sec-title">🗓️ Cases by Month (Filtered)</div>', unsafe_allow_html=True)
        month_order = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"]
        mon_f = filt["Month"].value_counts().reindex(month_order).dropna().reset_index()
        mon_f.columns = ["Month", "Cases"]
        fig_monf = px.bar(mon_f, x="Month", y="Cases", color_discrete_sequence=[COLORS["gold"]], text="Cases")
        fig_monf.update_traces(marker_line_width=0, textposition="outside", textfont=dict(size=11, color=COLORS["text_dark"]),
                                hovertemplate="<b>%{x}</b><br>Cases: %{y:,}<extra></extra>")
        fig_monf.update_layout(**base_layout(height=320, margin=dict(l=50, r=30, t=20, b=60), bargap=0.35))
        fig_monf.update_xaxes(**axis_x(-30, tickfont=dict(size=11, color=COLORS["text_mid"])))
        fig_monf.update_yaxes(**axis_y(tickfont=dict(size=11, color=COLORS["text_mid"])), title=dict(text="Number of Cases", font=dict(size=10, color=COLORS["text_mid"])))
        st.plotly_chart(fig_monf, use_container_width=True, key="chart_1168_fig_monf")

    with col_j:
        st.markdown('<div class="sec-title">🗂️ Cases per Category Share (Filtered)</div>', unsafe_allow_html=True)
        cc_don_f = collapse_top_n(cc_filt.sort_values("Cases", ascending=False), "Category", "Cases", top_n=6, other_label="Other")
        colors_ccd = [COLORS["navy"], COLORS["gold"], COLORS["teal"], COLORS["green_acc"], COLORS["red"], COLORS["orange"], COLORS["text_light"]]
        fig_ccdon = donut_chart(cc_don_f, "Category", "Cases", height=320, label_len=24, colors=colors_ccd, total_label="Cases")
        st.plotly_chart(fig_ccdon, use_container_width=True, key="chart_1175_fig_ccdon")

    st.markdown("<div style='height:8px'></div>", unsafe_allow_html=True)

    st.markdown('<div class="sec-title">📄 Full Cause List</div>', unsafe_allow_html=True)
    disp_cols = ["Sr_No", "Case_No", "Case_Category", "Petitioner", "Respondent",
                 "Petitioner_Advocate", "Respondent_Advocate", "Judge_Short", "Section_Clean", "Date_Str"]
    avail = [c for c in disp_cols if c in filt.columns]
    st.dataframe(
        filt[avail].rename(columns={
            "Judge_Short": "Judge", "Section_Clean": "Section", "Date_Str": "Date",
            "Petitioner_Advocate": "Pet. Advocate", "Respondent_Advocate": "Res. Advocate",
        }),
        use_container_width=True, height=460, hide_index=True,
    )
    st.download_button(
        "⬇  Export Filtered Data to CSV",
        filt[avail].to_csv(index=False).encode("utf-8"),
        f"SHC_DailyCauseList_{datetime.today().strftime('%d%b%Y')}.csv",
        "text/csv",
    )
    st.markdown("</div>", unsafe_allow_html=True)


# ═══════════════════════════════════════════════════════════════
# TAB 3 — JUDGE ANALYSIS
# ═══════════════════════════════════════════════════════════════
def build_judge_ranking(df: pd.DataFrame) -> pd.DataFrame:
    ranking = (
        df.groupby("Judge_Short")
        .agg(Cases=("Case_No", "count"), Sections=("Section_Clean", "nunique"),
             Categories=("Case_Category", "nunique"), Dates=("Date_Str", "nunique"))
        .reset_index()
        .sort_values("Cases", ascending=False)
        .reset_index(drop=True)
    )
    ranking.rename(columns={"Judge_Short": "Judge"}, inplace=True)
    ranking["Cases_Per_Day"] = (ranking["Cases"] / ranking["Dates"].replace(0, 1)).round(1)
    return ranking


def tab_judge_analysis(df: pd.DataFrame) -> None:
    st.markdown('<div class="content-area">', unsafe_allow_html=True)
    st.markdown('<div class="sec-title">Judge / Justice Analysis</div>', unsafe_allow_html=True)

    jrank = build_judge_ranking(df)

    col1, col2 = st.columns(2, gap="medium")
    with col1:
        st.markdown('<div class="sec-title">👨‍⚖️ Top 10 Judges by Caseload</div>', unsafe_allow_html=True)
        top10 = jrank.head(10)
        fig = horizontal_bar_ranked(top10, "Judge", "Cases", height=520, label_len=26, single_color=COLORS["dark_green"], right_margin=50)
        st.plotly_chart(fig, use_container_width=True, key="chart_1227_fig")

    with col2:
        st.markdown('<div class="sec-title">⚡ Workload Efficiency Bubble</div>', unsafe_allow_html=True)
        if len(jrank) > 3:
            wb = jrank.head(25).copy()
            wb["Judge_Disp"] = wb["Judge"].apply(lambda x: truncate(x, 22))
            fig2 = px.scatter(
                wb, x="Sections", y="Cases", size="Cases", color="Cases_Per_Day",
                hover_name="Judge_Disp", color_continuous_scale=[[0, COLORS["teal"]], [1, COLORS["dark_green"]]],
                size_max=40, labels={"Cases_Per_Day": "Cases / Day"},
            )
            fig2.update_traces(marker=dict(line=dict(width=1, color="#FFFFFF")),
                                hovertemplate="<b>%{hovertext}</b><br>Sections: %{x}<br>Cases: %{y}<extra></extra>")
            fig2.update_layout(**base_layout(
                height=520, margin=dict(l=50, r=20, t=10, b=50),
                coloraxis_colorbar=dict(title=dict(text="Cases/Day", font=dict(size=10)), thickness=12, len=0.6, tickfont=dict(size=9)),
            ))
            fig2.update_xaxes(**axis_x(tickfont=dict(size=10, color=COLORS["text_mid"])),
                              title=dict(text="Number of Sections Handled", font=dict(size=10, color=COLORS["text_mid"])), dtick=1)
            fig2.update_yaxes(**axis_y(), title=dict(text="Total Cases", font=dict(size=10, color=COLORS["text_mid"])))
            st.plotly_chart(fig2, use_container_width=True, key="chart_1248_fig2")
        else:
            st.info("Not enough judges in the current filter to plot a bubble chart.")

    st.markdown("<div style='height:8px'></div>", unsafe_allow_html=True)

    st.markdown('<div class="sec-title">📊 All Judges — Ranked by Caseload</div>', unsafe_allow_html=True)
    n_all = len(jrank)
    chart_h_all = min(max(n_all * 22, 400), 1200)
    fig_all = horizontal_bar_ranked(jrank, "Judge", "Cases", height=chart_h_all, label_len=36, gradient_desc=True, right_margin=60)
    fig_all.update_layout(bargap=0.25)
    st.plotly_chart(fig_all, use_container_width=True, key="chart_1259_fig_all")

    with st.expander("📄 View as sortable table"):
        jrank_tbl = jrank.sort_values("Cases", ascending=False).reset_index(drop=True)
        jrank_tbl.index = jrank_tbl.index + 1
        st.dataframe(jrank_tbl, use_container_width=True, height=420)

    st.markdown("<div style='height:8px'></div>", unsafe_allow_html=True)

    col_i, col_j = st.columns(2, gap="medium")
    with col_i:
        st.markdown('<div class="sec-title">📆 Avg. Cases / Day — Top 10 Judges</div>', unsafe_allow_html=True)
        top10_rate = jrank.sort_values("Cases_Per_Day", ascending=False).head(10)
        fig_rate = horizontal_bar_ranked(top10_rate, "Judge", "Cases_Per_Day", height=380, label_len=26, single_color=COLORS["gold"], right_margin=50)
        fig_rate.update_xaxes(title=dict(text="Cases per Day", font=dict(size=10, color=COLORS["text_mid"])))
        st.plotly_chart(fig_rate, use_container_width=True, key="chart_1274_fig_rate")

    with col_j:
        st.markdown('<div class="sec-title">📅 Days Active — Top 10 Judges</div>', unsafe_allow_html=True)
        top10_days = jrank.sort_values("Dates", ascending=False).head(10)
        fig_days = horizontal_bar_ranked(top10_days, "Judge", "Dates", height=380, label_len=26, single_color=COLORS["teal"], right_margin=50)
        fig_days.update_xaxes(title=dict(text="Days Active", font=dict(size=10, color=COLORS["text_mid"])))
        st.plotly_chart(fig_days, use_container_width=True, key="chart_1281_fig_days")

    st.markdown("<div style='height:8px'></div>", unsafe_allow_html=True)

    st.markdown('<div class="sec-title">📈 Caseload Trend — Top 5 Judges</div>', unsafe_allow_html=True)
    jt_df = df.dropna(subset=["Date"]).copy()
    if jt_df.empty:
        st.info("No dated records available for the current filter.")
    else:
        jt_df["Date_Label"] = jt_df["Date"].dt.strftime("%d %b")
        order_lbls_j = jt_df.drop_duplicates("Date").sort_values("Date")["Date_Label"].tolist()
        top5_judges = jrank.head(5)["Judge"].tolist()
        jt_line = (
            jt_df[jt_df["Judge_Short"].isin(top5_judges)]
            .groupby(["Date", "Date_Label", "Judge_Short"]).size()
            .reset_index(name="Cases").sort_values("Date")
        )
        line_colors3 = [COLORS["dark_green"], COLORS["gold"], COLORS["teal"], COLORS["red"], COLORS["purple"]]
        fig_jtrend = go.Figure()
        for i, jname in enumerate(top5_judges):
            jdata = jt_line[jt_line["Judge_Short"] == jname]
            label = truncate(jname, 26)
            fig_jtrend.add_trace(go.Scatter(
                x=jdata["Date_Label"], y=jdata["Cases"], mode="lines+markers", name=label,
                line=dict(width=2.5, color=line_colors3[i % len(line_colors3)]), marker=dict(size=6),
                hovertemplate=f"<b>{label}</b><br>%{{x}}<br>Cases: %{{y}}<extra></extra>",
            ))
        fig_jtrend.update_layout(**base_layout(
            height=340, margin=dict(l=50, r=20, t=40, b=50),
            legend=dict(orientation="h", x=0.5, y=1.15, xanchor="center", yanchor="bottom", font=dict(size=10), bgcolor="rgba(0,0,0,0)"),
        ))
        fig_jtrend.update_xaxes(**axis_x(-30 if len(order_lbls_j) > 10 else 0, tickfont=dict(size=10, color=COLORS["text_mid"])),
                                 type="category", categoryorder="array", categoryarray=order_lbls_j)
        fig_jtrend.update_yaxes(**axis_y(tickfont=dict(size=10, color=COLORS["text_mid"])), title=dict(text="Number of Cases", font=dict(size=10, color=COLORS["text_mid"])))
        st.plotly_chart(fig_jtrend, use_container_width=True, key="chart_1315_fig_jtrend")

    st.markdown("<div style='height:8px'></div>", unsafe_allow_html=True)

    st.markdown('<div class="sec-title">🧮 Top 10 Judges × Case Category</div>', unsafe_allow_html=True)
    top_j10 = jrank.head(10)["Judge"].tolist()
    jc_stack = df[df["Judge_Short"].isin(top_j10)].copy()
    jc_stack["Case_Category"] = jc_stack["Case_Category"].replace("", "Uncategorized")
    top_cats_stack = jc_stack["Case_Category"].value_counts().head(6).index.tolist()
    jc_stack = jc_stack[jc_stack["Case_Category"].isin(top_cats_stack)]
    jc_stack_g = jc_stack.groupby(["Judge_Short", "Case_Category"]).size().reset_index(name="Cases")
    jc_stack_g["Judge_Disp"] = jc_stack_g["Judge_Short"].apply(lambda x: truncate(x, 22))
    order_stack = jc_stack_g.groupby("Judge_Disp")["Cases"].sum().sort_values(ascending=False).index.tolist()
    fig_jcstack = px.bar(
        jc_stack_g, x="Judge_Disp", y="Cases", color="Case_Category", barmode="stack",
        category_orders={"Judge_Disp": order_stack},
        color_discrete_sequence=[COLORS["navy"], COLORS["gold"], COLORS["teal"], COLORS["green_acc"], COLORS["red"], COLORS["orange"]],
    )
    fig_jcstack.update_traces(marker_line_width=0)
    fig_jcstack.update_layout(**base_layout(
        height=420, margin=dict(l=50, r=10, t=40, b=110), bargap=0.25,
        legend=dict(orientation="h", x=0.5, y=1.12, xanchor="center", yanchor="bottom", font=dict(size=10), bgcolor="rgba(0,0,0,0)", title=None),
    ))
    fig_jcstack.update_xaxes(**axis_x(-35, tickfont=dict(size=10, color=COLORS["text_dark"])),
                              title=dict(text="Judge", font=dict(size=10, color=COLORS["text_mid"])), automargin=True)
    fig_jcstack.update_yaxes(**axis_y(tickfont=dict(size=10, color=COLORS["text_mid"])), title=dict(text="Number of Cases", font=dict(size=10, color=COLORS["text_mid"])))
    st.plotly_chart(fig_jcstack, use_container_width=True, key="chart_1341_fig_jcstack")

    st.markdown("<div style='height:8px'></div>", unsafe_allow_html=True)

    st.markdown('<div class="sec-title">📊 Judge Caseload Distribution</div>', unsafe_allow_html=True)
    fig_jhist = px.histogram(jrank, x="Cases", nbins=20, color_discrete_sequence=[COLORS["dark_green"]])
    fig_jhist.update_traces(marker_line_width=0, hovertemplate="Cases: %{x}<br>Judges: %{y}<extra></extra>")
    fig_jhist.update_layout(**base_layout(height=300, margin=dict(l=50, r=20, t=20, b=50), bargap=0.1))
    fig_jhist.update_xaxes(**axis_x(), title=dict(text="Cases Handled", font=dict(size=10, color=COLORS["text_mid"])))
    fig_jhist.update_yaxes(**axis_y(), title=dict(text="Number of Judges", font=dict(size=10, color=COLORS["text_mid"])))
    st.plotly_chart(fig_jhist, use_container_width=True, key="chart_1351_fig_jhist")

    st.markdown('<div class="sec-title">Individual Judge Deep Dive</div>', unsafe_allow_html=True)
    sel_jj = st.selectbox("Select Judge", sorted(df["Judge_Short"].unique()), key="jdd")
    jdf = df[df["Judge_Short"] == sel_jj]

    stats = [
        ("Total Cases", f"{len(jdf):,}", "Assigned hearings", "yellow", "cases", "#0F2E22"),
        ("Sections", str(jdf["Section_Clean"].nunique()), "Hearing types", "blue", "folder", "#16523C"),
        ("Categories", str(jdf["Case_Category"].nunique()), "Case types", "navy", "category", "#1FA463"),
        ("Days Active", str(jdf["Date_Str"].nunique()), "Days in data", "green", "calendar", "#2FBF71"),
    ]
    cols = st.columns(4, gap="small")
    for col, (label, value, sub, tone, icon, color) in zip(cols, stats):
        col.markdown(kpi_card_html(label, value, sub, tone, icon, color), unsafe_allow_html=True)

    colA, colB = st.columns(2, gap="medium")
    with colA:
        st.markdown('<div class="sec-title">🥧 By Section</div>', unsafe_allow_html=True)
        sc2_full = value_counts_df(jdf["Section_Clean"], "Section").sort_values("Cases", ascending=False)
        sc2 = collapse_top_n(sc2_full, "Section", "Cases", top_n=6, other_label="Other")
        colors = [COLORS["dark_green"], COLORS["gold"], COLORS["teal"], COLORS["green_acc"], COLORS["red"], COLORS["orange"], COLORS["text_light"]]
        fig3 = donut_chart(sc2, "Section", "Cases", height=340, label_len=24, colors=colors, total_label="Cases")
        fig3.update_layout(font=dict(family="Inter", color=COLORS["text_dark"], size=10),
                           legend=dict(orientation="h", x=0.5, y=-0.15, xanchor="center", yanchor="top", font=dict(size=9), bgcolor="rgba(0,0,0,0)"))
        fig3.update_layout(annotations=[dict(
            text=f"<b>{sc2['Cases'].sum():,}</b><br><span style='font-size:9px;color:{COLORS['text_mid']}'>Cases</span>",
            x=0.5, y=0.5, font=dict(size=16, color=COLORS["text_dark"]), showarrow=False)])
        st.plotly_chart(fig3, use_container_width=True, key="chart_1379_fig3")

    with colB:
        st.markdown('<div class="sec-title">📊 By Category</div>', unsafe_allow_html=True)
        cc2 = value_counts_df(jdf["Case_Category"], "Category", blank_label="Uncategorized").head(10)
        fig4 = horizontal_bar_ranked(cc2, "Category", "Cases", height=340, label_len=22, single_color=COLORS["teal"], right_margin=40, x_headroom=1.2)
        fig4.update_layout(font=dict(family="Inter", color=COLORS["text_dark"], size=10))
        st.plotly_chart(fig4, use_container_width=True, key="chart_1386_fig4")

    disp = ["Sr_No", "Case_No", "Case_Category", "Petitioner", "Respondent",
            "Petitioner_Advocate", "Respondent_Advocate", "Section_Clean", "Date_Str"]
    avail = [c for c in disp if c in jdf.columns]
    st.dataframe(
        jdf[avail].rename(columns={"Section_Clean": "Section", "Date_Str": "Date",
                                    "Petitioner_Advocate": "Pet. Adv", "Respondent_Advocate": "Res. Adv"}),
        use_container_width=True, height=360, hide_index=True,
    )
    st.markdown("</div>", unsafe_allow_html=True)


# ═══════════════════════════════════════════════════════════════
# TAB 4 — COURT INFRASTRUCTURE
# ═══════════════════════════════════════════════════════════════
def tab_court_infrastructure(df: pd.DataFrame) -> None:
    st.markdown('<div class="content-area">', unsafe_allow_html=True)
    st.markdown('<div class="sec-title">Court Infrastructure</div>', unsafe_allow_html=True)

    col1, col2 = st.columns(2, gap="medium")
    with col1:
        st.markdown('<div class="sec-title">📁 Cases per Section (Top 10)</div>', unsafe_allow_html=True)
        sec_c_full = value_counts_df(df["Section_Clean"], "Section")
        sec_c = collapse_top_n(sec_c_full, "Section", "Cases", top_n=10)
        sec_c["Pct"] = (sec_c["Cases"] / sec_c["Cases"].sum() * 100).round(1)
        sec_c["Label"] = sec_c.apply(lambda r: f"{r['Cases']:,}  ({r['Pct']}%)", axis=1)
        n_sc = len(sec_c)
        fig = horizontal_bar_ranked(
            sec_c, "Section", "Cases", height=min(max(n_sc * 40, 360), 500), label_len=32,
            single_color=COLORS["dark_green"], value_labels=sec_c.sort_values("Cases")["Label"],
            right_margin=195, x_headroom=1.3,
        )
        st.plotly_chart(fig, use_container_width=True, key="chart_1419_fig")

    with col2:
        st.markdown('<div class="sec-title">🥧 Section Share</div>', unsafe_allow_html=True)
        sec_p_full = sec_c.sort_values("Cases", ascending=False)[["Section", "Cases"]]
        sec_p = collapse_top_n(sec_p_full, "Section", "Cases", top_n=6, other_label="Other")
        colors = [COLORS["navy"], COLORS["gold"], COLORS["teal"], COLORS["green_acc"], COLORS["red"], COLORS["orange"], COLORS["text_light"]]
        fig2 = donut_chart(sec_p, "Section", "Cases", height=420, label_len=26, colors=colors, total_label="Total")
        fig2.update_layout(font=dict(family="Inter", color=COLORS["text_dark"], size=10),
                           legend=dict(orientation="h", x=0.5, y=-0.1, xanchor="center", yanchor="top", font=dict(size=9), bgcolor="rgba(0,0,0,0)"))
        fig2.update_layout(annotations=[dict(
            text=f"<b>{sec_p['Cases'].sum():,}</b><br><span style='font-size:9px;color:{COLORS['text_mid']}'>Total</span>",
            x=0.5, y=0.5, font=dict(size=18, color=COLORS["text_dark"]), showarrow=False)])
        st.plotly_chart(fig2, use_container_width=True, key="chart_1432_fig2")

    st.markdown("<div style='height:8px'></div>", unsafe_allow_html=True)

    st.markdown('<div class="sec-title">🗺️ Case Category × Section Heatmap</div>', unsafe_allow_html=True)
    top_cats = df["Case_Category"].replace("", "Uncategorized").value_counts().head(12).index.tolist()
    heat_df = df.copy()
    heat_df["Case_Category"] = heat_df["Case_Category"].replace("", "Uncategorized")
    heat_df = heat_df[heat_df["Case_Category"].isin(top_cats)]
    pivot = heat_df.groupby(["Section_Clean", "Case_Category"]).size().reset_index(name="Count")
    pivot_t = pivot.pivot(index="Section_Clean", columns="Case_Category", values="Count").fillna(0)
    pivot_t.index = [truncate(s, 34) for s in pivot_t.index]
    pivot_t.columns = [truncate(c, 20) for c in pivot_t.columns]

    heat_h = min(max(len(pivot_t.index) * 42, 420), 650)
    fig3 = px.imshow(pivot_t, color_continuous_scale=[[0, "#F0FAF4"], [0.5, COLORS["teal"]], [1, COLORS["dark_green"]]], aspect="auto", text_auto=True)
    fig3.update_traces(textfont=dict(size=11, color=COLORS["text_dark"]),
                        hovertemplate="Section: %{y}<br>Category: %{x}<br>Cases: %{z}<extra></extra>")
    fig3.update_layout(**base_layout(
        height=heat_h, margin=dict(l=10, r=10, t=10, b=110),
        coloraxis_colorbar=dict(title=dict(text="Cases", font=dict(size=10)), thickness=14, len=0.75, tickfont=dict(size=9)),
    ))
    fig3.update_xaxes(tickfont=dict(size=10, color=COLORS["text_dark"]), tickangle=-45, side="bottom", automargin=True)
    fig3.update_yaxes(tickfont=dict(size=10, color=COLORS["text_dark"]), automargin=True)
    st.plotly_chart(fig3, use_container_width=True, key="chart_1456_fig3")

    st.markdown("<div style='height:8px'></div>", unsafe_allow_html=True)

    st.markdown('<div class="sec-title">👨‍⚖️ Judge × Section Distribution</div>', unsafe_allow_html=True)
    top_j = df["Judge_Short"].value_counts().head(15).index.tolist()
    j_sec = df[df["Judge_Short"].isin(top_j)].groupby(["Judge_Short", "Section_Clean"]).size().reset_index(name="Cases")
    top_sections_j = j_sec.groupby("Section_Clean")["Cases"].sum().sort_values(ascending=False).head(8).index.tolist()
    j_sec["Section_Clean"] = j_sec["Section_Clean"].where(j_sec["Section_Clean"].isin(top_sections_j), "Other")
    j_sec = j_sec.groupby(["Judge_Short", "Section_Clean"], as_index=False)["Cases"].sum()
    j_sec["Judge_Disp"] = j_sec["Judge_Short"].apply(lambda x: truncate(x, 22))
    order_j = j_sec.groupby("Judge_Disp")["Cases"].sum().sort_values(ascending=False).index.tolist()

    fig4 = px.bar(
        j_sec, x="Judge_Disp", y="Cases", color="Section_Clean", barmode="stack",
        category_orders={"Judge_Disp": order_j},
        color_discrete_sequence=[COLORS["navy"], COLORS["gold"], COLORS["teal"], COLORS["green_acc"], COLORS["red"], COLORS["orange"], COLORS["purple"], "#E91E8C"],
    )
    fig4.update_traces(marker_line_width=0)
    fig4.update_layout(**base_layout(
        height=620, margin=dict(l=50, r=10, t=40, b=140), bargap=0.25,
        legend=dict(orientation="h", x=0.5, y=1.1, xanchor="center", yanchor="bottom", font=dict(size=10), bgcolor="rgba(0,0,0,0)", title=None),
    ))
    fig4.update_xaxes(**axis_x(-40, tickfont=dict(size=10, color=COLORS["text_dark"])),
                      title=dict(text="Judge", font=dict(size=10, color=COLORS["text_mid"])), automargin=True)
    fig4.update_yaxes(**axis_y(tickfont=dict(size=10, color=COLORS["text_mid"])), title=dict(text="Number of Cases", font=dict(size=10, color=COLORS["text_mid"])))
    st.plotly_chart(fig4, use_container_width=True, key="chart_1479_fig4")

    st.markdown("<div style='height:8px'></div>", unsafe_allow_html=True)

    st.markdown('<div class="sec-title">📈 Cases by Month per Section</div>', unsafe_allow_html=True)
    month_df = df.dropna(subset=["Date"]).copy()
    if month_df.empty:
        st.info("No dated records available for the current filter.")
    else:
        month_df["Month_Period"] = month_df["Date"].dt.to_period("M")
        month_df["Month_Label"] = month_df["Date"].dt.strftime("%b %Y")
        top_sections_line = month_df["Section_Clean"].value_counts().head(6).index.tolist()
        m_line = (
            month_df[month_df["Section_Clean"].isin(top_sections_line)]
            .groupby(["Month_Period", "Month_Label", "Section_Clean"]).size()
            .reset_index(name="Cases").sort_values("Month_Period")
        )
        line_colors = [COLORS["dark_green"], COLORS["gold"], COLORS["teal"], COLORS["green_acc"], COLORS["red"], COLORS["orange"]]
        fig5 = go.Figure()
        for i, sect in enumerate(top_sections_line):
            sdata = m_line[m_line["Section_Clean"] == sect]
            label = truncate(sect, 28)
            fig5.add_trace(go.Scatter(
                x=sdata["Month_Label"], y=sdata["Cases"], mode="lines+markers", name=label,
                line=dict(width=2.5, color=line_colors[i % len(line_colors)]), marker=dict(size=6),
                hovertemplate=f"<b>{label}</b><br>%{{x}}<br>Cases: %{{y}}<extra></extra>",
            ))
        fig5.update_layout(**base_layout(
            height=380, margin=dict(l=50, r=20, t=40, b=60),
            legend=dict(orientation="h", x=0.5, y=1.15, xanchor="center", yanchor="bottom", font=dict(size=10), bgcolor="rgba(0,0,0,0)"),
        ))
        fig5.update_xaxes(**axis_x(-30, tickfont=dict(size=10, color=COLORS["text_mid"])), title=dict(text="Month", font=dict(size=10, color=COLORS["text_mid"])))
        fig5.update_yaxes(**axis_y(tickfont=dict(size=10, color=COLORS["text_mid"])), title=dict(text="Number of Cases", font=dict(size=10, color=COLORS["text_mid"])))
        st.plotly_chart(fig5, use_container_width=True, key="chart_1512_fig5")

    st.markdown("<div style='height:8px'></div>", unsafe_allow_html=True)

    col_k, col_l = st.columns(2, gap="medium")
    with col_k:
        st.markdown('<div class="sec-title">🗂️ Cases by Category (Top 15)</div>', unsafe_allow_html=True)
        cat_full = value_counts_df(df["Case_Category"], "Category", blank_label="Uncategorized")
        cat15 = collapse_top_n(cat_full, "Category", "Cases", top_n=15)
        n_cat15 = len(cat15)
        fig_cat15 = horizontal_bar_ranked(cat15, "Category", "Cases", height=min(max(n_cat15 * 32, 380), 520), label_len=30, single_color=COLORS["navy"], right_margin=60)
        st.plotly_chart(fig_cat15, use_container_width=True, key="chart_1523_fig_cat15")

    with col_l:
        st.markdown('<div class="sec-title">🥧 Case Category Share</div>', unsafe_allow_html=True)
        cat_don = collapse_top_n(cat_full.sort_values("Cases", ascending=False), "Category", "Cases", top_n=6, other_label="Other")
        colors_cd = [COLORS["navy"], COLORS["gold"], COLORS["teal"], COLORS["green_acc"], COLORS["red"], COLORS["orange"], COLORS["text_light"]]
        fig_catdon = donut_chart(cat_don, "Category", "Cases", height=min(max(n_cat15 * 32, 380), 520), label_len=26, colors=colors_cd, total_label="Total")
        st.plotly_chart(fig_catdon, use_container_width=True, key="chart_1530_fig_catdon")

    st.markdown("<div style='height:8px'></div>", unsafe_allow_html=True)

    col_m, col_n = st.columns(2, gap="medium")
    with col_m:
        st.markdown('<div class="sec-title">📅 Days Active — Top 10 Sections</div>', unsafe_allow_html=True)
        sec_days = df.dropna(subset=["Date"]).groupby("Section_Clean")["Date_Str"].nunique().reset_index()
        sec_days.columns = ["Section", "Days"]
        sec_days = sec_days.sort_values("Days", ascending=False).head(10)
        fig_secdays = horizontal_bar_ranked(sec_days, "Section", "Days", height=380, label_len=30, single_color=COLORS["teal"], right_margin=50)
        fig_secdays.update_xaxes(title=dict(text="Days Active", font=dict(size=10, color=COLORS["text_mid"])))
        st.plotly_chart(fig_secdays, use_container_width=True, key="chart_1542_fig_secdays")

    with col_n:
        st.markdown('<div class="sec-title">⚖️ Avg. Cases / Day — Top 10 Sections</div>', unsafe_allow_html=True)
        sec_rate = df.dropna(subset=["Date"]).groupby("Section_Clean").agg(Cases=("Case_No", "count"), Days=("Date_Str", "nunique")).reset_index()
        sec_rate["Rate"] = (sec_rate["Cases"] / sec_rate["Days"].replace(0, 1)).round(1)
        sec_rate = sec_rate.rename(columns={"Section_Clean": "Section"}).sort_values("Rate", ascending=False).head(10)
        fig_secrate = horizontal_bar_ranked(sec_rate, "Section", "Rate", height=380, label_len=30, single_color=COLORS["gold"], right_margin=50)
        fig_secrate.update_xaxes(title=dict(text="Cases per Day", font=dict(size=10, color=COLORS["text_mid"])))
        st.plotly_chart(fig_secrate, use_container_width=True, key="chart_1551_fig_secrate")

    st.markdown("<div style='height:8px'></div>", unsafe_allow_html=True)

    st.markdown('<div class="sec-title">👨‍⚖️ Unique Judges per Section (Top 10)</div>', unsafe_allow_html=True)
    sec_judges = df.groupby("Section_Clean")["Judge_Short"].nunique().reset_index()
    sec_judges.columns = ["Section", "Judges"]
    sec_judges = sec_judges.sort_values("Judges", ascending=False).head(10)
    fig_secj = horizontal_bar_ranked(sec_judges, "Section", "Judges", height=380, label_len=32, single_color=COLORS["purple"], right_margin=50)
    fig_secj.update_xaxes(title=dict(text="Unique Judges", font=dict(size=10, color=COLORS["text_mid"])))
    st.plotly_chart(fig_secj, use_container_width=True, key="chart_1561_fig_secj")

    st.markdown("</div>", unsafe_allow_html=True)


# ═══════════════════════════════════════════════════════════════
# TAB 5 — LAWYER INTELLIGENCE
# ═══════════════════════════════════════════════════════════════
def tab_lawyer_intelligence(df: pd.DataFrame) -> None:
    st.markdown('<div class="content-area">', unsafe_allow_html=True)
    st.markdown('<div class="sec-title">Lawyer Intelligence</div>', unsafe_allow_html=True)

    pet_adv = value_counts_df(df.loc[df["Petitioner_Advocate"] != "", "Petitioner_Advocate"], "Advocate")
    res_adv = value_counts_df(df.loc[df["Respondent_Advocate"] != "", "Respondent_Advocate"], "Advocate")
    all_adv = pd.concat([pet_adv.assign(Role="Petitioner"), res_adv.assign(Role="Respondent")])
    total_adv = all_adv.groupby("Advocate")["Cases"].sum().reset_index().sort_values("Cases", ascending=False)

    col1, col2 = st.columns(2)
    with col1:
        st.markdown('<div class="sec-title">🧑‍💼 Top 10 Petitioner Advocates</div>', unsafe_allow_html=True)
        tp = pet_adv.head(10).copy()
        tp["Advocate"] = tp["Advocate"].apply(lambda x: truncate(x, 28))
        fig = px.bar(tp, x="Cases", y="Advocate", orientation="h", color_discrete_sequence=[COLORS["dark_green"]], text="Cases")
        fig.update_traces(marker_line_width=0, textposition="outside", textfont_size=8)
        fig.update_layout(**base_layout(height=480))
        fig.update_xaxes(**axis_x())
        fig.update_yaxes(**axis_y(), autorange="reversed")
        st.plotly_chart(fig, use_container_width=True, key="chart_1588_fig")

    with col2:
        st.markdown('<div class="sec-title">🧑‍💼 Top 10 Respondent Advocates</div>', unsafe_allow_html=True)
        tr = res_adv.head(10).copy()
        tr["Advocate"] = tr["Advocate"].apply(lambda x: truncate(x, 28))
        fig2 = px.bar(tr, x="Cases", y="Advocate", orientation="h", color_discrete_sequence=[COLORS["teal"]], text="Cases")
        fig2.update_traces(marker_line_width=0, textposition="outside", textfont_size=8)
        fig2.update_layout(**base_layout(height=480))
        fig2.update_xaxes(**axis_x())
        fig2.update_yaxes(**axis_y(), autorange="reversed")
        st.plotly_chart(fig2, use_container_width=True, key="chart_1599_fig2")

    st.markdown('<div class="sec-title">Lawyer Search — Find All Hearings by Advocate Name</div>', unsafe_allow_html=True)
    lq = st.text_input("", placeholder="Type advocate name...", key="lsearch", label_visibility="collapsed")
    if lq.strip():
        q = lq.strip().lower()
        lmask = (
            df["Petitioner_Advocate"].str.lower().str.contains(q, na=False)
            | df["Respondent_Advocate"].str.lower().str.contains(q, na=False)
        )
        lr = df[lmask]
        st.markdown(f'<div class="info-box">🧑‍💼 <b>{len(lr)}</b> hearing(s) found for "<b>{lq}</b>"</div>', unsafe_allow_html=True)
        if not lr.empty:
            disp = ["Date_Str", "Case_No", "Case_Category", "Petitioner", "Respondent",
                    "Petitioner_Advocate", "Respondent_Advocate", "Judge_Short", "Section_Clean"]
            avail = [c for c in disp if c in lr.columns]
            st.dataframe(
                lr[avail].rename(columns={"Date_Str": "Date", "Judge_Short": "Judge", "Section_Clean": "Section",
                                          "Petitioner_Advocate": "Pet. Adv", "Respondent_Advocate": "Res. Adv"}),
                use_container_width=True, height=350, hide_index=True,
            )

    st.markdown('<div class="sec-title">🏆 Top 10 Advocates — Combined Activity</div>', unsafe_allow_html=True)
    t10 = total_adv.head(10).copy()
    t10["Advocate"] = t10["Advocate"].apply(lambda x: truncate(x, 35))
    fig3 = px.bar(t10, x="Advocate", y="Cases", color_discrete_sequence=[COLORS["navy"]], text="Cases")
    fig3.update_traces(marker_line_width=0, textposition="outside", textfont_size=8)
    fig3.update_layout(**base_layout(height=340, margin=dict(l=50, r=20, t=20, b=110)))
    fig3.update_xaxes(**axis_x(-40), automargin=True)
    fig3.update_yaxes(**axis_y())
    st.plotly_chart(fig3, use_container_width=True, key="chart_1629_fig3")

    st.markdown("<div style='height:8px'></div>", unsafe_allow_html=True)

    col3, col4 = st.columns(2, gap="medium")
    with col3:
        st.markdown('<div class="sec-title">🥧 Petitioner vs Respondent Role Split</div>', unsafe_allow_html=True)
        role_split = all_adv.groupby("Role")["Cases"].sum().reset_index()
        fig_role = donut_chart(role_split, "Role", "Cases", height=340, label_len=20,
                                colors=[COLORS["dark_green"], COLORS["teal"]], total_label="Total")
        st.plotly_chart(fig_role, use_container_width=True, key="chart_1639_fig_role")

    with col4:
        st.markdown('<div class="sec-title">👨‍⚖️ Top 10 Advocates — Unique Judges Appeared Before</div>', unsafe_allow_html=True)
        adv_long = pd.concat([
            df.loc[df["Petitioner_Advocate"] != "", ["Petitioner_Advocate", "Judge_Short"]].rename(columns={"Petitioner_Advocate": "Advocate"}),
            df.loc[df["Respondent_Advocate"] != "", ["Respondent_Advocate", "Judge_Short"]].rename(columns={"Respondent_Advocate": "Advocate"}),
        ])
        adv_judges = adv_long.groupby("Advocate")["Judge_Short"].nunique().reset_index()
        adv_judges.columns = ["Advocate", "Judges"]
        adv_judges = adv_judges.sort_values("Judges", ascending=False).head(10)
        fig_advj = horizontal_bar_ranked(adv_judges, "Advocate", "Judges", height=340, label_len=28, single_color=COLORS["purple"], right_margin=50)
        fig_advj.update_xaxes(title=dict(text="Unique Judges", font=dict(size=10, color=COLORS["text_mid"])))
        st.plotly_chart(fig_advj, use_container_width=True, key="chart_1652_fig_advj")

    st.markdown("<div style='height:8px'></div>", unsafe_allow_html=True)

    col5, col6 = st.columns(2, gap="medium")
    with col5:
        st.markdown('<div class="sec-title">🗂️ Top 10 Advocates — Unique Case Categories</div>', unsafe_allow_html=True)
        adv_long2 = pd.concat([
            df.loc[df["Petitioner_Advocate"] != "", ["Petitioner_Advocate", "Case_Category"]].rename(columns={"Petitioner_Advocate": "Advocate"}),
            df.loc[df["Respondent_Advocate"] != "", ["Respondent_Advocate", "Case_Category"]].rename(columns={"Respondent_Advocate": "Advocate"}),
        ])
        adv_long2["Case_Category"] = adv_long2["Case_Category"].replace("", "Uncategorized")
        adv_cats = adv_long2.groupby("Advocate")["Case_Category"].nunique().reset_index()
        adv_cats.columns = ["Advocate", "Categories"]
        adv_cats = adv_cats.sort_values("Categories", ascending=False).head(10)
        fig_advc = horizontal_bar_ranked(adv_cats, "Advocate", "Categories", height=340, label_len=28, single_color=COLORS["red"], right_margin=50)
        fig_advc.update_xaxes(title=dict(text="Unique Categories", font=dict(size=10, color=COLORS["text_mid"])))
        st.plotly_chart(fig_advc, use_container_width=True, key="chart_1669_fig_advc")

    with col6:
        st.markdown('<div class="sec-title">📊 Advocate Caseload Distribution</div>', unsafe_allow_html=True)
        fig_hist = px.histogram(total_adv, x="Cases", nbins=20, color_discrete_sequence=[COLORS["gold"]])
        fig_hist.update_traces(marker_line_width=0, hovertemplate="Cases: %{x}<br>Advocates: %{y}<extra></extra>")
        fig_hist.update_layout(**base_layout(height=340, margin=dict(l=50, r=20, t=20, b=50), bargap=0.1))
        fig_hist.update_xaxes(**axis_x(), title=dict(text="Cases Handled", font=dict(size=10, color=COLORS["text_mid"])))
        fig_hist.update_yaxes(**axis_y(), title=dict(text="Number of Advocates", font=dict(size=10, color=COLORS["text_mid"])))
        st.plotly_chart(fig_hist, use_container_width=True, key="chart_1678_fig_hist")

    st.markdown("<div style='height:8px'></div>", unsafe_allow_html=True)

    st.markdown('<div class="sec-title">📈 Top 5 Advocates — Activity Trend</div>', unsafe_allow_html=True)
    adv_trend_df = df.dropna(subset=["Date"]).copy()
    if adv_trend_df.empty:
        st.info("No dated records available for the current filter.")
    else:
        adv_trend_df["Date_Label"] = adv_trend_df["Date"].dt.strftime("%d %b")
        order_lbls_a = adv_trend_df.drop_duplicates("Date").sort_values("Date")["Date_Label"].tolist()
        top5_adv = total_adv.head(5)["Advocate"].tolist()
        adv_long3 = pd.concat([
            adv_trend_df.loc[adv_trend_df["Petitioner_Advocate"] != "", ["Date", "Date_Label", "Petitioner_Advocate"]].rename(columns={"Petitioner_Advocate": "Advocate"}),
            adv_trend_df.loc[adv_trend_df["Respondent_Advocate"] != "", ["Date", "Date_Label", "Respondent_Advocate"]].rename(columns={"Respondent_Advocate": "Advocate"}),
        ])
        adv_line = (
            adv_long3[adv_long3["Advocate"].isin(top5_adv)]
            .groupby(["Date", "Date_Label", "Advocate"]).size()
            .reset_index(name="Cases").sort_values("Date")
        )
        line_colors4 = [COLORS["dark_green"], COLORS["gold"], COLORS["teal"], COLORS["red"], COLORS["purple"]]
        fig_advtrend = go.Figure()
        for i, aname in enumerate(top5_adv):
            adata = adv_line[adv_line["Advocate"] == aname]
            label = truncate(aname, 28)
            fig_advtrend.add_trace(go.Scatter(
                x=adata["Date_Label"], y=adata["Cases"], mode="lines+markers", name=label,
                line=dict(width=2.5, color=line_colors4[i % len(line_colors4)]), marker=dict(size=6),
                hovertemplate=f"<b>{label}</b><br>%{{x}}<br>Cases: %{{y}}<extra></extra>",
            ))
        fig_advtrend.update_layout(**base_layout(
            height=340, margin=dict(l=50, r=20, t=40, b=50),
            legend=dict(orientation="h", x=0.5, y=1.15, xanchor="center", yanchor="bottom", font=dict(size=10), bgcolor="rgba(0,0,0,0)"),
        ))
        fig_advtrend.update_xaxes(**axis_x(-30 if len(order_lbls_a) > 10 else 0, tickfont=dict(size=10, color=COLORS["text_mid"])),
                                   type="category", categoryorder="array", categoryarray=order_lbls_a)
        fig_advtrend.update_yaxes(**axis_y(tickfont=dict(size=10, color=COLORS["text_mid"])), title=dict(text="Number of Cases", font=dict(size=10, color=COLORS["text_mid"])))
        st.plotly_chart(fig_advtrend, use_container_width=True, key="chart_1716_fig_advtrend")

    st.markdown("<div style='height:8px'></div>", unsafe_allow_html=True)

    col7, col8 = st.columns(2, gap="medium")
    with col7:
        st.markdown('<div class="sec-title">📁 Top 10 Advocates — Unique Sections</div>', unsafe_allow_html=True)
        adv_long4 = pd.concat([
            df.loc[df["Petitioner_Advocate"] != "", ["Petitioner_Advocate", "Section_Clean"]].rename(columns={"Petitioner_Advocate": "Advocate"}),
            df.loc[df["Respondent_Advocate"] != "", ["Respondent_Advocate", "Section_Clean"]].rename(columns={"Respondent_Advocate": "Advocate"}),
        ])
        adv_secs = adv_long4.groupby("Advocate")["Section_Clean"].nunique().reset_index()
        adv_secs.columns = ["Advocate", "Sections"]
        adv_secs = adv_secs.sort_values("Sections", ascending=False).head(10)
        fig_advs = horizontal_bar_ranked(adv_secs, "Advocate", "Sections", height=340, label_len=28, single_color=COLORS["green_acc"], right_margin=50)
        fig_advs.update_xaxes(title=dict(text="Unique Sections", font=dict(size=10, color=COLORS["text_mid"])))
        st.plotly_chart(fig_advs, use_container_width=True, key="chart_1732_fig_advs")

    with col8:
        st.markdown('<div class="sec-title">🏆 Top Advocate — Category Breakdown</div>', unsafe_allow_html=True)
        top_adv_name = total_adv.iloc[0]["Advocate"] if not total_adv.empty else None
        if top_adv_name:
            ta_df = pd.concat([
                df.loc[(df["Petitioner_Advocate"] == top_adv_name), ["Case_Category"]],
                df.loc[(df["Respondent_Advocate"] == top_adv_name), ["Case_Category"]],
            ])
            ta_cats = value_counts_df(ta_df["Case_Category"], "Category", blank_label="Uncategorized").head(6)
            colors_ta = [COLORS["dark_green"], COLORS["gold"], COLORS["teal"], COLORS["red"], COLORS["orange"], COLORS["purple"]]
            fig_tacat = donut_chart(ta_cats, "Category", "Cases", height=340, label_len=24, colors=colors_ta, total_label=truncate(top_adv_name, 18))
            st.plotly_chart(fig_tacat, use_container_width=True, key="chart_1745_fig_tacat")
        else:
            st.info("No advocate data available.")

    st.markdown("</div>", unsafe_allow_html=True)


# ═══════════════════════════════════════════════════════════════
# TAB 6 — CASE SEARCH
# ═══════════════════════════════════════════════════════════════
def render_result_card(row: pd.Series, is_repeated: bool) -> str:
    rep_tag = '<span class="rbadge">🔁 Repeated</span>' if is_repeated else ""
    return f"""
    <div class="res-card">
        <div style='display:flex;justify-content:space-between;align-items:flex-start;'>
            <div style='flex:1;'>
                <div class="res-case">{row['Case_No']}</div>
                <div class="res-title">
                    {row['Petitioner'] or '—'}
                    <span style='color:#8FAFA0;font-weight:400;font-size:0.8rem;'> vs </span>
                    {row['Respondent'] or '—'}
                </div>
                <div class="res-meta" style='margin-top:0.3rem;'>
                    <span class='badge'>{row['Case_Category'] or 'N/A'}</span>{rep_tag}
                    <span style='margin-right:0.7rem;'>⚖️ {row['Judge_Short']}</span>
                    <span style='margin-right:0.7rem;'>📁 {row['Section_Clean']}</span>
                    <span>📅 {row['Date_Str']}</span>
                </div>
                <div class="res-meta" style='margin-top:0.2rem;'>
                    🧑‍💼 Pet. Adv: <b>{row['Petitioner_Advocate'] or '—'}</b>
                    &nbsp;|&nbsp;
                    Res. Adv: <b>{row['Respondent_Advocate'] or '—'}</b>
                </div>
            </div>
            <div style='font-size:1.2rem;opacity:0.25;margin-left:1rem;'>📄</div>
        </div>
    </div>"""


def tab_case_search(df: pd.DataFrame) -> None:
    st.markdown('<div class="content-area">', unsafe_allow_html=True)
    st.markdown('<div class="sec-title">Case Search & Lookup</div>', unsafe_allow_html=True)

    col1, col2 = st.columns([2.5, 1])
    with col1:
        query = st.text_input("", placeholder="Search by case number, party name, advocate, or judge...",
                              key="csearch", label_visibility="collapsed")
    with col2:
        field = st.selectbox(
            "Search in",
            ["All Fields", "Case Number", "Petitioner", "Respondent", "Petitioner Advocate", "Respondent Advocate", "Judge"],
            key="cfield", label_visibility="collapsed",
        )

    if not query.strip():
        st.markdown(
            """
            <div class="empty">
                <div class="ei">🔍</div>
                <div class="et">Ready to Search</div>
                <div class="es">Enter a case number, party name, advocate name, or judge above.</div>
            </div>""",
            unsafe_allow_html=True,
        )
        st.markdown("</div>", unsafe_allow_html=True)
        return

    q = query.strip().lower()
    field_map = {
        "Case Number": df["Case_No"].str.lower().str.contains(q, na=False),
        "Petitioner": df["Petitioner"].str.lower().str.contains(q, na=False),
        "Respondent": df["Respondent"].str.lower().str.contains(q, na=False),
        "Petitioner Advocate": df["Petitioner_Advocate"].str.lower().str.contains(q, na=False),
        "Respondent Advocate": df["Respondent_Advocate"].str.lower().str.contains(q, na=False),
        "Judge": df["Judge_Short"].str.lower().str.contains(q, na=False),
    }
    mask = pd.concat(list(field_map.values()), axis=1).any(axis=1) if field == "All Fields" else field_map[field]
    results = df[mask]
    repeated_cases = set(df["Case_No"].value_counts()[df["Case_No"].value_counts() > 1].index)

    st.markdown(f'<div class="info-box">🔍 <b>{len(results):,}</b> result(s) for "<b>{query}</b>"</div>', unsafe_allow_html=True)

    if not results.empty and len(results) > 1:
        st.markdown("<div style='height:8px'></div>", unsafe_allow_html=True)
        colx, coly, colz = st.columns(3, gap="medium")
        with colx:
            st.markdown('<div class="sec-title">🗂️ Results by Category</div>', unsafe_allow_html=True)
            rc = value_counts_df(results["Case_Category"], "Category", blank_label="Uncategorized").head(8)
            fig_rc = horizontal_bar_ranked(rc, "Category", "Cases", height=280, label_len=22, single_color=COLORS["navy"], right_margin=40)
            st.plotly_chart(fig_rc, use_container_width=True, key="chart_1834_fig_rc")
        with coly:
            st.markdown('<div class="sec-title">👨‍⚖️ Results by Judge</div>', unsafe_allow_html=True)
            rj = value_counts_df(results["Judge_Short"], "Judge").head(8)
            fig_rj = horizontal_bar_ranked(rj, "Judge", "Cases", height=280, label_len=22, single_color=COLORS["teal"], right_margin=40)
            st.plotly_chart(fig_rj, use_container_width=True, key="chart_1839_fig_rj")
        with colz:
            st.markdown('<div class="sec-title">📋 Results by Section</div>', unsafe_allow_html=True)
            rs = value_counts_df(results["Section_Clean"], "Section").head(8)
            fig_rs = horizontal_bar_ranked(rs, "Section", "Cases", height=280, label_len=22, single_color=COLORS["gold"], right_margin=40)
            st.plotly_chart(fig_rs, use_container_width=True, key="chart_1844_fig_rs")
        st.markdown("<div style='height:8px'></div>", unsafe_allow_html=True)

    if results.empty:
        st.markdown(
            """
            <div class="empty">
                <div class="ei">🔎</div>
                <div class="et">No Results Found</div>
                <div class="es">Try a different search term or check spelling.</div>
            </div>""",
            unsafe_allow_html=True,
        )
    else:
        for _, row in results.head(60).iterrows():
            st.markdown(render_result_card(row, row["Case_No"] in repeated_cases), unsafe_allow_html=True)
        if len(results) > 60:
            st.info(f"Showing first 60 of {len(results):,} results. Narrow your search for more specific results.")

    st.markdown("</div>", unsafe_allow_html=True)


# ═══════════════════════════════════════════════════════════════
# MAIN
# ═══════════════════════════════════════════════════════════════
def main() -> None:
    inject_css()
    df = render_sidebar()

    render_banner(df)
    render_kpis(df)

    tab_titles = ["Overview", "Daily Cause List", "Judge Analysis",
                  "Court Infrastructure", "Lawyer Intelligence", "Case Search"]
    tabs = st.tabs(tab_titles)

    with tabs[0]:
        tab_overview(df)
    with tabs[1]:
        tab_daily_cause_list(df)
    with tabs[2]:
        tab_judge_analysis(df)
    with tabs[3]:
        tab_court_infrastructure(df)
    with tabs[4]:
        tab_lawyer_intelligence(df)
    with tabs[5]:
        tab_case_search(df)


if __name__ == "__main__":
    main()