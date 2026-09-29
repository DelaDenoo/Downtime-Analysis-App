# ============================================================
# VISUAL DESIGN SYSTEM - matches a reference design (Downtime & OEE
# Intelligence, Barry Callebaut Ghana), applied on top of the app's real
# functionality below. Nothing in this block changes data, computation,
# or discovery logic - it only restyles how it's presented.
# ============================================================

import streamlit as st

_ICON_LINE_CHART = '<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M3 3v18h18"/><path d="M18.7 8.7 13 14.4l-3.5-3.5L3 18"/></svg>'
_ICON_BRANCH = '<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><line x1="6" y1="3" x2="6" y2="15"/><circle cx="18" cy="6" r="3"/><circle cx="6" cy="18" r="3"/><path d="M18 9a9 9 0 0 1-9 9"/></svg>'
_ICON_UPLOAD = '<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="17 8 12 3 7 8"/><line x1="12" y1="3" x2="12" y2="15"/></svg>'
_ICON_TARGET = '<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="10"/><circle cx="12" cy="12" r="6"/><circle cx="12" cy="12" r="2"/></svg>'
_ICON_CALENDAR_X = '<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect x="3" y="4" width="18" height="18" rx="2"/><line x1="16" y1="2" x2="16" y2="6"/><line x1="8" y1="2" x2="8" y2="6"/><line x1="3" y1="10" x2="21" y2="10"/><line x1="10" y1="14" x2="14" y2="18"/><line x1="14" y1="14" x2="10" y2="18"/></svg>'
_ICON_GAUGE = '<svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M12 15 15 9"/><path d="M3.5 19a9 9 0 1 1 17 0"/></svg>'
_ICON_UP = '<svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><line x1="12" y1="19" x2="12" y2="5"/><polyline points="5 12 12 5 19 12"/></svg>'
_ICON_DOWN = '<svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><line x1="12" y1="5" x2="12" y2="19"/><polyline points="19 12 12 19 5 12"/></svg>'
_ICON_CLOCK = '<svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="10"/><polyline points="12 6 12 12 16 14"/></svg>'
_ICON_SCALE = '<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M16 3h5v5"/><path d="M8 3H3v5"/><path d="M3 3l7.5 7.5"/><path d="M21 3l-7.5 7.5"/><path d="M6 13l-3 6a3 3 0 0 0 6 0z"/><path d="M18 13l-3 6a3 3 0 0 0 6 0z"/><line x1="12" y1="10" x2="12" y2="21"/></svg>'
_ICON_BAR_CHART = '<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><line x1="12" y1="20" x2="12" y2="10"/><line x1="18" y1="20" x2="18" y2="4"/><line x1="6" y1="20" x2="6" y2="16"/></svg>'
_ICON_ALARM = '<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="13" r="8"/><path d="M12 9v4l2 2"/><path d="M5 3 2 6"/><path d="M22 6l-3-3"/></svg>'
_ICON_PIN = '<svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M20 10c0 6-8 12-8 12s-8-6-8-12a8 8 0 0 1 16 0z"/><circle cx="12" cy="10" r="3"/></svg>'
_ICON_BOLT = '<svg width="12" height="12" viewBox="0 0 24 24" fill="currentColor"><path d="M13 2 3 14h7l-1 8 10-12h-7l1-8z"/></svg>'
_ICON_WARN = '<svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><path d="M12 9v4"/><path d="M12 17h.01"/><path d="M10.29 3.86 1.82 18a2 2 0 0 0 1.71 3h16.94a2 2 0 0 0 1.71-3L13.71 3.86a2 2 0 0 0-3.42 0z"/></svg>'

TOK = {
    "bg": "#FBF6EC", "card": "#FFFFFF", "border": "#EBE2D0",
    "ink": "#2A241C", "sub": "#8F8471",
    "orange": "#B9743D", "orange_tint": "#F3E1C8", "orange_deep": "#9C5F2E",
    "green": "#2F8F63", "green_tint": "#D9F0E1",
    "coral": "#B76552", "coral_tint": "#F2DFD7",
    "gray_tint": "#EFEAE0",
}

CSS_BLOCK = f"""
<style>
@import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800&display=swap');

html, body, [class*="css"] {{ font-family: 'Plus Jakarta Sans', sans-serif; }}

.stApp {{ background: {TOK['bg']}; }}
section[data-testid="stSidebar"] {{ background: {TOK['bg']}; border-right: 1px solid {TOK['border']}; }}
section[data-testid="stSidebar"] > div {{ padding-top: 1rem; }}

h1, h2, h3 {{ color: {TOK['ink']}; font-weight: 700 !important; }}
h2 {{ font-size: 20px !important; border-bottom: 2px solid {TOK['orange_tint']}; padding-bottom: 8px; margin-top: 8px !important; }}
h3 {{ font-size: 16.5px !important; }}
p, span, label, div {{ color: {TOK['ink']}; }}
.dc-sub {{ color: {TOK['sub']}; }}

/* ---- pill badges ---- */
.dc-pill {{
    display: inline-flex; align-items: center; gap: 6px;
    padding: 7px 14px; border-radius: 999px; font-size: 13.5px; font-weight: 600;
    white-space: nowrap;
}}
.dc-pill svg {{ flex-shrink: 0; }}
.dc-pill-orange {{ background: {TOK['orange_tint']}; color: {TOK['orange_deep']}; }}
.dc-pill-gray {{ background: {TOK['gray_tint']}; color: {TOK['ink']}; }}
.dc-pill-green {{ background: {TOK['green_tint']}; color: {TOK['green']}; }}
.dc-pill-coral {{ background: {TOK['coral_tint']}; color: {TOK['coral']}; }}
.dc-pill-orange-solid {{ background: {TOK['orange']}; color: white; }}

/* ---- top header bar ---- */
.dc-header {{
    display: flex; align-items: center; justify-content: space-between;
    padding: 4px 0 20px 0; flex-wrap: wrap; gap: 14px;
}}
.dc-header-left {{ display: flex; align-items: center; gap: 14px; }}
.dc-header-icon {{
    width: 48px; height: 48px; border-radius: 12px; background: {TOK['orange']};
    color: white; display: flex; align-items: center; justify-content: center; flex-shrink: 0;
}}
.dc-header-title {{ font-size: 22px; font-weight: 800; color: {TOK['ink']}; line-height: 1.2; }}
.dc-header-subtitle {{ font-size: 13.5px; color: {TOK['sub']}; margin-top: 1px; }}
.dc-header-right {{ display: flex; align-items: center; gap: 10px; flex-wrap: wrap; }}

/* ---- icon-square card headers, used throughout sidebar + main cards ---- */
.dc-card-head {{ display: flex; align-items: center; gap: 12px; margin-bottom: 4px; }}
.dc-icon-sq {{
    width: 38px; height: 38px; border-radius: 10px; display: flex;
    align-items: center; justify-content: center; flex-shrink: 0;
}}
.dc-icon-sq-orange {{ background: {TOK['orange_tint']}; color: {TOK['orange_deep']}; }}
.dc-icon-sq-green {{ background: {TOK['green_tint']}; color: {TOK['green']}; }}
.dc-icon-sq-coral {{ background: {TOK['coral_tint']}; color: {TOK['coral']}; }}
.dc-card-title {{ font-size: 15.5px; font-weight: 700; color: {TOK['ink']}; line-height: 1.25; }}
.dc-card-subtitle {{ font-size: 12.5px; color: {TOK['sub']}; margin-top: 1px; }}

/* ---- generic card wrapper ---- */
.dc-card {{
    background: {TOK['card']}; border: 1px solid {TOK['border']}; border-radius: 16px;
    padding: 18px 20px; margin-bottom: 14px;
}}

/* ---- selectable line cards (sidebar) ---- */
.dc-line-card {{
    border: 1.5px solid {TOK['border']}; border-radius: 12px; padding: 12px 14px;
    margin-bottom: 8px; display: flex; align-items: center; justify-content: space-between;
}}
.dc-line-card-selected {{ border-color: {TOK['orange']}; background: {TOK['orange_tint']}55; }}
.dc-line-card-name {{ font-weight: 700; font-size: 14.5px; color: {TOK['ink']}; }}
.dc-line-card-sub {{ font-size: 12px; color: {TOK['sub']}; margin-top: 1px; }}
.dc-radio-dot {{
    width: 20px; height: 20px; border-radius: 50%; border: 2px solid {TOK['border']};
    flex-shrink: 0; display: flex; align-items: center; justify-content: center;
}}
.dc-radio-dot-selected {{ border-color: {TOK['orange']}; }}
.dc-radio-dot-selected::after {{
    content: ""; width: 10px; height: 10px; border-radius: 50%; background: {TOK['orange']};
}}

/* ---- discovery unit rows ---- */
.dc-unit-row {{
    border: 1px solid {TOK['border']}; border-radius: 10px; padding: 9px 12px;
    margin-bottom: 6px; display: flex; align-items: flex-start; justify-content: space-between; gap: 8px;
}}
.dc-unit-name {{ font-weight: 600; font-size: 13.5px; color: {TOK['ink']}; }}
.dc-unit-loc {{ font-size: 11.5px; color: {TOK['sub']}; margin-top: 3px; display: flex; align-items: center; gap: 4px; }}

/* ---- metric cards ---- */
.dc-metric-label {{ font-size: 11px; font-weight: 700; letter-spacing: 0.04em; color: {TOK['sub']}; text-transform: uppercase; }}
.dc-metric-value {{ font-size: 30px; font-weight: 800; color: {TOK['ink']}; margin: 6px 0 2px 0; line-height: 1; }}
.dc-metric-sub {{ font-size: 12px; color: {TOK['sub']}; }}
.dc-metric-top {{ display: flex; align-items: center; justify-content: space-between; }}

/* ---- footer ---- */
.dc-footer {{
    text-align: center; padding: 16px; margin-top: 24px; border-top: 1px solid {TOK['border']};
    color: {TOK['sub']}; font-size: 12.5px;
}}

/* ---- Streamlit native widget reskin ---- */
.stTabs [data-baseweb="tab-list"] {{ gap: 6px; border-bottom: none; flex-wrap: wrap; }}
.stTabs [data-baseweb="tab"] {{
    background: {TOK['card']}; border: 1px solid {TOK['border']}; border-radius: 999px;
    padding: 8px 18px; font-weight: 600; font-size: 13.5px; color: {TOK['ink']};
}}
.stTabs [aria-selected="true"] {{
    background: {TOK['orange']} !important; color: white !important; border-color: {TOK['orange']} !important;
}}
.stTabs [data-baseweb="tab-highlight"] {{ display: none; }}

section[data-testid="stSidebar"] .stRadio, section[data-testid="stSidebar"] .stFileUploader,
section[data-testid="stSidebar"] .stMultiSelect {{ margin-top: 2px; }}

[data-testid="stFileUploaderDropzone"] {{
    background: {TOK['card']}; border: 2px dashed {TOK['border']}; border-radius: 14px;
}}

div[data-testid="stExpander"] {{
    background: {TOK['card']}; border: 1px solid {TOK['border']}; border-radius: 14px;
}}

.stMultiSelect [data-baseweb="tag"] {{
    background: {TOK['orange_tint']} !important; border-radius: 999px !important;
}}
.stMultiSelect [data-baseweb="tag"] span {{ color: {TOK['orange_deep']} !important; }}

.stDataFrame {{ border: 1px solid {TOK['border']}; border-radius: 12px; overflow: hidden; }}

button[kind="primary"], button[kind="secondary"] {{ border-radius: 999px !important; }}

#MainMenu {{visibility: hidden;}}
footer {{visibility: hidden;}}
</style>
"""


def dc_pill(icon_svg, text, variant="gray"):
    return f'<span class="dc-pill dc-pill-{variant}">{icon_svg}<span>{text}</span></span>'


def dc_header(title, subtitle, pills_html):
    st.markdown(
        f'<div class="dc-header"><div class="dc-header-left">'
        f'<div class="dc-header-icon">{_ICON_LINE_CHART}</div>'
        f'<div><div class="dc-header-title">{title}</div>'
        f'<div class="dc-header-subtitle">{subtitle}</div></div></div>'
        f'<div class="dc-header-right">{pills_html}</div></div>',
        unsafe_allow_html=True,
    )


def dc_card_head(icon_svg, title, subtitle, variant="orange"):
    return (
        f'<div class="dc-card-head"><div class="dc-icon-sq dc-icon-sq-{variant}">{icon_svg}</div>'
        f'<div><div class="dc-card-title">{title}</div>'
        f'<div class="dc-card-subtitle">{subtitle}</div></div></div>'
    )


def dc_metric_card(label, value, sub, icon_svg, variant="orange"):
    return (
        f'<div class="dc-card"><div class="dc-metric-top">'
        f'<span class="dc-metric-label">{label}</span>'
        f'<div class="dc-icon-sq dc-icon-sq-{variant}" style="width:30px;height:30px;">{icon_svg}</div>'
        f'</div><div class="dc-metric-value">{value}</div>'
        f'<div class="dc-metric-sub">{sub}</div></div>'
    )
