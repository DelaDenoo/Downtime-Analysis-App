"""
Barth Line & Buhler Line Downtime Analysis Dashboard
=================================================
Mirrors the flow, graphs, and reasoning of the two corrected Jupyter
notebooks (Barth Line, Buhler Line) as an interactive Streamlit app. Uses the
same fixed column letters already confirmed against real data in those
notebooks - no spatial/automatic column detection.

Run with:  streamlit run dashboard.py
"""

import re
import io
import os
import tempfile
import calendar
from datetime import datetime

import numpy as np
import pandas as pd
import streamlit as st
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
import seaborn as sns
from fpdf import FPDF
from PIL import Image
import openpyxl
from openpyxl.utils import get_column_letter, column_index_from_string

from design_system import (
    TOK, CSS_BLOCK, dc_pill, dc_header, dc_card_head, dc_metric_card,
    _ICON_BRANCH, _ICON_UPLOAD, _ICON_TARGET, _ICON_CALENDAR_X, _ICON_GAUGE,
    _ICON_UP, _ICON_DOWN, _ICON_CLOCK, _ICON_SCALE, _ICON_BAR_CHART,
    _ICON_ALARM, _ICON_PIN, _ICON_BOLT, _ICON_WARN, _ICON_LINE_CHART,
)

MONTH_ORDER = list(calendar.month_name)[1:]
DAY_ORDER = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]

# Matplotlib restyled to match the app's visual language - same chart
# logic and data everywhere below, just recolored fonts/spines/grid so
# charts don't look like a mismatched, separate tool bolted onto the page.
plt.rcParams.update({
    "font.family": "sans-serif",
    "font.sans-serif": ["Arial", "DejaVu Sans"],
    "axes.facecolor": "#FFFFFF", "figure.facecolor": "#FFFFFF",
    "axes.edgecolor": "#EBE2D0", "axes.labelcolor": "#2A241C",
    "text.color": "#2A241C", "xtick.color": "#8F8471", "ytick.color": "#8F8471",
    "grid.color": "#EFEAE0", "grid.alpha": 0.8,
    "axes.spines.top": False, "axes.spines.right": False,
    "legend.frameon": False,
})
_DC_ORANGE, _DC_GREEN, _DC_CORAL = "#B9743D", "#2F8F63", "#B76552"


def _show_logo_if_present(path="logo.png"):
    """st.image() raises MediaFileStorageError and crashes the whole app
    if the file doesn't exist (confirmed: a missing logo.png took down the
    entire page before the file uploader even rendered) - a branding image
    should never be able to do that. Skips silently if missing, rather
    than either crashing or nagging the user with a warning about a file
    they may not have wanted in the first place."""
    if os.path.exists(path):
        st.image(path, use_container_width=True)

# ============================================================
# DYNAMIC "BY CATEGORY" BLOCK DISCOVERY
# Finds each unit's reason block by scanning for header text (a "By
# Category" section label, then each unit's own "<Line>: <Unit>"
# sub-label, then the six field labels beneath it) instead of trusting
# fixed row/column config. Falls back to the hardcoded CATEGORY_COLS /
# CATEGORY_SKIPROWS config below if discovery finds nothing usable for a
# given unit - this does not replace that config, it supplements it.
# ============================================================

_FIELD_PATTERNS = {
    "Category": re.compile(r"category", re.IGNORECASE),
    "Reason": re.compile(r"reason", re.IGNORECASE),
    "Ratio": re.compile(r"ratio", re.IGNORECASE),
    "Duration_hours": re.compile(r"duration.*hour", re.IGNORECASE),
    "Duration_min": re.compile(r"duration.*min", re.IGNORECASE),
    "Count": re.compile(r"^count$|^#$", re.IGNORECASE),
}
# Category and Ratio are optional for detection - a re-sorted "Top Reasons"
# style duplicate can merge Category into the Reason text with no separate
# Category header label at all. Reason, Count, and at least one Duration
# field are the genuinely load-bearing ones.
_CORE_REQUIRED_FIELDS = {"Reason", "Count"}


def _cell_text(v):
    return str(v).strip() if v is not None else None


def _find_category_section_start(ws, max_row, max_col):
    """Finds the row of the 'By Category' marker, if present. This is a
    boundary, not a block anchor in its own right - the real per-unit
    6-column tables sit at their own '<Line>: <Unit>' labels somewhere
    below it, each at a different column, not directly next to this
    marker. Returns None if no such marker exists anywhere in the sheet."""
    marker = re.compile(r"by\s*category", re.IGNORECASE)
    for row in ws.iter_rows(min_row=1, max_row=max_row, max_col=max_col):
        for cell in row:
            text = _cell_text(cell.value)
            if text and marker.search(text):
                return cell.row
    return None


def _find_anchors(ws, max_row, max_col, min_row=1):
    """Scans for '<Line>: <Unit>' style per-unit labels, starting from
    min_row. The same label text (e.g. "Barth: DES11") appears once in the
    DAILY-data section and again in the CATEGORY-block section of the same
    sheet - restricting min_row to at/after the 'By Category' marker is
    what keeps the daily-section copy from being treated as a (wrong)
    category-block anchor. The prefix before the colon may contain spaces
    ("Liquor Moulding: TM41") - confirmed real, and a plain [A-Za-z0-9_]+
    prefix (no spaces allowed) silently misses these anchors entirely
    rather than finding-and-failing them, which is harder to notice."""
    unit_label = re.compile(r"^[A-Za-z0-9_][A-Za-z0-9_ ]*:\s*.+")
    anchors = []
    for row in ws.iter_rows(min_row=min_row, max_row=max_row, max_col=max_col):
        for cell in row:
            text = _cell_text(cell.value)
            if text and unit_label.search(text):
                anchors.append({"row": cell.row, "col": cell.column, "text": text})
    return anchors


def _find_header_row(ws, start_row, col_min, col_max, max_lookahead=10):
    for r in range(start_row, start_row + max_lookahead + 1):
        found = {}
        for c in range(col_min, col_max + 1):
            text = _cell_text(ws.cell(row=r, column=c).value)
            if text is None:
                continue
            for field, pat in _FIELD_PATTERNS.items():
                if field not in found and pat.search(text):
                    found[field] = c

        # Some real exports (confirmed on an actual file) label only the
        # numeric fields - Ratio, Duration, Count - and never put literal
        # "Category"/"Reason" text anywhere; those two columns go straight
        # into data starting the row right below, with no header text at
        # all. When Ratio is labeled but Reason isn't, infer Reason/
        # Category by DATA PRESENCE in the row right below, scanning
        # backward from Ratio's column for the two nearest columns that
        # actually hold a value there - not a fixed offset, since field
        # spacing isn't consistent between files (confirmed: sometimes 1
        # gap column, sometimes more).
        if "Reason" not in found and "Ratio" in found:
            other_label_cols = set(found.values())
            candidate_cols = []
            for c in range(found["Ratio"] - 1, col_min - 1, -1):
                if c in other_label_cols:
                    break
                # Check several rows below the header, not just the first -
                # a column can legitimately be blank on the very first data
                # row (confirmed: Category is ffilled and only set every
                # few rows, so it can be populated one row while the Reason
                # column right next to it is still empty there and only
                # starts the row after) while still being a real, active
                # data column overall.
                has_data = any(ws.cell(row=r + 1 + offset, column=c).value not in (None, "")
                               for offset in range(5))
                if has_data:
                    candidate_cols.append(c)
                if len(candidate_cols) == 2:
                    break
            if len(candidate_cols) >= 1:
                found["Reason"] = candidate_cols[0]
            if len(candidate_cols) >= 2:
                found["Category"] = candidate_cols[1]

        has_duration = "Duration_hours" in found or "Duration_min" in found
        if _CORE_REQUIRED_FIELDS.issubset(found.keys()) and has_duration:
            return r, found
    return None, None


def _find_block_extent(ws, header_row, col_range, other_anchor_rows, max_scan=500):
    blank_streak = 0
    last_data_row = header_row
    for r in range(header_row + 1, header_row + 1 + max_scan):
        if r in other_anchor_rows:
            break
        any_value = any(ws.cell(row=r, column=c).value not in (None, "") for c in col_range)
        if any_value:
            blank_streak = 0
            last_data_row = r
        else:
            blank_streak += 1
            if blank_streak >= 3:
                break
    return header_row + 1, last_data_row


@st.cache_data(show_spinner=False)
def discover_category_blocks(file_bytes):
    """Scans the whole workbook once and returns every discovered block,
    keyed by a unique label (anchor text + its row/col, since the same
    text - e.g. the same unit named twice - can legitimately appear at
    more than one location). NOT read_only=True: that mode trusts the
    file's cached <dimension> XML tag for max_row/max_column instead of
    scanning, and several real exports have that metadata missing or
    wrong, which makes max_row come back as None and crashes range(...,
    max_row + 1). A normal load forces a real scan and sidesteps this.

    Two-step search: 'By Category' is a single section marker, not a
    block itself - only unit labels at/after that row are treated as real
    category-block anchors, since the same "<Line>: <Unit>" text also
    appears earlier in the daily-data section, and that copy has no
    6-column table near it at all. Returns (results, meta) - meta records
    whether a 'By Category' marker was actually found, which is worth
    surfacing: its absence means the search fell back to scanning the
    whole sheet with nothing to narrow against."""
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)
    ws = wb.active
    max_row = ws.max_row or 2000
    max_col = ws.max_column or 200

    category_section_row = _find_category_section_start(ws, max_row, max_col)
    anchor_search_start = category_section_row if category_section_row is not None else 1
    anchors = _find_anchors(ws, max_row, max_col, min_row=anchor_search_start)
    anchor_rows = {a["row"] for a in anchors}

    # Process anchors in ROW BANDS - anchors within a few rows of each
    # other (like a unit's own anchor at row 383 and its neighbor's at row
    # 384, both converging on the same header row further down) belong to
    # the same logical section and should bound each other's search
    # windows. Anchors far apart in row (row 69 vs a duplicate "Top
    # Reasons" pivot at row 108 - a real, confirmed case) are genuinely
    # unrelated and must NOT bound each other: exact-row grouping broke
    # the 383-vs-384 case, and no row grouping at all let the row-108
    # duplicate corrupt row 69's window (both confirmed bugs). A gap
    # threshold between consecutive anchors, sorted by row, is what
    # separates a real duplicate/pivot block from anchors that are just a
    # row or two apart within the same section.
    ROW_BAND_GAP = 10
    sorted_by_row = sorted(anchors, key=lambda a: a["row"])
    row_bands = []
    current_band = []
    for a in sorted_by_row:
        if current_band and (a["row"] - current_band[-1]["row"]) > ROW_BAND_GAP:
            row_bands.append(current_band)
            current_band = []
        current_band.append(a)
    if current_band:
        row_bands.append(current_band)

    results = {}
    for band in row_bands:
        band_sorted_by_col = sorted(band, key=lambda a: a["col"])
        rightmost_used_col = 0
        for anchor in band_sorted_by_col:
            label = f"{anchor['text']} (row {anchor['row']}, col {get_column_letter(anchor['col'])})"
            col_min = max(1, rightmost_used_col + 1, anchor["col"] - 5)
            later_band_cols = [a["col"] for a in band if a["col"] > anchor["col"]]
            col_max = min(max_col, min(later_band_cols) - 1) if later_band_cols else min(max_col, anchor["col"] + 60)

            header_row, found_cols = _find_header_row(ws, anchor["row"], col_min, col_max)
            if header_row is None:
                results[label] = None
                continue

            combined_fields = ("Category" not in found_cols
                               or found_cols["Category"] == found_cols["Reason"])
            col_letters = {f: get_column_letter(c) for f, c in found_cols.items()}
            col_range = sorted(found_cols.values())
            data_start, data_end = _find_block_extent(ws, header_row, col_range, anchor_rows - {anchor["row"]})
            rightmost_used_col = max(rightmost_used_col, max(col_range))

            results[label] = {
                "header_row": header_row, "columns": col_letters,
                "data_start_row": data_start, "data_end_row": data_end,
                "combined_fields": combined_fields,
            }
    wb.close()
    meta = {"category_section_row": category_section_row}
    return results, meta


def load_discovered_block(file_bytes, block_info):
    """Reads a discovered block via openpyxl cell-by-cell (not pandas'
    usecols-as-letters, which determines the sheet's column bounds from
    row 1 of the WHOLE file and can wrongly flag a far-right block as out
    of bounds if row 1 happens to be narrower, regardless of whether the
    real data is fine)."""
    if block_info is None or block_info.get("combined_fields"):
        return None
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)
    ws = wb.active
    cols = block_info["columns"]
    field_order = ["Category", "Reason", "Ratio", "Duration_hours", "Duration_min", "Count"]
    col_idx = {f: (column_index_from_string(cols[f]) if f in cols else None) for f in field_order}

    rows_out = []
    for r in range(block_info["data_start_row"], block_info["data_end_row"] + 1):
        rows_out.append([ws.cell(row=r, column=col_idx[f]).value if col_idx[f] else None for f in field_order])
    wb.close()

    df = pd.DataFrame(rows_out, columns=["Category", "Reason", "Ratio", "Duration_hours", "Duration", "Count"])
    df["Category"] = df["Category"].replace(r"^\s*$", np.nan, regex=True).ffill()
    df["Reason"] = df["Reason"].replace(r"^\s*$", np.nan, regex=True)
    df = df[df["Reason"].notna()].reset_index(drop=True)
    for col in ["Duration_hours", "Duration", "Count"]:
        df[col] = pd.to_numeric(df[col], errors="coerce")
    df = df.dropna(subset=["Duration_hours"]).reset_index(drop=True)
    return df


def _extract_tag(unit_key):
    m = re.search(r"\(([^)]+)\)", unit_key)
    return m.group(1) if m else unit_key


def _normalize_tag(s):
    # Confirmed on a real file: the SAME roaster is labeled "R011" in this
    # app's config (from the historical file) but "RO11" in a newer export
    # of the same plant - the letter O and digit 0 used interchangeably for
    # the same equipment code across different file versions. Treating them
    # as equivalent for matching purposes is what makes both resolve to the
    # same unit. Also confirmed on Butter Moulding's real file: the config
    # key "Butter Moulding: TM11" (no space before the colon) didn't match
    # the file's real anchor text "Butter Moulding : TM11" (space before
    # the colon) until whitespace was stripped too - same principle as the
    # O/0 case, a superficial formatting difference that shouldn't block
    # an otherwise-correct match.
    return s.lower().replace("o", "0").replace(" ", "")


def resolve_unit_block(unit_key, discovered_blocks):
    """Matches a config unit key (e.g. 'Destoner (DES11)') against the
    discovered blocks by equipment code. Only ever returns a clean
    (non-combined-field) block, or None - never a combined-field one, even
    as a last resort, since force-parsing one produces wrong data
    (Category always blank, Reason containing merged text). None is the
    caller's signal to fall back to the hardcoded CATEGORY_COLS config."""
    tag = _normalize_tag(_extract_tag(unit_key))
    matches = [v for k, v in discovered_blocks.items()
              if v is not None and not v["combined_fields"] and tag in _normalize_tag(k)]
    return matches[0] if matches else None

# ============================================================
# SHARED HELPERS - reused by both lines' flows
# ============================================================

def load_daily(file, usecols, skiprows):
    """Loads one unit's daily downtime column and cleans it. Returns
    (dataframe, eq_title) - eq_title is the real column header text, shown
    as an identity check. No month exclusion here - that's applied
    uniformly afterward, once the real available date range is known, via
    apply_month_exclusion()."""
    d = pd.read_excel(file, usecols=usecols, skiprows=skiprows)
    eq_title = str(d.columns[1])
    d.columns = ["Day", "Duration (hours)"]
    d["Duration (hours)"] = pd.to_numeric(d["Duration (hours)"], errors="coerce")
    d = d.dropna(subset=["Duration (hours)"])
    d["Date"] = pd.to_datetime(d["Day"], errors="coerce")
    d = d.dropna(subset=["Date"]).sort_values("Date").reset_index(drop=True)
    return d, eq_title


def apply_month_exclusion(d, excluded_year_months):
    """excluded_year_months: list of 'YYYY-MM' strings to drop. Recomputes
    the calendar fields used throughout (month/week/day name) after
    filtering, so downstream groupbys never see the excluded days."""
    d = d.copy()
    if excluded_year_months:
        ym = d["Date"].dt.strftime("%Y-%m")
        d = d[~ym.isin(excluded_year_months)].reset_index(drop=True)
    d["Month_Name"] = d["Date"].dt.month_name()
    d["Week_of_Year"] = d["Date"].dt.isocalendar().week.astype(str)
    d["Day_of_Week"] = d["Date"].dt.day_name()
    return d


def load_category_block(file, usecols, skiprows, nrows=400):
    """Loads one unit's own 'By Category' reason block (Category, Reason,
    Ratio, Duration_hours, Duration, Count) at fixed column letters, using
    openpyxl cell-by-cell reads rather than pandas' usecols-as-letters.
    That pandas path determines the sheet's column bounds from ROW 1 of
    the whole file, not from where it's actually reading, and can raise
    "usecols out-of-bounds" for a real, valid far-right block if row 1
    happens to be narrower - confirmed to crash the app outright (not just
    read wrong data) when this fallback runs against a file shaped
    differently than the selected line expects. Reading cell-by-cell
    avoids that failure mode entirely, matching the fix already proven
    for the same issue elsewhere in this project."""
    wb = openpyxl.load_workbook(file, data_only=True)
    ws = wb.active
    col_idx = [column_index_from_string(c) for c in usecols.split(",")]
    max_row = ws.max_row or (skiprows + nrows)

    rows_out = []
    for r in range(skiprows + 1, min(skiprows + 1 + nrows, max_row + 1)):
        rows_out.append([ws.cell(row=r, column=c).value for c in col_idx])
    wb.close()

    df_cat = pd.DataFrame(rows_out, columns=["Category", "Reason", "Ratio", "Duration_hours", "Duration", "Count"])
    df_cat["Category"] = df_cat["Category"].replace(r"^\s*$", np.nan, regex=True).ffill()
    df_cat["Reason"] = df_cat["Reason"].replace(r"^\s*$", np.nan, regex=True)
    df_cat = df_cat[df_cat["Reason"].notna()].reset_index(drop=True)
    for col in ["Duration_hours", "Duration", "Count"]:
        df_cat[col] = pd.to_numeric(df_cat[col], errors="coerce")
    df_cat = df_cat.dropna(subset=["Duration_hours"]).reset_index(drop=True)
    return df_cat


def extract_equipment_references(text, equipment_patterns):
    hits = []
    for unit, patterns in equipment_patterns.items():
        if any(re.search(p, text, flags=re.IGNORECASE) for p in patterns):
            hits.append(unit)
    return hits


def classify_allocation(row, equipment_patterns, roaster_own_keywords):
    """The roaster's own comment-attribution classifier: does this reason
    name another unit, or is it the roaster's own fault, or neither?"""
    reason = str(row["Reason"]).strip()
    if reason.lower() == "unknown":
        return "Unknown"
    mentioned = extract_equipment_references(reason, equipment_patterns)
    if mentioned:
        return mentioned[0] if len(mentioned) == 1 else "Multiple units mentioned"
    if any(re.search(kw, reason, flags=re.IGNORECASE) for kw in roaster_own_keywords):
        return "Roaster (own fault)"
    return "Unattributable"


def build_allocation_summary(df_cat):
    alloc = df_cat.groupby("Allocated_To").agg(
        Total_Duration_hours=("Duration_hours", "sum"),
        Total_Count=("Count", "sum"),
    ).reset_index()
    total = alloc["Total_Duration_hours"].sum()
    alloc["Pct_of_Total"] = (alloc["Total_Duration_hours"] / total * 100) if total > 0 else 0.0
    alloc["Avg_Mins_per_Event"] = np.where(
        alloc["Total_Count"] > 0, (alloc["Total_Duration_hours"] * 60) / alloc["Total_Count"], 0.0)
    return alloc.sort_values("Total_Duration_hours", ascending=False).reset_index(drop=True)


def classify_theme(reason, theme_patterns):
    for theme, patterns in theme_patterns.items():
        if any(re.search(p, str(reason), flags=re.IGNORECASE) for p in patterns):
            return theme
    return None


def theme_breakdown_data(unit_reasons, theme_patterns):
    """Pure-data version of theme_breakdown() - classifies, groups, and
    returns (summary_df, unmapped_df) without plotting, so the caller
    decides how/where to render it in the Streamlit layout."""
    unit_reasons = unit_reasons.copy()
    unit_reasons["Theme"] = unit_reasons["Reason"].apply(lambda r: classify_theme(r, theme_patterns))
    unmapped = unit_reasons[unit_reasons["Theme"].isna()][["Reason", "Duration_hours", "Count"]]
    mapped = unit_reasons.dropna(subset=["Theme"])
    if mapped.empty:
        return pd.DataFrame(), unmapped
    summary = mapped.groupby("Theme").agg(
        Total_Duration_hours=("Duration_hours", "sum"),
        Total_Count=("Count", "sum"),
        Num_Reasons=("Reason", "count"),
    ).reset_index()
    total = summary["Total_Duration_hours"].sum()
    summary["Pct_of_Unit_Total"] = (summary["Total_Duration_hours"] / total * 100) if total > 0 else 0.0
    return summary.sort_values("Total_Duration_hours", ascending=False).reset_index(drop=True), unmapped


def plot_theme_bar(summary, title, color):
    fig, ax = plt.subplots(figsize=(10, max(3, 0.5 * len(summary))))
    bars = ax.barh(summary["Theme"], summary["Total_Duration_hours"], color=color)
    ax.bar_label(bars, fmt="%.1f hrs", padding=3)
    ax.set_xlim(0, max(summary["Total_Duration_hours"].max(), 0.1) * 1.15)  # room for the label text itself
    ax.set_title(title, fontsize=13, fontweight="bold")
    ax.set_xlabel("Total Duration (hours)")
    ax.invert_yaxis()
    plt.tight_layout()
    return fig


def render_theme_section(label, reasons_df, theme_patterns, color):
    """One unit's full theme-breakdown block: summary table, per-theme
    detail, unmapped audit, and chart - as its own Streamlit expander.
    Returns (summary, fig) so the caller can reuse the SAME figure for the
    PDF instead of regenerating it (this used to compute every theme chart
    twice - once here, once again for the PDF collection)."""
    with st.expander(label, expanded=False):
        if reasons_df is None or reasons_df.empty:
            st.info("No reasons available for this unit.")
            return None, None
        summary, unmapped = theme_breakdown_data(reasons_df, theme_patterns)
        if not unmapped.empty:
            st.warning(f"{len(unmapped)} reason(s) matched no known theme keyword:")
            st.dataframe(unmapped, use_container_width=True)
        if summary.empty:
            st.info("No themed reasons for this unit.")
            return None, None
        st.write(f"{len(reasons_df) - len(unmapped)} of {len(reasons_df)} reasons grouped into "
                 f"{len(summary)} of {len(theme_patterns)} possible themes")
        st.dataframe(summary.round(1), use_container_width=True)
        fig = plot_theme_bar(summary, f"{label}: Reasons Grouped by Theme", color)
        st.pyplot(fig)
        return summary, fig


def collision_aware_quadrant(q, title, label_col="Allocated_To"):
    """Frequency-vs-duration maintenance-strategy quadrant, log scale, with
    collision-aware label placement so bubbles and text don't overlap.
    label_col defaults to "Allocated_To" (Barth/Buhler's attributed-source
    column) but the Moulding lines pass "Unit" instead, since they have no
    cross-unit reattribution - each bubble is a whole analysed unit, not a
    source within one roaster's comments."""
    med_count = q["Total_Count"].median()
    med_dur = q["Avg_Mins_per_Event"].median()

    def strategy(row):
        hi_f = row["Total_Count"] >= med_count
        hi_d = row["Avg_Mins_per_Event"] >= med_dur
        if hi_f and hi_d:
            return "Critical Chronic"
        if not hi_f and hi_d:
            return "Major Breakdown"
        if hi_f and not hi_d:
            return "Micro-Stops"
        return "Low Priority"

    q = q.copy()
    q["Strategy"] = q.apply(strategy, axis=1)
    ACTION = {
        "Critical Chronic": "Engineering redesign / RCA",
        "Major Breakdown": "Time-based PMs / spares",
        "Micro-Stops": "Autonomous maint / pacing",
        "Low Priority": "Monitor",
    }
    q["Recommended Action"] = q["Strategy"].map(ACTION)
    QC = {"Critical Chronic": "#d64545", "Major Breakdown": "#e08a3c",
          "Micro-Stops": "#3f9142", "Low Priority": "#9aa5ad"}

    xmin, xmax = q["Total_Count"].min(), q["Total_Count"].max()
    ymin, ymax = q["Avg_Mins_per_Event"].min(), q["Avg_Mins_per_Event"].max()
    xlo, xhi = xmin / 3.5, xmax * 3.5
    ylo, yhi = ymin / 3.0, ymax * 3.0

    fig, ax = plt.subplots(figsize=(13, 8))
    ax.axvspan(med_count, xhi, color="#f2f6f8", zorder=0)
    ax.axhspan(med_dur, yhi, color="#f2f6f8", zorder=0)
    ax.set_xscale("log")
    ax.set_yscale("log")
    ax.set_xlim(xlo, xhi)
    ax.set_ylim(ylo, yhi)

    sizes = 120 + (q["Total_Duration_hours"] / q["Total_Duration_hours"].max()) * 800
    for (_, row), s in zip(q.iterrows(), sizes):
        ax.scatter(row["Total_Count"], row["Avg_Mins_per_Event"], s=s,
                   color=QC[row["Strategy"]], edgecolors="black", linewidth=1.2, alpha=0.9, zorder=3)

    CANDIDATES = [(0, 26), (0, -34), (46, 10), (-46, 10), (46, -24), (-46, -24), (0, 46), (0, -52)]

    def overlaps(box, others, pad=4):
        for o in others:
            if not (box[2] + pad < o[0] or box[0] - pad > o[2] or box[3] + pad < o[1] or box[1] - pad > o[3]):
                return True
        return False

    fig.canvas.draw()
    placed = []
    for idx in q.assign(_s=sizes).sort_values("_s", ascending=False).index:
        row = q.loc[idx]
        x, y = row["Total_Count"], row["Avg_Mins_per_Event"]
        txt = (f"{row[label_col]}\n{row['Total_Count']:.0f} ev \u00b7 "
               f"{row['Avg_Mins_per_Event']:.0f} min \u00b7 {row['Total_Duration_hours']:.0f} hrs")
        chosen = None
        for dx, dy in CANDIDATES:
            ha = "center" if dx == 0 else ("left" if dx > 0 else "right")
            va = "bottom" if dy > 0 else "top"
            ann = ax.annotate(txt, xy=(x, y), xytext=(dx, dy), textcoords="offset points",
                              ha=ha, va=va, fontsize=9,
                              bbox=dict(boxstyle="round,pad=0.35", facecolor="white",
                                        edgecolor=QC[row["Strategy"]], alpha=0.95), zorder=5)
            fig.canvas.draw()
            bb = ann.get_window_extent()
            box = (bb.x0, bb.y0, bb.x1, bb.y1)
            ax_bb = ax.get_window_extent()
            inside = (box[0] >= ax_bb.x0 - 2 and box[2] <= ax_bb.x1 + 2 and
                      box[1] >= ax_bb.y0 - 2 and box[3] <= ax_bb.y1 + 2)
            if inside and not overlaps(box, placed):
                chosen = box
                break
            ann.remove()
        if chosen is None:
            ann = ax.annotate(txt, xy=(x, y), xytext=(0, 26), textcoords="offset points",
                              ha="center", va="bottom", fontsize=9,
                              bbox=dict(boxstyle="round,pad=0.35", facecolor="white",
                                        edgecolor=QC[row["Strategy"]], alpha=0.95), zorder=5)
            fig.canvas.draw()
            bb = ann.get_window_extent()
            chosen = (bb.x0, bb.y0, bb.x1, bb.y1)
        placed.append(chosen)

    ax.axvline(med_count, color="black", linestyle="--", alpha=0.55, zorder=2)
    ax.axhline(med_dur, color="black", linestyle="--", alpha=0.55, zorder=2)
    ax.set_xlabel("Number of events (log scale)")
    ax.set_ylabel("Average minutes per event (log scale)")
    ax.set_title(f"{title}\nbubble size = total downtime hours", fontsize=14, fontweight="bold")
    for tx, ty, lab, col, va, ha in [
            (0.015, 0.975, "MAJOR BREAKDOWN\nrare but long", "#e08a3c", "top", "left"),
            (0.985, 0.975, "CRITICAL CHRONIC\nfrequent and long", "#d64545", "top", "right"),
            (0.015, 0.025, "LOW PRIORITY\nrare and short", "#9aa5ad", "bottom", "left"),
            (0.985, 0.025, "MICRO-STOPS\nfrequent but short", "#3f9142", "bottom", "right")]:
        ax.text(tx, ty, lab, transform=ax.transAxes, fontsize=9, fontweight="bold",
                color=col, va=va, ha=ha, alpha=0.7)
    handles = [mpatches.Patch(color=c, label=f"{k} - {ACTION[k]}") for k, c in QC.items()]
    handles.append(plt.Line2D([], [], color="black", linestyle="--", alpha=0.55,
                              label=f"Medians: {med_count:.0f} ev / {med_dur:.0f} min"))
    ax.legend(handles=handles, loc="upper left", bbox_to_anchor=(1.01, 1), fontsize=9)
    ax.grid(True, alpha=0.25, which="both")
    plt.tight_layout()
    return fig, q[[label_col, "Total_Count", "Avg_Mins_per_Event", "Total_Duration_hours",
                   "Strategy", "Recommended Action"]]


def bootstrap_oee_simulation(daily_downtime_hours, period_days=30, n_trials=10000, seed=42):
    """Vectorized: draws the full (n_trials, period_days) sample matrix in
    one call instead of looping n_trials times - ~20x faster, confirmed
    statistically equivalent to the original loop-based version (same
    mean/std across repeated runs, since both are the same bootstrap
    process, just drawn in a different order)."""
    rng = np.random.default_rng(seed)
    values = daily_downtime_hours.dropna().values
    samples = rng.choice(values, size=(n_trials, period_days), replace=True)
    return ((24 - samples) / 24 * 100).mean(axis=1)


def simulate_category_downtime(total_hours, total_count, observed_days, period_days=30,
                               n_trials=10000, rng=None):
    """Vectorized via a mathematical shortcut: the sum of N independent
    Exponential(scale) draws is exactly Gamma(shape=N, scale=scale). So
    instead of looping n_trials times (draw a Poisson event count, then
    draw and sum that many individual exponentials), this draws all
    n_trials event counts at once, then all n_trials summed durations at
    once via the Gamma distribution - ~27x faster, confirmed statistically
    equivalent to the original loop across repeated runs with different
    seeds (mean difference ~0.001, well within the ~0.045 run-to-run
    variation either version has)."""
    if rng is None:
        rng = np.random.default_rng(42)
    if total_count == 0 or total_hours == 0:
        return np.zeros(n_trials)
    rate_per_day = total_count / observed_days
    mean_minutes_per_event = (total_hours * 60) / total_count
    n_events = rng.poisson(rate_per_day * period_days, size=n_trials)
    sim_minutes = np.zeros(n_trials)
    nonzero = n_events > 0
    sim_minutes[nonzero] = rng.gamma(shape=n_events[nonzero], scale=mean_minutes_per_event)
    return sim_minutes / 60


def compute_weibull(alloc_like, timeframe_days, beta=1.5, top_n=3):
    """MTBF (uptime / failures) for the top N sources by total downtime,
    Weibull scale eta = MTBF / 0.9027 (valid for beta=1.5 specifically)."""
    rows = []
    for _, r in alloc_like.iterrows():
        if r["Total_Count"] <= 0:
            continue
        observation_hours = timeframe_days * 24
        uptime_hours = max(observation_hours - r["Total_Duration_hours"], 0.01)
        mtbf = uptime_hours / r["Total_Count"]
        rows.append({"Source": r.get("Allocated_To", r.get("Theme", "?")),
                     "Total_Downtime_hrs": r["Total_Duration_hours"],
                     "Total_Count": int(r["Total_Count"]), "MTBF_hrs": mtbf})
    if not rows:
        return pd.DataFrame()
    df = pd.DataFrame(rows).sort_values("Total_Downtime_hrs", ascending=False).head(top_n)
    df["Eta"] = df["MTBF_hrs"] / 0.9027
    df["Beta"] = beta
    return df.reset_index(drop=True)


def plot_weibull(weibull_df):
    fig, axes = plt.subplots(1, 2, figsize=(14, 5.5))
    colors = ["#d64545", "#e08a3c", "#4a7a96", "#3f9142", "#7a4f9e"]
    t_max = weibull_df["Eta"].max() * 2.5 if len(weibull_df) else 1
    t = np.linspace(0.1, max(t_max, 1), 400)
    for i, (_, row) in enumerate(weibull_df.iterrows()):
        eta, beta = row["Eta"], row["Beta"]
        color = colors[i % len(colors)]
        survival = np.exp(-((t / eta) ** beta)) * 100
        axes[0].plot(t, survival, linewidth=2.5, color=color, label=f"{row['Source']} (MTBF {row['MTBF_hrs']:.0f}h)")
        pdf = (beta / eta) * ((t / eta) ** (beta - 1)) * np.exp(-((t / eta) ** beta))
        axes[1].plot(t, pdf, linewidth=2.5, color=color, label=str(row["Source"]))
    axes[0].axhline(50, color="grey", linestyle="--", alpha=0.6, label="50% survival")
    axes[0].set_title("Weibull Survival Curve", fontsize=13, fontweight="bold")
    axes[0].set_xlabel("Continuous Operating Time (hours)")
    axes[0].set_ylabel("Survival Probability (%)")
    axes[0].set_ylim(0, 105)
    axes[0].legend(fontsize=8)
    axes[0].grid(alpha=0.3)
    axes[1].set_title("Failure Probability Density (\"Bell\" Curve)", fontsize=13, fontweight="bold")
    axes[1].set_xlabel("Continuous Operating Time (hours)")
    axes[1].set_ylabel("Probability Density")
    axes[1].legend(fontsize=8)
    axes[1].grid(alpha=0.3)
    axes[1].set_yticks([])
    plt.tight_layout()
    return fig


# ============================================================
# THEME PATTERN LIBRARIES
# Roaster/Winnower/Grinder theme content is IDENTICAL between the two
# notebooks (same process mechanics), so those three are shared constants.
# Destoner and Debacterizer only exist as their own theme libraries in the
# Buhler Line notebook (Barth's Destoner never got its own dedicated theme
# library - only the equipment-attribution keywords below). Cooler and
# PreGrinder theme libraries are Line-2-only for the same reason.
# ============================================================

ROASTER_THEME_PATTERNS = {
    "Combustion / Heating System": [
        r"\bburner", r"\bflame\b", r"\bignition\b", r"\bpilot light\b", r"\bfire ?box\b",
        r"\bcombustion\b", r"\bfuel\b", r"\bLPG\b", r"\bgas supply\b", r"\binfrared\b",
        r"\bIR\b", r"\bhot air\b", r"\bthermocouple\b", r"\btemperature\b", r"\boverheat",
        r"\bheating element\b", r"\broast(ing)?\b",
    ],
    "Utility Supply Interruption": [
        r"\bcompressed air\b", r"\bsteam\b", r"\butilit", r"\bpower\b", r"\belectric",
        r"\bwater supply\b", r"\bnitrogen\b",
    ],
    "Control System / Automation Fault": [
        r"\bPLC\b", r"\bmes-?plc\b", r"\bHMI\b", r"\bSCADA\b", r"\bsensor\b",
        r"\binstrument", r"\bcommunication\b", r"\bwiring\b", r"\bVFD\b",
        r"\binverter\b", r"\bsoftware\b",
    ],
    "Mechanical / Drive System": [
        r"\bmotor\b", r"\bgearbox\b", r"\bbearing\b", r"\bbelt\b", r"\bchain\b",
        r"\bdrum\b", r"\bcoupling\b", r"\bshaft\b", r"\bvibration\b", r"\bseal\b",
        r"\bgasket\b", r"\bmechanical\b", r"\broller\b",
    ],
    "Material Handling / Transfer": [
        r"\bfeed(er|ing)?\b", r"\bconveyor\b", r"\bauger\b", r"\bhopper\b",
        r"\bblockage\b", r"\bjam(med)?\b", r"\bchoke", r"\bdischarge\b", r"\bspillage\b",
        r"\btransfer\b", r"\belevator\b", r"\bpneumatic\b", r"\bchute\b",
        r"\bduct\b", r"\bpipe(line)?\b",
    ],
    "Quality / Process Deviation": [
        r"\bquality\b", r"\boff-?spec\b", r"\bmoisture\b", r"\broast profile\b",
        r"\bcolour?\b", r"\breject\b", r"\btolerance\b",
    ],
    "Safety / Emergency Stop": [
        r"\bE-?stop\b", r"\bemergency\b", r"\binterlock\b", r"\bguard\b",
        r"\btrip(ped)?\b", r"\block ?out\b", r"\bsafety\b",
    ],
    "Cleaning / Sanitation": [
        r"\bclean(ing)?\b", r"\bCIP\b", r"\bwash ?down\b", r"\bhygiene\b", r"\bsanitation\b",
    ],
    "Planned / Scheduled Downtime": [
        r"\bmaintenance\b", r"\bchangeover\b", r"\bproduction test\b", r"\bengineering\b",
        r"\bcivil\b", r"\bR&D\b", r"\bfeasibilit", r"\bcalibrat", r"\binspection\b",
        r"\baudit\b",
    ],
    "Other / Unclassified Technical": [
        r"\bother technical\b", r"\bunspecified\b",
    ],
}

WINNOWER_THEME_PATTERNS = {
    "Shelling / Cracking Mechanism": [
        r"\bshell", r"\bhusk", r"\bcrack(ed|ing)?\b", r"\bcracker\b", r"\bdeshell",
    ],
    "Air Classification / Aspiration": [
        r"\baspirat", r"\bair classif", r"\bcyclone\b", r"\bblower\b", r"\bfan\b",
    ],
    "Material Handling / Transfer": [
        r"\btransfer\b", r"\bconvey", r"\belevator\b", r"\bauger\b", r"\bpneumatic\b",
        r"\bchute\b", r"\bduct\b", r"\bpipe(line)?\b", r"\bblockage\b", r"\bjam(med)?\b",
    ],
    "Nib Quality / Yield": [
        r"\bnib yield\b", r"\bquality\b", r"\boff-?spec\b", r"\bcontaminat", r"\bshell content\b",
    ],
    "Raw Material Supply (Bean/Nibs Shortage)": [
        r"\bbean.*shortage\b", r"\bnibs?.*shortage\b", r"\braw material\b",
        r"\bsupply\b", r"\bunavailable\b",
    ],
    "Mechanical / Drive System": [
        r"\bmotor\b", r"\bbearing\b", r"\bbelt\b", r"\bvibration\b", r"\broller\b", r"\bgearbox\b",
    ],
    "Control System / Automation Fault": [
        r"\bPLC\b", r"\bmes-?plc\b", r"\bsensor\b", r"\binstrument", r"\bcommunication\b",
    ],
    "Safety / Emergency Stop": [
        r"\bE-?stop\b", r"\bemergency\b", r"\binterlock\b", r"\bguard\b", r"\btrip(ped)?\b",
    ],
    "Cleaning / Sanitation": [
        r"\bclean(ing)?\b", r"\bCIP\b", r"\bwash ?down\b",
    ],
    "Planned / Scheduled Downtime": [
        r"\bmaintenance\b", r"\bchangeover\b", r"\bproduction test\b", r"\bengineering\b",
        r"\bR&D\b", r"\bfeasibilit", r"\binspection\b",
    ],
}

GRINDER_THEME_PATTERNS = {
    "Milling / Grinding Mechanism": [
        r"\bmill(ing)?\b", r"\bgrind", r"\bdisc mill\b", r"\bball mill\b", r"\bstone mill\b",
        r"\brefin", r"\bconch", r"\bbeater blade\b",
    ],
    "Liquor Handling": [
        r"\bliquor\b", r"\bliquor\s*tank\b", r"\bcocoa mass\b", r"\bpump",
    ],
    "Material Handling / Transfer": [
        r"\btransfer\b", r"\bconvey", r"\belevator\b", r"\bauger\b", r"\bpneumatic\b",
        r"\bchute\b", r"\bduct\b", r"\bpipe(line)?\b", r"\bblockage\b", r"\bjam(med)?\b",
    ],
    "Particle Size / Quality": [
        r"\bparticle size\b", r"\bfineness\b", r"\bmicron", r"\bquality\b", r"\boff-?spec\b",
    ],
    "Mechanical / Drive System": [
        r"\bmotor\b", r"\bbearing\b", r"\bbelt\b", r"\bvibration\b", r"\bgearbox\b",
    ],
    "Control System / Automation Fault": [
        r"\bPLC\b", r"\bmes-?plc\b", r"\bsensor\b", r"\binstrument", r"\bcommunication\b",
    ],
    "Safety / Emergency Stop": [
        r"\bE-?stop\b", r"\bemergency\b", r"\binterlock\b", r"\bguard\b", r"\btrip(ped)?\b",
    ],
    "Cleaning / Sanitation": [
        r"\bclean(ing)?\b", r"\bCIP\b", r"\bwash ?down\b",
    ],
    "Planned / Scheduled Downtime": [
        r"\bmaintenance\b", r"\bchangeover\b", r"\bproduction test\b", r"\bengineering\b",
        r"\bR&D\b", r"\bfeasibilit", r"\binspection\b",
    ],
}

# Barth-only: what the roaster's Unattributable bucket actually contains,
# by the KIND of gap each reason represents (not by equipment).
UNATTRIBUTABLE_THEME_PATTERNS = {
    "Unspecified Quality Issue": [
        r"\bquality\b", r"\boff-?spec\b", r"\breject", r"\bdefect",
    ],
    "Unspecified Material / Supply Issue": [
        r"\bshortage\b", r"\bunavailable\b", r"\braw material\b", r"\bsupply\b",
    ],
    "Generic / Unspecified Waiting": [
        r"\bwait(ing)?\b",
    ],
    "Unspecified Technical Fault": [
        r"\bother technical\b", r"\bunspecified\b",
    ],
}

# Line-2-only theme libraries
DESTONER_THEME_PATTERNS = {
    "Utility Supply Interruption": [
        r"\bcompressed air\b", r"\bpressure switch\b", r"\bsteam\b", r"\butilit",
        r"\bpower\b", r"\belectric",
    ],
    "Air Classification / Aspiration": [
        r"\baspirat", r"\bair classif", r"\bcyclone\b",
    ],
    "Foreign Material / Stone Removal": [
        r"\bstone", r"\bgravel\b", r"\btramp metal\b", r"\bmetal detect", r"\bmagnet",
        r"\brotary magnet\b", r"\bforeign material\b", r"\bforeign bod", r"\bimpurit",
        r"\bcontaminant",
    ],
    "Density / Gravity Separation": [
        r"\bdensity separat", r"\bgravity separat", r"\bbean cleaning\b", r"\bpre-?clean",
    ],
    "Raw Material Supply (Bean Shortage)": [
        r"\bbean.*\bshortage\b", r"\bshortage\b.*\bbean",
    ],
    "Downstream Blockage / Material Backup": [
        r"\bhopper full\b", r"\btank full\b", r"\bbin full\b", r"\bfull\b.*\bhopper\b",
    ],
    "Material Handling / Transfer": [
        r"\btransfer\b", r"\bconvey", r"\belevator\b", r"\bauger\b", r"\bpneumatic\b",
        r"\bchute\b", r"\bduct\b", r"\bpipe(line)?\b", r"\bblockage\b", r"\bjam(med)?\b",
    ],
    "Control System / Automation Fault": [
        r"\bPLC\b", r"\bmes-?plc\b", r"\bsensor\b", r"\binstrument", r"\bcommunication\b",
        r"\bvalve\b", r"\bfeedback\b",
    ],
    "Mechanical / Drive System": [
        r"\bmotor\b", r"\bbearing\b", r"\bbelt\b", r"\bvibration\b", r"\broller\b", r"\bgearbox\b",
    ],
    "Safety / Emergency Stop": [
        r"\bE-?stop\b", r"\bemergency\b", r"\binterlock\b", r"\bguard\b", r"\btrip(ped)?\b",
    ],
    "Cleaning / Sanitation": [
        r"\bclean(ing)?\b", r"\bCIP\b", r"\bwash ?down\b",
    ],
    "Planned / Scheduled Downtime": [
        r"\bmaintenance\b", r"\bchangeover\b", r"\bproduction test\b", r"\bengineering\b",
        r"\bR&D\b", r"\bfeasibilit", r"\binspection\b",
    ],
    "No Specific Cause Given": [
        r"\bunknown\b", r"\bother\b.*\bplease comment\b", r"\bAU Destoner\s*-\s*BP_PL02",
    ],
}

DEBACTERIZER_THEME_PATTERNS = {
    "Decontamination / Pressure Vessel Process Fault": [
        r"\bdebacteri[sz]", r"\bpasteuri[sz]", r"\bsterili[sz]", r"\bdecontaminat",
        r"\bmicrobial\b", r"\bbacteria\b", r"\bpathogen", r"\bautoclave\b", r"\bmoist heat\b",
        r"\bvessel\b", r"\bpressure\b", r"\bfilling\b",
    ],
    "Steam / Heat Supply": [
        r"\bsteam\b", r"\btemperature\b", r"\bheat(ing)?\b", r"\boverheat",
    ],
    "Upstream Starvation (Waiting for Roaster)": [
        r"\bwaiting product from roaster\b", r"\bwaiting\b.*\broaster\b",
    ],
    "Downstream Blockage / Material Backup": [
        r"\bhopper full\b", r"\bdownstream not ready\b", r"\btank full\b", r"\bbin full\b",
    ],
    "Raw Material Supply (Bean Shortage)": [
        r"\bbean.*\bshortage\b", r"\bshortage\b.*\bbean",
    ],
    "Planned / Scheduled Downtime": [
        r"\bmaintenance\b", r"\bchangeover\b", r"\bproduction test\b", r"\bengineering\b",
        r"\bR&D\b", r"\bfeasibilit", r"\binspection\b", r"\boverhaul\b", r"\bannual\b",
        r"\bpest\b",
    ],
    "Utility Supply Interruption": [
        r"\bpower\b", r"\belectric", r"\bsurge\b", r"\boutage\b", r"\bcompressed air\b", r"\butilit",
    ],
    "Material Handling / Transfer": [
        r"\btransfer\b", r"\bconvey", r"\belevator\b", r"\bauger\b", r"\bpneumatic\b",
        r"\bchute\b", r"\bduct\b", r"\bpipe(line)?\b", r"\bblockage\b", r"\bjam(med)?\b",
        r"\bdischarge\b", r"\bdelay\b",
    ],
    "Control System / Automation Fault": [
        r"\bPLC\b", r"\bmes-?plc\b", r"\bsensor\b", r"\binstrument", r"\bcommunication\b",
        r"\bvalve\b", r"\bthreshold\b", r"\bfeedback\b",
    ],
    "Mechanical / Drive System": [
        r"\bmotor\b", r"\bbearing\b", r"\bbelt\b", r"\bvibration\b", r"\bgearbox\b",
    ],
    "Safety / Emergency Stop": [
        r"\bE-?stop\b", r"\bemergency\b", r"\binterlock\b", r"\bguard\b", r"\btrip(ped)?\b",
        r"\bfire\b",
    ],
    "Cleaning / Sanitation": [
        r"\bclean(ing)?\b", r"\bCIP\b", r"\bwash ?down\b", r"\bcontaminat",
    ],
    "Other / Unclassified Technical": [
        r"\bunknown\b", r"\bother\b.*\bplease comment\b", r"\bother technical\b", r"\bunspecified\b",
    ],
}

COOLER_THEME_PATTERNS = {
    "Cooling System Fault": [
        r"\bcooler\b", r"\bcooling\b", r"\bchiller\b", r"\bchilled\b", r"\bbed cooler\b",
        r"\bfluidi[sz]ed bed\b", r"\bcooling tower\b", r"\bquench",
    ],
    "Material Handling / Transfer": [
        r"\btransfer\b", r"\bconvey", r"\belevator\b", r"\bauger\b", r"\bpneumatic\b",
        r"\bchute\b", r"\bduct\b", r"\bpipe(line)?\b", r"\bblockage\b", r"\bjam(med)?\b",
    ],
    "Mechanical / Drive System": [
        r"\bmotor\b", r"\bbearing\b", r"\bbelt\b", r"\bvibration\b", r"\bfan\b", r"\bgearbox\b",
    ],
    "Control System / Automation Fault": [
        r"\bPLC\b", r"\bmes-?plc\b", r"\bsensor\b", r"\binstrument", r"\bcommunication\b",
    ],
    "Safety / Emergency Stop": [
        r"\bE-?stop\b", r"\bemergency\b", r"\binterlock\b", r"\bguard\b", r"\btrip(ped)?\b",
    ],
    "Cleaning / Sanitation": [
        r"\bclean(ing)?\b", r"\bCIP\b", r"\bwash ?down\b",
    ],
    "Planned / Scheduled Downtime": [
        r"\bmaintenance\b", r"\bchangeover\b", r"\bproduction test\b", r"\bengineering\b",
        r"\bR&D\b", r"\bfeasibilit", r"\binspection\b",
    ],
}

PREGRINDER_THEME_PATTERNS = {
    "Pre-Grinding / Crushing Mechanism": [
        r"\bpre-?grind", r"\bpre-?mill", r"\bcrush(ing)?\b", r"\bcoarse (grind|reduction)\b",
    ],
    "Material Handling / Transfer": [
        r"\btransfer\b", r"\bconvey", r"\belevator\b", r"\bauger\b", r"\bpneumatic\b",
        r"\bchute\b", r"\bduct\b", r"\bpipe(line)?\b", r"\bblockage\b", r"\bjam(med)?\b",
    ],
    "Mechanical / Drive System": [
        r"\bmotor\b", r"\bbearing\b", r"\bbelt\b", r"\bvibration\b", r"\bgearbox\b",
    ],
    "Control System / Automation Fault": [
        r"\bPLC\b", r"\bmes-?plc\b", r"\bsensor\b", r"\binstrument", r"\bcommunication\b",
    ],
    "Safety / Emergency Stop": [
        r"\bE-?stop\b", r"\bemergency\b", r"\binterlock\b", r"\bguard\b", r"\btrip(ped)?\b",
    ],
    "Cleaning / Sanitation": [
        r"\bclean(ing)?\b", r"\bCIP\b", r"\bwash ?down\b",
    ],
    "Planned / Scheduled Downtime": [
        r"\bmaintenance\b", r"\bchangeover\b", r"\bproduction test\b", r"\bengineering\b",
        r"\bR&D\b", r"\bfeasibilit", r"\binspection\b",
    ],
}

# Presses theme library. Ported unchanged from the Presses notebook, where it
# took three rounds of correction against real PR2 reasons. One addition made
# here after checking it against the real PR1 log: "fuse" (an electrical
# fault) was the only one of 39 real PR1 reasons that matched no theme.
PRESS_THEME_PATTERNS = {
    "Hydraulic System Fault": [
        r"\bhydraul", r"\bpump\b", r"\boil\b", r"\bleak", r"\bseal\b", r"\bram\b",
        r"\bhose\b", r"\bcylinder\b", r"\bvalve\b", r"\bcounter.?pot\b", r"\bGDO\b",
    ],
    "Liquor Feed & Conditioning (LCS)": [
        r"\bliquor\b", r"\bfeed(er|ing)?\b", r"\bconditioner\b", r"\blcs\b",
        r"\bblockage\b", r"\bchoke", r"\bviscosity\b", r"\bheating\b", r"\bcooling\b",
    ],
    "Cake Discharge / Handling": [
        r"\bcake\b", r"\bbreaker\b", r"\bdischarge\b", r"\bconveyor\b",
        r"\bjam(med)?\b", r"\bchute\b", r"\bauger\b", r"\bsilo\b", r"\bpackaging\b",
    ],
    "Butter Handling / Pumping": [
        r"\bbutter(ing)?\b", r"\bscale\b", r"\bweigh", r"\bbutter line\b",
    ],
    "Filter Cloth / Screens / Felts": [
        r"\bcloth\b", r"\bfilter\b", r"\bscreen\b", r"\bmesh\b", r"\btear\b",
        r"\bblind", r"\bfelt\b", r"\bfelt cord\b", r"\bplate\b",
    ],
    "Quality / Process Deviation": [
        r"\bquality\b", r"\boff-?spec\b", r"\breject\b", r"\btolerance\b"
    ],
    "Mechanical / Drive System": [
        r"\bmotor\b", r"\bgearbox\b", r"\bbearing\b", r"\bbelt\b", r"\bshaft\b",
        r"\bvibration\b", r"\bmechanical\b"
    ],
    "Control System / Automation": [
        r"\bPLC\b", r"\bsensor\b", r"\binstrument", r"\bcommunication\b", r"\binterlock\b",
        r"\bfuse\b",
    ],
    "Safety / Emergency Stop": [
        r"\bE-?stop\b", r"\bemergency\b", r"\bguard\b", r"\btrip(ped)?\b"
    ],
    "Upstream Starvation / Material Shortage": [
        r"\braw material\b", r"\bshortage\b", r"\bunavailable\b", r"\bstarvat"
    ],
    "Utility Supply Interruption": [
        r"\bcompressed air\b", r"\bsteam\b", r"\butilit", r"\bpower\b", r"\belectric"
    ],
    "Cleaning / Sanitation / Changeover": [
        r"\bclean(ing)?\b", r"\bwash ?down\b", r"\bhygiene\b", r"\bchange.?over\b",
        r"\bpress settings\b",
    ],
    "Planned / Scheduled Maintenance": [
        r"\bmaintenance\b", r"\bengineering\b", r"\bproduction test\b", r"\binspection\b",
        r"\bplanned downtime\b", r"\bshutdown\b", r"\bweekly\b", r"\bunnamed\b", r"\binventory\b",
    ],
    "Demand-Driven / No Production Need": [
        r"\bno demand\b",
    ],
    "Staffing / Labor Availability": [
        r"\babsenteeism\b",
    ],
    "Minor Stops / Micro-Stoppages": [
        r"\bminor stop",
    ],
    "No Specific Cause Given": [
        r"\bunknown\b", r"\bplease comment\b", r"\bunclas"
    ],
}


# ============================================================
# EQUIPMENT-ATTRIBUTION PATTERNS (for classifying the roaster's own
# comments by which unit they name) - separate from the theme libraries
# above, and different per line since the equipment sets differ.
# ============================================================

BARTH_EQUIPMENT_PATTERNS = {
    "Destoner (DES11)": [
        r"\bDES1?1\b", r"\bdestoner\b", r"\bdestoning\b", r"\bde-?stoner\b",
        r"\bstone", r"\bgravel\b", r"\btramp metal\b", r"\bmetal detect", r"\bmagnet",
        r"\bforeign material\b", r"\bforeign bod", r"\bimpurit", r"\bcontaminant",
        r"\bdensity separat", r"\bgravity separat", r"\bbean cleaning\b", r"\bpre-?clean",
    ],
    "Winnower (WIN11)": [
        r"\bWIN1?1\b", r"\bwinnow", r"\bshell", r"\bhusk", r"\bnib\b", r"\bnibs\b",
        r"\bcracker\b", r"\bcracking\b", r"\bcrack(ed|ing)? bean", r"\baspirat",
        r"\bair classif", r"\bcyclone\b", r"\bnib yield\b", r"\bshelling\b", r"\bdeshell",
    ],
    "Potash": [
        r"\bpotash\b", r"\bpotassium carbonate\b", r"\bK2CO3\b", r"\balkali",
        r"\bdutch(ing|ed)?\b", r"\bneutrali[sz]", r"\bsolution prep", r"\bcaustic\b",
    ],
    "Mixer": [
        r"\bmixer\b", r"\bBP_PL01_FB11\b", r"\bmixing\b", r"\bblend", r"\bagitat",
        r"\bhomogeni[sz]", r"\bbatch mix", r"\bpaste\b", r"\brecipe\b",
    ],
    "Cooler (CO11)": [
        r"\bCO1?1\b", r"\bcooler\b", r"\bcooling\b", r"\bchiller\b", r"\bchilled\b",
        r"\bbed cooler\b", r"\bfluidi[sz]ed bed\b", r"\bcooling tower\b", r"\bquench",
    ],
    "Grinder (GR11)": [
        r"\bGR1?1\b", r"\bgrind", r"\bliquor\b", r"\bliquor\s*tank\b", r"\bcocoa mass\b",
        r"\bmill(ing)?\b", r"\bball mill\b", r"\bdisc mill\b", r"\bstone mill\b",
        r"\bbeater blade\b", r"\brefin", r"\bconch", r"\bparticle size\b",
        r"\bfineness\b", r"\bmicron", r"\bpre-?grind",
    ],
}

BARTH_ROASTER_OWN_KEYWORDS = [
    r"\bburner", r"\bflame\b", r"\bcombustion\b", r"\broast", r"\bdrum\b",
    r"\binfrared\b", r"\bIR\b", r"\bhot air\b", r"\bpilot light\b", r"\bthermocouple\b",
    r"\btemperature profile\b", r"\bfire ?box\b", r"\bBP_PL01_RN11\b",
    r"\bcompressed air\b", r"\bsteam\b", r"\bPLC\b", r"\bmaintenance\b",
    r"\bchangeover\b", r"\butilities\b", r"\bengineering\b", r"\bproduction test\b",
]

LINE2_EQUIPMENT_PATTERNS = {
    "Destoner (DES201)": [
        r"\bDES2?01\b", r"\bdestoner\b", r"\bdestoning\b", r"\bde-?stoner\b",
        r"\bstone", r"\bgravel\b", r"\btramp metal\b", r"\bmetal detect", r"\bmagnet",
        r"\bforeign material\b", r"\bforeign bod", r"\bimpurit", r"\bcontaminant",
        r"\bdensity separat", r"\bgravity separat", r"\bbean cleaning\b", r"\bpre-?clean",
    ],
    "Debacterizer (DR201)": [
        r"\bDR2?01\b", r"\bdebacteri[sz]", r"\bpasteuri[sz]", r"\bsterili[sz]",
        r"\bdecontaminat", r"\bmicrobial\b", r"\bbacteria\b", r"\bpathogen",
        r"\bautoclave\b", r"\bmoist heat\b",
    ],
    "Winnower (WIN201)": [
        r"\bWIN2?01\b", r"\bwinnow", r"\bshell", r"\bhusk", r"\bnib\b", r"\bnibs\b",
        r"\bcracker\b", r"\bcracking\b", r"\bcrack(ed|ing)? bean", r"\baspirat",
        r"\bair classif", r"\bcyclone\b", r"\bnib yield\b", r"\bshelling\b", r"\bdeshell",
    ],
    "Cooler (CO201)": [
        r"\bCO2?01\b", r"\bcooler\b", r"\bcooling\b", r"\bchiller\b", r"\bchilled\b",
        r"\bbed cooler\b", r"\bfluidi[sz]ed bed\b", r"\bcooling tower\b", r"\bquench",
    ],
    "PreGrinder": [
        r"\bpre-?grind", r"\bpre-?mill", r"\bcrush(ing)?\b", r"\bcoarse (grind|reduction)\b",
    ],
    "Grinding (GR21/GR31)": [
        r"\bGR2?1\b", r"\bGR3?1\b", r"\bgrind", r"\bliquor\b", r"\bliquor\s*tank\b",
        r"\bcocoa mass\b", r"\bmill(ing)?\b", r"\bball mill\b", r"\bdisc mill\b",
        r"\bstone mill\b", r"\bbeater blade\b", r"\brefin", r"\bconch",
        r"\bparticle size\b", r"\bfineness\b", r"\bmicron",
    ],
}

LINE2_ROASTER_OWN_KEYWORDS = [
    r"\bburner", r"\bflame\b", r"\bcombustion\b", r"\broast", r"\bdrum\b",
    r"\binfrared\b", r"\bIR\b", r"\bhot air\b", r"\bpilot light\b", r"\bthermocouple\b",
    r"\btemperature\b", r"\btemperature profile\b", r"\bfire ?box\b", r"\boverheat",
    r"\bcompressed air\b", r"\bsteam\b", r"\bPLC\b", r"\bmaintenance\b",
    r"\bchangeover\b", r"\butilities\b", r"\bengineering\b", r"\bproduction test\b",
]

# ============================================================
# PER-LINE CONFIGURATION - fixed column letters confirmed against real
# data in the two notebooks.
# ============================================================

BARTH = {
    "roaster": "Roaster (R011)",
    "file_label_prefix": "Barth",
    "sequence": ["Destoner (DES11)", "Winnower (WIN11)", "Potash", "Mixer",
                 "Roaster (R011)", "Cooler (CO11)", "Grinder (GR11)"],
    "daily_cols": {
        "Destoner (DES11)": {"usecols": "E,H", "skiprows": 6},
        "Winnower (WIN11)": {"usecols": "E,J", "skiprows": 8},
        "Potash": {"usecols": "E,P", "skiprows": 7},
        "Mixer": {"usecols": "E,X", "skiprows": 7},
        "Roaster (R011)": {"usecols": "E,AF", "skiprows": 7},
        "Cooler (CO11)": {"usecols": "E,AP", "skiprows": 7},
        "Grinder (GR11)": {"usecols": "E,AU", "skiprows": 7},
    },
    "roaster_category_cols": "CR,CS,CU,CY,DA,DC",
    "category_skiprows": 386,
    "equipment_patterns": BARTH_EQUIPMENT_PATTERNS,
    "roaster_own_keywords": BARTH_ROASTER_OWN_KEYWORDS,
    # (label, theme_patterns, color) - only these 3 units + Unattributable
    # ever got dedicated theme libraries in the Barth notebook.
    "theme_targets": [
        ("Roaster (own fault)", ROASTER_THEME_PATTERNS, "#d64545"),
        ("Winnower (WIN11)", WINNOWER_THEME_PATTERNS, "#4a7a96"),
        ("Grinder (GR11)", GRINDER_THEME_PATTERNS, "#3f8f6d"),
        ("Unattributable", UNATTRIBUTABLE_THEME_PATTERNS, "#9aa5ad"),
    ],
    "exclude_march_default": True,
}

LINE2 = {
    "roaster": "Roaster (RO201)",
    "file_label_prefix": "Buhler",
    "sequence": ["Destoner (DES201)", "Roaster (RO201)", "Debacterizer (DR201)",
                 "Winnower (WIN201)", "Cooler (CO201)", "PreGrinder",
                 "Grinder 1 (GR21)", "Grinder 2 (GR31)"],
    "daily_cols": {
        "Destoner (DES201)": {"usecols": "E,H", "skiprows": 7},
        "Roaster (RO201)": {"usecols": "E,J", "skiprows": 7},
        "Debacterizer (DR201)": {"usecols": "E,P", "skiprows": 7},
        "Winnower (WIN201)": {"usecols": "E,X", "skiprows": 7},
        "Cooler (CO201)": {"usecols": "E,AF", "skiprows": 7},
        "PreGrinder": {"usecols": "E,AP", "skiprows": 7},
        "Grinder 1 (GR21)": {"usecols": "E,AU", "skiprows": 7},
        "Grinder 2 (GR31)": {"usecols": "E,AY", "skiprows": 7},
    },
    "category_cols": {
        "Destoner (DES201)": "F,G,M,W,Z,AE",
        "Roaster (RO201)": "AN,AO,AX,BD,BH,BJ",
        "Debacterizer (DR201)": "BO,BP,BR,BV,BX,BZ",
        "Winnower (WIN201)": "CE,CF,CH,CL,CN,CP",
        "Cooler (CO201)": "CU,CV,CX,DB,DD,DF",
        "PreGrinder": "DK,DL,DN,DR,DT,DV",
        "Grinder 1 (GR21)": "EA,EB,ED,EH,EJ,EL",
        "Grinder 2 (GR31)": "EQ,ER,ET,EX,EZ,FB",
    },
    "category_skiprows": 386,
    "equipment_patterns": LINE2_EQUIPMENT_PATTERNS,
    "roaster_own_keywords": LINE2_ROASTER_OWN_KEYWORDS,
    "theme_targets": [
        ("Roaster (own fault)", ROASTER_THEME_PATTERNS, "#d64545"),
        ("Destoner (DES201)", DESTONER_THEME_PATTERNS, "#8c6d46"),
        ("Debacterizer (DR201)", DEBACTERIZER_THEME_PATTERNS, "#7a4f9e"),
        ("Winnower (WIN201)", WINNOWER_THEME_PATTERNS, "#4a7a96"),
        ("Cooler (CO201)", COOLER_THEME_PATTERNS, "#2e8b8b"),
        ("PreGrinder", PREGRINDER_THEME_PATTERNS, "#c98a3a"),
        ("Grinding (GR21/GR31)", GRINDER_THEME_PATTERNS, "#3f8f6d"),
    ],
    # Units whose own "By Category" log needs pulling separately, since the
    # roaster's own comments never name them (confirmed: zero attribution).
    "own_log_units": ["Destoner (DES201)", "Debacterizer (DR201)"],
    "exclude_march_default": False,
}

# Presses are a different shape from Barth/Buhler: there is no single roaster
# whose comments get reattributed to other units. Each of the three presses
# has its OWN "By Category" reason log (on its "PRx: PRx" unit), paired with
# an LCS stage that only has daily downtime. So this config has its own
# "presses" map and render_presses(), rather than the "roaster" keys the
# other two lines use. The daily column letters are the Barth first-six-unit
# letters the Presses notebook was built on; on the real file each press's
# daily total matched its own reason-log total exactly, which is strong
# evidence the PRx:PRx columns are right (the LCS columns follow the same
# pattern but are not separately confirmed). The OEE tab prints each unit's
# real Excel header as a check.
PRESSES_CFG = {
    "kind": "presses",
    "file_label_prefixes": ["PR1", "PR2", "PR3"],
    "sequence": ["PR1: LCS1", "PR1: PR1", "PR2: LCS2", "PR2: PR2", "PR3: LCS3", "PR3: PR3"],
    "daily_cols": {
        "PR1: LCS1": {"usecols": "E,H", "skiprows": 6},
        "PR1: PR1": {"usecols": "E,J", "skiprows": 8},
        "PR2: LCS2": {"usecols": "E,P", "skiprows": 7},
        "PR2: PR2": {"usecols": "E,X", "skiprows": 7},
        "PR3: LCS3": {"usecols": "E,AF", "skiprows": 7},
        "PR3: PR3": {"usecols": "E,AP", "skiprows": 7},
    },
    "presses": {
        "PR1": {"oee_source": "PR1: PR1", "paired_lcs": "PR1: LCS1", "color": "#d64545"},
        "PR2": {"oee_source": "PR2: PR2", "paired_lcs": "PR2: LCS2", "color": "#4a7a96"},
        "PR3": {"oee_source": "PR3: PR3", "paired_lcs": "PR3: LCS3", "color": "#3f8f6d"},
    },
    "category_cols": {
        "PR1: PR1": "AN,AO,AW,BA,BC,BE",
        "PR2: PR2": "BZ,CA,CC,CG,CI,CK",
        "PR3: PR3": "DF,DG,DI,DM,DO,DQ",
    },
    "category_skiprows": 386,
    "theme_patterns": PRESS_THEME_PATTERNS,
    "exclude_march_default": False,
}

LIQUOR_MOULDING_THEME_PATTERNS = {
    "Cleaning / Sanitation": ['\\bclean(ing)?\\b', '\\bfilter check\\b', '\\bmagnet cleaning\\b', '\\bpest treatment\\b'],
    "Raw Material / Liquor Supply Shortage": ['\\bliquor shortage\\b', '\\braw material shortage\\b', '\\bunavailable\\b', '\\btruck availability\\b', '\\bstorage capacity\\b'],
    "Demand-Driven / No Production Need": ['\\bno demand\\b'],
    "Dosing System Fault": ['\\bdosing\\b'],
    "Tempering / Heating System Fault": ['\\breheating\\b', '\\btempering\\b', '\\bmass temperature\\b', '\\btemperature\\b', '\\bhigh pressure\\b', '\\bheating zone\\b'],
    "Cooling Tunnel Fault": ['\\btunnel\\b'],
    "Weighing / Metal Detection Fault": ['\\bweight system\\b', '\\bmetal detector\\b'],
    "Quality Issue": ['\\bquality\\b'],
    "Line Sequencing / Conflict": ['\\bconflict\\b'],
    "Packaging Equipment Fault": ['\\bpackaging equipment\\b'],
    "Changeover / Product Change": ['\\bpackaging change\\b', '\\bproduct change\\b', '\\bchange.?over\\b'],
    "Material Handling / Conveyor": ['\\bconveyor\\b'],
    "Mechanical / Drive System Fault": ['\\bmotor\\b', '\\bchain\\b', '\\bgearbox\\b', '\\bbearing\\b'],
    "Control System / Automation Fault": ['\\bMES-?PLC\\b', '\\bcommunication failure\\b', '\\bcontrol voltage\\b', '\\blocal isolator\\b', '\\bautomation\\b'],
    "Safety / Emergency Stop": ['\\bemergency stop\\b', '\\bE-?stop\\b', '\\bemergency\\b'],
    "Feed System / Machine Readiness": ['\\bfeed pump\\b', '\\bmachine not ready\\b'],
    "Planned / Scheduled Maintenance": ['\\bmaintenance\\b', '\\bengineering\\b', '\\bcivil works\\b', '\\bweekly shutdown\\b', '\\bweekly start ?up\\b', '\\bproduction test\\b', '\\binventory\\b'],
    "Staffing / Labor Availability": ['\\babsenteeism\\b', '\\bprocess knowledge\\b', '\\bteam meeting\\b', '\\bteam brief\\b'],
    "Minor Stops / Micro-Stoppages": ['\\bminor stop'],
    "Generic / Unspecified Waiting": ['\\bwaiting time\\b'],
    "No Specific Cause Given": ['\\bunknown\\b', '\\bplease comment\\b', '\\bother technical issue\\b', '\\bgeneral alarm\\b'],
}

LIQUOR_MOULDING_UNITS = {
    "TM41": {"oee_source": "Liquor Moulding: TM41",
             "other_units": ["Liquor Moulding: ML", "Liquor Moulding: PK Group"], "color": "#d64545"},
    "ML41": {"oee_source": "ML41: ML41", "other_units": ["ML41: PK41"], "color": "#4a7a96"},
    "ML42": {"oee_source": "ML42: ML42", "other_units": ["ML42: PK42"], "color": "#3f8f6d"},
}

# Moulding lines are a different shape again from both Barth/Buhler and
# Presses: multiple fully-analysed units (own OEE, own reason log, own
# themes - no single roaster whose comments get reattributed), but each
# unit's number of supporting units VARIES (TM41 has two, ML41/ML42 have
# one each) rather than Presses' fixed one-LCS-per-press pattern. So this
# uses "units" with a LIST of other_units per entry, and its own
# render_moulding_line(), rather than reusing render_presses() as-is.
LIQUOR_MOULDING_CFG = {
    "kind": "moulding",
    "file_label_prefixes": ["Liquor Moulding", "ML41", "ML42"],
    "sequence": ["Liquor Moulding: TM41", "Liquor Moulding: ML", "Liquor Moulding: PK Group",
                "ML41: ML41", "ML41: PK41", "ML42: ML42", "ML42: PK42"],
    "daily_cols": {
        "Liquor Moulding: TM41":     {"usecols": "E,H",  "skiprows": 6},
        "Liquor Moulding: ML":       {"usecols": "E,J",  "skiprows": 8},
        "Liquor Moulding: PK Group": {"usecols": "E,P",  "skiprows": 7},
        "ML41: ML41":                {"usecols": "E,X",  "skiprows": 7},
        "ML41: PK41":                {"usecols": "E,AF", "skiprows": 7},
        "ML42: ML42":                {"usecols": "E,AP", "skiprows": 7},
        "ML42: PK42":                {"usecols": "E,AU", "skiprows": 7},
    },
    "units": LIQUOR_MOULDING_UNITS,
    "category_cols": {
        "Liquor Moulding: TM41": "AN,AO,AW,BA,BC,BE",
        "ML41: ML41":            "BZ,CA,CC,CG,CI,CK",
        "ML42: ML42":            "DF,DG,DI,DM,DO,DQ",
    },
    "category_skiprows": 386,
    "theme_patterns": LIQUOR_MOULDING_THEME_PATTERNS,
    "exclude_march_default": False,
}

BUTTER_MOULDING_THEME_PATTERNS = {
    "Cleaning / Sanitation": ['\\bclean(ing)?\\b', '\\bfilter check\\b', '\\bmagnet cleaning\\b', '\\bpest treatment\\b'],
    "Raw Material / Supply Shortage": ['\\bbutter shortage\\b', '\\braw material shortage\\b', '\\bunavailable\\b'],
    "Demand-Driven / No Production Need": ['\\bno demand\\b', '\\bahead of schedule\\b'],
    "Dosing System Fault": ['\\bdosing\\b', '\\bpipe not clear\\b'],
    "Tempering / Heating System Fault": ['\\breheating\\b', '\\btempering\\b', '\\bmass temperature\\b', '\\btemperature\\b', '\\bhigh pressure\\b', '\\bthermistor\\b'],
    "Cooling Tunnel Fault": ['\\btunnel\\b'],
    "Weighing / Metal Detection Fault": ['\\bweight system\\b', '\\bmetal detector\\b'],
    "Quality Issue": ['\\bquality\\b'],
    "Line Sequencing / Conflict": ['\\bconflict\\b'],
    "Packaging Equipment Fault": ['\\bpackaging equipment\\b'],
    "Changeover / Product Change": ['\\bpackaging change\\b', '\\bproduct change\\b', '\\bchange.?over\\b'],
    "Material Handling / Conveyor": ['\\bconveyor\\b'],
    "Control System / Automation Fault": ['\\bMES-?PLC\\b', '\\bcommunication failure\\b', '\\bcontrol voltage\\b', '\\blocal isolator\\b', '\\bautomation\\b', '\\bground fault interrupt\\b'],
    "Safety / Emergency Stop": ['\\bemergency stop\\b', '\\bE-?stop\\b', '\\bemergency\\b'],
    "Feed System / Machine Readiness": ['\\bfeed pump\\b', '\\bmachine not ready\\b'],
    "Planned / Scheduled Maintenance": ['\\bmaintenance\\b', '\\bengineering\\b', '\\bcivil works\\b', '\\bweekly shutdown\\b', '\\bweekly start ?up\\b', '\\bproduction test\\b', '\\binventory\\b'],
    "Staffing / Breaks / Shift Handover": ['\\babsenteeism\\b', '\\bprocess knowledge\\b', '\\bteam meeting\\b', '\\bteam brief\\b', '\\blunch break\\b', '\\bprayer\\b', '\\bhandover\\b', '\\bshift change\\b'],
    "Force Majeure / Utilities": ['\\butilit', '\\bforce majeure\\b'],
    "Minor Stops / Micro-Stoppages": ['\\bminor stop'],
    "Generic / Unspecified Waiting": ['\\bwaiting time\\b'],
    "No Specific Cause Given": ['\\bunknown\\b', '\\bplease comment\\b', '\\bother technical issue\\b'],
}

BUTTER_MOULDING_UNITS = {
    "TM11": {"oee_source": "Butter Moulding: TM11", "other_units": ["Butter Moulding: ML"], "color": "#d64545"},
    "ML11": {"oee_source": "ML11: ML11", "other_units": [], "color": "#4a7a96"},
    "ML12": {"oee_source": "ML12: ML12", "other_units": [], "color": "#3f8f6d"},
}

BUTTER_MOULDING_CFG = {
    "kind": "moulding",
    "file_label_prefixes": ["Butter Moulding", "ML11", "ML12"],
    "sequence": ["Butter Moulding: TM11", "Butter Moulding: ML", "ML11: ML11", "ML12: ML12"],
    "daily_cols": {
        "Butter Moulding: TM11": {"usecols": "E,H", "skiprows": 6},
        "Butter Moulding: ML":   {"usecols": "E,J", "skiprows": 8},
        "ML11: ML11":            {"usecols": "E,P", "skiprows": 7},
        "ML12: ML12":            {"usecols": "E,X", "skiprows": 7},
    },
    "units": BUTTER_MOULDING_UNITS,
    "category_cols": {
        "Butter Moulding: TM11": "AN,AO,AW,BA,BC,BE",
        "ML11: ML11":            "BZ,CA,CC,CG,CI,CK",
        "ML12: ML12":            "DF,DG,DI,DM,DO,DQ",
    },
    "category_skiprows": 386,
    "theme_patterns": BUTTER_MOULDING_THEME_PATTERNS,
    "exclude_march_default": False,
}


# Every selectable line, keyed by the name shown in the sidebar dropdown.
# Adding a line to the app means adding one entry here.
LINES = {
    "Barth Line": BARTH,
    "Buhler Line": LINE2,
    "Presses": PRESSES_CFG,
    "Liquor Moulding": LIQUOR_MOULDING_CFG,
    "Butter Moulding": BUTTER_MOULDING_CFG,
}


# ============================================================
# MAIN LINE FLOW - mirrors the notebook's section order for whichever
# line's config dict is passed in.
# ============================================================

# ============================================================
# PDF REPORT
# ============================================================

_PDF_CHAR_MAP = {
    "\u2013": "-", "\u2014": "-", "\u2212": "-", "\u2018": "'", "\u2019": "'",
    "\u201c": '"', "\u201d": '"', "\u2026": "...", "\u2192": "->", "\u2265": ">=",
    "\u2264": "<=", "\u2022": "*", "\u00a0": " ",
}


def _pdf_safe(text):
    """fpdf2's built-in Helvetica only covers Latin-1 and raises on anything
    else, so one en dash or arrow in a real reason text would crash the whole
    export. Common typographic characters are mapped to plain equivalents
    first; whatever is still unencodable becomes '?'."""
    text = str(text)
    for src, dst in _PDF_CHAR_MAP.items():
        text = text.replace(src, dst)
    return text.encode("latin-1", "replace").decode("latin-1")


def _wrap_text(pdf, text, max_width):
    """Greedy word-wrap using the PDF's real font metrics (the current font
    must already be set). Written by hand instead of using multi_cell,
    because multi_cell's cursor behaviour differs between fpdf2 versions."""
    lines, current = [], ""
    for word in str(text).split():
        trial = f"{current} {word}".strip()
        if pdf.get_string_width(trial) <= max_width:
            current = trial
            continue
        if current:
            lines.append(current)
        current = word
        while pdf.get_string_width(current) > max_width and len(current) > 1:
            cut = len(current) - 1
            while cut > 1 and pdf.get_string_width(current[:cut]) > max_width:
                cut -= 1
            lines.append(current[:cut])
            current = current[cut:]
    if current or not lines:
        lines.append(current)
    return lines


_COMMENT_COLUMNS = ["Group", "Reason", "Hours", "Events"]


def build_comments_table(reasons_df, groups):
    """One comments table: every reason row with the group it belongs to
    (a theme, or the unit it was attributed to), sorted by group total then
    by hours. `groups` is a Series aligned with reasons_df; missing values
    become "Unclassified" so no reason is silently dropped."""
    if reasons_df is None or reasons_df.empty:
        return pd.DataFrame(columns=_COMMENT_COLUMNS)
    t = pd.DataFrame({
        "Group": pd.Series(groups).fillna("Unclassified").astype(str).values,
        "Reason": reasons_df["Reason"].astype(str).values,
        "Hours": pd.to_numeric(reasons_df["Duration_hours"], errors="coerce").fillna(0.0).values,
        "Events": pd.to_numeric(reasons_df["Count"], errors="coerce").fillna(0.0).values,
    })
    t["_group_total"] = t["Group"].map(t.groupby("Group")["Hours"].sum())
    t = t.sort_values(["_group_total", "Group", "Hours"], ascending=[False, True, False])
    return t.drop(columns="_group_total").reset_index(drop=True)


def themed_comments(reasons_df, theme_patterns):
    """build_comments_table() grouped by theme."""
    if reasons_df is None or reasons_df.empty:
        return build_comments_table(None, None)
    return build_comments_table(
        reasons_df, reasons_df["Reason"].apply(lambda r: classify_theme(r, theme_patterns)))


def _draw_comments_table(pdf, tbl, page_bottom_mm):
    """Draws a Group/Reason/Hours/Events table with wrapped reason text and
    a shaded, subtotalled row per group. Drawn cell by cell from measured
    text heights so a long reason wraps instead of being cut off."""
    widths = (135, 22, 23)  # Reason | Hours | Events = 180 mm
    line_h = 4.5
    x0 = pdf.l_margin

    def row(cells, bold=False, fill=None):
        pdf.set_font("Helvetica", "B" if bold else "", 8)
        lines = _wrap_text(pdf, cells[0], widths[0] - 2)
        h = line_h * len(lines) + 1
        if pdf.get_y() + h > page_bottom_mm:
            pdf.add_page()
            pdf.set_font("Helvetica", "B" if bold else "", 8)
        y0 = pdf.get_y()
        if fill:
            pdf.set_fill_color(*fill)
        x = x0
        for w in widths:
            pdf.rect(x, y0, w, h, style="DF" if fill else "D")
            x += w
        for i, text in enumerate(lines):
            pdf.set_xy(x0 + 1, y0 + 0.5 + i * line_h)
            pdf.cell(widths[0] - 2, line_h, text)
        pdf.set_xy(x0 + widths[0], y0 + 0.5)
        pdf.cell(widths[1] - 1, line_h, cells[1], align="R")
        pdf.set_xy(x0 + widths[0] + widths[1], y0 + 0.5)
        pdf.cell(widths[2] - 1, line_h, cells[2], align="R")
        pdf.set_xy(x0, y0 + h)

    row(["Reason", "Hours", "Events"], bold=True, fill=(225, 225, 225))
    for group, sub in tbl.groupby("Group", sort=False):
        if pdf.get_y() + 20 > page_bottom_mm:  # keep a group heading with its first rows
            pdf.add_page()
        row([_pdf_safe(group), f"{sub['Hours'].sum():.1f}", f"{int(round(sub['Events'].sum()))}"],
            bold=True, fill=(243, 225, 200))
        for _, r in sub.iterrows():
            row([_pdf_safe(r["Reason"]), f"{r['Hours']:.1f}", f"{int(round(r['Events']))}"])


def generate_pdf_report(line_name, key_metrics, report_figures, comments=None,
                        include_comments=False, sections=None):
    """Builds the PDF: a cover page (key metrics + contents), then the chart
    sections in the order they were produced, then - only if asked - a
    comments appendix.

      report_figures   list of (section, title, matplotlib figure)
      comments         list of (section, title, table) from build_comments_table()
      include_comments False -> graphs and figures only
      sections         section names to keep; None keeps every section

    Uses the BytesIO -> temp-file -> pdf.image() pattern because fpdf2
    embeds from a file path, not an in-memory buffer."""
    figures = [(sec, title, fig) for sec, title, fig in report_figures
               if fig is not None and (sections is None or sec in sections)]
    comment_tables = ([(sec, title, tbl) for sec, title, tbl in (comments or [])
                       if tbl is not None and len(tbl)] if include_comments else [])

    pdf = FPDF()
    pdf.set_auto_page_break(auto=True, margin=15)
    page_bottom_mm = pdf.h - 15
    text_w = pdf.w - pdf.l_margin - pdf.r_margin

    def line(text, size=10, bold=False, h=6):
        pdf.set_font("Helvetica", "B" if bold else "", size)
        for chunk in _wrap_text(pdf, _pdf_safe(text), text_w):
            pdf.cell(0, h, chunk)
            pdf.ln(h)

    # ---- cover ----
    pdf.add_page()
    line("Downtime Analysis Report", 18, True, 10)
    line(line_name, 12, False, 8)
    line("Generated " + datetime.now().strftime("%d %b %Y, %H:%M"), 9, False, 6)
    pdf.ln(4)
    line("Key Metrics", 12, True, 8)
    for label, value in key_metrics:
        line(f"{label}: {value}", 10)
    pdf.ln(4)
    line("Contents", 12, True, 8)
    section_counts = {}
    for sec, _, _ in figures:
        section_counts[sec] = section_counts.get(sec, 0) + 1
    for sec, n in section_counts.items():
        line(f"- {sec} ({n} chart{'s' if n != 1 else ''})", 10)
    if not section_counts:
        line("- No chart sections selected", 10)
    if include_comments and comment_tables:
        n_rows = sum(len(t) for _, _, t in comment_tables)
        line(f"- Comments appendix ({n_rows} rows in {len(comment_tables)} tables)", 10)
    elif include_comments:
        line("- Comments were requested, but no comment rows were found in this file", 10)
    else:
        pdf.ln(2)
        line("Comments are not included in this report.", 9)

    # ---- chart sections ----
    temp_files = []
    current_section = None
    try:
        for section, title, fig in figures:
            if section != current_section:
                pdf.add_page()
                line(section, 15, True, 10)
                pdf.ln(2)
                current_section = section

            buf = io.BytesIO()
            fig.savefig(buf, format="png", dpi=150, bbox_inches="tight")
            buf.seek(0)
            # A genuinely unique temp path - a name built from section+title
            # text would collide if two users generated a PDF at once.
            fd, tmp_path = tempfile.mkstemp(suffix=".png")
            os.close(fd)
            with open(tmp_path, "wb") as f:
                f.write(buf.read())
            temp_files.append(tmp_path)

            # Decide on the page break from the image's REAL scaled height, not
            # a flat guess - a tall multi-panel figure needs far more room
            # than a short wide one, and a flat threshold left the title and
            # image stranded on different pages.
            with Image.open(tmp_path) as im:
                img_w_px, img_h_px = im.size
            target_w_mm = 180
            scaled_h_mm = target_w_mm * (img_h_px / img_w_px)
            if pdf.get_y() + 10 + scaled_h_mm > page_bottom_mm:
                pdf.add_page()

            line(title, 11, True, 7)
            pdf.image(tmp_path, w=target_w_mm)
            pdf.ln(4)
    finally:
        for tf in temp_files:
            if os.path.exists(tf):
                os.remove(tf)

    # ---- comments appendix (only when asked for) ----
    current_section = None
    for section, title, tbl in comment_tables:
        if section != current_section:
            pdf.add_page()
            line(f"Comments: {section}", 15, True, 10)
            pdf.ln(2)
            current_section = section
        elif pdf.get_y() + 30 > page_bottom_mm:
            pdf.add_page()
        line(title, 11, True, 7)
        _draw_comments_table(pdf, tbl, page_bottom_mm)
        pdf.ln(4)

    try:
        return bytes(pdf.output())
    except TypeError:
        return pdf.output(dest="S").encode("latin-1")


def render_report_tab(line_name, key_metrics, report_figures, report_comments):
    """The Report Export tab, shared by every line: pick which chart sections
    go in, choose whether the comments are included, build the PDF. The built
    PDF is kept in session state so the download button survives the rerun
    that clicking it triggers (a button created inside `if st.button(...)`
    vanishes on that rerun)."""
    st.header("Generate PDF Report")
    figs = [(s, t, f) for s, t, f in report_figures if f is not None]
    section_counts = {}
    for s, _, _ in figs:
        section_counts[s] = section_counts.get(s, 0) + 1
    sections = list(section_counts)

    st.write(
        f"The report is built from the {len(figs)} charts shown across the tabs above, "
        f"grouped into {len(sections)} sections. Choose what goes in."
    )
    chosen = st.multiselect(
        "Sections to include (graphs)", options=sections, default=sections,
        format_func=lambda s: f"{s} ({section_counts[s]} chart{'s' if section_counts[s] != 1 else ''})",
    )
    comments_choice = st.radio(
        "Include comments (the reason text from the log) in the report?",
        ["Yes", "No"], index=1, horizontal=True,
        help="Comments are the reason entries logged against each unit. Yes adds an appendix "
             "listing every reason with its hours and event count, grouped by theme or by "
             "the unit it was attributed to. No keeps the report to graphs and figures only.",
    )
    include_comments = comments_choice == "Yes"
    n_comment_rows = sum(len(t) for _, _, t in report_comments if t is not None)
    if include_comments:
        if n_comment_rows:
            st.caption(f"{n_comment_rows} comment rows in {len(report_comments)} tables will be added as an appendix.")
        else:
            st.caption("No comment rows were found in this file, so there is nothing to append.")

    signature = (line_name, tuple(chosen), include_comments, str(key_metrics))
    if st.button("Generate PDF Report", type="primary", disabled=(not chosen and not include_comments)):
        with st.spinner("Building PDF..."):
            st.session_state["_pdf_report"] = {
                "signature": signature,
                "bytes": generate_pdf_report(line_name, key_metrics, report_figures,
                                             comments=report_comments,
                                             include_comments=include_comments, sections=chosen),
            }
    stored = st.session_state.get("_pdf_report")
    if stored and stored["signature"] == signature:
        st.download_button(
            "Download PDF Report", data=stored["bytes"],
            file_name=f"{line_name.replace(' ', '_')}_downtime_report.pdf", mime="application/pdf",
        )
    elif stored and stored["signature"][0] == line_name:
        st.caption("The options changed since the last PDF was built. Click Generate PDF Report again.")


@st.cache_data(show_spinner="Reading Excel file...")
def load_all_data(file_bytes, line_choice):
    """Every Excel read this app needs, in one cached call. Previously
    load_daily()/load_category_block() were called directly from inside
    render_line(), which meant every one of them re-ran on EVERY Streamlit
    interaction (checking a box, clicking a button - Streamlit reruns the
    whole script top to bottom on any of these). That's 8 separate full
    reads of the file for Barth, 11 for Buhler Line, repeated on every single
    interaction, which is almost certainly the actual cause of both the
    slow load and the slow PDF generation (the PDF button click itself
    triggers a full rerun before the PDF-specific work even starts).
    Caching this on (file_bytes, line_choice) means the file is genuinely
    read once per upload, not once per interaction.

    Presses has no single roaster - each press has its own reason log, so
    the fourth return value is a dict keyed by press name ("PR1"->df) for
    that line, instead of the single roaster df_cat every other line
    returns. main() checks cfg.get("kind") to know which shape it got back
    before choosing render_presses() vs render_line()."""
    cfg = LINES[line_choice]

    # Discovery runs FIRST, before any daily-column reads are attempted -
    # if the file doesn't match the selected line at all, stop here rather
    # than reading daily columns from the wrong positions (silently wrong
    # data) and then hitting the SAME class of pandas usecols-out-of-bounds
    # crash the category-block fallback used to hit further down. Each
    # line's real files use a distinct label prefix ("Barth: <unit>" vs
    # "Buhler: <unit>") - if none of the discovered labels match what the
    # selected line expects, every column position in that line's config
    # is very likely wrong for this file. Presses uses several prefixes
    # (PR1/PR2/PR3, one per press) rather than one, so this checks for
    # ANY overlap with the expected set, not an exact single match.
    discovered, discovery_meta = discover_category_blocks(file_bytes)
    discovered_prefixes = {label.split(":")[0].strip() for label in discovered.keys() if ":" in label}
    expected_prefixes = cfg.get("file_label_prefixes")
    if expected_prefixes is None:
        single = cfg.get("file_label_prefix")
        expected_prefixes = [single] if single is not None else []
    prefix_mismatch = (bool(discovered_prefixes) and bool(expected_prefixes)
                       and not (discovered_prefixes & set(expected_prefixes)))
    discovery_meta["prefix_mismatch"] = prefix_mismatch
    discovery_meta["discovered_prefixes"] = sorted(discovered_prefixes)
    discovery_meta["expected_prefix"] = expected_prefixes[0] if expected_prefixes else None
    discovery_meta["expected_prefixes"] = expected_prefixes

    if prefix_mismatch:
        return {}, {}, pd.DataFrame(), {}, {"__meta__": discovery_meta}

    dfs_raw, eq_titles = {}, {}
    for unit, c in cfg["daily_cols"].items():
        d, eq_title = load_daily(io.BytesIO(file_bytes), c["usecols"], c["skiprows"])
        dfs_raw[unit] = d
        eq_titles[unit] = eq_title

    discovery_status = {"__meta__": discovery_meta}

    if cfg.get("kind") == "presses":
        press_category_data = {}
        for press, roles in cfg["presses"].items():
            unit = roles["oee_source"]
            unit_block = resolve_unit_block(unit, discovered)
            if unit_block is not None:
                press_category_data[press] = load_discovered_block(file_bytes, unit_block)
                discovery_status[unit] = {"source": "discovered", "detail": unit_block}
            else:
                press_category_data[press] = load_category_block(
                    io.BytesIO(file_bytes), cfg["category_cols"][unit], cfg["category_skiprows"])
                discovery_status[unit] = {"source": "hardcoded fallback", "detail": cfg["category_cols"][unit]}
        return dfs_raw, eq_titles, press_category_data, {}, discovery_status

    if cfg.get("kind") == "moulding":
        unit_category_data = {}
        for unit_label, roles in cfg["units"].items():
            unit = roles["oee_source"]
            unit_block = resolve_unit_block(unit, discovered)
            if unit_block is not None:
                unit_category_data[unit_label] = load_discovered_block(file_bytes, unit_block)
                discovery_status[unit] = {"source": "discovered", "detail": unit_block}
            else:
                unit_category_data[unit_label] = load_category_block(
                    io.BytesIO(file_bytes), cfg["category_cols"][unit], cfg["category_skiprows"])
                discovery_status[unit] = {"source": "hardcoded fallback", "detail": cfg["category_cols"][unit]}
        return dfs_raw, eq_titles, unit_category_data, {}, discovery_status

    roaster_key = cfg["roaster"]
    roaster_block = resolve_unit_block(roaster_key, discovered)
    if roaster_block is not None:
        df_cat = load_discovered_block(file_bytes, roaster_block)
        discovery_status[roaster_key] = {"source": "discovered", "detail": roaster_block}
    else:
        roaster_cat_cols = cfg["roaster_category_cols"] if "roaster_category_cols" in cfg else cfg["category_cols"][roaster_key]
        df_cat = load_category_block(io.BytesIO(file_bytes), roaster_cat_cols, cfg["category_skiprows"])
        discovery_status[roaster_key] = {"source": "hardcoded fallback", "detail": roaster_cat_cols}

    own_category_data = {}
    if "own_log_units" in cfg:
        for unit in cfg["own_log_units"]:
            unit_block = resolve_unit_block(unit, discovered)
            if unit_block is not None:
                own_category_data[unit] = load_discovered_block(file_bytes, unit_block)
                discovery_status[unit] = {"source": "discovered", "detail": unit_block}
            else:
                own_category_data[unit] = load_category_block(
                    io.BytesIO(file_bytes), cfg["category_cols"][unit], cfg["category_skiprows"])
                discovery_status[unit] = {"source": "hardcoded fallback", "detail": cfg["category_cols"][unit]}

    return dfs_raw, eq_titles, df_cat, own_category_data, discovery_status


def _show_prefix_mismatch_error(line_name, meta_check):
    """Shows the wrong-file error and returns True when the uploaded file's
    own "<Line>: <Unit>" labels don't match the line selected in the sidebar
    (every column position for that line is then very likely wrong). Shared
    by every line's render function."""
    if not meta_check.get("prefix_mismatch"):
        return False
    found = ", ".join(f"'{p}:'" for p in meta_check.get("discovered_prefixes", []))
    expected_list = meta_check.get("expected_prefixes") or [meta_check.get("expected_prefix")]
    expected = " or ".join(f"'{p}:'" for p in expected_list)
    st.error(
        f"\u26a0\ufe0f **This file doesn't look like a {line_name} file.** "
        f"You've selected **{line_name}**, but this file's own labels use "
        f"{found} rather than the {expected} prefix {line_name} normally uses. "
        f"Nothing below has been read from this file, since every column position "
        f"for {line_name} is very likely wrong for it. If this is actually a "
        f"different line's file, switch **Select Line** in the sidebar and re-upload."
    )
    return True


def _render_discovery_sidebar(discovery_status):
    """Sidebar 'Block Discovery' panel: one row per unit saying whether its
    reason block was auto-discovered or fell back to the hardcoded config."""
    if discovery_status:
        unit_entries = {k: v for k, v in discovery_status.items() if k != "__meta__"}
        meta = discovery_status.get("__meta__", {})
        n_fallback = sum(1 for v in unit_entries.values() if v["source"] == "hardcoded fallback")
        st.sidebar.markdown('<div style="height:10px"></div>', unsafe_allow_html=True)
        st.sidebar.markdown(dc_card_head(_ICON_TARGET, "Block Discovery", "Reason data location check"), unsafe_allow_html=True)
        n_auto = len(unit_entries) - n_fallback
        st.sidebar.markdown(
            f'<div class="dc-card" style="padding:10px 14px;display:flex;justify-content:space-between;'
            f'align-items:center;margin-top:8px;"><span class="dc-sub" style="font-size:12.5px;">Auto-discovered</span>'
            f'<span style="font-weight:700;">{n_auto}/{len(unit_entries)} units</span></div>',
            unsafe_allow_html=True,
        )
        if meta.get("category_section_row") is None:
            st.sidebar.caption("⚠️ No 'By Category' marker found anywhere in this file - search fell "
                              "back to scanning the whole sheet.")
        with st.sidebar.container(height=280 if len(unit_entries) > 4 else "content", border=False):
            for unit, info in unit_entries.items():
                if info["source"] == "discovered":
                    d = info["detail"]
                    col_indices = [column_index_from_string(c) for c in d["columns"].values()]
                    lo, hi = get_column_letter(min(col_indices)), get_column_letter(max(col_indices))
                    pill = dc_pill(_ICON_BOLT, "Auto", "green")
                    loc = f"row {d['header_row']}, cols {lo}\u2013{hi}"
                else:
                    pill = dc_pill(_ICON_WARN, "Manual", "coral")
                    loc = "hardcoded fallback config"
                st.sidebar.markdown(
                    f'<div class="dc-unit-row"><div><div class="dc-unit-name">{unit}</div>'
                    f'<div class="dc-unit-loc">{_ICON_PIN}<span>{loc}</span></div></div>{pill}</div>',
                    unsafe_allow_html=True,
                )


def _render_month_exclusion(date_frames, exclude_march_default):
    """Sidebar month-exclusion control. Offers every month present in any of
    the given daily dataframes (not just a hardcoded March window). Returns
    (excluded 'YYYY-MM' list, {ym: 'Month YYYY'} labels)."""
    available_ym = sorted({ym for d in date_frames for ym in d["Date"].dt.strftime("%Y-%m").unique()})
    month_labels = {ym: pd.Timestamp(ym + "-01").strftime("%B %Y") for ym in available_ym}
    default_excl = ["2026-03"] if exclude_march_default and "2026-03" in available_ym else []

    st.sidebar.markdown('<div style="height:14px"></div>', unsafe_allow_html=True)
    st.sidebar.markdown(dc_card_head(_ICON_CALENDAR_X, "Exclude Month(s)", "Drop known bad-data windows"), unsafe_allow_html=True)
    excluded = st.sidebar.multiselect(
        "Months to exclude from analysis", options=available_ym,
        default=default_excl, format_func=lambda ym: month_labels[ym],
        label_visibility="collapsed",
    )
    return excluded, month_labels


def render_line(cfg, dfs_raw, eq_titles, df_cat_loaded, own_category_data_cached, line_name, discovery_status=None):
    roaster_key = cfg["roaster"]

    if discovery_status and _show_prefix_mismatch_error(line_name, discovery_status.get("__meta__", {})):
        return

    if dfs_raw[roaster_key].empty:
        st.error(f"No data loaded for the roaster ({roaster_key}) - check skiprows/usecols against the real file.")
        return

    if discovery_status:
        _render_discovery_sidebar(discovery_status)

    excluded, month_labels = _render_month_exclusion([dfs_raw[roaster_key]], cfg.get("exclude_march_default"))

    dfs = {unit: apply_month_exclusion(d, excluded) for unit, d in dfs_raw.items()}
    roaster = dfs[roaster_key].copy()

    # OEE_Percentage and line_oee computed HERE, once, on the real
    # post-exclusion roaster data - the header pill below and the
    # "Average OEE" metric card in the OEE tab both read this SAME value,
    # rather than each computing their own (confirmed real bug: the header
    # used to compute a preliminary figure from dfs_raw BEFORE month
    # exclusion was applied, so it silently included excluded months like
    # a default-excluded March contamination window - two numbers on the
    # same page both labeled "Avg OEE" that quietly meant different
    # things). The rest of the OEE tab's feature columns (SMA, EWMA, lag,
    # cumulative) are computed later, where they're actually used - only
    # what the header needs is pulled forward.
    roaster["OEE_Percentage"] = ((24 - roaster["Duration (hours)"]) / 24) * 100
    line_oee = roaster["OEE_Percentage"].mean()

    _show_logo_if_present()

    window_start = roaster["Date"].min()
    window_end = roaster["Date"].max()

    dc_header(
        "Downtime & OEE Intelligence", "Barry Callebaut Ghana · Cocoa Processing",
        dc_pill(_ICON_BRANCH, line_name, "orange")
        + dc_pill(_ICON_CALENDAR_X, f"{window_start.day} {window_start.strftime('%b %Y')} \u2013 "
                                     f"{window_end.day} {window_end.strftime('%b %Y')}", "gray")
        + dc_pill(_ICON_GAUGE, f"Avg OEE {line_oee:.1f}%", "green"),
    )

    st.markdown(
        f"**{len(roaster)} days analysed** after exclusions"
        + (f" (excluded: {', '.join(month_labels[ym] for ym in excluded)})" if excluded else "")
    )

    # ---- Figure collector for the PDF report ----
    report_figures = []  # (section, title, fig)
    report_comments = []  # (section, title, comments table) - only used if the user opts in

    def emit(section, title, fig):
        st.pyplot(fig)
        report_figures.append((section, title, fig))

    tab_names = ["OEE & Data Integrity", "Roaster Attribution", "Strategy & Improvement", "Theme Breakdown"]
    has_own_logs = "own_log_units" in cfg
    if has_own_logs:
        tab_names.append("Own Reason Logs")
    tab_names += ["Raw Logs", "Predictive Analytics", "Report Export"]
    tabs = st.tabs(tab_names)

    i = 0
    tab_overview = tabs[i]; i += 1
    tab_attribution = tabs[i]; i += 1
    tab_strategy = tabs[i]; i += 1
    tab_themes = tabs[i]; i += 1
    tab_ownlogs = None
    if has_own_logs:
        tab_ownlogs = tabs[i]; i += 1
    tab_raw = tabs[i]; i += 1
    tab_predictive = tabs[i]; i += 1
    tab_report = tabs[i]; i += 1

    # ================= TAB: OEE & Data Integrity =================
    with tab_overview:
        st.markdown(
            dc_card_head(_ICON_UPLOAD, "Load Daily Downtime Data",
                        "Real column headers checked against the config, not just trusted"),
            unsafe_allow_html=True,
        )
        st.caption(
            "eq_title is read straight from each column's own Excel header - an identity "
            "check that this confirms (or corrects) the unit names against the real "
            "workbook, not just the names typed into the config."
        )
        id_rows = []
        for unit, d in dfs.items():
            date_range = f"{d['Date'].min().date()} to {d['Date'].max().date()}" if len(d) else "N/A"
            id_rows.append({"Config name": unit, "Real name from Excel header": eq_titles[unit],
                            "Rows": len(d), "Date range": date_range,
                            "Total hours": round(d["Duration (hours)"].sum(), 1)})
        st.dataframe(pd.DataFrame(id_rows), use_container_width=True)
        st.markdown('<div style="height:8px"></div>', unsafe_allow_html=True)

        st.markdown(
            dc_card_head(_ICON_SCALE, "Data Integrity Check",
                        "Do the per-unit numbers reconcile with real, independent measurement?"),
            unsafe_allow_html=True,
        )
        st.caption(
            "Do the other units' daily downtime sum to anything like the roaster's own, "
            "and how tightly does each track it day to day? High correlation despite a "
            "large sum-vs-roaster gap points to the same real stoppages being logged "
            "redundantly at multiple stations, not independent per-unit measurements."
        )
        others = [u for u in cfg["daily_cols"] if u != roaster_key]
        merged = roaster[["Date", "Duration (hours)"]].rename(columns={"Duration (hours)": roaster_key})
        for u in others:
            merged = merged.merge(
                dfs[u][["Date", "Duration (hours)"]].rename(columns={"Duration (hours)": u}),
                on="Date", how="inner")
        merged["Sum_of_units"] = merged[others].sum(axis=1)

        ratio = merged["Sum_of_units"].mean() / merged[roaster_key].mean() if merged[roaster_key].mean() else float("nan")
        c1, c2, c3 = st.columns(3)
        with c1:
            st.markdown(
                f'<div class="dc-card"><span class="dc-metric-label">Mean sum across units</span>'
                f'<div class="dc-metric-value" style="font-size:24px;">{merged["Sum_of_units"].mean():.2f} hrs/day</div></div>',
                unsafe_allow_html=True)
        with c2:
            st.markdown(
                f'<div class="dc-card"><span class="dc-metric-label">Mean roaster downtime</span>'
                f'<div class="dc-metric-value" style="font-size:24px;">{merged[roaster_key].mean():.2f} hrs/day</div></div>',
                unsafe_allow_html=True)
        with c3:
            st.markdown(
                f'<div class="dc-card"><span class="dc-metric-label">Ratio</span>'
                f'<div class="dc-metric-value" style="font-size:24px;">{ratio:.2f}x</div></div>',
                unsafe_allow_html=True)
        st.markdown('<div style="height:10px"></div>', unsafe_allow_html=True)

        corr_rows = [{"Unit": u, "Corr. with Roaster": round(merged[u].corr(merged[roaster_key]), 3)} for u in others]
        corr_df = pd.DataFrame(corr_rows).sort_values("Corr. with Roaster", ascending=False)
        fig, ax = plt.subplots(figsize=(9, max(3, 0.4 * len(corr_df))))
        bars = ax.barh(corr_df["Unit"], corr_df["Corr. with Roaster"], color=_DC_ORANGE)
        ax.bar_label(bars, fmt="%.3f", padding=3)
        ax.set_xlim(0, 1.15)
        ax.set_xlabel("Pearson r vs roaster daily downtime")
        ax.set_title("Each Unit vs the Roaster", fontsize=13, fontweight="bold")
        ax.invert_yaxis()
        plt.tight_layout()
        emit("OEE & Data Integrity", "Each Unit vs the Roaster (correlation)", fig)

        st.markdown(
            dc_card_head(_ICON_GAUGE, "Line-Level OEE", f"Measured at {roaster_key}"),
            unsafe_allow_html=True,
        )
        # OEE_Percentage and line_oee already computed above (shared with
        # the header pill) - only the extra feature columns are new here.
        roaster["OEE_7Day_SMA"] = roaster["OEE_Percentage"].rolling(7, min_periods=1).mean()
        roaster["OEE_7Day_EWMA"] = roaster["OEE_Percentage"].ewm(span=7, adjust=False).mean()
        roaster["Downtime_Lag_1"] = roaster["Duration (hours)"].shift(1)
        roaster["Cumulative_Downtime"] = roaster["Duration (hours)"].cumsum()

        best_row = roaster.loc[roaster["OEE_Percentage"].idxmax()]
        worst_row = roaster.loc[roaster["OEE_Percentage"].idxmin()]
        total_dt_minutes = roaster["Duration (hours)"].sum() * 60
        dt_h, dt_m = divmod(int(round(total_dt_minutes)), 60)
        n_events = int((roaster["Duration (hours)"] > 0).sum())

        c1, c2, c3, c4 = st.columns(4)
        with c1:
            st.markdown(dc_metric_card("Average OEE", f"{line_oee:.1f}%", "filtered window", _ICON_GAUGE, "orange"), unsafe_allow_html=True)
        with c2:
            st.markdown(dc_metric_card("Best Day", f"{best_row['OEE_Percentage']:.1f}%", f"{best_row['Date'].day} {best_row['Date'].strftime('%b')}", _ICON_UP, "green"), unsafe_allow_html=True)
        with c3:
            st.markdown(dc_metric_card("Worst Day", f"{worst_row['OEE_Percentage']:.1f}%", f"{worst_row['Date'].day} {worst_row['Date'].strftime('%b')}", _ICON_DOWN, "coral"), unsafe_allow_html=True)
        with c4:
            st.markdown(dc_metric_card("Logged Downtime", f"{dt_h}h {dt_m:02d}m", f"{n_events} events", _ICON_CLOCK, "orange"), unsafe_allow_html=True)
        st.markdown('<div style="height:14px"></div>', unsafe_allow_html=True)

        pct_mean = roaster["OEE_Percentage"].mean()
        pct_std = roaster["OEE_Percentage"].std()
        pct_ucl = min(pct_mean + 3 * pct_std, 100)
        pct_lcl = max(pct_mean - 3 * pct_std, 0)
        pct_flagged = int(((roaster["OEE_Percentage"] > pct_ucl) | (roaster["OEE_Percentage"] < pct_lcl)).sum())

        sign_minus = "\u2212"
        sign_dots = "\u22ef"

        st.markdown(
            dc_card_head(_ICON_LINE_CHART, "Daily OEE with statistical control limits",
                        "Control chart flags days beyond \u00b13\u03c3 as statistically unusual")
            + f'<div style="margin:10px 0 6px 0;">{dc_pill(sign_minus, f"Mean {pct_mean:.1f}%", "gray")}'
            f'{dc_pill(sign_dots, f"UCL {pct_ucl:.1f}%", "gray")}{dc_pill(sign_dots, f"LCL {pct_lcl:.1f}%", "gray")}'
            f'{dc_pill(_ICON_ALARM, f"{pct_flagged} flagged", "coral" if pct_flagged else "green")}</div>',
            unsafe_allow_html=True,
        )

        fig, ax = plt.subplots(figsize=(13, 6))
        ax.plot(roaster["Date"], roaster["OEE_Percentage"], label="Daily OEE (raw)", alpha=0.3, color="grey", marker=".")
        ax.plot(roaster["Date"], roaster["OEE_7Day_SMA"], label="7-day SMA", color=_DC_ORANGE, linewidth=2, linestyle="--")
        ax.plot(roaster["Date"], roaster["OEE_7Day_EWMA"], label="7-day EWMA", color=_DC_CORAL, linewidth=2)
        ax.axhline(line_oee, color=_DC_GREEN, linestyle=":", linewidth=2, label=f"Mean ({line_oee:.1f}%)")
        ax.axhline(pct_ucl, color=_DC_CORAL, linestyle="--", linewidth=1.2, alpha=0.7, label=f"UCL ({pct_ucl:.1f}%)")
        ax.axhline(pct_lcl, color=_DC_CORAL, linestyle="--", linewidth=1.2, alpha=0.7, label=f"LCL ({pct_lcl:.1f}%)")
        ax.set_title("OEE Trend", fontsize=14, fontweight="bold")
        ax.set_ylim(0, 105)
        ax.legend()
        ax.grid(alpha=0.3)
        plt.tight_layout()
        emit("OEE & Data Integrity", "OEE Trend", fig)

        mean_dt = roaster["Duration (hours)"].mean()
        ucl = min(mean_dt + 3 * roaster["Duration (hours)"].std(), 24)
        out = roaster[roaster["Duration (hours)"] >= ucl]
        fig, ax = plt.subplots(figsize=(13, 5.5))
        ax.plot(roaster["Date"], roaster["Duration (hours)"], marker="o", color="grey", alpha=0.6, markersize=4, label="Daily downtime")
        ax.axhline(mean_dt, color=_DC_GREEN, linestyle="--", linewidth=2, label=f"Mean ({mean_dt:.2f} hrs)")
        ax.axhline(ucl, color=_DC_CORAL, linewidth=2, label=f"UCL +3\u03c3 ({ucl:.2f} hrs)")
        ax.scatter(out["Date"], out["Duration (hours)"], color=_DC_CORAL, s=90, zorder=5, label=f"Special cause ({len(out)} days)")
        ax.set_title("SPC Control Chart (downtime hours)", fontsize=13, fontweight="bold")
        ax.legend()
        ax.grid(alpha=0.3)
        plt.tight_layout()
        emit("OEE & Data Integrity", "SPC Control Chart", fig)
        st.caption(f"{len(out)} special-cause days ({len(out)/len(roaster)*100:.1f}% of days)")

        fig, axes = plt.subplots(3, 1, figsize=(13, 16), gridspec_kw={"height_ratios": [1, 2, 1]})
        m = roaster.groupby("Month_Name")["Duration (hours)"].sum().reset_index()
        sns.barplot(data=m, x="Month_Name", y="Duration (hours)", order=[mo for mo in MONTH_ORDER if mo in m["Month_Name"].values],
                    ax=axes[0], color=_DC_ORANGE)
        axes[0].set_title("Total Downtime per Month")
        plt.setp(axes[0].xaxis.get_majorticklabels(), rotation=45, ha="right")
        w = roaster.groupby("Week_of_Year")["Duration (hours)"].sum().reset_index()
        sns.barplot(data=w, x="Duration (hours)", y="Week_of_Year", ax=axes[1], color=_DC_GREEN)
        axes[1].set_title("Total Downtime per Week of Year")
        dw = roaster.groupby("Day_of_Week")["Duration (hours)"].sum().reset_index()
        sns.barplot(data=dw, x="Day_of_Week", y="Duration (hours)", order=[d for d in DAY_ORDER if d in dw["Day_of_Week"].values],
                    ax=axes[2], color=_DC_CORAL)
        axes[2].set_title("Total Downtime per Day of Week")
        plt.setp(axes[2].xaxis.get_majorticklabels(), rotation=45, ha="right")
        plt.tight_layout()
        emit("OEE & Data Integrity", "Downtime Distribution (Month/Week/Day)", fig)

        fig, ax = plt.subplots(figsize=(11, 5.5))
        sns.boxplot(data=roaster, x="Month_Name", y="Duration (hours)",
                    order=[mo for mo in MONTH_ORDER if mo in roaster["Month_Name"].values], color="#F3E1C8", ax=ax)
        ax.set_title("Daily Downtime Distribution per Month", fontsize=13, fontweight="bold")
        plt.setp(ax.xaxis.get_majorticklabels(), rotation=45, ha="right")
        plt.tight_layout()
        emit("OEE & Data Integrity", "Boxplot by Month", fig)

        r_val = roaster["Downtime_Lag_1"].corr(roaster["Duration (hours)"])
        fig, ax = plt.subplots(figsize=(7, 7))
        sns.regplot(data=roaster, x="Downtime_Lag_1", y="Duration (hours)",
                    scatter_kws={"alpha": 0.5, "color": "steelblue"}, line_kws={"color": "red"}, ax=ax)
        ax.annotate(f"r = {r_val:.3f}\nr\u00b2 = {r_val**2:.3f}", xy=(0.03, 0.97), xycoords="axes fraction",
                    ha="left", va="top", fontsize=11, bbox=dict(boxstyle="round,pad=0.5", facecolor="white", alpha=0.9))
        ax.set_title("1-Day Autocorrelation", fontsize=13, fontweight="bold")
        ax.set_xlabel("Yesterday's Downtime (hours)")
        ax.set_ylabel("Today's Downtime (hours)")
        plt.tight_layout()
        emit("OEE & Data Integrity", "1-Day Autocorrelation", fig)

        fig, ax = plt.subplots(figsize=(11, 5))
        ax.plot(roaster["Date"], roaster["Cumulative_Downtime"], color=_DC_ORANGE, linewidth=2)
        ax.fill_between(roaster["Date"], roaster["Cumulative_Downtime"], color=_DC_ORANGE, alpha=0.1)
        ax.set_title("Cumulative Downtime Over the Period", fontsize=13, fontweight="bold")
        plt.tight_layout()
        emit("OEE & Data Integrity", "Cumulative Downtime", fig)

    # ================= TAB: Roaster Attribution =================
    with tab_attribution:
        st.header("Attributing the Roaster's Downtime")
        df_cat = df_cat_loaded.copy()
        st.write(f"{len(df_cat)} roaster reason rows loaded ({df_cat['Duration_hours'].sum():.1f} hrs total)")
        with st.expander("Preview raw reason rows"):
            st.dataframe(df_cat.head(10), use_container_width=True)

        if df_cat.empty:
            st.warning(
                "⚠️ No reason rows were found in the roaster's category block for this file. "
                "Every table and chart in the tabs below that depend on attribution "
                "(Strategy & Improvement, Theme Breakdown, Predictive Analytics, Weibull) "
                "will be empty, not broken - there's simply nothing to attribute yet. "
                "Check the '📍 Category block discovery' panel in the sidebar to see "
                "whether the roaster's block was actually found, or check the file itself."
            )

        df_cat["Allocated_To"] = df_cat.apply(
            lambda row: classify_allocation(row, cfg["equipment_patterns"], cfg["roaster_own_keywords"]), axis=1)
        alloc = build_allocation_summary(df_cat)
        st.dataframe(alloc.round(2), use_container_width=True)

        fig, axes = plt.subplots(1, 2, figsize=(15, 6))
        colors = ["#9aa5ad" if a in ("Unattributable", "Unknown") else "#d64545" if a == "Roaster (own fault)" else "#4a7a96"
                  for a in alloc["Allocated_To"]]
        b1 = axes[0].barh(alloc["Allocated_To"], alloc["Total_Duration_hours"], color=colors)
        axes[0].bar_label(b1, fmt="%.1f hrs", padding=3, fontsize=8)
        axes[0].set_xlim(0, max(alloc["Total_Duration_hours"].max(), 0.1) * 1.15)  # room for the label text itself
        axes[0].set_title("Downtime by Attributed Source", fontsize=12, fontweight="bold")
        axes[0].invert_yaxis()
        known = alloc[~alloc["Allocated_To"].isin(["Unattributable", "Unknown"])]
        if len(known):
            b2 = axes[1].barh(known["Allocated_To"], known["Avg_Mins_per_Event"], color="#e08a3c")
            axes[1].bar_label(b2, fmt="%.1f min", padding=3, fontsize=8)
            axes[1].set_xlim(0, max(known["Avg_Mins_per_Event"].max(), 0.1) * 1.15)  # room for the label text itself
        axes[1].set_title("Average Duration per Event (attributed only)", fontsize=12, fontweight="bold")
        axes[1].invert_yaxis()
        plt.tight_layout()
        emit("Roaster Attribution", "Downtime by Attributed Source", fig)

    # ================= TAB: Strategy & Improvement =================
    with tab_strategy:
        st.header("Maintenance Strategy Quadrant")
        q = alloc[~alloc["Allocated_To"].isin(["Unattributable", "Unknown"])].copy()
        q = q[(q["Total_Count"] > 0) & (q["Avg_Mins_per_Event"] > 0)].reset_index(drop=True)
        if len(q) >= 2:
            fig, q_labeled = collision_aware_quadrant(q, "Maintenance Strategy by Attributed Source")
            emit("Strategy & Improvement", "Maintenance Strategy Quadrant", fig)
            st.dataframe(q_labeled.round(1), use_container_width=True)
        else:
            st.info("Fewer than 2 attributed sources with valid data - not enough to plot a quadrant.")

        st.header("Improvement Potential")
        total_dt = roaster["Duration (hours)"].sum()
        total_time = len(roaster) * 24
        rows = []
        for _, r in alloc.iterrows():
            hrs = r["Total_Duration_hours"]
            new_oee = ((total_time - max(total_dt - hrs, 0)) / total_time) * 100
            rows.append({"Source": r["Allocated_To"], "Downtime (hrs)": round(hrs, 1),
                         "Line OEE if eliminated (%)": round(new_oee, 2), "OEE gain (pp)": round(new_oee - line_oee, 2)})
        if not rows:
            st.info("No attributed sources to show - see the note above in Roaster Attribution.")
        else:
            potential = pd.DataFrame(rows).sort_values("OEE gain (pp)", ascending=False).reset_index(drop=True)
            st.write(f"Current Line OEE: {line_oee:.2f}% | total downtime {total_dt:.1f} hrs over {len(roaster)} days")
            st.dataframe(potential, use_container_width=True)
            pl = potential[~potential["Source"].isin(["Unattributable", "Unknown"])]
            if len(pl):
                fig, ax = plt.subplots(figsize=(10, 5))
                bars = ax.barh(pl["Source"], pl["OEE gain (pp)"], color="#3f8f6d")
                ax.bar_label(bars, fmt="+%.2f pp", padding=3)
                ax.set_xlim(0, max(pl["OEE gain (pp)"].max(), 0.1) * 1.15)  # room for the label text itself
                ax.set_title("Potential OEE Gain if Each Source Were Fully Eliminated", fontsize=12, fontweight="bold")
                ax.invert_yaxis()
                plt.tight_layout()
                emit("Strategy & Improvement", "Potential OEE Gain by Source", fig)

    # ================= TAB: Theme Breakdown =================
    with tab_themes:
        st.header("Fault-Type Theme Grouping")
        st.caption(
            "Explicit keyword-based themes rather than clustering - too few distinct "
            "reasons per unit for clustering to have real statistical footing, and an "
            "explicit library lets every match be audited."
        )
        for label, patterns, color in cfg["theme_targets"]:
            source = df_cat[df_cat["Allocated_To"] == label]
            summary, fig = render_theme_section(label, source, patterns, color)
            if summary is not None:
                report_figures.append(("Theme Breakdown", f"{label}: Theme Breakdown", fig))
        report_comments.append((
            "Roaster Comments", f"All {roaster_key} comments, grouped by the unit they were attributed to",
            build_comments_table(df_cat, df_cat["Allocated_To"]),
        ))

    # ================= TAB: Own Reason Logs (Buhler Line only) =================
    if tab_ownlogs is not None:
        with tab_ownlogs:
            st.header("Units With Their Own Reason Logs")
            st.caption(
                "The roaster's comments never name these units directly (confirmed: zero "
                "attribution above), so their downtime picture comes from their own "
                "\"By Category\" block instead."
            )
            theme_lookup = {label: (patterns, color) for label, patterns, color in cfg["theme_targets"]}
            for unit in cfg["own_log_units"]:
                df_own = own_category_data_cached.get(unit)
                st.subheader(unit)
                st.write(f"{len(df_own)} own reason rows loaded ({df_own['Duration_hours'].sum():.1f} hrs total)")
                patterns, color = theme_lookup.get(unit, (ROASTER_THEME_PATTERNS, "#4a7a96"))
                summary, fig = render_theme_section(f"{unit} (own log)", df_own, patterns, color)
                if summary is not None:
                    report_figures.append(("Own Reason Logs", f"{unit}: Theme Breakdown", fig))
                report_comments.append(("Own Reason Logs", f"{unit}: comments grouped by theme",
                                        themed_comments(df_own, patterns)))

    # ================= TAB: Raw Logs =================
    with tab_raw:
        st.header("Per-Unit Raw Logs (reference only)")
        st.caption("Not independent measurements - see the Data Integrity Check tab above.")
        raw_rows = []
        for unit in cfg["sequence"]:
            d = dfs.get(unit)
            if d is None or d.empty:
                continue
            raw_rows.append({"Unit": unit, "Total downtime (hrs)": round(d["Duration (hours)"].sum(), 1),
                             "Mean daily (hrs)": round(d["Duration (hours)"].mean(), 2), "Days": len(d)})
        raw_df = pd.DataFrame(raw_rows)
        st.dataframe(raw_df, use_container_width=True)
        fig, ax = plt.subplots(figsize=(12, 5.5))
        cols = ["#d64545" if u == roaster_key else "#4a7a96" for u in raw_df["Unit"]]
        bars = ax.bar(raw_df["Unit"], raw_df["Total downtime (hrs)"], color=cols)
        ax.bar_label(bars, fmt="%.0f", padding=3, fontsize=8)
        ax.set_title("Self-Reported Downtime per Unit, in Process Order\n(roaster in red)", fontsize=12, fontweight="bold")
        plt.setp(ax.xaxis.get_majorticklabels(), rotation=30, ha="right")
        plt.tight_layout()
        emit("Raw Logs", "Self-Reported Downtime per Unit", fig)

    # ================= TAB: Predictive Analytics =================
    with tab_predictive:
        st.header("Predictive Analytics")
        timeframe_days = len(roaster)
        st.caption(f"Historical timeframe used for scaling: {timeframe_days} days (from the roaster's own daily data).")

        st.subheader("Monte Carlo: 30-Day Line OEE Forecast (bootstrap on real daily data)")
        sim_oee = bootstrap_oee_simulation(roaster["Duration (hours)"])
        p10, p50, p90 = np.percentile(sim_oee, [10, 50, 90])
        c1, c2, c3 = st.columns(3)
        c1.metric("P10 (worst)", f"{p10:.1f}%")
        c2.metric("P50 (typical)", f"{p50:.1f}%")
        c3.metric("P90 (best)", f"{p90:.1f}%")
        fig, ax = plt.subplots(figsize=(10, 5.5))
        ax.hist(sim_oee, bins=60, color=_DC_ORANGE, edgecolor="white", alpha=0.85)
        ax.axvline(p50, color="green", linestyle="--", linewidth=2, label=f"P50 ({p50:.1f}%)")
        ax.axvline(p10, color="red", linestyle=":", linewidth=2, label=f"P10 ({p10:.1f}%)")
        ax.axvline(p90, color="red", linestyle=":", linewidth=2, label=f"P90 ({p90:.1f}%)")
        ax.set_title("Simulated 30-Day Average OEE Distribution", fontsize=13, fontweight="bold")
        ax.legend()
        plt.tight_layout()
        emit("Predictive Analytics", "Monte Carlo: 30-Day OEE Forecast", fig)

        st.subheader("Monte Carlo: Downtime Risk by Theme (Roaster's Own Faults)")
        roaster_theme_summary, _ = theme_breakdown_data(df_cat[df_cat["Allocated_To"] == "Roaster (own fault)"], ROASTER_THEME_PATTERNS)
        if not roaster_theme_summary.empty:
            rng = np.random.default_rng(42)
            risk_rows = []
            for _, row in roaster_theme_summary.iterrows():
                sim = simulate_category_downtime(row["Total_Duration_hours"], row["Total_Count"], timeframe_days, rng=rng)
                risk_rows.append({"Theme": row["Theme"], "Historical (hrs)": row["Total_Duration_hours"],
                                  "Sim P10 (30d)": np.percentile(sim, 10), "Sim P50 (30d)": np.percentile(sim, 50),
                                  "Sim P90 (30d)": np.percentile(sim, 90)})
            risk_df = pd.DataFrame(risk_rows).sort_values("Sim P90 (30d)", ascending=False)
            st.dataframe(risk_df.round(2), use_container_width=True)
            fig, ax = plt.subplots(figsize=(10, max(3, 0.5 * len(risk_df))))
            y = np.arange(len(risk_df))
            ax.barh(y, risk_df["Sim P90 (30d)"], color="#e8a0a0", label="P90")
            ax.barh(y, risk_df["Sim P50 (30d)"], color="#d64545", label="P50")
            ax.set_yticks(y)
            ax.set_yticklabels(risk_df["Theme"], fontsize=9)
            ax.set_xlabel("Simulated hours over 30 days")
            ax.legend()
            ax.invert_yaxis()
            plt.tight_layout()
            emit("Predictive Analytics", "Monte Carlo: Roaster Theme Risk", fig)

        st.subheader("Weibull Reliability - Top 3 Problem Sources")
        weibull_df = compute_weibull(alloc[~alloc["Allocated_To"].isin(["Unattributable", "Unknown"])], timeframe_days)
        if not weibull_df.empty:
            st.dataframe(weibull_df.round(2), use_container_width=True)
            fig = plot_weibull(weibull_df)
            emit("Predictive Analytics", "Weibull Reliability Curves", fig)
        else:
            st.info("Not enough attributed data to compute Weibull parameters.")

    # ================= TAB: Report Export =================
    with tab_report:
        key_metrics = [
            ("Line", line_name),
            ("Days analysed", len(roaster)),
            ("Months excluded", ", ".join(month_labels[ym] for ym in excluded) if excluded else "None"),
            ("Average Line OEE", f"{line_oee:.2f}%"),
            ("Total downtime", f"{roaster['Duration (hours)'].sum():.1f} hrs"),
        ]
        render_report_tab(line_name, key_metrics, report_figures, report_comments)


def render_presses(cfg, dfs_raw, eq_titles, press_category_data, line_name, discovery_status=None):
    """Presses is a different shape from Barth/Buhler: three independent
    press-lines (PR1/PR2/PR3), each measured at its own 'PRx: PRx' unit
    with its own 'By Category' reason log - no single roaster whose
    comments get reattributed elsewhere. So there's no Attribution or
    Strategy quadrant here (that mechanism doesn't exist for this line);
    everything below loops over the three presses instead."""
    presses = cfg["presses"]
    first_press, first_roles = next(iter(presses.items()))
    ref_unit = first_roles["oee_source"]

    if discovery_status and _show_prefix_mismatch_error(line_name, discovery_status.get("__meta__", {})):
        return

    if dfs_raw[ref_unit].empty:
        st.error(f"No data loaded for {ref_unit} - check skiprows/usecols against the real file.")
        return

    if discovery_status:
        _render_discovery_sidebar(discovery_status)

    excluded, month_labels = _render_month_exclusion(
        [dfs_raw[roles["oee_source"]] for roles in presses.values()], cfg.get("exclude_march_default"))

    dfs = {unit: apply_month_exclusion(d, excluded) for unit, d in dfs_raw.items()}

    # Per-press OEE, computed once up front (shared by the header pill and
    # the OEE tab below) - same reasoning as render_line()'s single-roaster
    # version: one number computed once, never two competing calculations.
    press_data, press_oee = {}, {}
    for press, roles in presses.items():
        d = dfs[roles["oee_source"]].copy()
        d["OEE_Percentage"] = ((24 - d["Duration (hours)"]) / 24) * 100
        press_data[press] = d
        press_oee[press] = d["OEE_Percentage"].mean()
    line_oee = float(np.mean(list(press_oee.values())))
    window_start = press_data[first_press]["Date"].min()
    window_end = press_data[first_press]["Date"].max()

    _show_logo_if_present()

    # No line-level "mean of N presses" OEE pill here on purpose: with each
    # press's own "Average OEE" card visible right below (the first press's
    # expander is open by default), a second, differently-scoped "Average
    # OEE" figure in the header read as a contradiction, not a summary -
    # confirmed confusing in practice, not just in theory. line_oee is
    # still computed and still shown in the Report Export tab, where it's
    # the only OEE figure on the page and the "(mean of presses)" label
    # has no competing per-press card next to it to conflict with.
    dc_header(
        "Downtime & OEE Intelligence", "Barry Callebaut Ghana · Cocoa Processing",
        dc_pill(_ICON_BRANCH, line_name, "orange")
        + dc_pill(_ICON_CALENDAR_X, f"{window_start.day} {window_start.strftime('%b %Y')} \u2013 "
                                     f"{window_end.day} {window_end.strftime('%b %Y')}", "gray"),
    )
    st.markdown(
        f"**{len(press_data[first_press])} days analysed** after exclusions"
        + (f" (excluded: {', '.join(month_labels[ym] for ym in excluded)})" if excluded else "")
    )

    report_figures = []
    report_comments = []

    def emit(section, title, fig):
        st.pyplot(fig)
        report_figures.append((section, title, fig))

    tab_names = ["OEE & Data Integrity", "Own Reason Logs", "Raw Logs", "Predictive Analytics", "Report Export"]
    tab_overview, tab_ownlogs, tab_raw, tab_predictive, tab_report = st.tabs(tab_names)

    # ================= TAB: OEE & Data Integrity =================
    with tab_overview:
        st.markdown(
            dc_card_head(_ICON_UPLOAD, "Load Daily Downtime Data",
                        "Real column headers checked against the config, not just trusted"),
            unsafe_allow_html=True,
        )
        id_rows = []
        for unit, d in dfs.items():
            date_range = f"{d['Date'].min().date()} to {d['Date'].max().date()}" if len(d) else "N/A"
            id_rows.append({"Config name": unit, "Real name from Excel header": eq_titles[unit],
                            "Rows": len(d), "Date range": date_range,
                            "Total hours": round(d["Duration (hours)"].sum(), 1)})
        st.dataframe(pd.DataFrame(id_rows), use_container_width=True)
        st.markdown('<div style="height:8px"></div>', unsafe_allow_html=True)

        for i, (press, roles) in enumerate(presses.items()):
            oee_unit, lcs_unit, color = roles["oee_source"], roles["paired_lcs"], roles["color"]
            d = press_data[press]
            with st.expander(f"{press} ({oee_unit})", expanded=(i == 0)):
                st.markdown(
                    dc_card_head(_ICON_SCALE, "Data Integrity Check", f"Does {lcs_unit} track {oee_unit}?"),
                    unsafe_allow_html=True,
                )
                merged = d[["Date", "Duration (hours)"]].rename(columns={"Duration (hours)": oee_unit}).merge(
                    dfs[lcs_unit][["Date", "Duration (hours)"]].rename(columns={"Duration (hours)": lcs_unit}),
                    on="Date", how="inner")
                r = merged[lcs_unit].corr(merged[oee_unit])
                c1, c2, c3 = st.columns(3)
                c1.markdown(f'<div class="dc-card"><span class="dc-metric-label">Mean {lcs_unit}</span>'
                           f'<div class="dc-metric-value" style="font-size:22px;">{merged[lcs_unit].mean():.2f} hrs/day</div></div>',
                           unsafe_allow_html=True)
                c2.markdown(f'<div class="dc-card"><span class="dc-metric-label">Mean {oee_unit}</span>'
                           f'<div class="dc-metric-value" style="font-size:22px;">{merged[oee_unit].mean():.2f} hrs/day</div></div>',
                           unsafe_allow_html=True)
                c3.markdown(f'<div class="dc-card"><span class="dc-metric-label">Correlation (r)</span>'
                           f'<div class="dc-metric-value" style="font-size:22px;">{r:.3f}</div></div>',
                           unsafe_allow_html=True)
                st.markdown('<div style="height:10px"></div>', unsafe_allow_html=True)

                d["OEE_7Day_SMA"] = d["OEE_Percentage"].rolling(7, min_periods=1).mean()
                d["OEE_7Day_EWMA"] = d["OEE_Percentage"].ewm(span=7, adjust=False).mean()
                d["Downtime_Lag_1"] = d["Duration (hours)"].shift(1)
                d["Cumulative_Downtime"] = d["Duration (hours)"].cumsum()
                best_row = d.loc[d["OEE_Percentage"].idxmax()]
                worst_row = d.loc[d["OEE_Percentage"].idxmin()]

                m1, m2, m3 = st.columns(3)
                m1.markdown(dc_metric_card("Average OEE", f"{press_oee[press]:.1f}%", "filtered window", _ICON_GAUGE, "orange"), unsafe_allow_html=True)
                m2.markdown(dc_metric_card("Best Day", f"{best_row['OEE_Percentage']:.1f}%", f"{best_row['Date'].day} {best_row['Date'].strftime('%b')}", _ICON_UP, "green"), unsafe_allow_html=True)
                m3.markdown(dc_metric_card("Worst Day", f"{worst_row['OEE_Percentage']:.1f}%", f"{worst_row['Date'].day} {worst_row['Date'].strftime('%b')}", _ICON_DOWN, "coral"), unsafe_allow_html=True)
                st.markdown('<div style="height:10px"></div>', unsafe_allow_html=True)

                fig, ax = plt.subplots(figsize=(12, 5))
                ax.plot(d["Date"], d["OEE_Percentage"], label="Daily OEE (raw)", alpha=0.3, color="grey", marker=".")
                ax.plot(d["Date"], d["OEE_7Day_SMA"], label="7-day SMA", color="#e0a458", linewidth=2, linestyle="--")
                ax.plot(d["Date"], d["OEE_7Day_EWMA"], label="7-day EWMA", color=color, linewidth=2)
                ax.axhline(press_oee[press], color=color, linestyle=":", linewidth=2, alpha=0.7, label=f"Mean ({press_oee[press]:.1f}%)")
                ax.set_title(f"{press}: OEE Trend", fontsize=13, fontweight="bold")
                ax.set_ylim(0, 105); ax.legend(); ax.grid(alpha=0.3)
                plt.tight_layout()
                emit("OEE & Data Integrity", f"{press}: OEE Trend", fig)

                mean_dt = d["Duration (hours)"].mean()
                ucl = min(mean_dt + 3 * d["Duration (hours)"].std(), 24)
                out = d[d["Duration (hours)"] >= ucl]
                fig, ax = plt.subplots(figsize=(12, 4.5))
                ax.plot(d["Date"], d["Duration (hours)"], marker="o", color="grey", alpha=0.6, markersize=4, label="Daily downtime")
                ax.axhline(mean_dt, color="#e0a458", linestyle="--", linewidth=2, label=f"Mean ({mean_dt:.2f} hrs)")
                ax.axhline(ucl, color=color, linewidth=2, label=f"UCL +3\u03c3 ({ucl:.2f} hrs)")
                ax.scatter(out["Date"], out["Duration (hours)"], color=color, s=80, zorder=5, label=f"Special cause ({len(out)} days)")
                ax.set_title(f"{press}: SPC Control Chart", fontsize=13, fontweight="bold")
                ax.legend(); ax.grid(alpha=0.3)
                plt.tight_layout()
                emit("OEE & Data Integrity", f"{press}: SPC Control Chart", fig)

                r_val = d["Downtime_Lag_1"].corr(d["Duration (hours)"])
                fig, ax = plt.subplots(figsize=(6, 6))
                sns.regplot(data=d, x="Downtime_Lag_1", y="Duration (hours)",
                           scatter_kws={"alpha": 0.5, "color": color}, line_kws={"color": "#e0a458"}, ax=ax)
                ax.annotate(f"r = {r_val:.3f}", xy=(0.03, 0.97), xycoords="axes fraction", ha="left", va="top",
                           fontsize=10, bbox=dict(boxstyle="round,pad=0.4", facecolor="white", alpha=0.9))
                ax.set_title(f"{press}: 1-Day Autocorrelation", fontsize=12, fontweight="bold")
                plt.tight_layout()
                emit("OEE & Data Integrity", f"{press}: 1-Day Autocorrelation", fig)

    # ================= TAB: Own Reason Logs =================
    with tab_ownlogs:
        st.header("Each Press's Own Reason Log")
        st.caption("Each press has its own \"By Category\" block - no reattribution needed, unlike Barth/Buhler.")
        theme_patterns = cfg["theme_patterns"]
        for press, roles in presses.items():
            df_own = press_category_data.get(press)
            oee_unit, color = roles["oee_source"], roles["color"]
            st.write(f"**{press}** ({oee_unit}): "
                    f"{0 if df_own is None else len(df_own)} reason rows loaded "
                    f"({0.0 if df_own is None else df_own['Duration_hours'].sum():.1f} hrs total)")
            summary, fig = render_theme_section(press, df_own, theme_patterns, color)
            if summary is not None:
                report_figures.append(("Own Reason Logs", f"{press}: Theme Breakdown", fig))
            report_comments.append((press, f"{press}: comments grouped by theme", themed_comments(df_own, theme_patterns)))

    # ================= TAB: Raw Logs =================
    with tab_raw:
        st.header("Per-Unit Raw Logs (reference only)")
        st.caption("The LCS units especially are not independent measurements - see the Data Integrity Check above.")
        raw_rows = []
        for unit in cfg["sequence"]:
            d = dfs.get(unit)
            if d is None or d.empty:
                continue
            raw_rows.append({"Unit": unit, "Total downtime (hrs)": round(d["Duration (hours)"].sum(), 1),
                             "Mean daily (hrs)": round(d["Duration (hours)"].mean(), 2), "Days": len(d)})
        raw_df = pd.DataFrame(raw_rows)
        st.dataframe(raw_df, use_container_width=True)
        oee_units = {roles["oee_source"] for roles in presses.values()}
        fig, ax = plt.subplots(figsize=(11, 5))
        cols = ["#d64545" if u in oee_units else "#9aa5ad" for u in raw_df["Unit"]]
        bars = ax.bar(raw_df["Unit"], raw_df["Total downtime (hrs)"], color=cols)
        ax.bar_label(bars, fmt="%.0f", padding=3, fontsize=8)
        ax.set_title("Self-Reported Downtime per Sub-Unit\n(press OEE-source units in red)", fontsize=12, fontweight="bold")
        plt.setp(ax.xaxis.get_majorticklabels(), rotation=30, ha="right")
        plt.tight_layout()
        emit("Raw Logs", "Self-Reported Downtime per Sub-Unit", fig)

    # ================= TAB: Predictive Analytics =================
    with tab_predictive:
        st.header("Predictive Analytics")
        theme_patterns = cfg["theme_patterns"]

        st.subheader("Equipment-Level Reliability: PR1 vs PR2 vs PR3")
        eq_rows = []
        for press, roles in presses.items():
            df_own = press_category_data.get(press)
            if df_own is None or df_own.empty:
                continue
            total_count = df_own["Count"].sum()
            if total_count <= 0:
                continue
            timeframe_days = len(press_data[press])
            uptime_hours = max(timeframe_days * 24 - df_own["Duration_hours"].sum(), 0.01)
            mtbf = uptime_hours / total_count
            eq_rows.append({"Allocated_To": press, "Total_Duration_hours": df_own["Duration_hours"].sum(),
                            "Total_Count": int(total_count), "MTBF_hrs": mtbf,
                            "Eta": mtbf / 0.9027, "Beta": 1.5})
        if eq_rows:
            eq_df = pd.DataFrame(eq_rows)
            st.dataframe(eq_df.round(2), use_container_width=True)
            fig = plot_weibull(eq_df.rename(columns={"Allocated_To": "Source"}))
            emit("Predictive Analytics", "Equipment-Level Reliability", fig)

        for press, roles in presses.items():
            oee_unit, color = roles["oee_source"], roles["color"]
            d = press_data[press]
            timeframe_days = len(d)
            with st.expander(f"{press}: forecasts", expanded=False):
                st.subheader("Monte Carlo: 30-Day OEE Forecast")
                sim_oee = bootstrap_oee_simulation(d["Duration (hours)"])
                p10, p50, p90 = np.percentile(sim_oee, [10, 50, 90])
                c1, c2, c3 = st.columns(3)
                c1.metric("P10 (worst)", f"{p10:.1f}%")
                c2.metric("P50 (typical)", f"{p50:.1f}%")
                c3.metric("P90 (best)", f"{p90:.1f}%")
                fig, ax = plt.subplots(figsize=(9, 4.5))
                ax.hist(sim_oee, bins=60, color=color, edgecolor="white", alpha=0.85)
                ax.axvline(p50, color="green", linestyle="--", linewidth=2, label=f"P50 ({p50:.1f}%)")
                ax.axvline(p10, color="red", linestyle=":", linewidth=2, label=f"P10 ({p10:.1f}%)")
                ax.axvline(p90, color="red", linestyle=":", linewidth=2, label=f"P90 ({p90:.1f}%)")
                ax.set_title(f"{press}: Simulated 30-Day OEE Distribution", fontsize=12, fontweight="bold")
                ax.legend()
                plt.tight_layout()
                emit("Predictive Analytics", f"{press}: Monte Carlo OEE Forecast", fig)

                theme_summary, _ = theme_breakdown_data(press_category_data.get(press, pd.DataFrame()), theme_patterns)
                if not theme_summary.empty:
                    rng = np.random.default_rng(42)
                    risk_rows = []
                    for _, row in theme_summary.iterrows():
                        sim = simulate_category_downtime(row["Total_Duration_hours"], row["Total_Count"], timeframe_days, rng=rng)
                        risk_rows.append({"Theme": row["Theme"], "Historical (hrs)": row["Total_Duration_hours"],
                                          "Sim P10 (30d)": np.percentile(sim, 10), "Sim P50 (30d)": np.percentile(sim, 50),
                                          "Sim P90 (30d)": np.percentile(sim, 90)})
                    risk_df = pd.DataFrame(risk_rows).sort_values("Sim P90 (30d)", ascending=False)
                    st.dataframe(risk_df.round(2), use_container_width=True)

                st.subheader("Weibull Reliability - Top 3 Themes")
                weibull_df = compute_weibull(theme_summary, timeframe_days)
                if not weibull_df.empty:
                    st.dataframe(weibull_df.round(2), use_container_width=True)
                    fig = plot_weibull(weibull_df)
                    emit("Predictive Analytics", f"{press}: Weibull Reliability", fig)
                else:
                    st.info("Not enough themed data to compute Weibull parameters.")

    # ================= TAB: Report Export =================
    with tab_report:
        total_dt = sum(press_data[p]["Duration (hours)"].sum() for p in presses)
        key_metrics = [
            ("Line", line_name),
            ("Days analysed", len(press_data[first_press])),
            ("Months excluded", ", ".join(month_labels[ym] for ym in excluded) if excluded else "None"),
            ("Average OEE (mean of presses)", f"{line_oee:.2f}%"),
            ("Total downtime (all presses)", f"{total_dt:.1f} hrs"),
        ]
        render_report_tab(line_name, key_metrics, report_figures, report_comments)


def render_moulding_line(cfg, dfs_raw, eq_titles, unit_category_data, line_name, discovery_status=None):
    """Moulding lines (Liquor Moulding, Butter Moulding) are a third shape,
    different from both Barth/Buhler's single-roaster-with-reattribution
    and Presses' fixed one-LCS-per-press pattern: multiple fully-analysed
    units (own OEE, own reason log, own themes), each with its OWN LIST of
    supporting units - which can be empty (ML11/ML12 have none), one
    (Presses' shape), or two (TM41 has ML and PK Group). A shared theme
    library covers every unit on the line. Adds a Maintenance Quadrant
    combining every unit into one chart (one bubble per unit, not per
    theme), the direct equivalent of Barth's Allocated_To quadrant with no
    cross-unit reattribution to drive it."""
    units = cfg["units"]
    first_unit, first_roles = next(iter(units.items()))
    ref_unit = first_roles["oee_source"]

    if discovery_status and _show_prefix_mismatch_error(line_name, discovery_status.get("__meta__", {})):
        return

    if dfs_raw[ref_unit].empty:
        st.error(f"No data loaded for {ref_unit} - check skiprows/usecols against the real file.")
        return

    if discovery_status:
        _render_discovery_sidebar(discovery_status)

    excluded, month_labels = _render_month_exclusion(
        [dfs_raw[roles["oee_source"]] for roles in units.values()], cfg.get("exclude_march_default"))

    dfs = {unit: apply_month_exclusion(d, excluded) for unit, d in dfs_raw.items()}

    unit_data, unit_oee = {}, {}
    for unit_label, roles in units.items():
        d = dfs[roles["oee_source"]].copy()
        d["OEE_Percentage"] = ((24 - d["Duration (hours)"]) / 24) * 100
        unit_data[unit_label] = d
        unit_oee[unit_label] = d["OEE_Percentage"].mean()
    line_oee = float(np.mean(list(unit_oee.values())))
    window_start = unit_data[first_unit]["Date"].min()
    window_end = unit_data[first_unit]["Date"].max()

    _show_logo_if_present()

    # No line-level "mean of N units" OEE pill here on purpose: with each
    # unit's own "Average OEE" card visible right below (the first unit's
    # expander is open by default), a second, differently-scoped "Average
    # OEE" figure in the header read as a contradiction, not a summary -
    # confirmed confusing in practice, not just in theory. line_oee is
    # still computed and still shown in the Report Export tab, where it's
    # the only OEE figure on the page and the "(mean of units)" label has
    # no competing per-unit card next to it to conflict with.
    dc_header(
        "Downtime & OEE Intelligence", "Barry Callebaut Ghana · Cocoa Processing",
        dc_pill(_ICON_BRANCH, line_name, "orange")
        + dc_pill(_ICON_CALENDAR_X, f"{window_start.day} {window_start.strftime('%b %Y')} \u2013 "
                                     f"{window_end.day} {window_end.strftime('%b %Y')}", "gray"),
    )
    st.markdown(
        f"**{len(unit_data[first_unit])} days analysed** after exclusions"
        + (f" (excluded: {', '.join(month_labels[ym] for ym in excluded)})" if excluded else "")
    )

    report_figures = []
    report_comments = []

    def emit(section, title, fig):
        st.pyplot(fig)
        report_figures.append((section, title, fig))

    tab_names = ["OEE & Data Integrity", "Own Reason Logs", "Raw Logs", "Predictive Analytics",
                "Maintenance Quadrant", "Report Export"]
    tab_overview, tab_ownlogs, tab_raw, tab_predictive, tab_quadrant, tab_report = st.tabs(tab_names)

    # ================= TAB: OEE & Data Integrity =================
    with tab_overview:
        st.markdown(
            dc_card_head(_ICON_UPLOAD, "Load Daily Downtime Data",
                        "Real column headers checked against the config, not just trusted"),
            unsafe_allow_html=True,
        )
        id_rows = []
        for unit, d in dfs.items():
            date_range = f"{d['Date'].min().date()} to {d['Date'].max().date()}" if len(d) else "N/A"
            id_rows.append({"Config name": unit, "Real name from Excel header": eq_titles[unit],
                            "Rows": len(d), "Date range": date_range,
                            "Total hours": round(d["Duration (hours)"].sum(), 1)})
        st.dataframe(pd.DataFrame(id_rows), use_container_width=True)
        st.markdown('<div style="height:8px"></div>', unsafe_allow_html=True)

        for i, (unit_label, roles) in enumerate(units.items()):
            oee_unit, other_units, color = roles["oee_source"], roles["other_units"], roles["color"]
            d = unit_data[unit_label]
            with st.expander(f"{unit_label} ({oee_unit})", expanded=(i == 0)):
                st.markdown(
                    dc_card_head(_ICON_SCALE, "Data Integrity Check",
                                f"Do {', '.join(other_units)} track {oee_unit}?" if other_units
                                else f"{oee_unit} has no supporting unit to compare against"),
                    unsafe_allow_html=True,
                )
                if other_units:
                    merged = d[["Date", "Duration (hours)"]].rename(columns={"Duration (hours)": oee_unit})
                    for u in other_units:
                        merged = merged.merge(
                            dfs[u][["Date", "Duration (hours)"]].rename(columns={"Duration (hours)": u}),
                            on="Date", how="inner")
                    merged["Sum_of_others"] = merged[other_units].sum(axis=1)
                    r = merged["Sum_of_others"].corr(merged[oee_unit])
                    c1, c2, c3 = st.columns(3)
                    c1.markdown(f'<div class="dc-card"><span class="dc-metric-label">Mean {oee_unit}</span>'
                               f'<div class="dc-metric-value" style="font-size:22px;">{merged[oee_unit].mean():.2f} hrs/day</div></div>',
                               unsafe_allow_html=True)
                    c2.markdown(f'<div class="dc-card"><span class="dc-metric-label">Mean sum of others</span>'
                               f'<div class="dc-metric-value" style="font-size:22px;">{merged["Sum_of_others"].mean():.2f} hrs/day</div></div>',
                               unsafe_allow_html=True)
                    c3.markdown(f'<div class="dc-card"><span class="dc-metric-label">Correlation (r)</span>'
                               f'<div class="dc-metric-value" style="font-size:22px;">{r:.3f}</div></div>',
                               unsafe_allow_html=True)
                    st.markdown('<div style="height:10px"></div>', unsafe_allow_html=True)
                else:
                    st.info(f"{unit_label} has no supporting units in this config - nothing to compare here.")

                d["OEE_7Day_SMA"] = d["OEE_Percentage"].rolling(7, min_periods=1).mean()
                d["OEE_7Day_EWMA"] = d["OEE_Percentage"].ewm(span=7, adjust=False).mean()
                d["Downtime_Lag_1"] = d["Duration (hours)"].shift(1)
                d["Cumulative_Downtime"] = d["Duration (hours)"].cumsum()
                best_row = d.loc[d["OEE_Percentage"].idxmax()]
                worst_row = d.loc[d["OEE_Percentage"].idxmin()]

                m1, m2, m3 = st.columns(3)
                m1.markdown(dc_metric_card("Average OEE", f"{unit_oee[unit_label]:.1f}%", "filtered window", _ICON_GAUGE, "orange"), unsafe_allow_html=True)
                m2.markdown(dc_metric_card("Best Day", f"{best_row['OEE_Percentage']:.1f}%", f"{best_row['Date'].day} {best_row['Date'].strftime('%b')}", _ICON_UP, "green"), unsafe_allow_html=True)
                m3.markdown(dc_metric_card("Worst Day", f"{worst_row['OEE_Percentage']:.1f}%", f"{worst_row['Date'].day} {worst_row['Date'].strftime('%b')}", _ICON_DOWN, "coral"), unsafe_allow_html=True)
                st.markdown('<div style="height:10px"></div>', unsafe_allow_html=True)

                fig, ax = plt.subplots(figsize=(12, 5))
                ax.plot(d["Date"], d["OEE_Percentage"], label="Daily OEE (raw)", alpha=0.3, color="grey", marker=".")
                ax.plot(d["Date"], d["OEE_7Day_SMA"], label="7-day SMA", color="#e0a458", linewidth=2, linestyle="--")
                ax.plot(d["Date"], d["OEE_7Day_EWMA"], label="7-day EWMA", color=color, linewidth=2)
                ax.axhline(unit_oee[unit_label], color=color, linestyle=":", linewidth=2, alpha=0.7, label=f"Mean ({unit_oee[unit_label]:.1f}%)")
                ax.set_title(f"{unit_label}: OEE Trend", fontsize=13, fontweight="bold")
                ax.set_ylim(0, 105); ax.legend(); ax.grid(alpha=0.3)
                plt.tight_layout()
                emit("OEE & Data Integrity", f"{unit_label}: OEE Trend", fig)

                mean_dt = d["Duration (hours)"].mean()
                ucl = min(mean_dt + 3 * d["Duration (hours)"].std(), 24)
                out = d[d["Duration (hours)"] >= ucl]
                fig, ax = plt.subplots(figsize=(12, 4.5))
                ax.plot(d["Date"], d["Duration (hours)"], marker="o", color="grey", alpha=0.6, markersize=4, label="Daily downtime")
                ax.axhline(mean_dt, color="#e0a458", linestyle="--", linewidth=2, label=f"Mean ({mean_dt:.2f} hrs)")
                ax.axhline(ucl, color=color, linewidth=2, label=f"UCL +3\u03c3 ({ucl:.2f} hrs)")
                ax.scatter(out["Date"], out["Duration (hours)"], color=color, s=80, zorder=5, label=f"Special cause ({len(out)} days)")
                ax.set_title(f"{unit_label}: SPC Control Chart", fontsize=13, fontweight="bold")
                ax.legend(); ax.grid(alpha=0.3)
                plt.tight_layout()
                emit("OEE & Data Integrity", f"{unit_label}: SPC Control Chart", fig)

                r_val = d["Downtime_Lag_1"].corr(d["Duration (hours)"])
                fig, ax = plt.subplots(figsize=(6, 6))
                sns.regplot(data=d, x="Downtime_Lag_1", y="Duration (hours)",
                           scatter_kws={"alpha": 0.5, "color": color}, line_kws={"color": "#e0a458"}, ax=ax)
                ax.annotate(f"r = {r_val:.3f}", xy=(0.03, 0.97), xycoords="axes fraction", ha="left", va="top",
                           fontsize=10, bbox=dict(boxstyle="round,pad=0.4", facecolor="white", alpha=0.9))
                ax.set_title(f"{unit_label}: 1-Day Autocorrelation", fontsize=12, fontweight="bold")
                plt.tight_layout()
                emit("OEE & Data Integrity", f"{unit_label}: 1-Day Autocorrelation", fig)

    # ================= TAB: Own Reason Logs =================
    with tab_ownlogs:
        st.header("Each Unit's Own Reason Log")
        st.caption("Every analysed unit has its own \"By Category\" block - no reattribution needed, unlike Barth/Buhler.")
        theme_patterns = cfg["theme_patterns"]
        for unit_label, roles in units.items():
            df_own = unit_category_data.get(unit_label)
            oee_unit, color = roles["oee_source"], roles["color"]
            st.write(f"**{unit_label}** ({oee_unit}): "
                    f"{0 if df_own is None else len(df_own)} reason rows loaded "
                    f"({0.0 if df_own is None else df_own['Duration_hours'].sum():.1f} hrs total)")
            summary, fig = render_theme_section(unit_label, df_own, theme_patterns, color)
            if summary is not None:
                report_figures.append(("Own Reason Logs", f"{unit_label}: Theme Breakdown", fig))
            report_comments.append((unit_label, f"{unit_label}: comments grouped by theme",
                                    themed_comments(df_own, theme_patterns)))

    # ================= TAB: Raw Logs =================
    with tab_raw:
        st.header("Per-Unit Raw Logs (reference only)")
        st.caption("Supporting units especially are not independent measurements - see the Data Integrity Check above.")
        raw_rows = []
        for unit in cfg["sequence"]:
            d = dfs.get(unit)
            if d is None or d.empty:
                continue
            raw_rows.append({"Unit": unit, "Total downtime (hrs)": round(d["Duration (hours)"].sum(), 1),
                             "Mean daily (hrs)": round(d["Duration (hours)"].mean(), 2), "Days": len(d)})
        raw_df = pd.DataFrame(raw_rows)
        st.dataframe(raw_df, use_container_width=True)
        oee_units = {roles["oee_source"] for roles in units.values()}
        fig, ax = plt.subplots(figsize=(11, 5))
        cols = ["#d64545" if u in oee_units else "#9aa5ad" for u in raw_df["Unit"]]
        bars = ax.bar(raw_df["Unit"], raw_df["Total downtime (hrs)"], color=cols)
        ax.bar_label(bars, fmt="%.0f", padding=3, fontsize=8)
        ax.set_title("Self-Reported Downtime per Sub-Unit\n(analysed OEE-source units in red)", fontsize=12, fontweight="bold")
        plt.setp(ax.xaxis.get_majorticklabels(), rotation=30, ha="right")
        plt.tight_layout()
        emit("Raw Logs", "Self-Reported Downtime per Sub-Unit", fig)

    # ================= TAB: Predictive Analytics =================
    with tab_predictive:
        st.header("Predictive Analytics")
        theme_patterns = cfg["theme_patterns"]

        for unit_label, roles in units.items():
            oee_unit, color = roles["oee_source"], roles["color"]
            d = unit_data[unit_label]
            timeframe_days = len(d)
            with st.expander(f"{unit_label}: forecasts", expanded=False):
                st.subheader("Monte Carlo: 30-Day OEE Forecast")
                sim_oee = bootstrap_oee_simulation(d["Duration (hours)"])
                p10, p50, p90 = np.percentile(sim_oee, [10, 50, 90])
                c1, c2, c3 = st.columns(3)
                c1.metric("P10 (worst)", f"{p10:.1f}%")
                c2.metric("P50 (typical)", f"{p50:.1f}%")
                c3.metric("P90 (best)", f"{p90:.1f}%")
                fig, ax = plt.subplots(figsize=(9, 4.5))
                ax.hist(sim_oee, bins=60, color=color, edgecolor="white", alpha=0.85)
                ax.axvline(p50, color="green", linestyle="--", linewidth=2, label=f"P50 ({p50:.1f}%)")
                ax.axvline(p10, color="red", linestyle=":", linewidth=2, label=f"P10 ({p10:.1f}%)")
                ax.axvline(p90, color="red", linestyle=":", linewidth=2, label=f"P90 ({p90:.1f}%)")
                ax.set_title(f"{unit_label}: Simulated 30-Day OEE Distribution", fontsize=12, fontweight="bold")
                ax.legend()
                plt.tight_layout()
                emit("Predictive Analytics", f"{unit_label}: Monte Carlo OEE Forecast", fig)

                theme_summary, _ = theme_breakdown_data(unit_category_data.get(unit_label, pd.DataFrame()), theme_patterns)
                if not theme_summary.empty:
                    rng = np.random.default_rng(42)
                    risk_rows = []
                    for _, row in theme_summary.iterrows():
                        sim = simulate_category_downtime(row["Total_Duration_hours"], row["Total_Count"], timeframe_days, rng=rng)
                        risk_rows.append({"Theme": row["Theme"], "Historical (hrs)": row["Total_Duration_hours"],
                                          "Sim P10 (30d)": np.percentile(sim, 10), "Sim P50 (30d)": np.percentile(sim, 50),
                                          "Sim P90 (30d)": np.percentile(sim, 90)})
                    risk_df = pd.DataFrame(risk_rows).sort_values("Sim P90 (30d)", ascending=False)
                    st.dataframe(risk_df.round(2), use_container_width=True)

                st.subheader("Weibull Reliability - Top 3 Themes")
                weibull_df = compute_weibull(theme_summary, timeframe_days)
                if not weibull_df.empty:
                    st.dataframe(weibull_df.round(2), use_container_width=True)
                    fig = plot_weibull(weibull_df)
                    emit("Predictive Analytics", f"{unit_label}: Weibull Reliability", fig)
                else:
                    st.info("Not enough themed data to compute Weibull parameters.")

    # ================= TAB: Maintenance Quadrant =================
    with tab_quadrant:
        st.header("Maintenance Strategy Quadrant")
        st.caption(
            "One bubble per analysed unit, aggregated across each unit's whole reason log - "
            "the direct equivalent of Barth's Allocated_To quadrant, with no cross-unit "
            "reattribution to drive it (each bubble is a whole unit, not a theme within one)."
        )
        q_rows = []
        for unit_label, roles in units.items():
            df_own = unit_category_data.get(unit_label)
            if df_own is None or df_own.empty:
                continue
            total_count = df_own["Count"].sum()
            total_hours = df_own["Duration_hours"].sum()
            if total_count <= 0:
                continue
            q_rows.append({"Unit": unit_label, "Total_Count": total_count,
                           "Total_Duration_hours": total_hours,
                           "Avg_Mins_per_Event": total_hours * 60 / total_count})
        q = pd.DataFrame(q_rows)
        if len(q) >= 2:
            fig, q_labeled = collision_aware_quadrant(q, f"{line_name}: Maintenance Strategy by Unit", label_col="Unit")
            emit("Maintenance Quadrant", f"{line_name}: Maintenance Strategy by Unit", fig)
            st.dataframe(q_labeled, use_container_width=True)
        else:
            st.info("Fewer than 2 units with logged events - not enough for a quadrant chart.")

    # ================= TAB: Report Export =================
    with tab_report:
        total_dt = sum(unit_data[u]["Duration (hours)"].sum() for u in units)
        key_metrics = [
            ("Line", line_name),
            ("Days analysed", len(unit_data[first_unit])),
            ("Months excluded", ", ".join(month_labels[ym] for ym in excluded) if excluded else "None"),
            ("Average OEE (mean of units)", f"{line_oee:.2f}%"),
            ("Total downtime (all units)", f"{total_dt:.1f} hrs"),
        ]
        render_report_tab(line_name, key_metrics, report_figures, report_comments)


# ============================================================
# SIDEBAR + DISPATCH
# ============================================================
def main():
    """Wrapped in a function (not bare module-level code) specifically so
    this file can be safely imported by a test suite - importing it used to
    immediately try to render Streamlit UI and call st.stop(), which made
    it impossible to import dashboard.py's pure functions for testing
    without a live Streamlit session."""
    st.set_page_config(page_title="Downtime & OEE Intelligence",page_icon="favicon_barry.jpg" , layout="wide")
    st.markdown(CSS_BLOCK, unsafe_allow_html=True)

    st.sidebar.markdown(dc_card_head(_ICON_BRANCH, "Line Setup", "Configure the analysis source"), unsafe_allow_html=True)
    st.sidebar.markdown('<div style="height:6px"></div>', unsafe_allow_html=True)
    line_choice = st.sidebar.selectbox("Select line", list(LINES.keys()), label_visibility="collapsed")
    cfg = LINES[line_choice]
    n_units = len(cfg["daily_cols"])
    st.sidebar.caption(f"Cocoa primary processing · {n_units} units")

    st.sidebar.markdown('<div style="height:10px"></div>', unsafe_allow_html=True)
    st.sidebar.markdown(dc_card_head(_ICON_UPLOAD, "Upload Downtime Log", "Drag and drop the .xlsx file"), unsafe_allow_html=True)
    uploaded_file = st.sidebar.file_uploader("Upload Downtime Log (.xlsx)", type=["xlsx"], label_visibility="collapsed")

    # Run Analysis gate: analysis only starts on an explicit click, not the
    # instant a file is uploaded. Tracked by (file identity, line) in
    # session_state rather than just the button's own return value, since
    # a Streamlit button only returns True on the exact rerun it was
    # clicked - every later interaction (changing month exclusion, which
    # lives inside render_line() further down and reruns the whole script)
    # needs the activation to persist, not require re-clicking Run
    # Analysis every time. Switching to a different file or line DOES
    # reset it, deliberately - stale results from a previous file should
    # never linger under a new one without an explicit re-run.
    run_key = (uploaded_file.name, uploaded_file.size, line_choice) if uploaded_file is not None else None
    if st.session_state.get("_analysis_run_key") != run_key:
        st.session_state["_analysis_active"] = False
        st.session_state["_analysis_run_key"] = run_key

    st.sidebar.markdown('<div style="height:14px"></div>', unsafe_allow_html=True)
    run_clicked = st.sidebar.button(
        "\u25b6 Run Analysis", type="primary", use_container_width=True, disabled=uploaded_file is None,
    )
    if run_clicked:
        st.session_state["_analysis_active"] = True

    if uploaded_file is None:
        _show_logo_if_present()

        dc_header("Downtime & OEE Intelligence", "Barry Callebaut Ghana · Cocoa Processing",
                  dc_pill(_ICON_BRANCH, line_choice, "orange"))
        st.info("Upload a downtime log Excel file in the sidebar to begin.")
        st.stop()

    if not st.session_state.get("_analysis_active"):
        _show_logo_if_present()

        dc_header("Downtime & OEE Intelligence", "Barry Callebaut Ghana · Cocoa Processing",
                  dc_pill(_ICON_BRANCH, line_choice, "orange"))
        st.info(f"**{uploaded_file.name}** is ready. Click **\u25b6 Run Analysis** in the sidebar to begin.")
        st.stop()

    dfs_raw, eq_titles, cat_data, own_category_data_cached, discovery_status = load_all_data(uploaded_file.getvalue(), line_choice)
    if cfg.get("kind") == "presses":
        render_presses(cfg, dfs_raw, eq_titles, cat_data, line_choice, discovery_status)
    elif cfg.get("kind") == "moulding":
        render_moulding_line(cfg, dfs_raw, eq_titles, cat_data, line_choice, discovery_status)
    else:
        render_line(cfg, dfs_raw, eq_titles, cat_data, own_category_data_cached, line_choice, discovery_status)

    st.markdown(
        '<div class="dc-footer">Downtime &amp; OEE Intelligence · Barry Callebaut Ghana reliability '
        'workspace · analysis runs on your uploaded file for this session only.</div>',
        unsafe_allow_html=True,
    )


if __name__ == "__main__":
    main()
