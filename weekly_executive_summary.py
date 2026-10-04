"""
Weekly Executive Summary Dashboard
────────────────────────────────────────────────────────────────────
A self-contained Streamlit app that transforms raw sales data into
a polished, executive-grade weekly summary — no analyst required.

Run:  streamlit run weekly_executive_summary.py
Deps: pip install streamlit pandas numpy openpyxl
"""

import io
from datetime import timedelta

import numpy as np
import pandas as pd
import plotly.graph_objects as go
import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# ══════════════════════════════════════════════════════════════════════════════
# PAGE CONFIG  (must be the very first Streamlit call)
# ══════════════════════════════════════════════════════════════════════════════

st.set_page_config(
    page_title="Weekly Executive Summary",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ══════════════════════════════════════════════════════════════════════════════
# GLOBAL STYLES
# ══════════════════════════════════════════════════════════════════════════════

st.markdown(
    """
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800&display=swap');
html, body, [class*="css"] { font-family: 'Inter', sans-serif !important; }

/* ── Force light background on main content area ── */
.stApp                        { background-color: #f8fafc !important; }
.stApp > .main                { background-color: #f8fafc !important; }
.block-container              { background-color: #f8fafc !important; }

/* ── Sidebar ── */
section[data-testid="stSidebar"] {
    background: linear-gradient(180deg, #0f172a 0%, #1a2744 100%) !important;
}
section[data-testid="stSidebar"] .stMarkdown *,
section[data-testid="stSidebar"] label,
section[data-testid="stSidebar"] span,
section[data-testid="stSidebar"] p      { color: #94a3b8 !important; }
section[data-testid="stSidebar"] h2,
section[data-testid="stSidebar"] h3     { color: #f1f5f9 !important; }
section[data-testid="stSidebar"] hr     { border-color: #1e3a5f !important; }
section[data-testid="stSidebar"] .stDownloadButton button {
    background: #1e3a8a !important; color: #bfdbfe !important;
    border: 1px solid #2563eb !important; font-size: 13px !important;
}

/* ── Metric cards ── */
[data-testid="metric-container"] {
    background: #ffffff !important; border: 1px solid #e2e8f0 !important;
    border-radius: 10px !important; padding: 14px 18px !important;
    box-shadow: 0 1px 3px rgba(0,0,0,0.06) !important;
}
[data-testid="stMetricValue"]  { font-size: 20px !important; font-weight: 700 !important; color: #0f172a !important; }
[data-testid="stMetricLabel"]  { font-size: 11px !important; font-weight: 600 !important; color: #475569 !important; text-transform: uppercase; letter-spacing: .05em; }
[data-testid="stMetricDelta"]  { font-size: 13px !important; font-weight: 600 !important; }

/* ── Tabs ── */
.stTabs [data-baseweb="tab-list"] {
    background: #e2e8f0 !important; border-radius: 10px; padding: 4px; gap: 2px;
}
.stTabs [data-baseweb="tab"] {
    border-radius: 7px; font-size: 13px; font-weight: 500;
    color: #334155 !important; padding: 6px 18px;
}
.stTabs [aria-selected="true"] {
    background: #fff !important; color: #0f172a !important;
    box-shadow: 0 1px 4px rgba(0,0,0,0.12); font-weight: 600;
}

/* ── Text ── */
.section-title { font-size: 18px; font-weight: 700; color: #0f172a !important; margin-bottom: 2px; }
.section-sub   { font-size: 13px; color: #475569 !important; margin-bottom: 14px; }
p, label, .stMarkdown p { color: #1e293b !important; }

/* ── Date range info box ── */
.dr-box {
    background: #dbeafe !important; border: 1px solid #93c5fd !important;
    border-radius: 8px; padding: 10px 14px; font-size: 13px;
    color: #1e3a8a !important; margin-bottom: 14px; font-weight: 500;
}

/* ── Empty state cards ── */
.empty-card {
    background: #fff !important; border: 1px solid #e2e8f0; border-radius: 12px;
    padding: 28px 24px; text-align: center;
    box-shadow: 0 2px 6px rgba(0,0,0,0.04); height: 155px;
    display: flex; flex-direction: column; align-items: center; justify-content: center;
}
</style>
""",
    unsafe_allow_html=True,
)

# ══════════════════════════════════════════════════════════════════════════════
# PASSWORD GATE
# ══════════════════════════════════════════════════════════════════════════════

def _check_password():
    """Block access until the correct password is entered. Uses st.secrets."""
    if st.session_state.get("authenticated"):
        return

    _, mid, _ = st.columns([1, 1.2, 1])
    with mid:
        st.markdown("""
        <div style="text-align:center; padding:48px 0 28px;">
            <div style="font-size:44px; margin-bottom:14px">📊</div>
            <div style="font-size:24px; font-weight:800; color:#0f172a; margin-bottom:6px;">
                Weekly Executive Summary
            </div>
            <div style="font-size:14px; color:#64748b; margin-bottom:32px;">
                Enter your access password to continue
            </div>
        </div>
        """, unsafe_allow_html=True)

        pw = st.text_input(
            "Password", type="password",
            label_visibility="collapsed",
            placeholder="Enter your password"
        )
        if st.button("Access Dashboard →", use_container_width=True, type="primary"):
            correct = st.secrets.get("APP_PASSWORD", "")
            if pw == correct and correct != "":
                st.session_state.authenticated = True
                st.rerun()
            else:
                st.error("Incorrect password. Please check your purchase confirmation and try again.")

        st.markdown("""
        <div style="text-align:center; margin-top:16px; font-size:12px; color:#94a3b8;">
            Purchased this tool? Your password was included in your receipt.
        </div>
        """, unsafe_allow_html=True)

    st.stop()

_check_password()

# ══════════════════════════════════════════════════════════════════════════════
# CONSTANTS
# ══════════════════════════════════════════════════════════════════════════════

COLUMN_ALIASES: dict[str, list[str]] = {
    "Date": ["date"],
    "Item Name": ["item name", "item", "product name", "product", "sku", "sku name", "description"],
    "Sales Revenue": ["sales revenue", "revenue", "sales", "total revenue", "net revenue", "amount", "total sales"],
    "Units Sold": ["units sold", "units", "qty", "quantity", "quantity sold", "volume"],
}

TIMEFRAMES: list[tuple[str, str]] = [
    ("L1W",  "Latest Week"),
    ("L4W",  "Latest 4 Weeks"),
    ("L13W", "Latest 13 Weeks"),
    ("L52W", "Latest 52 Weeks"),
    ("YTD",  "Year to Date"),
]

# ══════════════════════════════════════════════════════════════════════════════
# HELPERS — DATA LOADING & VALIDATION
# ══════════════════════════════════════════════════════════════════════════════

def _resolve_columns(df: pd.DataFrame) -> tuple[dict, list[str]]:
    lower_map = {c.strip().lower(): c for c in df.columns}
    rename: dict[str, str] = {}
    missing: list[str] = []
    for standard, aliases in COLUMN_ALIASES.items():
        matched = next((lower_map[a] for a in aliases if a in lower_map), None)
        if matched:
            rename[matched] = standard
        else:
            missing.append(standard)
    return rename, missing


@st.cache_data(show_spinner=False)
def load_file(file_bytes: bytes, filename: str) -> pd.DataFrame:
    buf = io.BytesIO(file_bytes)
    raw = (
        pd.read_csv(buf)
        if filename.lower().endswith(".csv")
        else pd.read_excel(buf)
    )
    rename, missing = _resolve_columns(raw)
    if missing:
        raise ValueError(f"Missing required column(s): **{', '.join(missing)}**")
    df = raw.rename(columns=rename)[list(COLUMN_ALIASES)].copy()
    df["Date"] = pd.to_datetime(df["Date"])
    df["Sales Revenue"] = pd.to_numeric(df["Sales Revenue"], errors="coerce").fillna(0)
    df["Units Sold"]    = pd.to_numeric(df["Units Sold"],    errors="coerce").fillna(0)
    return df.dropna(subset=["Date"]).sort_values("Date").reset_index(drop=True)

# ══════════════════════════════════════════════════════════════════════════════
# HELPERS — DATE WINDOWS & METRIC COMPUTATION
# ══════════════════════════════════════════════════════════════════════════════

def _window(ref: pd.Timestamp, tf: str) -> tuple[pd.Timestamp, pd.Timestamp, float]:
    if tf == "L1W":  return ref - timedelta(days=6),   ref, 1.0
    if tf == "L4W":  return ref - timedelta(days=27),  ref, 4.0
    if tf == "L13W": return ref - timedelta(days=90),  ref, 13.0
    if tf == "L52W": return ref - timedelta(days=363), ref, 52.0
    if tf == "YTD":
        start = pd.Timestamp(ref.year, 1, 1)
        return start, ref, max((ref - start).days + 1, 7) / 7
    raise ValueError(tf)


def _agg_items(df: pd.DataFrame, start, end) -> pd.DataFrame:
    mask = (df["Date"] >= start) & (df["Date"] <= end)
    sub = df[mask]
    if sub.empty:
        return pd.DataFrame({"Item Name": pd.Series(dtype=str),
                             "Revenue":   pd.Series(dtype=float),
                             "Units":     pd.Series(dtype=float)})
    return sub.groupby("Item Name", as_index=False).agg(
        Revenue=("Sales Revenue", "sum"),
        Units=("Units Sold", "sum"),
    )


@st.cache_data(show_spinner=False)
def compute_all_timeframes(df: pd.DataFrame) -> dict:
    """Compute CY + LY aggregates for all 5 timeframes. Cached per unique df."""
    max_date = df["Date"].max()
    out = {}
    for tf, _ in TIMEFRAMES:
        cy_s, cy_e, cy_w = _window(max_date, tf)
        ly_s, ly_e, ly_w = _window(max_date - timedelta(weeks=52), tf)
        cy = _agg_items(df, cy_s, cy_e)
        ly = _agg_items(df, ly_s, ly_e)
        m = cy.merge(ly, on="Item Name", how="outer", suffixes=("_CY", "_LY")).fillna(0)
        tot_cy = m["Revenue_CY"].sum()
        tot_ly = m["Revenue_LY"].sum()
        m["Pct_CY"]  = m["Revenue_CY"] / tot_cy if tot_cy else 0.0
        m["Vel_CY"]  = m["Units_CY"] / cy_w
        m["Vel_LY"]  = m["Units_LY"] / ly_w
        m["YoY_Rev"] = np.where(m["Revenue_LY"] > 0,
                                (m["Revenue_CY"] - m["Revenue_LY"]) / m["Revenue_LY"],
                                np.nan)
        out[tf] = (m, {"cy_s": cy_s, "cy_e": cy_e, "cy_w": cy_w,
                       "ly_s": ly_s, "ly_e": ly_e, "ly_w": ly_w,
                       "tot_cy": tot_cy, "tot_ly": tot_ly})
    return out

# ══════════════════════════════════════════════════════════════════════════════
# HELPERS — DISPLAY FORMATTING
# ══════════════════════════════════════════════════════════════════════════════

def _usd(v)  -> str: return f"${v:,.0f}" if (pd.notna(v) and v != 0) else "$0"
def _pct(v, sign=False) -> str:
    if pd.isna(v): return "—"
    return f"{'+'if(sign and v>0)else''}{v:.1%}"
def _units(v) -> str: return f"{int(v):,}" if pd.notna(v) else "0"
def _vel(v)   -> str: return f"{v:,.1f}"  if (pd.notna(v) and v > 0) else "0.0"
def _yoy(v) -> str:
    if pd.isna(v): return "—"
    if abs(v) < 0.005: return f"● {v:.1%}"          # within ±0.5% = flat
    return ("▲ " if v > 0 else "▼ ") + f"{abs(v):.1%}"

def _style_yoy(val: str) -> str:
    """CSS for a YoY cell: green ▲, red ▼, yellow ● flat, blank for —."""
    if not isinstance(val, str): return ""
    if val.startswith("▲"): return "background-color: #dcfce7; color: #166534; font-weight: 600"
    if val.startswith("▼"): return "background-color: #fee2e2; color: #991b1b; font-weight: 600"
    if val.startswith("●"): return "background-color: #fef9c3; color: #854d0e; font-weight: 600"
    return ""

# ══════════════════════════════════════════════════════════════════════════════
# EXCEL EXPORT
# ══════════════════════════════════════════════════════════════════════════════

# ── openpyxl style helpers ────────────────────────────────────────────────────
_NAVY  = "0F172A"; _BLUE  = "1E3A8A"; _SLATE = "475569"
_WHITE = "FFFFFF"; _LIGHT = "F8FAFC"; _LGRAY = "E2E8F0"
_GREEN = "16A34A"; _RED   = "DC2626"; _LBLUE = "EFF6FF"
_FMT_USD = "$#,##0";  _FMT_PCT = "0.0%"; _FMT_INT = "#,##0"; _FMT_VEL = "0.0"

def _xl_fill(hex_c: str) -> PatternFill:
    return PatternFill("solid", fgColor=hex_c)

def _xl_font(bold=False, color="1E293B", size=10) -> Font:
    return Font(name="Arial", bold=bold, color=color, size=size)

def _xl_border() -> Border:
    t = Side(style="thin", color="CBD5E1")
    return Border(left=t, right=t, top=t, bottom=t)

_CTR  = Alignment(horizontal="center", vertical="center")
_LEFT = Alignment(horizontal="left",   vertical="center")

def _xl_title(ws, text: str, row: int, ncols: int, bg=_NAVY, size=12):
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=ncols)
    c = ws.cell(row=row, column=1, value=text)
    c.fill = _xl_fill(bg); c.font = _xl_font(bold=True, color=_WHITE, size=size)
    c.alignment = _LEFT
    ws.row_dimensions[row].height = 26

def _xl_headers(ws, headers: list[str], row: int, bg=_SLATE):
    for col, h in enumerate(headers, 1):
        c = ws.cell(row=row, column=col, value=h)
        c.fill = _xl_fill(bg); c.font = _xl_font(bold=True, color=_WHITE)
        c.alignment = _CTR; c.border = _xl_border()
    ws.row_dimensions[row].height = 22

def _xl_cell(ws, row: int, col: int, value, fmt=None,
             bold=False, bg=None, color="1E293B", align=_CTR):
    c = ws.cell(row=row, column=col, value=value)
    c.font = _xl_font(bold=bold, color=color)
    c.alignment = align; c.border = _xl_border()
    if bg:  c.fill = _xl_fill(bg)
    if fmt and isinstance(value, (int, float)): c.number_format = fmt
    return c

def _xl_yoy_cell(ws, row: int, col: int, raw_val, bg):
    """Write a YoY % cell: numeric+format+green/red color, or '—' if NaN."""
    if pd.notna(raw_val):
        color = _GREEN if raw_val > 0 else _RED
        _xl_cell(ws, row, col, raw_val, _FMT_PCT, bg=bg, color=color)
    else:
        _xl_cell(ws, row, col, "—", bg=bg, color="94A3B8")

def _xl_colwidths(ws, widths: dict[str, float]):
    for letter, w in widths.items():
        ws.column_dimensions[letter].width = w


@st.cache_data(show_spinner=False)
def build_excel_export(df: pd.DataFrame) -> bytes:
    """
    Generate a 6-sheet formatted Excel workbook from the sales DataFrame.
    Returns raw bytes (cached per unique df hash).

    Sheet layout:
      1. Executive Summary  — all-timeframe totals, CY vs LY
      2-6. Latest Week / 4W / 13W / 52W / YTD — item-level detail
    """
    results = compute_all_timeframes(df)
    max_date = df["Date"].max()

    wb = Workbook()
    wb.remove(wb.active)                        # discard default blank sheet

    # ── Sheet 1 : Executive Summary ──────────────────────────────────────────
    ws = wb.create_sheet("Executive Summary")

    _xl_title(ws, "WEEKLY EXECUTIVE SUMMARY", 1, 8, _NAVY, size=13)
    ws.merge_cells("A2:H2")
    c = ws.cell(row=2, column=1,
                value=(f"Period ending: {max_date.strftime('%B %d, %Y')}"
                       "     |     All metrics auto-calculated from uploaded data"))
    c.fill = _xl_fill(_BLUE); c.font = _xl_font(color=_WHITE); c.alignment = _LEFT
    ws.row_dimensions[2].height = 20

    ws.row_dimensions[3].height = 5            # spacer

    _xl_headers(ws, ["Timeframe", "CY Revenue", "LY Revenue", "Rev YoY %",
                     "CY Units", "LY Units", "Units YoY %", "CY Velocity (u/wk)"], 4)

    for i, (tf, label) in enumerate(TIMEFRAMES):
        r   = 5 + i
        m, meta = results[tf]
        tot_cy_u = m["Units_CY"].sum()
        tot_ly_u = m["Units_LY"].sum()
        yoy_r    = (meta["tot_cy"] - meta["tot_ly"]) / meta["tot_ly"] if meta["tot_ly"] else np.nan
        yoy_u    = (tot_cy_u - tot_ly_u) / tot_ly_u                   if tot_ly_u       else np.nan
        bg       = _WHITE if i % 2 == 0 else _LIGHT

        c = ws.cell(row=r, column=1, value=label)
        c.font = _xl_font(bold=True); c.alignment = _LEFT
        c.fill = _xl_fill(bg); c.border = _xl_border()

        _xl_cell(ws, r, 2, meta["tot_cy"],   _FMT_USD, bg=bg)
        _xl_cell(ws, r, 3, meta["tot_ly"],   _FMT_USD, bg=bg)
        _xl_yoy_cell(ws, r, 4, yoy_r, bg)
        _xl_cell(ws, r, 5, int(tot_cy_u),   _FMT_INT, bg=bg)
        _xl_cell(ws, r, 6, int(tot_ly_u),   _FMT_INT, bg=bg)
        _xl_yoy_cell(ws, r, 7, yoy_u, bg)
        _xl_cell(ws, r, 8, tot_cy_u / meta["cy_w"], _FMT_VEL, bg=bg)
        ws.row_dimensions[r].height = 20

    # LY note
    ws.merge_cells("A11:H11")
    c = ws.cell(row=11, column=1,
                value="Note: LY = same rolling window shifted 52 weeks back for consistent day-of-week alignment.")
    c.font = Font(name="Arial", italic=True, color="64748B", size=9)
    c.alignment = _LEFT

    _xl_colwidths(ws, {"A":22,"B":16,"C":16,"D":13,"E":13,"F":13,"G":13,"H":20})
    ws.freeze_panes = "B5"

    # ── Sheets 2-6 : Item Detail per Timeframe ────────────────────────────────
    ITEM_HDRS = ["Item", "CY Revenue", "LY Revenue", "YoY Rev %",
                 "CY Units", "LY Units", "% of Total (CY)",
                 "CY Velocity (u/wk)", "LY Velocity (u/wk)"]

    for tf, label in TIMEFRAMES:
        ws = wb.create_sheet(label)
        m, meta = results[tf]

        # ── title + date range ────────────────────────────────────────────────
        _xl_title(ws, f"ITEM DETAIL — {label.upper()}", 1, 9, _NAVY)

        ws.merge_cells("A2:I2")
        c = ws.cell(row=2, column=1,
                    value=(f"CY: {meta['cy_s'].strftime('%b %d, %Y')} → "
                           f"{meta['cy_e'].strftime('%b %d, %Y')}"
                           f"   |   LY: {meta['ly_s'].strftime('%b %d, %Y')} → "
                           f"{meta['ly_e'].strftime('%b %d, %Y')}"))
        c.fill = _xl_fill(_BLUE); c.font = _xl_font(color="BFDBFE"); c.alignment = _LEFT
        ws.row_dimensions[2].height = 20

        # ── KPI summary bar ───────────────────────────────────────────────────
        tot_cy_u = m["Units_CY"].sum(); tot_ly_u = m["Units_LY"].sum()
        yoy_r    = (meta["tot_cy"] - meta["tot_ly"]) / meta["tot_ly"] if meta["tot_ly"] else np.nan
        ws.merge_cells("A3:I3")
        kpi_text = (f"CY Revenue: ${meta['tot_cy']:,.0f}     |     "
                    f"LY Revenue: ${meta['tot_ly']:,.0f}     |     "
                    f"Rev YoY: {f'{yoy_r:+.1%}' if pd.notna(yoy_r) else '—'}     |     "
                    f"CY Units: {int(tot_cy_u):,}")
        c = ws.cell(row=3, column=1, value=kpi_text)
        c.fill = _xl_fill(_LBLUE); c.font = _xl_font(bold=True, color="1E40AF"); c.alignment = _LEFT
        ws.row_dimensions[3].height = 20

        ws.row_dimensions[4].height = 5           # spacer

        # ── column headers ────────────────────────────────────────────────────
        _xl_headers(ws, ITEM_HDRS, 5)

        # ── data rows ─────────────────────────────────────────────────────────
        display = m.sort_values("Revenue_CY", ascending=False).reset_index(drop=True)

        for i, row_data in display.iterrows():
            r  = 6 + i
            bg = _WHITE if i % 2 == 0 else _LIGHT
            ws.row_dimensions[r].height = 19

            c = ws.cell(row=r, column=1, value=row_data["Item Name"])
            c.font = _xl_font(); c.alignment = _LEFT
            c.fill = _xl_fill(bg); c.border = _xl_border()

            _xl_cell(ws, r, 2, row_data["Revenue_CY"],    _FMT_USD, bg=bg)
            _xl_cell(ws, r, 3, row_data["Revenue_LY"],    _FMT_USD, bg=bg)
            _xl_yoy_cell(ws, r, 4, row_data["YoY_Rev"],   bg)
            _xl_cell(ws, r, 5, int(row_data["Units_CY"]), _FMT_INT, bg=bg)
            _xl_cell(ws, r, 6, int(row_data["Units_LY"]), _FMT_INT, bg=bg)
            _xl_cell(ws, r, 7, row_data["Pct_CY"],        _FMT_PCT, bg=bg)
            _xl_cell(ws, r, 8, row_data["Vel_CY"],        _FMT_VEL, bg=bg)
            _xl_cell(ws, r, 9, row_data["Vel_LY"],        _FMT_VEL, bg=bg)

        # ── TOTAL row ─────────────────────────────────────────────────────────
        tr = 6 + len(display)
        ws.row_dimensions[tr].height = 22
        c = ws.cell(row=tr, column=1, value="TOTAL")
        c.font = _xl_font(bold=True, size=10); c.alignment = _LEFT
        c.fill = _xl_fill(_LGRAY); c.border = _xl_border()

        _xl_cell(ws, tr, 2, meta["tot_cy"],     _FMT_USD, bold=True, bg=_LGRAY)
        _xl_cell(ws, tr, 3, meta["tot_ly"],     _FMT_USD, bold=True, bg=_LGRAY)
        _xl_yoy_cell(ws, tr, 4, yoy_r,          _LGRAY)
        _xl_cell(ws, tr, 5, int(tot_cy_u),      _FMT_INT, bold=True, bg=_LGRAY)
        _xl_cell(ws, tr, 6, int(tot_ly_u),      _FMT_INT, bold=True, bg=_LGRAY)
        _xl_cell(ws, tr, 7, 1.0,                _FMT_PCT, bold=True, bg=_LGRAY)
        _xl_cell(ws, tr, 8, tot_cy_u / meta["cy_w"], _FMT_VEL, bold=True, bg=_LGRAY)
        _xl_cell(ws, tr, 9, tot_ly_u / meta["ly_w"], _FMT_VEL, bold=True, bg=_LGRAY)

        _xl_colwidths(ws, {"A":26,"B":15,"C":15,"D":12,"E":12,"F":12,"G":17,"H":19,"I":19})
        ws.freeze_panes = "B6"

    # ── save ─────────────────────────────────────────────────────────────────
    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    return out.read()                           # return bytes, not BytesIO (cache-safe)


# ══════════════════════════════════════════════════════════════════════════════
# SAMPLE DATA GENERATOR
# ══════════════════════════════════════════════════════════════════════════════

@st.cache_data(show_spinner=False)
def _make_sample_csv() -> bytes:
    rng = np.random.default_rng(42)
    catalog = {
        "Organic Granola Bar":   (3.49, 60),
        "Sparkling Water 12pk":  (8.99, 45),
        "Trail Mix 8oz":         (5.99, 35),
        "Protein Shake Vanilla": (2.79, 80),
        "Green Tea Extract":     (14.99, 20),
    }
    today = pd.Timestamp.today().normalize()
    rows = []
    for day in pd.date_range(today - timedelta(days=730), today, freq="D"):
        for name, (price, avg_u) in catalog.items():
            if rng.random() > 0.22:
                u = max(0, int(rng.normal(avg_u, avg_u * 0.25)))
                rows.append({"Date": day.strftime("%Y-%m-%d"), "Item Name": name,
                             "Sales Revenue": round(u * price, 2), "Units Sold": u})
    return pd.DataFrame(rows).to_csv(index=False).encode()


# ══════════════════════════════════════════════════════════════════════════════
# SIDEBAR
# ══════════════════════════════════════════════════════════════════════════════

with st.sidebar:
    st.markdown("## 📊 Weekly Summary")
    st.caption("Automated retail analytics · no analyst needed")
    st.markdown("---")
    st.markdown("### Upload Sales Data")
    uploaded = st.file_uploader(
        "CSV or Excel",
        type=["csv", "xlsx", "xls"],
        label_visibility="collapsed",
    )
    with st.expander("📥 No data yet? Download a sample"):
        st.caption("Two years of synthetic grocery data — perfect for testing.")
        st.download_button("⬇️ sample_sales.csv", _make_sample_csv(),
                           file_name="sample_sales.csv", mime="text/csv")
    st.markdown("---")
    st.markdown("**Required columns**")
    for col in COLUMN_ALIASES:
        st.markdown(f"• `{col}`")
    st.markdown("---")
    st.markdown("**Timeframes analyzed**")
    for _, label in TIMEFRAMES:
        st.markdown(f"• {label}")
    st.markdown("---")
    st.caption("LY = same rolling window shifted 52 weeks back for day-of-week alignment.")

# ══════════════════════════════════════════════════════════════════════════════
# HEADER
# ══════════════════════════════════════════════════════════════════════════════

st.markdown("""
<div style="background:linear-gradient(135deg,#0f172a 0%,#1a2e6b 55%,#1d4ed8 100%);
            border-radius:14px; padding:28px 34px; margin-bottom:28px;">
  <div style="font-size:29px; font-weight:800; color:#fff; margin-bottom:5px; letter-spacing:-0.3px;">
    📊 Weekly Executive Summary
  </div>
  <div style="font-size:14px; color:#93c5fd; font-weight:400;">
    Upload your sales file → get instant CY vs LY intelligence across 5 timeframes
  </div>
</div>""", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# EMPTY STATE
# ══════════════════════════════════════════════════════════════════════════════

if uploaded is None:
    c1, c2, c3 = st.columns(3)
    _cards = [
        ("📁", "Upload Your Sales File",
         "CSV or Excel with Date, Item, Revenue & Units columns"),
        ("⚡", "Zero Manual Math",
         "Revenue, units, velocity, YoY %, and % of total — all auto-calculated"),
        ("📈", "5 Timeframes × CY & LY",
         "Week / 4W / 13W / 52W / YTD compared against the same window last year"),
    ]
    for col, (icon, title, body) in zip([c1, c2, c3], _cards):
        with col:
            st.markdown(f"""
<div class="empty-card">
  <div style="font-size:36px; margin-bottom:10px">{icon}</div>
  <div style="font-weight:600; font-size:15px; color:#1e293b; margin-bottom:5px">{title}</div>
  <div style="font-size:13px; color:#64748b; line-height:1.5">{body}</div>
</div>""", unsafe_allow_html=True)
    st.markdown("<br>", unsafe_allow_html=True)
    st.info("👈 **Get started:** Upload your sales data in the sidebar, "
            "or download the sample file to see the dashboard in action.")
    st.stop()

# ══════════════════════════════════════════════════════════════════════════════
# LOAD DATA + COMPUTE (cached — fast on every re-render)
# ══════════════════════════════════════════════════════════════════════════════

with st.spinner("Reading your file…"):
    try:
        file_bytes = uploaded.read()
        df = load_file(file_bytes, uploaded.name)
    except ValueError as exc:
        st.error(f"❌ {exc}\n\nPlease check your column names match the required format.")
        st.stop()
    except Exception as exc:
        st.error(f"❌ Could not read the file: {exc}")
        st.stop()

max_date = df["Date"].max()
min_date = df["Date"].min()

# Pre-compute everything upfront (both results + Excel are cached)
results = compute_all_timeframes(df)

# ══════════════════════════════════════════════════════════════════════════════
# INFO BAR  +  EXCEL EXPORT BUTTON
# ══════════════════════════════════════════════════════════════════════════════

ic1, ic2, ic3, ic4 = st.columns(4)
ic1.metric("📅 Latest Date in File",  max_date.strftime("%b %d, %Y"))
ic2.metric("📆 Earliest Date",        min_date.strftime("%b %d, %Y"))
ic3.metric("🏷️ Unique Products",      df["Item Name"].nunique())
ic4.metric("💰 All-Time Revenue",     _usd(df["Sales Revenue"].sum()))

st.markdown("<div style='height:12px'></div>", unsafe_allow_html=True)

# Export button — right-aligned, full report in one click
_, btn_col = st.columns([3, 1])
with btn_col:
    excel_bytes = build_excel_export(df)
    st.download_button(
        label="📥 Export Full Report to Excel",
        data=excel_bytes,
        file_name=f"Weekly_Summary_{max_date.strftime('%Y%m%d')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
        type="primary",
    )

st.divider()

# ══════════════════════════════════════════════════════════════════════════════
# MASTER SUMMARY TABLE
# ══════════════════════════════════════════════════════════════════════════════

st.markdown('<div class="section-title">📋 Master Business Summary</div>', unsafe_allow_html=True)
st.markdown('<div class="section-sub">Total revenue & units across all timeframes — Current Year (CY) vs. Last Year (LY)</div>', unsafe_allow_html=True)

_summary_rows = []
for tf, label in TIMEFRAMES:
    m, meta = results[tf]
    tot_cy_u = m["Units_CY"].sum(); tot_ly_u = m["Units_LY"].sum()
    yoy_r = (meta["tot_cy"] - meta["tot_ly"]) / meta["tot_ly"] if meta["tot_ly"] else np.nan
    yoy_u = (tot_cy_u - tot_ly_u) / tot_ly_u                   if tot_ly_u       else np.nan
    _summary_rows.append({
        "Timeframe":           label,
        "CY Revenue":          _usd(meta["tot_cy"]),
        "LY Revenue":          _usd(meta["tot_ly"]),
        "Rev YoY %":           _yoy(yoy_r),
        "CY Units":            _units(tot_cy_u),
        "LY Units":            _units(tot_ly_u),
        "Units YoY %":         _yoy(yoy_u),
        "CY Velocity (u/wk)":  _vel(tot_cy_u / meta["cy_w"]),
    })

st.dataframe(
    pd.DataFrame(_summary_rows).style.map(
        _style_yoy, subset=["Rev YoY %", "Units YoY %"]
    ),
    use_container_width=True, hide_index=True, height=222
)
st.divider()

# ══════════════════════════════════════════════════════════════════════════════
# ITEM-LEVEL DETAIL — ONE TAB PER TIMEFRAME
# ══════════════════════════════════════════════════════════════════════════════

st.markdown('<div class="section-title">🔍 Item-Level Detail by Timeframe</div>', unsafe_allow_html=True)
st.markdown('<div class="section-sub">Click any column header to sort · Default: ranked by CY Revenue</div>', unsafe_allow_html=True)

_tabs = st.tabs([lbl for _, lbl in TIMEFRAMES])

for (tf, label), tab in zip(TIMEFRAMES, _tabs):
    with tab:
        m, meta = results[tf]

        dc1, dc2 = st.columns(2)
        with dc1:
            st.markdown(
                f'<div class="dr-box">🗓️ <strong>CY:</strong> '
                f'{meta["cy_s"].strftime("%b %d, %Y")} → {meta["cy_e"].strftime("%b %d, %Y")}'
                f'&nbsp;&nbsp;·&nbsp;&nbsp;{meta["cy_w"]:.0f} week{"s" if meta["cy_w"] != 1 else ""}</div>',
                unsafe_allow_html=True)
        with dc2:
            st.markdown(
                f'<div class="dr-box">🗓️ <strong>LY:</strong> '
                f'{meta["ly_s"].strftime("%b %d, %Y")} → {meta["ly_e"].strftime("%b %d, %Y")}'
                f'&nbsp;&nbsp;·&nbsp;&nbsp;{meta["ly_w"]:.0f} week{"s" if meta["ly_w"] != 1 else ""}</div>',
                unsafe_allow_html=True)

        tot_cy_u = m["Units_CY"].sum()
        yoy_r    = (meta["tot_cy"] - meta["tot_ly"]) / meta["tot_ly"] if meta["tot_ly"] else 0.0

        k1, k2, k3, k4 = st.columns(4)
        k1.metric("CY Revenue",    _usd(meta["tot_cy"]))
        k2.metric("LY Revenue",    _usd(meta["tot_ly"]))
        k3.metric("Revenue YoY %", _pct(yoy_r, sign=True),
                  delta=f"{'↑' if yoy_r >= 0 else '↓'} {abs(yoy_r):.1%}" if meta["tot_ly"] else None)
        k4.metric("CY Units Sold", _units(tot_cy_u))

        st.markdown("<br>", unsafe_allow_html=True)

        if m.empty or (meta["tot_cy"] == 0 and meta["tot_ly"] == 0):
            st.warning(f"⚠️ No sales data found in the **{label}** window.")
            continue

        display = m.sort_values("Revenue_CY", ascending=False).reset_index(drop=True)

        tbl = pd.DataFrame({
            "Item":                   display["Item Name"],
            "CY Revenue":             display["Revenue_CY"].apply(_usd),
            "LY Revenue":             display["Revenue_LY"].apply(_usd),
            "YoY Rev %":              display["YoY_Rev"].apply(_yoy),
            "CY Units":               display["Units_CY"].apply(_units),
            "LY Units":               display["Units_LY"].apply(_units),
            "% of Total (CY)":        display["Pct_CY"].apply(_pct),
            "CY Velocity (u/wk)":     display["Vel_CY"].apply(_vel),
            "LY Velocity (u/wk)":     display["Vel_LY"].apply(_vel),
        })

        st.dataframe(
            tbl.style.map(_style_yoy, subset=["YoY Rev %"]),
            use_container_width=True, hide_index=True,
            height=min(80 + 40 * len(tbl), 540)
        )

        if not display.empty and display["Revenue_CY"].iloc[0] > 0:
            top = display.iloc[0]
            st.success(
                f"🏆 **Top Performer — {top['Item Name']}:** "
                f"{_usd(top['Revenue_CY'])} CY revenue · "
                f"{_pct(top['Pct_CY'])} of total · "
                f"{_vel(top['Vel_CY'])} units/week"
            )

        cy_items = display[display["Revenue_CY"] > 0]
        if len(cy_items) > 1:
            bot = cy_items.iloc[-1]
            st.caption(
                f"📉 Lowest CY revenue: **{bot['Item Name']}** — "
                f"{_usd(bot['Revenue_CY'])} · {_pct(bot['Pct_CY'])} of total · "
                f"{_vel(bot['Vel_CY'])} units/week"
            )

# ══════════════════════════════════════════════════════════════════════════════
# VISUAL INSIGHTS — CHARTS
# ══════════════════════════════════════════════════════════════════════════════

st.divider()
st.markdown('<div class="section-title">📊 Visual Insights</div>', unsafe_allow_html=True)
st.markdown('<div class="section-sub">Charts update automatically with your data</div>', unsafe_allow_html=True)

_CHART_BG   = "#ffffff"
_CHART_FONT = dict(family="Inter, Arial", color="#1e293b", size=12)
_CHART_GRID = dict(gridcolor="#f1f5f9", zerolinecolor="#e2e8f0")

# ── Chart 1: CY vs LY Revenue across all timeframes (full width) ──────────────
_tf_labels = [label for _, label in TIMEFRAMES]
_cy_revs   = [results[tf][1]["tot_cy"] for tf, _ in TIMEFRAMES]
_ly_revs   = [results[tf][1]["tot_ly"] for tf, _ in TIMEFRAMES]

fig_trend = go.Figure()
fig_trend.add_bar(
    name="Current Year", x=_tf_labels, y=_cy_revs,
    marker_color="#1d4ed8",
    text=[_usd(v) for v in _cy_revs], textposition="outside",
    textfont=dict(size=11, color="#1e293b"),
)
fig_trend.add_bar(
    name="Last Year", x=_tf_labels, y=_ly_revs,
    marker_color="#93c5fd",
    text=[_usd(v) for v in _ly_revs], textposition="outside",
    textfont=dict(size=11, color="#64748b"),
)
fig_trend.update_layout(
    title=dict(text="Revenue: Current Year vs Last Year — All Timeframes",
               font=dict(size=15, color="#0f172a")),
    barmode="group",
    plot_bgcolor=_CHART_BG, paper_bgcolor=_CHART_BG,
    font=_CHART_FONT,
    yaxis=dict(tickprefix="$", tickformat=",.0f", **_CHART_GRID),
    xaxis=dict(gridcolor="#f1f5f9"),
    legend=dict(orientation="h", yanchor="bottom", y=1.02,
                xanchor="right", x=1, bgcolor="rgba(0,0,0,0)"),
    height=400,
    margin=dict(t=70, b=40, l=70, r=20),
)
st.plotly_chart(fig_trend, use_container_width=True)

# ── Timeframe selector for item-level charts ──────────────────────────────────
st.markdown(
    '<div style="font-size:13px; font-weight:600; color:#475569; margin-bottom:6px;">'
    'Select timeframe for item charts</div>',
    unsafe_allow_html=True,
)
_chart_tf_label = st.selectbox(
    "Timeframe", _tf_labels, index=0,
    label_visibility="collapsed", key="chart_tf_selector"
)
_chart_tf_code  = next(tf for tf, lbl in TIMEFRAMES if lbl == _chart_tf_label)
_cm, _cmeta     = results[_chart_tf_code]
_chart_display  = _cm.sort_values("Revenue_CY", ascending=True).reset_index(drop=True)

ch1, ch2 = st.columns(2)

# ── Chart 2: CY Revenue by Item (horizontal bar) ──────────────────────────────
with ch1:
    fig_items = go.Figure(go.Bar(
        x=_chart_display["Revenue_CY"],
        y=_chart_display["Item Name"],
        orientation="h",
        marker_color="#1d4ed8",
        marker_line_color="#1e40af", marker_line_width=0.5,
        text=[_usd(v) for v in _chart_display["Revenue_CY"]],
        textposition="outside",
        textfont=dict(size=10, color="#1e293b"),
    ))
    fig_items.update_layout(
        title=dict(text=f"CY Revenue by Item — {_chart_tf_label}",
                   font=dict(size=14, color="#0f172a")),
        plot_bgcolor=_CHART_BG, paper_bgcolor=_CHART_BG,
        font=_CHART_FONT,
        xaxis=dict(tickprefix="$", tickformat=",.0f", **_CHART_GRID),
        yaxis=dict(gridcolor="#f1f5f9", automargin=True),
        height=360,
        margin=dict(t=55, b=30, l=10, r=80),
    )
    st.plotly_chart(fig_items, use_container_width=True)

# ── Chart 3: YoY % by Item (horizontal bar, green/red) ────────────────────────
with ch2:
    _yoy_chart = _chart_display.dropna(subset=["YoY_Rev"]).sort_values("YoY_Rev", ascending=True)
    _bar_colors = [
        "#dc2626" if v < -0.005 else
        "#16a34a" if v >  0.005 else
        "#eab308"
        for v in _yoy_chart["YoY_Rev"]
    ]
    fig_yoy = go.Figure(go.Bar(
        x=_yoy_chart["YoY_Rev"] * 100,
        y=_yoy_chart["Item Name"],
        orientation="h",
        marker_color=_bar_colors,
        marker_line_width=0,
        text=[f"{v:+.1f}%" for v in _yoy_chart["YoY_Rev"] * 100],
        textposition="outside",
        textfont=dict(size=10, color="#1e293b"),
    ))
    fig_yoy.add_vline(x=0, line_color="#94a3b8", line_width=1.5, line_dash="dot")
    fig_yoy.update_layout(
        title=dict(text=f"YoY Revenue % by Item — {_chart_tf_label}",
                   font=dict(size=14, color="#0f172a")),
        plot_bgcolor=_CHART_BG, paper_bgcolor=_CHART_BG,
        font=_CHART_FONT,
        xaxis=dict(ticksuffix="%", **_CHART_GRID, automargin=True),
        yaxis=dict(gridcolor="#f1f5f9", automargin=True),
        height=360,
        margin=dict(t=55, b=30, l=10, r=80),
    )
    st.plotly_chart(fig_yoy, use_container_width=True)

# ══════════════════════════════════════════════════════════════════════════════
# FOOTER
# ══════════════════════════════════════════════════════════════════════════════

st.markdown("---")
st.markdown("""
<div style="text-align:center; color:#94a3b8; font-size:12px; padding:10px 0 18px">
  Weekly Executive Summary &nbsp;·&nbsp;
  All metrics auto-calculated from uploaded data &nbsp;·&nbsp;
  LY = same rolling window shifted 52 weeks back for day-of-week alignment
</div>""", unsafe_allow_html=True)
