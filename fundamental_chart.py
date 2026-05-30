"""
Fundamental Chart Viewer — EPS · Gross Margin · ROE · CF/Share
==============================================================
Chạy độc lập : python fundamental_chart.py
Import vào screener: from fundamental_chart import _col, _compute, _build_chart, ChartWorker
"""
import sys
import pandas as pd
import yfinance as yf
import plotly.graph_objects as go
from plotly.subplots import make_subplots

from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget,
    QVBoxLayout, QHBoxLayout, QLineEdit, QPushButton, QLabel, QFrame,
)
from PySide6.QtWebEngineWidgets import QWebEngineView
from PySide6.QtCore import Qt, QThread, Signal, QTimer

# ─────────────────────────────────────────────────────────────────────────────
# Design tokens
# ─────────────────────────────────────────────────────────────────────────────
BG       = "#06080F"
SURFACE  = "#0B0E18"
PANEL    = "#0E1220"
INPUT_BG = "#131928"
BORDER   = "#1A2133"
BORDER2  = "#243044"
BLUE     = "#3D8EF0"
BLUE_HV  = "#2463B4"
TEXT1    = "#DCE4EE"
TEXT2    = "#7A8899"
TEXT3    = "#3D4D60"
GREEN    = "#34C472"
RED      = "#E8483D"
AMBER    = "#E8A93D"

SPINNER = ["⠋", "⠙", "⠹", "⠸", "⠼", "⠴", "⠦", "⠧", "⠇", "⠏"]

# ─────────────────────────────────────────────────────────────────────────────
# HTML placeholders
# ─────────────────────────────────────────────────────────────────────────────
EMPTY_HTML = f"""<!DOCTYPE html><html><head><style>
  * {{ margin:0; padding:0; box-sizing:border-box; }}
  body {{
    background:{BG}; height:100vh;
    display:flex; align-items:center; justify-content:center;
    font-family:'Segoe UI',system-ui,sans-serif; user-select:none;
  }}
  .wrap {{ text-align:center; }}
  .icon  {{ font-size:42px; margin-bottom:24px; opacity:.18; letter-spacing:6px;
             color:{TEXT1}; font-family:'Consolas',monospace; }}
  .title {{ font-size:13px; font-weight:300; letter-spacing:5px;
             color:{TEXT2}; margin-bottom:10px; text-transform:uppercase; }}
  .sub   {{ font-size:11px; color:{TEXT3}; letter-spacing:1px; margin-bottom:36px; }}
  .chips {{ display:flex; gap:10px; justify-content:center; }}
  .chip  {{ font-family:'Consolas',monospace; font-size:11px; color:{TEXT3};
             padding:5px 14px; border:1px solid {BORDER}; border-radius:3px;
             letter-spacing:2px; }}
  .divider {{ width:40px; height:1px; background:{BORDER}; margin:0 auto 28px; }}
</style></head><body>
  <div class="wrap">
    <div class="icon">▦ ▧ ▨</div>
    <div class="title">Equity Research Terminal</div>
    <div class="divider"></div>
    <div class="sub">Enter a ticker symbol above to begin fundamental analysis</div>
    <div class="chips">
      <span class="chip">AAPL</span><span class="chip">NVDA</span>
      <span class="chip">MSFT</span><span class="chip">META</span>
      <span class="chip">TSLA</span>
    </div>
  </div>
</body></html>"""

# Placeholder dùng cho ChartPanel trong screener_gui
CHART_EMPTY_HTML = f"""<!DOCTYPE html><html><head><style>
  * {{ margin:0; padding:0; box-sizing:border-box; }}
  body {{
    background:{BG}; height:100vh;
    display:flex; align-items:center; justify-content:center;
    font-family:'Segoe UI',system-ui,sans-serif; user-select:none;
  }}
  .wrap {{ text-align:center; }}
  .icon {{ font-size:32px; color:{TEXT3}; opacity:.35; margin-bottom:16px; }}
  .msg  {{ color:{TEXT3}; font-size:11px; letter-spacing:2px; }}
  .sub  {{ color:{TEXT3}; font-size:9px; letter-spacing:1px; margin-top:6px; opacity:.5; }}
</style></head><body>
  <div class="wrap">
    <div class="icon">▦</div>
    <div class="msg">Click a row to load chart</div>
    <div class="sub">EPS · Gross Margin · ROE · CF/Share</div>
  </div>
</body></html>"""


# ─────────────────────────────────────────────────────────────────────────────
# Chart engine — shared between standalone GUI and screener_gui
# ─────────────────────────────────────────────────────────────────────────────
def _col(df, *keys):
    """Return first matching column as numeric Series, or empty Series."""
    for k in keys:
        if k in df.columns:
            return pd.to_numeric(df[k], errors="coerce")
    return pd.Series(dtype=float)


def _compute(fin, cf, bal, info, ttm=False):
    """Return (eps, gm, roe, rev) from pre-transposed DataFrames.

    rev = Total Revenue in absolute value (used for Revenue growth panel).
    ttm=True: ROE uses trailing-12-month rolling sum.
    """
    dil_sh = _col(fin, "Diluted Average Shares", "DilutedAverageShares",
                  "Basic Average Shares")

    eps = _col(fin, "Diluted EPS", "Basic EPS")
    eps = eps.round(2) if not eps.empty else (
        (_col(fin, "Net Income", "NetIncome",
              "Net Income Common Stockholders") / dil_sh).round(2)
        if not dil_sh.empty else pd.Series(dtype=float))

    net = _col(cf, "Net Income From Continuing Operations")
    if net.empty:
        net = _col(fin, "Net Income", "NetIncome",
                   "Net Income Common Stockholders",
                   "Net Income From Continuing Operation Net Minority Interest",
                   "Net Income Including Noncontrolling Interests")

    gp  = _col(fin, "Gross Profit", "GrossProfit")
    rev = _col(fin, "Total Revenue", "TotalRevenue", "Operating Revenue")
    gm  = (gp / rev * 100).round(1) if not gp.empty else pd.Series(dtype=float)

    eq  = _col(bal, "Common Stock Equity", "Stockholders Equity", "StockholdersEquity",
               "Total Stockholders Equity", "Total Equity Gross Minority Interest")
    net_roe = net.rolling(4, min_periods=1).sum() if (ttm and not net.empty) else net
    if not eq.empty:
        shift_n = 4 if ttm else 1
        eq_avg     = (eq + eq.shift(shift_n)) / 2
        eq_for_roe = eq_avg.where(eq_avg.notna(), eq)
    else:
        eq_for_roe = eq
    roe = (net_roe / eq_for_roe * 100).round(1) if not eq.empty else pd.Series(dtype=float)

    return eps, gm, roe, rev


def _build_figure(ticker, name, labels, eps, gm, roe, rev,
                  mode="quarterly", height=860, dark_mode=True):
    """Build and return a go.Figure (2×2 grid). Shared by HTML and PNG paths."""
    if dark_mode:
        C_BG    = "#06080F"; C_PANEL = "#0B0E18"; C_GRID = "#151E2E"
        C_ZERO  = "#1F2D40"; C_TEXT  = "#B8C4D0"; C_DIM  = "#3D4D60"
        _hover_bg = "#0F1622"; _hover_border = "#1F2D40"
    else:
        C_BG    = "#FFFFFF"; C_PANEL = "#F8FAFC"; C_GRID = "#E2E8F0"
        C_ZERO  = "#CBD5E0"; C_TEXT  = "#1A202C"; C_DIM  = "#64748B"
        _hover_bg = "#FFFFFF"; _hover_border = "#CBD5E0"
    C_BLUE  = "#4A9EFF"; C_GOLD  = "#E8A93D"; C_GREEN = "#34C472"
    C_ORG   = "#E87C3D"; C_RED   = "#E8483D"; C_TEAL  = "#2EBFA5"

    yoy_periods = 4 if mode == "quarterly" else 1

    def bar_color(val, hi, lo, c_hi, c_lo, c_neg):
        if val is None or pd.isna(val): return C_DIM
        return c_hi if val >= hi else c_lo if val >= lo else c_neg

    def yoy_color(pct):
        if pct is None or pd.isna(pct): return C_DIM
        return C_GREEN if pct >= 0 else C_RED

    def with_yoy(vals, fmt):
        """Return bar text list: 'value\\n+YoY%' for each bar."""
        shifted = vals.shift(yoy_periods)
        pct = ((vals - shifted) / shifted.abs() * 100).round(1)
        texts = []
        for v, p in zip(vals, pct):
            main = fmt(v) if pd.notna(v) else ""
            if pd.notna(p) and main:
                sign = "+" if p >= 0 else ""
                c = C_GREEN if p >= 0 else C_RED
                texts.append(f"{main}<br><span style='font-size:8px;color:{c}'>{sign}{p:.0f}%</span>")
            else:
                texts.append(main)
        return texts

    def fmt_rev(v):
        if not pd.notna(v): return ""
        if abs(v) >= 1e9:  return f"${v/1e9:.1f}B"
        if abs(v) >= 1e6:  return f"${v/1e6:.0f}M"
        return f"${v/1e3:.0f}K"

    fig = make_subplots(
        rows=2, cols=2,
        subplot_titles=(
            "EPS  ·  Earnings Per Share",
            "Gross Margin",
            "Return on Equity  (ROE TTM)" if mode == "quarterly"
                else "Return on Equity  (ROE)",
            "Revenue",
        ),
        vertical_spacing=0.16, horizontal_spacing=0.10,
    )

    txt_size = 10 if height >= 800 else 9

    def panel(row, col, vals, bar_cols, texts):
        fig.add_trace(go.Bar(
            x=labels, y=vals, marker_color=bar_cols, marker_opacity=0.88,
            text=texts, textposition="outside",
            textfont=dict(size=txt_size, color=C_TEXT), showlegend=False,
            hovertemplate="%{text}<extra></extra>",
        ), row=row, col=col)

    panel(1, 1, eps,
          [C_BLUE if (v or 0) >= 0 else C_RED for v in eps.fillna(0)],
          with_yoy(eps, lambda v: f"${v:.2f}"))
    panel(1, 2, gm,
          [bar_color(v, 40, 25, C_GREEN, C_ORG, C_RED) for v in gm],
          with_yoy(gm, lambda v: f"{v:.1f}%"))
    panel(2, 1, roe,
          [bar_color(v, 17, 10, C_BLUE, C_ORG, C_RED) for v in roe],
          with_yoy(roe, lambda v: f"{v:.1f}%"))

    # Revenue panel — color by YoY growth rate
    rev_shifted = rev.shift(yoy_periods)
    rev_pct = ((rev - rev_shifted) / rev_shifted.abs() * 100).round(1)
    rev_colors = [bar_color(p, 20, 5, C_GREEN, C_BLUE, C_RED) for p in rev_pct.fillna(0)]
    panel(2, 2, rev, rev_colors, with_yoy(rev, fmt_rev))

    fig.add_hline(y=40, line_dash="dot", line_color=C_GREEN, line_width=1,
                  row=1, col=2, secondary_y=False,
                  annotation_text="GM 40%",
                  annotation_font=dict(size=8, color=C_GREEN),
                  annotation_position="top right")
    fig.add_hline(y=17, line_dash="dot", line_color=C_BLUE, line_width=1,
                  row=2, col=1, secondary_y=False,
                  annotation_text="ROE 17%",
                  annotation_font=dict(size=8, color=C_BLUE),
                  annotation_position="top right")

    fig.update_layout(
        paper_bgcolor=C_BG, plot_bgcolor=C_PANEL,
        font=dict(family="Segoe UI, system-ui, sans-serif", color=C_DIM, size=10),
        title=dict(
            text=(f'<span style="color:{C_TEXT};font-size:14px;font-weight:600;'
                  f'letter-spacing:1px">{ticker}</span>'
                  f'<span style="color:{C_DIM};font-size:12px;font-weight:300">'
                  f'  ·  {name}</span>'),
            x=0.0, xanchor="left", pad=dict(l=14, t=4),
        ),
        showlegend=False, height=height,
        margin=dict(l=56, r=56, t=52, b=40),
        hovermode="x unified",
        hoverlabel=dict(bgcolor=_hover_bg, bordercolor=_hover_border,
                        font=dict(family="Segoe UI", size=12, color=C_TEXT)),
        bargap=0.28,
    )
    fig.update_xaxes(gridcolor=C_GRID, gridwidth=1, zeroline=False,
                     showline=True, linecolor=C_GRID,
                     tickfont=dict(size=9, color=C_DIM))
    fig.update_yaxes(gridcolor=C_GRID, gridwidth=1,
                     zerolinecolor=C_ZERO, zerolinewidth=1,
                     tickfont=dict(size=9, color=C_DIM))
    for ann in fig.layout.annotations:
        if ann.text and ann.text.startswith(("EPS", "Gross", "Return", "Operating")):
            ann.update(font=dict(size=10, color=C_DIM, family="Segoe UI"))

    mode_note = ("Full-year values" if mode == "annual"
                 else "EPS & GM = single quarter  ·  ROE & CF/Share = TTM (trailing 4 quarters)")
    fig.add_annotation(
        text=f"<span style='color:{C_DIM}'>■ Bars = value  ·  {mode_note}</span>",
        xref="paper", yref="paper", x=0.5, y=-0.05, showarrow=False,
        font=dict(size=9, color=C_DIM, family="Segoe UI"), xanchor="center",
    )
    return fig


def _build_chart(ticker, name, labels, eps, gm, roe, rev,
                 mode="quarterly", height=860, dark_mode=True):
    """Return interactive Plotly HTML string — fills 100% of the WebEngineView viewport."""
    fig = _build_figure(ticker, name, labels, eps, gm, roe, rev,
                        mode=mode, height=height, dark_mode=dark_mode)
    fig.update_layout(height=None, autosize=True)
    div = fig.to_html(
        include_plotlyjs="cdn", full_html=False,
        config={"displayModeBar": True, "displaylogo": False,
                "modeBarButtonsToRemove": ["select2d", "lasso2d"]},
    )
    bg = "#06080F" if dark_mode else "#FFFFFF"
    return (
        f'<!DOCTYPE html><html><head><meta charset="utf-8"><style>'
        f'html,body{{margin:0;padding:0;width:100%;height:100%;overflow:hidden;background:{bg};}}'
        f'.plotly-graph-div{{height:100vh!important;width:100%!important;}}'
        f'</style></head><body>{div}</body></html>'
    )


def _build_chart_png(ticker, name, labels, eps, gm, roe, rev,
                     mode="quarterly", height=680, width=920, dark_mode=True):
    """Return PNG bytes (legacy, requires kaleido)."""
    fig = _build_figure(ticker, name, labels, eps, gm, roe, rev,
                        mode=mode, height=height, dark_mode=dark_mode)
    return fig.to_image(format="png", width=width, height=height, scale=2)


# ─────────────────────────────────────────────────────────────────────────────
# Shared data-fetch helper (used by both worker classes)
# ─────────────────────────────────────────────────────────────────────────────
def _fetch_data_for_chart(ticker, msg_cb=None):
    """Fetch yfinance data and return a dict with all chart-ready series.

    Raises ValueError / any exception on failure so callers can forward to
    the failed signal.
    """
    def _msg(t):
        if msg_cb: msg_cb(t)

    _msg("Fetching market data…")
    tk   = yf.Ticker(ticker)
    info = tk.info
    name = info.get("longName", ticker)

    _msg("Loading quarterly data…")
    q_fin = tk.quarterly_income_stmt

    # Vietnam stocks on yfinance need a ".VN" suffix (e.g. VIC → VIC.VN)
    if (q_fin is None or q_fin.empty) and not ticker.upper().endswith(".VN"):
        _msg(f"Trying {ticker}.VN (Vietnam market)…")
        tk_vn = yf.Ticker(ticker + ".VN")
        q_fin_vn = tk_vn.quarterly_income_stmt
        if q_fin_vn is not None and not q_fin_vn.empty:
            tk    = tk_vn
            info  = tk_vn.info
            name  = info.get("longName", ticker)
            q_fin = q_fin_vn

    if q_fin is None or q_fin.empty:
        raise ValueError(f"No data for '{ticker}' — check the ticker symbol.")
    q_fin = q_fin.T.sort_index()
    q_cf  = tk.quarterly_cashflow.T.sort_index()
    q_bal = tk.quarterly_balance_sheet.T.sort_index()

    q_eps, q_gm, q_roe, q_rev = _compute(q_fin, q_cf, q_bal, info, ttm=True)

    q_dates = []
    for s in (q_eps, q_gm, q_roe, q_rev):
        if not s.empty: q_dates.extend(s.index.tolist())
    q_idx = pd.DatetimeIndex(sorted(set(q_dates))) if q_dates else q_fin.index

    def q_align(s):
        return s.reindex(q_idx) if not s.empty \
               else pd.Series([None]*len(q_idx), index=q_idx)
    q_eps, q_gm, q_roe, q_rev = (
        q_align(q_eps), q_align(q_gm), q_align(q_roe), q_align(q_rev))
    q_labels = [d.strftime("%b %Y") if hasattr(d, "strftime") else str(d)[:7]
                for d in q_idx]

    _msg("Loading annual data…")
    a_labels = []; a_fin = None
    a_eps = a_gm = a_roe = a_rev = pd.Series(dtype=float)
    a_idx = pd.DatetimeIndex([])
    try:
        a_fin = tk.income_stmt
        a_cf  = tk.cashflow
        a_bal = tk.balance_sheet
        if a_fin is not None and not a_fin.empty:
            a_fin = a_fin.T.sort_index()
            a_cf  = a_cf.T.sort_index()  if not a_cf.empty  else pd.DataFrame()
            a_bal = a_bal.T.sort_index() if not a_bal.empty else pd.DataFrame()
            a_eps, a_gm, a_roe, a_rev = _compute(a_fin, a_cf, a_bal, info)
            a_idx = a_fin.index

            def a_align(s):
                return s.reindex(a_idx) if not s.empty \
                       else pd.Series([None]*len(a_idx), index=a_idx)
            a_eps, a_gm, a_roe, a_rev = (
                a_align(a_eps), a_align(a_gm), a_align(a_roe), a_align(a_rev))
            a_labels = [f"FY{d.year}" for d in a_idx]
    except Exception:
        pass

    # Pass 2: improve quarterly ROE accuracy via annual equity interpolation
    try:
        _eq_cols = ["Common Stock Equity", "Stockholders Equity",
                    "StockholdersEquity", "Total Stockholders Equity",
                    "Total Equity Gross Minority Interest"]
        q_eq_raw = _col(q_bal, *_eq_cols)
        a_eq_raw = (_col(a_bal, *_eq_cols)
                    if (a_fin is not None and not a_fin.empty
                        and not a_bal.empty) else pd.Series(dtype=float))

        if not q_eq_raw.empty:
            q_eq_filled = q_eq_raw.copy().astype(float)
            if not a_eq_raw.empty:
                a_eq_clean = a_eq_raw.dropna()
                for qd in q_eq_raw.index:
                    if pd.isna(q_eq_filled.at[qd]):
                        prev_a = a_eq_clean[a_eq_clean.index <= qd]
                        next_a = a_eq_clean[a_eq_clean.index > qd]
                        if len(prev_a) > 0 and len(next_a) > 0:
                            d1, d2 = prev_a.index[-1], next_a.index[0]
                            t = max(0.0, min(1.0,
                                (pd.Timestamp(qd) - pd.Timestamp(d1)).days /
                                max((pd.Timestamp(d2) - pd.Timestamp(d1)).days, 1)))
                            q_eq_filled.at[qd] = (float(prev_a.iloc[-1])
                                                  + t * float(next_a.iloc[0]
                                                              - prev_a.iloc[-1]))
            q_net_raw = _col(q_cf, "Net Income From Continuing Operations")
            if q_net_raw.empty:
                q_net_raw = _col(q_fin, "Net Income", "NetIncome",
                                 "Net Income Common Stockholders")
            if not q_net_raw.empty:
                q_net_ttm = q_net_raw.rolling(4, min_periods=1).sum()
                q_eq_avg  = (q_eq_filled + q_eq_filled.shift(4)) / 2
                q_eq_use  = q_eq_avg.where(q_eq_avg.notna(), q_eq_filled)
                q_roe = (q_net_ttm / q_eq_use * 100).round(1).reindex(q_idx)

        if len(a_idx) > 0 and not a_roe.empty:
            for a_date in a_idx:
                ts_a  = pd.Timestamp(a_date)
                diffs = [(abs((pd.Timestamp(qd) - ts_a).days), qd) for qd in q_idx]
                diff_days, closest_q = min(diffs)
                if diff_days <= 15 and a_date in a_roe.index \
                        and pd.notna(a_roe.at[a_date]):
                    q_roe.at[closest_q] = float(a_roe.at[a_date])
    except Exception:
        pass

    return dict(
        name=name,
        q_labels=q_labels, q_eps=q_eps, q_gm=q_gm, q_roe=q_roe, q_rev=q_rev,
        a_labels=a_labels, a_eps=a_eps, a_gm=a_gm, a_roe=a_roe, a_rev=a_rev,
    )


# ─────────────────────────────────────────────────────────────────────────────
# Background workers
# ─────────────────────────────────────────────────────────────────────────────
class ChartWorker(QThread):
    """Emits interactive HTML strings — used by the standalone viewer and screener panel."""
    done   = Signal(str, str, str)   # html_annual, html_quarterly, summary
    failed = Signal(str)
    msg    = Signal(str)

    def __init__(self, ticker: str, height: int = 860, dark_mode: bool = True):
        super().__init__()
        self.ticker     = ticker.upper()
        self._height    = height
        self._dark_mode = dark_mode

    def run(self):
        try:
            d = _fetch_data_for_chart(self.ticker, self.msg.emit)
            self.msg.emit("Rendering chart…")
            html_q = _build_chart(self.ticker, d["name"], d["q_labels"],
                                  d["q_eps"], d["q_gm"], d["q_roe"], d["q_rev"],
                                  mode="quarterly", height=self._height,
                                  dark_mode=self._dark_mode)
            html_a = (_build_chart(self.ticker, d["name"], d["a_labels"],
                                   d["a_eps"], d["a_gm"], d["a_roe"], d["a_rev"],
                                   mode="annual", height=self._height,
                                   dark_mode=self._dark_mode)
                      if d["a_labels"] else "")
            summary = (f"{self.ticker}  ·  {d['name']}  ·  "
                       f"{len(d['a_labels'])} annual  /  {len(d['q_labels'])} quarterly")
            self.done.emit(html_a, html_q, summary)
        except Exception as e:
            self.failed.emit(str(e))


class ChartWorkerPng(QThread):
    """Emits PNG bytes — used by the screener side panel (no QWebEngineView)."""
    done       = Signal(bytes, bytes, str)   # png_annual, png_quarterly, summary
    data_ready = Signal(object)              # raw data dict (for caching)
    failed     = Signal(str)
    msg        = Signal(str)

    def __init__(self, ticker: str, width: int = 920, height: int = 680,
                 dark_mode: bool = True):
        super().__init__()
        self.ticker     = ticker.upper()
        self._width     = width
        self._height    = height
        self._dark_mode = dark_mode

    def run(self):
        try:
            d = _fetch_data_for_chart(self.ticker, self.msg.emit)
            self.data_ready.emit(d)
            self.msg.emit("Rendering chart…")
            png_q = _build_chart_png(self.ticker, d["name"], d["q_labels"],
                                     d["q_eps"], d["q_gm"], d["q_roe"], d["q_rev"],
                                     mode="quarterly",
                                     height=self._height, width=self._width,
                                     dark_mode=self._dark_mode)
            png_a = (_build_chart_png(self.ticker, d["name"], d["a_labels"],
                                      d["a_eps"], d["a_gm"], d["a_roe"], d["a_rev"],
                                      mode="annual",
                                      height=self._height, width=self._width,
                                      dark_mode=self._dark_mode)
                     if d["a_labels"] else b"")
            summary = (f"{self.ticker}  ·  {d['name']}  ·  "
                       f"{len(d['a_labels'])} annual  /  {len(d['q_labels'])} quarterly")
            self.done.emit(png_a, png_q, summary)
        except Exception as e:
            self.failed.emit(str(e))


class _RenderWorkerPng(QThread):
    """Re-renders PNGs from a cached data dict — no yfinance fetch."""
    done   = Signal(bytes, bytes, str)
    failed = Signal(str)
    msg    = Signal(str)

    def __init__(self, ticker: str, d: dict, width: int = 920, height: int = 680,
                 dark_mode: bool = True):
        super().__init__()
        self.ticker     = ticker
        self._d         = d
        self._width     = width
        self._height    = height
        self._dark_mode = dark_mode

    def run(self):
        try:
            d = self._d
            self.msg.emit("Rendering chart…")
            png_q = _build_chart_png(self.ticker, d["name"], d["q_labels"],
                                     d["q_eps"], d["q_gm"], d["q_roe"], d["q_rev"],
                                     mode="quarterly",
                                     height=self._height, width=self._width,
                                     dark_mode=self._dark_mode)
            png_a = (_build_chart_png(self.ticker, d["name"], d["a_labels"],
                                      d["a_eps"], d["a_gm"], d["a_roe"], d["a_rev"],
                                      mode="annual",
                                      height=self._height, width=self._width,
                                      dark_mode=self._dark_mode)
                     if d["a_labels"] else b"")
            summary = (f"{self.ticker}  ·  {d['name']}  ·  "
                       f"{len(d['a_labels'])} annual  /  {len(d['q_labels'])} quarterly")
            self.done.emit(png_a, png_q, summary)
        except Exception as e:
            self.failed.emit(str(e))


# ─────────────────────────────────────────────────────────────────────────────
# Standalone chart viewer — runs when executed directly
# ─────────────────────────────────────────────────────────────────────────────
class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Fundamental Chart Viewer")
        self.resize(1400, 980)
        self._worker   = None
        self._spin_idx = 0
        self._spin_msg = ""
        self._spin_tmr = QTimer(self)
        self._spin_tmr.timeout.connect(self._tick)
        self._html_a   = ""
        self._html_q   = ""
        self._mode     = "quarterly"

        self.setStyleSheet(f"""
            QMainWindow {{ background:{BG}; }}
            * {{ background:transparent; }}
        """)

        root = QWidget()
        self.setCentralWidget(root)
        vbox = QVBoxLayout(root)
        vbox.setContentsMargins(0, 0, 0, 0)
        vbox.setSpacing(0)

        vbox.addWidget(self._header())
        vbox.addWidget(self._sep())
        vbox.addWidget(self._search_bar())
        vbox.addWidget(self._sep())
        vbox.addWidget(self._mode_bar())
        vbox.addWidget(self._sep())

        self.web = QWebEngineView()
        self.web.setHtml(EMPTY_HTML)
        vbox.addWidget(self.web, stretch=1)

        vbox.addWidget(self._sep())
        vbox.addWidget(self._status_bar())

    def _header(self):
        w = QWidget(); w.setFixedHeight(50)
        w.setStyleSheet(f"background:{SURFACE};")
        h = QHBoxLayout(w); h.setContentsMargins(22, 0, 22, 0)
        brand = QLabel()
        brand.setTextFormat(Qt.RichText)
        brand.setText(
            f'<span style="color:{TEXT1};font-size:14px;font-weight:700;'
            f'letter-spacing:4px">FUNDAMENTAL</span>'
            f'<span style="color:{BLUE};font-size:14px;font-weight:700;'
            f'letter-spacing:4px"> CHART</span>'
        )
        h.addWidget(brand); h.addStretch()
        tag = QLabel("Equity Research Terminal")
        tag.setStyleSheet(f"color:{TEXT3}; font-size:10px; letter-spacing:2px;")
        h.addWidget(tag)
        return w

    def _sep(self):
        f = QFrame(); f.setFrameShape(QFrame.HLine); f.setFixedHeight(1)
        f.setStyleSheet(f"background:{BORDER}; border:none;"); return f

    def _search_bar(self):
        w = QWidget(); w.setFixedHeight(58)
        w.setStyleSheet(f"background:{BG};")
        h = QHBoxLayout(w); h.setContentsMargins(22, 10, 22, 10); h.setSpacing(10)

        lbl = QLabel("TICKER"); lbl.setFixedWidth(50)
        lbl.setStyleSheet(f"color:{TEXT3}; font-size:9px; font-weight:600;"
                          f" letter-spacing:2px; font-family:'Segoe UI',sans-serif;")
        lbl.setAlignment(Qt.AlignVCenter | Qt.AlignRight)
        h.addWidget(lbl)

        self.inp = QLineEdit()
        self.inp.setPlaceholderText("e.g.  AAPL  ·  NVDA  ·  MSFT  ·  TSLA")
        self.inp.setFixedHeight(36); self.inp.setMaxLength(12)
        self.inp.setStyleSheet(f"""
            QLineEdit {{
                background:{INPUT_BG}; color:{TEXT1}; border:1px solid {BORDER};
                border-radius:4px; padding:0 14px; font-size:14px;
                font-family:'Consolas','Courier New',monospace; letter-spacing:3px;
                selection-background-color:{BLUE};
            }}
            QLineEdit:focus {{ border-color:{BLUE}; }}
            QLineEdit:disabled {{ color:{TEXT3}; background:{SURFACE}; border-color:{BORDER}; }}
        """)
        self.inp.returnPressed.connect(self._fetch)
        h.addWidget(self.inp)

        self.btn = QPushButton("ANALYZE")
        self.btn.setFixedSize(100, 36); self.btn.setCursor(Qt.PointingHandCursor)
        self.btn.setStyleSheet(f"""
            QPushButton {{
                background:{BLUE}; color:#FFFFFF; border:none; border-radius:4px;
                font-size:11px; font-weight:700; letter-spacing:2.5px;
                font-family:'Segoe UI',sans-serif;
            }}
            QPushButton:hover   {{ background:{BLUE_HV}; }}
            QPushButton:pressed {{ background:#1A4C9E; }}
            QPushButton:disabled {{ background:{PANEL}; color:{TEXT3}; }}
        """)
        self.btn.clicked.connect(self._fetch)
        h.addWidget(self.btn)
        return w

    def _mode_bar(self):
        w = QWidget(); w.setFixedHeight(38)
        w.setStyleSheet(f"background:{BG};")
        h = QHBoxLayout(w); h.setContentsMargins(22, 0, 22, 0); h.setSpacing(6)
        h.addStretch()

        def _make_btn(label, mode):
            b = QPushButton(label); b.setFixedSize(110, 26)
            b.setCursor(Qt.PointingHandCursor)
            b.clicked.connect(lambda: self._switch_mode(mode)); return b

        self._btn_quarterly = _make_btn("Quarterly", "quarterly")
        self._btn_annual    = _make_btn("Annual",    "annual")
        self._btn_annual.setEnabled(False)
        for b in (self._btn_quarterly, self._btn_annual): h.addWidget(b)
        self._refresh_mode_buttons()
        return w

    def _toggle_style(self, active, enabled=True):
        if not enabled:
            return (f"QPushButton {{ background:transparent; color:{TEXT3};"
                    f" border:1px solid {BORDER}; border-radius:3px;"
                    f" font-size:10px; font-weight:500; letter-spacing:1.5px;"
                    f" font-family:'Segoe UI',sans-serif; }}")
        if active:
            return (f"QPushButton {{ background:{BLUE}; color:#FFFFFF;"
                    f" border:1px solid {BLUE}; border-radius:3px;"
                    f" font-size:10px; font-weight:700; letter-spacing:1.5px;"
                    f" font-family:'Segoe UI',sans-serif; }}")
        return (f"QPushButton {{ background:transparent; color:{TEXT2};"
                f" border:1px solid {BORDER}; border-radius:3px;"
                f" font-size:10px; font-weight:500; letter-spacing:1.5px;"
                f" font-family:'Segoe UI',sans-serif; }}"
                f"QPushButton:hover {{ color:{TEXT1}; border-color:{BORDER2}; }}")

    def _refresh_mode_buttons(self):
        annual_ok = bool(self._html_a)
        self._btn_quarterly.setEnabled(True)
        self._btn_annual.setEnabled(annual_ok)
        self._btn_quarterly.setStyleSheet(self._toggle_style(self._mode == "quarterly"))
        self._btn_annual.setStyleSheet(self._toggle_style(self._mode == "annual", annual_ok))

    def _switch_mode(self, mode):
        if mode == self._mode: return
        html = self._html_a if mode == "annual" else self._html_q
        if not html: return
        self._mode = mode
        self.web.setHtml(html)
        self._refresh_mode_buttons()

    def _status_bar(self):
        w = QWidget(); w.setFixedHeight(28)
        w.setStyleSheet(f"background:{SURFACE};")
        h = QHBoxLayout(w); h.setContentsMargins(22, 0, 22, 0)
        self._dot = QLabel("●"); self._dot.setFixedWidth(14)
        self._dot.setStyleSheet(f"color:{TEXT3}; font-size:7px;")
        h.addWidget(self._dot)
        self.status_lbl = QLabel("Ready  ·  Enter a ticker to begin analysis")
        self.status_lbl.setStyleSheet(
            f"color:{TEXT3}; font-size:11px; font-family:'Segoe UI',sans-serif;")
        h.addWidget(self.status_lbl); h.addStretch()
        ver = QLabel("yfinance  ·  plotly  ·  PySide6")
        ver.setStyleSheet(f"color:{TEXT3}; font-size:9px; letter-spacing:1px;")
        h.addWidget(ver)
        return w

    def _fetch(self):
        ticker = self.inp.text().strip().upper()
        if not ticker:
            self._set_status("Enter a ticker symbol first", AMBER); return
        self.btn.setEnabled(False); self.inp.setEnabled(False)
        self._spin_idx = 0; self._spin_msg = "Connecting…"
        self._spin_tmr.start(80)

        self._worker = ChartWorker(ticker, height=860)
        self._worker.msg.connect(lambda t: setattr(self, "_spin_msg", t))
        self._worker.done.connect(self._on_done)
        self._worker.failed.connect(self._on_err)
        self._worker.start()

    def _tick(self):
        self._spin_idx = (self._spin_idx + 1) % len(SPINNER)
        self._set_status(f"{SPINNER[self._spin_idx]}  {self._spin_msg}", AMBER)

    def _on_done(self, html_a, html_q, summary):
        self._spin_tmr.stop()
        self._html_a = html_a; self._html_q = html_q
        self._mode   = "quarterly"
        self.web.setHtml(html_q)
        self._refresh_mode_buttons()
        self._set_status(f"✓  {summary}", GREEN)
        self.btn.setEnabled(True); self.inp.setEnabled(True)
        self.inp.selectAll(); self.inp.setFocus()

    def _on_err(self, msg):
        self._spin_tmr.stop()
        self._set_status(f"✕  {msg}", RED)
        self.btn.setEnabled(True); self.inp.setEnabled(True)

    def _set_status(self, text, color=TEXT3):
        self.status_lbl.setText(text)
        self.status_lbl.setStyleSheet(
            f"color:{color}; font-size:11px; font-family:'Segoe UI',sans-serif;")
        self._dot.setStyleSheet(f"color:{color}; font-size:7px;")


# ─────────────────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    win = MainWindow()
    win.show()
    sys.exit(app.exec())
