"""
Fundamental Screener — PySide6 GUI
Wraps python screen_top100.py
"""
import sys, os, time, importlib.util
import pandas as pd

from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLineEdit, QPushButton, QLabel, QFrame, QSplitter,
    QTableWidget, QTableWidgetItem, QHeaderView, QAbstractItemView,
    QComboBox, QSpinBox, QCheckBox, QFileDialog,
    QDialog, QTextBrowser, QTabWidget, QProgressBar,
)
from PySide6.QtCore import Qt, QThread, Signal, QTimer, QSortFilterProxyModel
from PySide6.QtGui import QFont, QColor, QKeySequence, QShortcut, QPixmap, QPainter
from PySide6.QtWidgets import QHeaderView
#pyinstaller StockScreener.spec --clean#
# ─────────────────────────────────────────────────────────────────────────────
# Load python screen_top100.py via importlib (filename has a space)
# ─────────────────────────────────────────────────────────────────────────────
_dir  = os.path.dirname(os.path.abspath(__file__))
_spec = importlib.util.spec_from_file_location(
    "screen_top100", os.path.join(_dir, "python screen_top100.py"))
_m = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(_m)

fetch_data             = _m.fetch_data
clean_df               = _m.clean_df
score_canslim          = _m.score_canslim
fetch_moat_yfinance    = _m.fetch_moat_yfinance
fetch_market_direction = _m.fetch_market_direction
_MKT_INDEX          = _m._MKT_INDEX
write_excel            = _m.write_excel
CANSLIM             = _m.CANSLIM
CS_KEYS             = _m.CS_KEYS
N_CS                = _m.N_CS
MOAT_SCORE_STYLE    = _m.MOAT_SCORE_STYLE
MOAT_PROXY_MAP      = _m.MOAT_PROXY_MAP
FETCH_COLS          = _m.FETCH_COLS
RENAME              = _m.RENAME

# ─────────────────────────────────────────────────────────────────────────────
# Design tokens
# ─────────────────────────────────────────────────────────────────────────────
BG       = "#F0F4F9"
SURFACE  = "#FFFFFF"
PANEL    = "#E8EDF5"
INPUT_BG = "#FFFFFF"
BORDER   = "#D1D9E6"
BORDER2  = "#A8B8CC"
BLUE     = "#2563EB"
BLUE_HV  = "#1D4ED8"
TEXT1    = "#1A202C"
TEXT2    = "#4A5568"
TEXT3    = "#94A3B8"
GREEN    = "#16A34A"
RED      = "#DC2626"
AMBER    = "#D97706"
SPINNER  = ["⠋","⠙","⠹","⠸","⠼","⠴","⠦","⠧","⠇","⠏"]

DARK_THEME = dict(
    BG="#06080F", SURFACE="#0B0E18", PANEL="#0E1220",
    INPUT_BG="#131928", BORDER="#1A2133", BORDER2="#243044",
    BLUE="#3D8EF0", BLUE_HV="#2463B4",
    TEXT1="#DCE4EE", TEXT2="#7A8899", TEXT3="#3D4D60",
    GREEN="#34C472", RED="#E8483D", AMBER="#E8A93D",
)
LIGHT_THEME = dict(
    BG="#F0F4F9", SURFACE="#FFFFFF", PANEL="#E8EDF5",
    INPUT_BG="#FFFFFF", BORDER="#D1D9E6", BORDER2="#A8B8CC",
    BLUE="#2563EB", BLUE_HV="#1D4ED8",
    TEXT1="#1A202C", TEXT2="#4A5568", TEXT3="#94A3B8",
    GREEN="#276749", RED="#C53030", AMBER="#975A16",
)

# Signal badge colors  (text, background)
SIGNAL_COLORS = {
    "🟢 STRONG BUY": ("#1A5C2B", "#C6EFCE"),
    "🔵 BUY":        ("#1B3A5C", "#DDEEFF"),
    "🟡 WATCH":      ("#7D6608", "#FFF2CC"),
    "🔴 SKIP":       ("#9C0006", "#FFC7CE"),
}

# Table columns: (header, data-key, width, align)
TABLE_COLS = [
    ("No",          "_no_",         36,  Qt.AlignCenter | Qt.AlignVCenter),
    ("Ticker",      "Ticker",       65,  Qt.AlignLeft   | Qt.AlignVCenter),
    ("Company",     "Tên Công Ty", 175,  Qt.AlignLeft   | Qt.AlignVCenter),
    ("Sector",      "Sector",      105,  Qt.AlignLeft   | Qt.AlignVCenter),
    ("Price($)",    "Price ($)",    70,  Qt.AlignRight  | Qt.AlignVCenter),
    ("MCap($B)",    "MCap ($B)",    72,  Qt.AlignRight  | Qt.AlignVCenter),
    ("Moat",        "Moat Score",  110,  Qt.AlignCenter | Qt.AlignVCenter),
    ("Moat Proxy",  "Moat Proxy",  180,  Qt.AlignLeft   | Qt.AlignVCenter),
    ("EPS Q%",      "EPS Qtr%",     62,  Qt.AlignRight  | Qt.AlignVCenter),
    ("ACC",         "EPS_Acc",      48,  Qt.AlignCenter | Qt.AlignVCenter),
    ("EPS A%",      "EPS Annual%",  62,  Qt.AlignRight  | Qt.AlignVCenter),
    ("Rev Q%",      "Rev Qtr%",     62,  Qt.AlignRight  | Qt.AlignVCenter),
    ("GM%",         "Gross Margin%",60,  Qt.AlignRight  | Qt.AlignVCenter),
    ("ROE%",        "ROE%",         60,  Qt.AlignRight  | Qt.AlignVCenter),
    ("D/E",         "D/E",          52,  Qt.AlignRight  | Qt.AlignVCenter),
    ("P/E",         "P/E",          52,  Qt.AlignRight  | Qt.AlignVCenter),
    ("PEG",         "PEG",          55,  Qt.AlignRight  | Qt.AlignVCenter),
    ("3M%",         "3M%",          52,  Qt.AlignRight  | Qt.AlignVCenter),
    ("1Y%",         "1Y%",          52,  Qt.AlignRight  | Qt.AlignVCenter),
    ("C",           "CS_C",         32,  Qt.AlignCenter | Qt.AlignVCenter),
    ("A",           "CS_A",         32,  Qt.AlignCenter | Qt.AlignVCenter),
    ("S",           "CS_S",         32,  Qt.AlignCenter | Qt.AlignVCenter),
    ("L",           "CS_L",         32,  Qt.AlignCenter | Qt.AlignVCenter),
    ("Q",           "CS_Q",         32,  Qt.AlignCenter | Qt.AlignVCenter),
    ("R",           "CS_R",         32,  Qt.AlignCenter | Qt.AlignVCenter),
    ("M",           "CS_M",         32,  Qt.AlignCenter | Qt.AlignVCenter),
    ("D",           "CS_D",         32,  Qt.AlignCenter | Qt.AlignVCenter),
    ("N",           "CS_N",         32,  Qt.AlignCenter | Qt.AlignVCenter),
    ("MKT",         "CS_MKT",       40,  Qt.AlignCenter | Qt.AlignVCenter),
    ("Score",       "CS_Score",     52,  Qt.AlignCenter | Qt.AlignVCenter),
    ("Conviction",  "Conviction",   70,  Qt.AlignCenter | Qt.AlignVCenter),
    ("Signal",      "CS_Signal",   118,  Qt.AlignCenter | Qt.AlignVCenter),
]

# ─────────────────────────────────────────────────────────────────────────────
# Quality Compounder — criteria, scoring, columns
# ─────────────────────────────────────────────────────────────────────────────
QC_CRITERIA = {
    "ROIC": {"label": "ROIC%",         "field": "ROIC%",         "op": "gt", "thr": 15,  "desc": "ROIC > 15%"},
    "OPGM": {"label": "Op Margin%",    "field": "Op Margin%",    "op": "gt", "thr": 15,  "desc": "Op Mgn > 15%"},
    "GM":   {"label": "Gross Margin%", "field": "Gross Margin%", "op": "gt", "thr": 40,  "desc": "GM > 40%"},
    "FCF":  {"label": "FCF/sh",        "field": "FCF/sh",        "op": "gt", "thr": 0,   "desc": "FCF/sh > 0"},
    "DE":   {"label": "D/E",           "field": "D/E",           "op": "lt", "thr": 1.0, "desc": "D/E < 1.0"},
    "MOAT": {"label": "Moat Score",    "field": "Moat Score",    "op": "in", "thr": None, "desc": "Wide/Narrow Moat"},
}
QC_KEYS = list(QC_CRITERIA.keys())
N_QC    = len(QC_CRITERIA)

QC_SIGNAL_COLORS = {
    "🏆 COMPOUNDER": ("#1A5C2B", "#C6EFCE"),
    "⭐ QUALITY":     ("#1B3A5C", "#DDEEFF"),
    "○ AVERAGE":     ("#7D6608", "#FFF2CC"),
    "✗ WEAK":        ("#9C0006", "#FFC7CE"),
}

QC_COLS = [
    ("No",        "_no_",          36,  Qt.AlignCenter | Qt.AlignVCenter),
    ("Ticker",    "Ticker",        65,  Qt.AlignLeft   | Qt.AlignVCenter),
    ("Company",   "Tên Công Ty",  175,  Qt.AlignLeft   | Qt.AlignVCenter),
    ("Sector",    "Sector",       110,  Qt.AlignLeft   | Qt.AlignVCenter),
    ("Price($)",  "Price ($)",     70,  Qt.AlignRight  | Qt.AlignVCenter),
    ("MCap($B)",  "MCap ($B)",     72,  Qt.AlignRight  | Qt.AlignVCenter),
    ("GM%",       "Gross Margin%", 60,  Qt.AlignRight  | Qt.AlignVCenter),
    ("Op Mgn%",   "Op Margin%",    70,  Qt.AlignRight  | Qt.AlignVCenter),
    ("ROIC%",     "ROIC%",         62,  Qt.AlignRight  | Qt.AlignVCenter),
    ("D/E",       "D/E",           52,  Qt.AlignRight  | Qt.AlignVCenter),
    ("Net Cash",  "Net Cash ($B)", 75,  Qt.AlignRight  | Qt.AlignVCenter),
    ("FCF/sh",    "FCF/sh",        65,  Qt.AlignRight  | Qt.AlignVCenter),
    ("FCF Mgn%",  "FCF_Margin%",   68,  Qt.AlignRight  | Qt.AlignVCenter),
    ("Curr.R",    "Current Ratio", 60,  Qt.AlignRight  | Qt.AlignVCenter),
    ("EV/EBITDA", "EV/EBITDA",     78,  Qt.AlignRight  | Qt.AlignVCenter),
    ("P/E",       "P/E",           52,  Qt.AlignRight  | Qt.AlignVCenter),
    ("1Y%",       "1Y%",           55,  Qt.AlignRight  | Qt.AlignVCenter),
    ("Moat",      "Moat Score",   110,  Qt.AlignCenter | Qt.AlignVCenter),
    ("ROIC",      "QC_ROIC",       42,  Qt.AlignCenter | Qt.AlignVCenter),
    ("OpMgn",     "QC_OPGM",       44,  Qt.AlignCenter | Qt.AlignVCenter),
    ("GM",        "QC_GM",         38,  Qt.AlignCenter | Qt.AlignVCenter),
    ("FCF",       "QC_FCF",        38,  Qt.AlignCenter | Qt.AlignVCenter),
    ("DE",        "QC_DE",         38,  Qt.AlignCenter | Qt.AlignVCenter),
    ("Moat✓",     "QC_MOAT",       45,  Qt.AlignCenter | Qt.AlignVCenter),
    ("Score",     "QC_Score",      55,  Qt.AlignCenter | Qt.AlignVCenter),
    ("Conviction","Conviction",    70,  Qt.AlignCenter | Qt.AlignVCenter),
    ("Quality",   "QC_Signal",    138,  Qt.AlignCenter | Qt.AlignVCenter),
    ("EQ",        "EQ_Badge",     118,  Qt.AlignCenter | Qt.AlignVCenter),
]


def compute_qc_score(row: dict) -> dict:
    """Return QC pass/fail flags + Score + Signal for one stock row dict.

    Criteria (all sourced from TradingView):
      ROIC > 15%   — earns above cost of capital (non-financial)
      ROE  > 12%   — replaces ROIC for Financial Services (leverage cấu trúc)
      Op Mgn > 15% — operational leverage / pricing power
      GM > 40%     — structural margin advantage
      FCF/sh > 0   — actually generating free cash (not just accounting profit)
      D/E < 1.0    — clean balance sheet (skip for Financial Services)
      Moat Wide/Narrow — confirmed competitive advantage
    """
    result = {}
    score  = 0
    moat_good = {"WIDE  ★★★", "NARROW ★★"}
    is_financial = row.get("Sector", "") in ("Financial Services", "Finance")

    for key, field, op, thr in [
        ("QC_ROIC", "ROIC%",         "gt", 15),
        ("QC_OPGM", "Op Margin%",    "gt", 15),
        ("QC_GM",   "Gross Margin%", "gt", 40),
        ("QC_FCF",  "FCF/sh",        "gt", 0),
        ("QC_DE",   "D/E",           "lt", 1.0),
    ]:
        if key == "QC_DE" and is_financial:
            result[key] = None  # D/E không áp dụng cho banks/insurance
            continue
        if key == "QC_ROIC" and is_financial:
            # ROIC bị méo bởi leverage cấu trúc — dùng ROE > 12% thay thế
            val = row.get("ROE%")
            if val is None or (isinstance(val, float) and pd.isna(val)):
                result[key] = None
            else:
                result[key] = val > 12
                if result[key]: score += 1
            continue
        val = row.get(field)
        if val is None or (isinstance(val, float) and pd.isna(val)):
            result[key] = None
        else:
            passed = (val > thr) if op == "gt" else (val < thr)
            result[key] = passed
            if passed:
                score += 1

    moat = row.get("Moat Score")
    if moat is None or (isinstance(moat, float) and pd.isna(moat)):
        result["QC_MOAT"] = None
    else:
        passed = moat in moat_good
        result["QC_MOAT"] = passed
        if passed:
            score += 1

    result["QC_Score"] = score
    if   score >= 5: result["QC_Signal"] = "🏆 COMPOUNDER"
    elif score >= 4: result["QC_Signal"] = "⭐ QUALITY"    # ≥4/6 = 67% (trước là ≥3 = 50%)
    elif score >= 1: result["QC_Signal"] = "○ AVERAGE"
    else:            result["QC_Signal"] = "✗ WEAK"
    return result


from fundamental_chart import ChartWorker
from PySide6.QtWebEngineWidgets import QWebEngineView


# ─────────────────────────────────────────────────────────────────────────────
# Chart panel — Plotly rendered to PNG, displayed as a native QLabel pixmap
# ─────────────────────────────────────────────────────────────────────────────
class ColoredHeader(QHeaderView):
    """Header tự vẽ để hỗ trợ màu nền khác nhau theo nhóm cột."""
    _DEF_BG    = QColor("#0A1628")
    _DEF_FG    = QColor("#7EB8D4")
    _BORDER_B  = QColor("#3D8EF0")
    _BORDER_R  = QColor("#1A2133")

    def __init__(self, parent=None):
        super().__init__(Qt.Horizontal, parent)
        self.setHighlightSections(False)
        self.setSectionsClickable(True)

    def paintSection(self, painter, rect, idx):
        painter.save()
        painter.setClipRect(rect)

        bg_raw = self.model().headerData(idx, Qt.Horizontal, Qt.BackgroundRole)
        fg_raw = self.model().headerData(idx, Qt.Horizontal, Qt.ForegroundRole)
        bg = bg_raw.color() if hasattr(bg_raw, "color") else (bg_raw if isinstance(bg_raw, QColor) else self._DEF_BG)
        fg = fg_raw.color() if hasattr(fg_raw, "color") else (fg_raw if isinstance(fg_raw, QColor) else self._DEF_FG)

        painter.fillRect(rect, bg)
        painter.setPen(self._BORDER_B)
        painter.drawLine(rect.left(), rect.bottom(), rect.right(), rect.bottom())
        painter.setPen(self._BORDER_R)
        painter.drawLine(rect.right(), rect.top(), rect.right(), rect.bottom() - 2)

        # draw sort arrow
        sort_col = self.sortIndicatorSection()
        sort_asc = self.sortIndicatorOrder() == Qt.AscendingOrder
        arrow_w = 0
        if sort_col == idx:
            arrow_w = 10
            ax = rect.right() - arrow_w - 3
            ay = rect.center().y()
            painter.setPen(fg)
            painter.setBrush(fg)
            if sort_asc:
                pts = [(ax, ay + 3), (ax + arrow_w, ay + 3), (ax + arrow_w // 2, ay - 2)]
            else:
                pts = [(ax, ay - 3), (ax + arrow_w, ay - 3), (ax + arrow_w // 2, ay + 2)]
            from PySide6.QtGui import QPolygon
            from PySide6.QtCore import QPoint
            poly = QPolygon([QPoint(x, y) for x, y in pts])
            painter.drawPolygon(poly)

        text = self.model().headerData(idx, Qt.Horizontal, Qt.DisplayRole) or ""
        painter.setPen(fg)
        f = QFont("Segoe UI", 9, QFont.Bold)
        f.setLetterSpacing(QFont.AbsoluteSpacing, 0.8)
        painter.setFont(f)
        text_rect = rect.adjusted(3, 0, -(3 + arrow_w + 4), -2)
        painter.drawText(text_rect, Qt.AlignCenter | Qt.AlignVCenter, str(text))
        painter.restore()


class ChartPanel(QWidget):
    """Chart side panel — uses QWebEngineView + HTML (no kaleido dependency)."""
    def __init__(self, parent=None):
        super().__init__(parent)
        self._dark     = True
        self._html_q   = ""
        self._html_a   = ""
        self._mode     = "quarterly"
        self._worker   = None
        self._ticker   = ""
        self._spin_idx = 0
        self._spin_msg = ""
        self._spin_tmr = QTimer(self)
        self._spin_tmr.timeout.connect(self._tick)
        self.setStyleSheet(f"background:{BG};")

        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)
        self._bar_w    = self._build_bar()
        self._sep_line = self._sep()
        layout.addWidget(self._bar_w)
        layout.addWidget(self._sep_line)

        self._web = QWebEngineView()
        self._web.setHtml(self._placeholder())
        layout.addWidget(self._web, stretch=1)

    def _placeholder(self):
        bg = "#0A1628" if self._dark else "#F0F4F9"
        fg = "#3D4D60" if self._dark else "#94A3B8"
        return (f'<html><body style="background:{bg};margin:0;display:flex;'
                f'align-items:center;justify-content:center;height:100vh;">'
                f'<p style="color:{fg};font-family:Segoe UI,sans-serif;'
                f'font-size:11px;letter-spacing:1px;">Click a row to load chart</p>'
                f'</body></html>')

    def _build_bar(self):
        w = QWidget(); w.setFixedHeight(36)
        w.setStyleSheet(f"background:{SURFACE};")
        h = QHBoxLayout(w); h.setContentsMargins(14, 0, 14, 0); h.setSpacing(8)
        self._chart_lbl = QLabel("—")
        self._chart_lbl.setStyleSheet(
            f"color:{TEXT1}; font-size:13px; font-weight:700;"
            f" font-family:'Consolas',monospace; letter-spacing:2px;")
        h.addWidget(self._chart_lbl); h.addStretch()
        self._chart_status = QLabel("Click a row to load chart")
        self._chart_status.setStyleSheet(f"color:{TEXT3}; font-size:9px; letter-spacing:1px;")
        h.addWidget(self._chart_status); h.addSpacing(10)
        self._btn_q = QPushButton("Quarterly")
        self._btn_a = QPushButton("Annual")
        for btn in (self._btn_q, self._btn_a):
            btn.setFixedSize(84, 22); btn.setCursor(Qt.PointingHandCursor)
        self._btn_q.clicked.connect(lambda: self._switch_mode("quarterly"))
        self._btn_a.clicked.connect(lambda: self._switch_mode("annual"))
        h.addWidget(self._btn_q); h.addWidget(self._btn_a)
        self._refresh_btns()
        return w

    def _sep(self):
        f = QFrame(); f.setFrameShape(QFrame.HLine); f.setFixedHeight(1)
        f.setStyleSheet(f"background:{BORDER}; border:none;"); return f

    def load_ticker(self, ticker: str):
        ticker = ticker.upper()
        if ticker == self._ticker and (self._html_q or self._html_a):
            return
        self._ticker = ticker
        self._html_q = ""; self._html_a = ""
        self._mode   = "quarterly"
        self._chart_lbl.setText(ticker)
        self._web.setHtml(self._placeholder())
        self._refresh_btns()
        if self._worker and self._worker.isRunning():
            self._worker.terminate(); self._worker.wait()
        self._spin_msg = f"Loading {ticker}…"
        self._spin_tmr.start(80)
        self._worker = ChartWorker(ticker, height=580, dark_mode=self._dark)
        self._worker.done.connect(self._on_done)
        self._worker.failed.connect(self._on_failed)
        self._worker.msg.connect(lambda m: setattr(self, "_spin_msg", m))
        self._worker.start()

    def _on_done(self, html_a, html_q, summary):
        self._spin_tmr.stop()
        self._html_q = html_q
        self._html_a = html_a
        self._mode   = "quarterly"
        self._web.setHtml(html_q)
        parts = summary.split("·")
        self._set_status(f"✓  {parts[-1].strip()}" if len(parts) > 1 else "✓", GREEN)
        self._refresh_btns()

    def _on_failed(self, msg):
        self._spin_tmr.stop()
        bg = "#0A1628" if self._dark else "#F0F4F9"
        self._web.setHtml(
            f'<html><body style="background:{bg};margin:0;display:flex;'
            f'align-items:center;justify-content:center;height:100vh;">'
            f'<p style="color:#DC2626;font-family:Segoe UI,sans-serif;'
            f'font-size:11px;">Error: {msg}</p></body></html>')
        self._set_status(f"✕  {msg}", RED)

    def _switch_mode(self, mode):
        if mode == self._mode: return
        html = self._html_a if mode == "annual" else self._html_q
        if not html: return
        self._mode = mode
        self._web.setHtml(html)
        self._refresh_btns()

    def _set_status(self, text, color=TEXT3):
        self._chart_status.setText(text)
        self._chart_status.setStyleSheet(f"color:{color}; font-size:9px; letter-spacing:1px;")

    def _tick(self):
        self._spin_idx = (self._spin_idx + 1) % len(SPINNER)
        self._set_status(f"{SPINNER[self._spin_idx]}  {self._spin_msg}", AMBER)

    def _refresh_btns(self):
        has_q = bool(self._html_q)
        has_a = bool(self._html_a)
        act = (f"QPushButton {{ background:{BLUE}; color:#FFF; border:none;"
               f" border-radius:3px; font-size:9px; font-weight:700;"
               f" letter-spacing:1px; font-family:'Segoe UI',sans-serif; }}")
        off = (f"QPushButton {{ background:transparent; color:{TEXT2};"
               f" border:1px solid {BORDER}; border-radius:3px; font-size:9px;"
               f" font-weight:500; letter-spacing:1px; font-family:'Segoe UI',sans-serif; }}"
               f"QPushButton:hover {{ color:{TEXT1}; border-color:{BORDER2}; }}")
        dis = (f"QPushButton {{ background:transparent; color:{TEXT3};"
               f" border:1px solid {BORDER}; border-radius:3px; font-size:9px;"
               f" font-family:'Segoe UI',sans-serif; }}")
        self._btn_q.setEnabled(has_q)
        self._btn_a.setEnabled(has_a)
        self._btn_q.setStyleSheet(act if (self._mode == "quarterly" and has_q) else off if has_q else dis)
        self._btn_a.setStyleSheet(act if (self._mode == "annual"    and has_a) else off if has_a else dis)

    def set_theme(self, dark: bool):
        self._dark = dark
        self.setStyleSheet(f"background:{BG};")
        self._bar_w.setStyleSheet(f"background:{SURFACE};")
        self._sep_line.setStyleSheet(f"background:{BORDER}; border:none;")
        self._chart_lbl.setStyleSheet(
            f"color:{TEXT1}; font-size:13px; font-weight:700;"
            f" font-family:'Consolas',monospace; letter-spacing:2px;")
        self._chart_status.setStyleSheet(f"color:{TEXT3}; font-size:9px; letter-spacing:1px;")
        self._refresh_btns()
        if self._ticker:
            prev = self._ticker; self._ticker = ""
            self.load_ticker(prev)


# ─────────────────────────────────────────────────────────────────────────────
# Custom table item with numeric sort
# ─────────────────────────────────────────────────────────────────────────────
class NumItem(QTableWidgetItem):
    def __init__(self, val, fmt=""):
        self._num = float(val) if val is not None and val == val else float("-inf")
        if val is None or (isinstance(val, float) and val != val):
            text = "—"
        elif fmt == "pct":
            text = f"{val:+.1f}%" if val != 0 else "—"
        elif fmt == "price":
            text = f"${val:,.2f}"
        elif fmt == "vnd":
            text = f"₫{val:,.0f}"
        elif fmt == "mcap":
            text = f"{val:,.1f}"
        elif fmt == "mcap_vnd":
            text = f"₫{val:,.0f} tỷ"
        elif fmt == "int":
            text = str(int(val))
        else:
            text = f"{val:.2f}" if isinstance(val, float) else str(val)
        super().__init__(text)

    def __lt__(self, other):
        if isinstance(other, NumItem):
            return self._num < other._num
        return super().__lt__(other)


# ─────────────────────────────────────────────────────────────────────────────
# Background workers
# ─────────────────────────────────────────────────────────────────────────────
class ScanWorker(QThread):
    progress      = Signal(str)
    ticker_update = Signal(int, int, str, str)   # idx, total, ticker, score
    done          = Signal(object)               # DataFrame
    failed        = Signal(str)

    def __init__(self, market, top, use_yf):
        super().__init__()
        self.market  = market
        self.top     = top
        self.use_yf  = use_yf
        self._cancel = False

    def cancel(self): self._cancel = True

    def run(self):
        try:
            from tradingview_screener import Query
            self.progress.emit(f"Querying TradingView — {self.market.upper()} top {self.top}…")
            _, raw = (Query().set_markets(self.market.lower())
                      .select(*FETCH_COLS)
                      .order_by("market_cap_basic", ascending=False)
                      .limit(self.top)
                      .get_scanner_data())
            df = clean_df(raw)
            self.progress.emit(f"Got {len(df)} stocks.")

            _US_MARKETS = {"america", "nasdaq", "nyse"}
            moat_cache = None
            if self.use_yf and self.market.lower() in _US_MARKETS:
                tickers = df["Ticker"].tolist()
                sectors = dict(zip(df["Ticker"], df["Sector"].fillna("")))
                moat_cache = {}
                total = len(tickers)
                for i, ticker in enumerate(tickers, 1):
                    if self._cancel:
                        self.failed.emit("Cancelled.")
                        return
                    sector = sectors.get(ticker, "")
                    proxy, score, w52_pct, eq_badge, eps_acc, fcf_margin, net_cash = fetch_moat_yfinance(ticker, sector)
                    moat_cache[ticker] = (proxy, score, w52_pct, eq_badge, eps_acc, fcf_margin, net_cash)
                    self.ticker_update.emit(i, total, ticker, score)
                    time.sleep(0.3)

            self.progress.emit(f"Fetching market direction ({_MKT_INDEX.get(self.market.lower(), ('','SP:SPX'))[1]})…")
            market_ok, index_1y = fetch_market_direction(self.market)
            self.progress.emit("Applying CAN SLIM scoring…")
            df = score_canslim(df, moat_cache, market_ok=market_ok, index_1y=index_1y)
            self.done.emit(df)
        except Exception as e:
            self.failed.emit(str(e))


class TickerWorker(QThread):
    done   = Signal(object)   # single-row DataFrame
    failed = Signal(str)

    def __init__(self, ticker, market="america"):
        super().__init__()
        self.ticker = ticker.upper()
        self.market = market

    def run(self):
        try:
            from tradingview_screener import Query, Column
            _, raw = (Query().set_markets(self.market.lower())
                      .select(*FETCH_COLS)
                      .where(Column("name") == self.ticker)
                      .get_scanner_data())
            if raw.empty:
                self.failed.emit(f"'{self.ticker}' not found on TradingView.")
                return
            df = clean_df(raw)
            proxy, score, w52_pct, eq_badge, eps_acc, fcf_margin, net_cash = fetch_moat_yfinance(
                self.ticker, df["Sector"].iloc[0] if "Sector" in df.columns else "")
            moat_cache = {self.ticker: (proxy, score, w52_pct, eq_badge, eps_acc, fcf_margin, net_cash)}
            market_ok, index_1y = fetch_market_direction(self.market)
            df = score_canslim(df, moat_cache, market_ok=market_ok, index_1y=index_1y)
            self.done.emit(df)
        except Exception as e:
            self.failed.emit(str(e))


# ─────────────────────────────────────────────────────────────────────────────
# Detail card — shows CAN SLIM breakdown for one stock
# ─────────────────────────────────────────────────────────────────────────────
class DetailCard(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setStyleSheet(f"background:{PANEL};")
        layout = QVBoxLayout(self)
        layout.setContentsMargins(18, 10, 18, 10)
        layout.setSpacing(8)

        # Top info row
        self._lbl_ticker = QLabel("—")
        self._lbl_ticker.setStyleSheet(
            f"color:{TEXT1}; font-size:20px; font-weight:700;"
            f" font-family:'Consolas',monospace; letter-spacing:3px;")
        self._lbl_info = QLabel("")
        self._lbl_info.setWordWrap(True)
        self._lbl_info.setStyleSheet(
            f"color:{TEXT1}; font-size:13px; font-family:'Segoe UI',sans-serif;")
        self._lbl_moat = QLabel("")
        self._lbl_moat.setStyleSheet(
            f"font-size:12px; font-weight:700; font-family:'Segoe UI',sans-serif;"
            f" padding:3px 12px; border-radius:3px;")
        self._lbl_signal = QLabel("")
        self._lbl_signal.setStyleSheet(
            f"font-size:12px; font-weight:700; font-family:'Segoe UI',sans-serif;"
            f" padding:3px 12px; border-radius:3px;")

        top_row = QHBoxLayout()
        top_row.addWidget(self._lbl_ticker)
        top_row.addWidget(self._lbl_info)
        top_row.addStretch()
        top_row.addWidget(self._lbl_moat)
        top_row.addWidget(self._lbl_signal)
        layout.addLayout(top_row)

        # CAN SLIM criteria row
        cs_widget = QWidget()
        self._cs_widget = cs_widget
        cs_widget.setStyleSheet(f"background:{SURFACE}; border-radius:4px;")
        cs_layout = QHBoxLayout(cs_widget)
        cs_layout.setContentsMargins(12, 8, 12, 8)
        cs_layout.setSpacing(6)

        self._cs_labels = {}
        for key in CS_KEYS:
            cfg = CANSLIM[key]
            col = QVBoxLayout()
            col.setSpacing(3)
            lbl_key = QLabel(key)
            lbl_key.setAlignment(Qt.AlignCenter)
            lbl_key.setStyleSheet(
                f"color:{TEXT2}; font-size:10px; font-weight:600;"
                f" letter-spacing:1px; font-family:'Segoe UI',sans-serif;")
            lbl_val = QLabel("—")
            lbl_val.setAlignment(Qt.AlignCenter)
            lbl_val.setFixedSize(54, 28)
            lbl_val.setStyleSheet(
                f"color:{TEXT1}; font-size:13px; font-weight:700;"
                f" background:{PANEL}; border-radius:3px; font-family:'Consolas',monospace;")
            col.addWidget(lbl_key)
            col.addWidget(lbl_val)
            cs_layout.addLayout(col)
            self._cs_labels[key] = lbl_val

        cs_layout.addStretch()

        # Score badge
        self._lbl_score = QLabel("—/8")
        self._lbl_score.setStyleSheet(
            f"color:{TEXT1}; font-size:18px; font-weight:700;"
            f" font-family:'Consolas',monospace;")
        cs_layout.addWidget(self._lbl_score)
        layout.addWidget(cs_widget)

        self._lbl_empty = QLabel("Click a row or look up a ticker to see details")
        self._lbl_empty.setAlignment(Qt.AlignCenter)
        self._lbl_empty.setStyleSheet(
            f"color:{TEXT3}; font-size:12px; font-style:italic;"
            f" font-family:'Segoe UI',sans-serif;")
        layout.addWidget(self._lbl_empty)

    def show_row(self, row: dict):
        self._lbl_empty.hide()
        ticker = row.get("Ticker", "—")
        name   = row.get("Tên Công Ty", "")
        sector = row.get("Sector", "")
        price  = row.get("Price ($)")
        mcap   = row.get("MCap ($B)")

        self._lbl_ticker.setText(ticker)
        parts = [p for p in [name, sector] if p]
        price_str = f"  ${price:,.2f}" if isinstance(price, (int, float)) else ""
        mcap_str  = f"  MCap ${mcap:,.1f}B" if isinstance(mcap, (int, float)) else ""
        self._lbl_info.setText("  ·  ".join(parts) + price_str + mcap_str)

        moat = row.get("Moat Score", "")
        if moat in MOAT_SCORE_STYLE:
            fg, bg = MOAT_SCORE_STYLE[moat]
            self._lbl_moat.setText(f"🏰 {moat}")
            self._lbl_moat.setStyleSheet(
                f"color:#{fg}; background:#{bg}; font-size:10px; font-weight:700;"
                f" font-family:'Segoe UI',sans-serif; padding:2px 10px; border-radius:3px;")
        else:
            self._lbl_moat.setText("")

        sig = row.get("CS_Signal", "")
        if sig in SIGNAL_COLORS:
            fg, bg = SIGNAL_COLORS[sig]
            self._lbl_signal.setText(sig)
            self._lbl_signal.setStyleSheet(
                f"color:{fg}; background:{bg}; font-size:10px; font-weight:700;"
                f" font-family:'Segoe UI',sans-serif; padding:2px 10px; border-radius:3px;")
        else:
            self._lbl_signal.setText("")

        score = row.get("CS_Score", 0)
        self._lbl_score.setText(f"{score}/{N_CS}")
        pct = score / N_CS
        sc = GREEN if pct >= 0.875 else BLUE if pct >= 0.625 else AMBER if pct >= 0.375 else RED
        self._lbl_score.setStyleSheet(
            f"color:{sc}; font-size:14px; font-weight:700; font-family:'Consolas',monospace;")

        for key in CS_KEYS:
            cfg = CANSLIM[key]
            val = row.get(cfg["field"])
            passed = row.get(f"CS_{key}")
            lbl = self._cs_labels[key]
            if passed is True:
                display = f"+{val:.0f}%" if isinstance(val, float) and cfg["field"].endswith("%") else f"{val:.1f}"
                lbl.setText("✓")
                lbl.setStyleSheet(
                    f"color:#1A5C2B; font-size:15px; font-weight:700;"
                    f" background:#C6EFCE; border-radius:3px; font-family:'Consolas',monospace;")
                lbl.setToolTip(f"{cfg['label']}: {display}")
            elif passed is False:
                display = f"{val:+.0f}%" if isinstance(val, float) and cfg["field"].endswith("%") else f"{val:.1f}" if isinstance(val, float) else "—"
                lbl.setText("✗")
                lbl.setStyleSheet(
                    f"color:#9C0006; font-size:15px; font-weight:700;"
                    f" background:#FFC7CE; border-radius:3px; font-family:'Consolas',monospace;")
                lbl.setToolTip(f"{cfg['label']}: {display}")
            else:
                lbl.setText("—")
                lbl.setStyleSheet(
                    f"color:{TEXT2}; font-size:13px; font-weight:700;"
                    f" background:{PANEL}; border-radius:3px; font-family:'Consolas',monospace;")
                lbl.setToolTip(f"{cfg['label']}: N/A")

    def clear(self):
        self._lbl_ticker.setText("—")
        self._lbl_info.setText("")
        self._lbl_moat.setText("")
        self._lbl_signal.setText("")
        self._lbl_score.setText("—/8")
        for lbl in self._cs_labels.values():
            lbl.setText("—")
            lbl.setStyleSheet(
                f"color:{TEXT1}; font-size:13px; font-weight:700;"
                f" background:{PANEL}; border-radius:3px; font-family:'Consolas',monospace;")
        self._lbl_empty.show()

    def apply_theme(self):
        self.setStyleSheet(f"background:{PANEL};")
        self._cs_widget.setStyleSheet(f"background:{SURFACE}; border-radius:4px;")
        self._lbl_ticker.setStyleSheet(
            f"color:{TEXT1}; font-size:20px; font-weight:700;"
            f" font-family:'Consolas',monospace; letter-spacing:3px;")
        self._lbl_info.setStyleSheet(
            f"color:{TEXT1}; font-size:13px; font-family:'Segoe UI',sans-serif;")
        self._lbl_empty.setStyleSheet(
            f"color:{TEXT3}; font-size:12px; font-style:italic;"
            f" font-family:'Segoe UI',sans-serif;")
        for lbl in self._cs_labels.values():
            # Update tất cả labels không phân biệt text — đảm bảo theme luôn đúng
            if lbl.text() in ("—", "✓", "✗"):
                lbl.setStyleSheet(
                    f"color:{TEXT1}; font-size:13px; font-weight:700;"
                    f" background:{PANEL}; border-radius:3px; font-family:'Consolas',monospace;")


# ─────────────────────────────────────────────────────────────────────────────
# Quality Compounder detail card
# ─────────────────────────────────────────────────────────────────────────────
class QualityDetailCard(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setStyleSheet(f"background:{PANEL};")
        layout = QVBoxLayout(self)
        layout.setContentsMargins(18, 10, 18, 10)
        layout.setSpacing(8)

        self._lbl_ticker = QLabel("—")
        self._lbl_ticker.setStyleSheet(
            f"color:{TEXT1}; font-size:20px; font-weight:700;"
            f" font-family:'Consolas',monospace; letter-spacing:3px;")
        self._lbl_info = QLabel("")
        self._lbl_info.setWordWrap(True)
        self._lbl_info.setStyleSheet(
            f"color:{TEXT1}; font-size:13px; font-family:'Segoe UI',sans-serif;")
        self._lbl_moat = QLabel("")
        self._lbl_moat.setStyleSheet(
            f"font-size:12px; font-weight:700; font-family:'Segoe UI',sans-serif;"
            f" padding:3px 12px; border-radius:3px;")
        self._lbl_signal = QLabel("")
        self._lbl_signal.setStyleSheet(
            f"font-size:12px; font-weight:700; font-family:'Segoe UI',sans-serif;"
            f" padding:3px 12px; border-radius:3px;")
        self._lbl_eq = QLabel("")
        self._lbl_eq.setStyleSheet(
            f"font-size:10px; font-weight:700; font-family:'Segoe UI',sans-serif;"
            f" padding:2px 8px; border-radius:3px;")
        self._lbl_eq.hide()

        top_row = QHBoxLayout()
        top_row.addWidget(self._lbl_ticker)
        top_row.addWidget(self._lbl_info)
        top_row.addStretch()
        top_row.addWidget(self._lbl_moat)
        top_row.addWidget(self._lbl_signal)
        top_row.addWidget(self._lbl_eq)
        layout.addLayout(top_row)

        qc_widget = QWidget()
        self._qc_widget = qc_widget
        qc_widget.setStyleSheet(f"background:{SURFACE}; border-radius:4px;")
        qc_layout = QHBoxLayout(qc_widget)
        qc_layout.setContentsMargins(12, 8, 12, 8)
        qc_layout.setSpacing(6)

        self._qc_labels = {}
        for key in QC_KEYS:
            cfg = QC_CRITERIA[key]
            col = QVBoxLayout()
            col.setSpacing(3)
            lbl_key = QLabel(cfg["desc"])
            lbl_key.setAlignment(Qt.AlignCenter)
            lbl_key.setStyleSheet(
                f"color:{TEXT2}; font-size:9px; font-weight:600;"
                f" letter-spacing:0.5px; font-family:'Segoe UI',sans-serif;")
            lbl_val = QLabel("—")
            lbl_val.setAlignment(Qt.AlignCenter)
            lbl_val.setFixedSize(62, 28)
            lbl_val.setStyleSheet(
                f"color:{TEXT1}; font-size:13px; font-weight:700;"
                f" background:{PANEL}; border-radius:3px; font-family:'Consolas',monospace;")
            col.addWidget(lbl_key)
            col.addWidget(lbl_val)
            qc_layout.addLayout(col)
            self._qc_labels[key] = lbl_val

        qc_layout.addStretch()
        self._lbl_score = QLabel(f"—/{N_QC}")
        self._lbl_score.setStyleSheet(
            f"color:{TEXT1}; font-size:18px; font-weight:700;"
            f" font-family:'Consolas',monospace;")
        qc_layout.addWidget(self._lbl_score)
        layout.addWidget(qc_widget)

        self._lbl_empty = QLabel("Click a row to see Quality Compounder breakdown")
        self._lbl_empty.setAlignment(Qt.AlignCenter)
        self._lbl_empty.setStyleSheet(
            f"color:{TEXT3}; font-size:12px; font-style:italic;"
            f" font-family:'Segoe UI',sans-serif;")
        layout.addWidget(self._lbl_empty)

    def show_row(self, row: dict):
        self._lbl_empty.hide()
        ticker = row.get("Ticker", "—")
        name   = row.get("Tên Công Ty", "")
        sector = row.get("Sector", "")
        price  = row.get("Price ($)")
        mcap   = row.get("MCap ($B)")

        self._lbl_ticker.setText(ticker)
        parts = [p for p in [name, sector] if p]
        price_str = f"  ${price:,.2f}" if isinstance(price, (int, float)) else ""
        mcap_str  = f"  MCap ${mcap:,.1f}B" if isinstance(mcap, (int, float)) else ""
        self._lbl_info.setText("  ·  ".join(parts) + price_str + mcap_str)

        moat = row.get("Moat Score", "")
        if moat in MOAT_SCORE_STYLE:
            fg, bg = MOAT_SCORE_STYLE[moat]
            self._lbl_moat.setText(f"🏰 {moat}")
            self._lbl_moat.setStyleSheet(
                f"color:#{fg}; background:#{bg}; font-size:10px; font-weight:700;"
                f" font-family:'Segoe UI',sans-serif; padding:2px 10px; border-radius:3px;")
        else:
            self._lbl_moat.setText("")

        sig = row.get("QC_Signal", "")
        if sig in QC_SIGNAL_COLORS:
            fg, bg = QC_SIGNAL_COLORS[sig]
            self._lbl_signal.setText(sig)
            self._lbl_signal.setStyleSheet(
                f"color:{fg}; background:{bg}; font-size:10px; font-weight:700;"
                f" font-family:'Segoe UI',sans-serif; padding:2px 10px; border-radius:3px;")
        else:
            self._lbl_signal.setText("")

        eq_badge = row.get("EQ_Badge")
        _eq_colors = {
            "💚 Cash Backed":   ("#1A5C2B", "#C6EFCE"),
            "🟡 Mixed":         ("#7D6608", "#FFF2CC"),
            "🔴 Accrual Heavy": ("#9C0006", "#FFC7CE"),
        }
        if eq_badge and eq_badge in _eq_colors:
            efg, ebg = _eq_colors[eq_badge]
            self._lbl_eq.setText(eq_badge)
            self._lbl_eq.setStyleSheet(
                f"color:{efg}; background:{ebg}; font-size:10px; font-weight:700;"
                f" font-family:'Segoe UI',sans-serif; padding:2px 8px; border-radius:3px;")
            self._lbl_eq.show()
        else:
            self._lbl_eq.setText("")
            self._lbl_eq.hide()

        score = row.get("QC_Score", 0)
        self._lbl_score.setText(f"{score}/{N_QC}")
        pct = score / N_QC
        sc = GREEN if pct >= 5/6 else BLUE if pct >= 3/6 else AMBER if pct >= 1/6 else RED
        self._lbl_score.setStyleSheet(
            f"color:{sc}; font-size:14px; font-weight:700; font-family:'Consolas',monospace;")

        for key in QC_KEYS:
            cfg   = QC_CRITERIA[key]
            passed = row.get(f"QC_{key}")
            val   = row.get(cfg["field"])
            lbl   = self._qc_labels[key]
            if passed is True:
                lbl.setText("✓")
                lbl.setStyleSheet(
                    f"color:#1A5C2B; font-size:15px; font-weight:700;"
                    f" background:#C6EFCE; border-radius:3px; font-family:'Consolas',monospace;")
                tip = f"{val:.1f}" if isinstance(val, float) else str(val)
                lbl.setToolTip(f"{cfg['desc']}: {tip}")
            elif passed is False:
                lbl.setText("✗")
                lbl.setStyleSheet(
                    f"color:#9C0006; font-size:15px; font-weight:700;"
                    f" background:#FFC7CE; border-radius:3px; font-family:'Consolas',monospace;")
                tip = f"{val:.1f}" if isinstance(val, float) else str(val) if val is not None else "—"
                lbl.setToolTip(f"{cfg['desc']}: {tip}")
            else:
                lbl.setText("—")
                lbl.setStyleSheet(
                    f"color:{TEXT2}; font-size:13px; font-weight:700;"
                    f" background:{PANEL}; border-radius:3px; font-family:'Consolas',monospace;")
                lbl.setToolTip(f"{cfg['desc']}: N/A")

    def clear(self):
        self._lbl_ticker.setText("—")
        self._lbl_info.setText("")
        self._lbl_moat.setText("")
        self._lbl_signal.setText("")
        self._lbl_eq.setText("")
        self._lbl_eq.hide()
        self._lbl_score.setText(f"—/{N_QC}")
        for lbl in self._qc_labels.values():
            lbl.setText("—")
            lbl.setStyleSheet(
                f"color:{TEXT1}; font-size:13px; font-weight:700;"
                f" background:{PANEL}; border-radius:3px; font-family:'Consolas',monospace;")
        self._lbl_empty.show()

    def apply_theme(self):
        self.setStyleSheet(f"background:{PANEL};")
        self._qc_widget.setStyleSheet(f"background:{SURFACE}; border-radius:4px;")
        self._lbl_ticker.setStyleSheet(
            f"color:{TEXT1}; font-size:20px; font-weight:700;"
            f" font-family:'Consolas',monospace; letter-spacing:3px;")
        self._lbl_info.setStyleSheet(
            f"color:{TEXT1}; font-size:13px; font-family:'Segoe UI',sans-serif;")
        self._lbl_empty.setStyleSheet(
            f"color:{TEXT3}; font-size:12px; font-style:italic;"
            f" font-family:'Segoe UI',sans-serif;")
        for lbl in self._qc_labels.values():
            if lbl.text() in ("—", "✓", "✗"):
                lbl.setStyleSheet(
                    f"color:{TEXT1}; font-size:13px; font-weight:700;"
                    f" background:{PANEL}; border-radius:3px; font-family:'Consolas',monospace;")


# ─────────────────────────────────────────────────────────────────────────────
# Help dialog — F1
# ─────────────────────────────────────────────────────────────────────────────
class HelpDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Help  ·  F1")
        self.resize(800, 660)
        self.setStyleSheet(f"""
            QDialog      {{ background:{BG}; }}
            *            {{ background:transparent; color:{TEXT1}; }}
            QTextBrowser {{
                background:{SURFACE}; color:{TEXT1};
                border:none; padding:20px 26px;
                font-family:'Segoe UI',sans-serif;
            }}
            QPushButton  {{
                background:{BLUE}; color:#FFFFFF;
                border:none; border-radius:3px;
                font-size:10px; font-weight:700; letter-spacing:1.5px;
                font-family:'Segoe UI',sans-serif;
            }}
            QPushButton:hover {{ background:{BLUE_HV}; }}
        """)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)

        browser = QTextBrowser()
        browser.setOpenExternalLinks(False)
        browser.setHtml(self._html())
        layout.addWidget(browser, stretch=1)

        foot = QWidget()
        foot.setFixedHeight(44)
        foot.setStyleSheet(f"background:{SURFACE}; border-top:1px solid {BORDER};")
        fh = QHBoxLayout(foot)
        fh.setContentsMargins(18, 0, 18, 0)
        hint = QLabel("F1 / Esc  to close")
        hint.setStyleSheet(f"color:{TEXT3}; font-size:9px; letter-spacing:1px;")
        btn = QPushButton("CLOSE")
        btn.setFixedSize(80, 28)
        btn.setCursor(Qt.PointingHandCursor)
        btn.clicked.connect(self.close)
        fh.addWidget(hint); fh.addStretch(); fh.addWidget(btn)
        layout.addWidget(foot)

    def keyPressEvent(self, event):
        if event.key() in (Qt.Key_Escape, Qt.Key_F1):
            self.close()
        else:
            super().keyPressEvent(event)

    def _html(self):  # noqa: C901
        ops = {"gt": ">", "lt": "<"}
        # ── Light-theme palette ───────────────────────────────────────────────
        H_ROW_E = "#F0F4FA"; H_ROW_O = "#FFFFFF"
        H_TEXT  = "#1A202C"; H_DIM   = "#64748B"; H_SEC  = "#374151"
        H_BLUE  = "#1E40AF"; H_GREEN = "#16A34A"; H_WARN = "#D97706"
        H_HDR_BG = "#E2E8F0"; H_HDR_FG = "#2D3748"
        H_CODE_BG = "#EBF4FF"

        full_desc = {
            "C": (f"<b>Chỉ số:</b> EPS Quý YoY% — lợi nhuận mỗi cổ phiếu so cùng quý năm trước.<br>"
                  f"<b>Ngưỡng:</b> &gt;25% — loại bỏ công ty tăng trưởng bình thường.<br>"
                  f"<b>Lý do:</b> EPS quý là tín hiệu sớm nhất. Công ty đang tăng tốc sẽ thể hiện ngay ở quý gần nhất "
                  f"trước khi số năm bắt kịp. O'Neil thống kê phần lớn cổ phiếu tăng mạnh nhất trong lịch sử "
                  f"đều có EPS quý tăng ≥25–50% trước breakout.<br>"
                  f"<span style='color:{H_DIM};font-size:11px;'>Nguồn: TradingView — EPS diluted YoY growth FQ (so cùng quý năm trước).</span>"),
            "A": (f"<b>Chỉ số:</b> EPS Năm YoY% — lợi nhuận mỗi cổ phiếu tăng trưởng so với năm trước.<br>"
                  f"<b>Ngưỡng:</b> &gt;20% — thấp hơn C vì tăng trưởng năm ổn định hơn quý.<br>"
                  f"<b>Lý do:</b> Một quý tốt có thể là may mắn; nhiều năm tốt liên tiếp là cấu trúc. "
                  f"Tiêu chí A xác nhận C không phải điểm bất thường mà là xu hướng bền vững.<br>"
                  f"<span style='color:{H_WARN};'>⚠</span> EPS annual dễ bị méo bởi base effect (phục hồi từ năm lỗ → growth ảo 500%) "
                  f"— kết hợp xem cả C và S để xác nhận."),
            "S": (f"<b>Chỉ số:</b> Revenue YoY% — doanh thu tăng so với cùng kỳ năm trước.<br>"
                  f"<b>Ngưỡng:</b> &gt;20%.<br>"
                  f"<b>Lý do:</b> EPS có thể tăng giả tạo nhờ cắt giảm chi phí, buyback, hoặc thay đổi kế toán — nhưng doanh thu thì không. "
                  f"S đảm bảo tăng trưởng đến từ top-line thực sự: khách hàng đang mua nhiều hơn. "
                  f"Công ty có C+S cùng tăng mạnh = tăng trưởng toàn diện, không phải 'tối ưu hóa sổ sách'."),
            "L": (f"<b>Chỉ số:</b> Relative Strength — hiệu suất giá so với thị trường.<br>"
                  f"<b>Ngưỡng:</b> Outperform index (1Y% cổ phiếu &gt; 1Y% index).<br>"
                  f"<b>Lý do:</b> O'Neil's L = Leader vs Laggard. Cổ phiếu dẫn đầu thị trường thường tiếp tục dẫn đầu "
                  f"— mua laggard là cưỡi ngựa chậm. App dùng proxy: 1Y% cổ phiếu trừ 1Y% index thị trường; "
                  f"nếu dương = đang outperform."),
            "Q": (f"<b>Chỉ số:</b> Gross Margin% — biên lợi nhuận gộp = (Revenue − COGS) / Revenue.<br>"
                  f"<b>Ngưỡng:</b> &gt;40%.<br>"
                  f"<b>Lý do:</b> GM cao = pricing power thực sự — công ty có thể định giá cao hơn chi phí sản xuất, "
                  f"còn tiền cho R&amp;D và marketing. Dưới 40% thường là commodity business (cạnh tranh bằng giá thấp). "
                  f"Trên 40% thường có lợi thế khác biệt bền vững.<br>"
                  f"<span style='color:{H_WARN};'>⚠ Lưu ý sector:</span> Tech/Software có GM 50–80% là bình thường; "
                  f"Industrials/Retail hiếm vượt 40% — đặc thù ngành, không phải điểm yếu."),
            "R": (f"<b>Chỉ số:</b> ROE% — Return on Equity = Lợi nhuận ròng / Vốn chủ sở hữu.<br>"
                  f"<b>Ngưỡng:</b> &gt;17%.<br>"
                  f"<b>Lý do:</b> ROE đo hiệu quả sinh lời trên vốn cổ đông. ROE &gt;17% = công ty không chỉ tăng trưởng "
                  f"mà còn tạo ra giá trị thực. ROE thấp + EPS growth cao = tăng trưởng nhờ pha loãng vốn "
                  f"(phát hành cổ phiếu mới) — không bền vững."),
            "M": (f"<b>Chỉ số:</b> EPS Acceleration — tốc độ tăng trưởng EPS đang tăng dần qua các quý.<br>"
                  f"<b>Ngưỡng:</b> ≥2 quý liên tiếp YoY% tăng dần (khi bật yfinance), hoặc EPS Qtr &gt; EPS Annual &gt;0 (khi tắt).<br>"
                  f"<b>Lý do:</b> Đây là leading indicator mạnh nhất: thị trường mất 1–2 quý để định giá lại khi earnings "
                  f"acceleration bắt đầu — phát hiện sớm có lợi thế lớn.<br>"
                  f"<b>Cột ACC trong bảng:</b> ↑3Q = 3 quý liên tiếp (mạnh nhất) · ↑2Q = 2 quý · ↑1Q = mới bắt đầu · "
                  f"↓ = đang giảm tốc (cờ đỏ dù EPS vẫn tăng tuyệt đối).<br>"
                  f"<span style='color:{H_DIM};font-size:11px;'>Trong O'Neil gốc M = Market Direction — đã tách thành MKT. Slot M dùng cho acceleration.</span>"),
            "D": (f"<b>Chỉ số:</b> D/E Ratio = Tổng nợ / Vốn chủ sở hữu.<br>"
                  f"<b>Ngưỡng:</b> &lt;1.5 — nợ không vượt quá 1.5× vốn tự có.<br>"
                  f"<b>Lý do:</b> Lãi vay phải trả dù doanh thu giảm — đòn bẩy khuếch đại lợi nhuận khi tốt "
                  f"nhưng khuếch đại thua lỗ khi xấu. D/E cao + lãi suất tăng = rủi ro thanh khoản nghiêm trọng.<br>"
                  f"<span style='color:{H_WARN};'>⚠ Financial Services:</span> Ngân hàng/bảo hiểm có D/E 8–15x là cấu trúc bình thường "
                  f"(tiền gửi khách hàng tính là 'nợ') — tiêu chí D bị bỏ qua hoàn toàn cho nhóm này."),
            "N": (f"<b>Chỉ số:</b> 52W High% — giá hiện tại so với đỉnh 52 tuần.<br>"
                  f"<b>Ngưỡng:</b> ≥90% — đang tích lũy gần đỉnh.<br>"
                  f"<b>Lý do:</b> O'Neil quan sát cổ phiếu tăng mạnh nhất thường phá đỉnh 52 tuần trước khi tăng tiếp "
                  f"— không phải 'mua rẻ' từ vùng đáy. Khi giá gần đỉnh 52W, tất cả người mua trước đều đang có lời "
                  f"→ không có áp lực bán cắt lỗ → giá dễ tiếp tục tăng. "
                  f"Cổ phiếu ở &lt;70% của 52W High thường đang trong downtrend."),
            "MKT": (f"<b>Chỉ số:</b> Market Direction — xu hướng thị trường chung (index vs MA50 &amp; MA200).<br>"
                    f"<b>Ngưỡng:</b> Index &gt; MA50 &amp; MA200 (uptrend xác nhận).<br>"
                    f"<b>Lý do:</b> O'Neil nghiên cứu lịch sử và kết luận: 3 trong 4 cổ phiếu đi cùng xu hướng thị trường. "
                    f"Mua cổ phiếu tốt trong bear market = bơi ngược dòng. Điều kiện này lọc bỏ những lần bounce ngắn "
                    f"không phải uptrend thực sự.<br>"
                    f"<span style='color:{H_DIM};font-size:11px;'>Index theo market: America=S&amp;P 500 · NASDAQ=NDX · Vietnam=VNINDEX · HongKong=HSI.</span>"),
        }
        canslim_rows = ""
        for i, key in enumerate(CS_KEYS):
            cfg  = CANSLIM[key]
            if cfg["op"] == "accel":
                thr = "Qtr&nbsp;&gt;&nbsp;Ann&nbsp;&amp;&nbsp;&gt;0%"
            elif key == "D":
                thr = f"&lt;&nbsp;{cfg['thr']}"
            else:
                op  = ops[cfg["op"]]
                thr = f"{op}&nbsp;{cfg['thr']}%"
            name = cfg["label"].split("—")[1].strip() if "—" in cfg["label"] else cfg["label"]
            bg   = H_ROW_E if i % 2 == 0 else H_ROW_O
            canslim_rows += f"""
              <tr style="background-color:{bg};">
                <td style="color:{H_BLUE};font-weight:700;font-family:Consolas,monospace;
                           font-size:15px;padding:9px 13px;width:36px;">{key}</td>
                <td style="color:{H_TEXT};padding:9px 13px;width:110px;">{name}</td>
                <td style="color:{H_SEC};font-family:Consolas,monospace;font-size:11px;
                           padding:9px 13px;width:110px;">{cfg['field']}</td>
                <td style="color:{H_GREEN};font-weight:700;padding:9px 13px;width:70px;">{thr}</td>
                <td style="color:{H_TEXT};font-size:12px;padding:9px 13px;line-height:1.6;">{full_desc[key]}</td>
              </tr>"""

        signal_rows = "".join([
            f"""<tr style="background-color:{bg};">
                  <td style="color:{fg};font-weight:700;font-size:13px;padding:10px 14px;width:140px;">{icon}&nbsp;{label}</td>
                  <td style="color:{fg};font-weight:700;font-size:13px;padding:10px 14px;width:85px;">{score}</td>
                  <td style="color:{fg};font-size:12px;padding:10px 14px;">{desc}</td>
                </tr>"""
            for icon, label, score, fg, bg, desc in [
                ("🟢", "STRONG BUY", f"≥{round(N_CS*0.80)} / {N_CS}", "#1A5C2B", "#C6EFCE",
                 f"Đạt ≥80% tiêu chí (≥{round(N_CS*0.80)}/{N_CS} điểm). Nền tảng cơ bản + kỹ thuật đều mạnh. Đây là nhóm ưu tiên cao nhất — tất cả tiêu chí quan trọng đều đạt."),
                ("🔵", "BUY",        f"≥{round(N_CS*0.625)} / {N_CS}", "#1B3A5C", "#DDEEFF",
                 f"Đạt ≥62.5% tiêu chí (≥{round(N_CS*0.625)}/{N_CS} điểm). Nền tảng tốt, một vài tiêu chí chưa đạt. Đáng xem xét nhưng cần kiểm tra kỹ tiêu chí còn thiếu trước khi vào tiền."),
                ("🟡", "WATCH",      f"≥{round(N_CS*0.375)} / {N_CS}", "#7D6608", "#FFF2CC",
                 f"Đạt ≥37.5% tiêu chí (≥{round(N_CS*0.375)}/{N_CS} điểm). Tiềm năng nhưng chưa đủ điều kiện. Đưa vào watchlist, scan lại sau 2–4 tuần để xem có cải thiện không."),
                ("🔴", "SKIP",       f"< {round(N_CS*0.375)} / {N_CS}", "#9C0006", "#FFC7CE",
                 f"Dưới {round(N_CS*0.375)}/{N_CS} điểm. Không đủ điều kiện — bỏ qua. Không nên ép mua chỉ vì thích tên công ty; chờ fundamental xoay chiều rõ ràng."),
            ]
        ])

        moat_rows = "".join([
            f"""<tr style="background-color:{bg};">
                  <td style="color:{fg};font-weight:700;font-size:13px;padding:10px 14px;width:160px;">{label}</td>
                  <td style="color:{fg};font-size:12px;padding:10px 14px;width:240px;">{cond}</td>
                  <td style="color:{fg};font-size:12px;padding:10px 14px;">{desc}</td>
                </tr>"""
            for label, cond, fg, bg, desc in [
                ("🏰 WIDE  ★★★",    "ROE 5yr ≥20%  &amp;  GM 5yr ≥50%",             "#1A5C2B", "#C6EFCE",
                 "Lợi thế cạnh tranh bền vững rộng. Rất khó bị xói mòn trong 10+ năm. Ví dụ: Apple, NVIDIA, Visa."),
                ("🏰 NARROW ★★",   "ROE ≥15%  &amp;  GM ≥35%  (hoặc Utilities/RE)", "#1A3A5C", "#DDEEFF",
                 "Lợi thế hẹp hơn, duy trì được 5–10 năm. Cần giám sát cạnh tranh định kỳ."),
                ("🏰 UNCERTAIN ★", "ROE ≥10%  (GM bất kỳ)",                           "#7D6608", "#FFF2CC",
                 "Chưa rõ có moat bền vững. Cần phân tích định tính thêm: sản phẩm, thị trường, ban lãnh đạo."),
                ("🏰 WEAK",        "ROE &lt;10%",                                     "#9C0006", "#FFC7CE",
                 "Không có dấu hiệu lợi thế cạnh tranh. Dễ bị cạnh tranh hoặc biên lợi nhuận xói mòn."),
            ]
        ])

        proxy_sector_rows = ""
        for i, (sector, proxy) in enumerate(MOAT_PROXY_MAP.items()):
            bg = H_ROW_E if i % 2 == 0 else H_ROW_O
            proxy_sector_rows += (
                f'<tr style="background-color:{bg};">'
                f'<td style="color:{H_TEXT};font-size:12px;padding:7px 13px;">{sector}</td>'
                f'<td style="color:{H_SEC};font-size:12px;padding:7px 13px;">{proxy}</td>'
                f'</tr>'
            )

        proxy_types = [
            ("Switching Cost",
             "Khách hàng khó chuyển sang đối thủ vì chi phí chuyển đổi cao (thời gian, tiền bạc, rủi ro). "
             "Ví dụ: phần mềm ERP (SAP, Oracle), cloud platform (AWS, Azure), thiết bị y tế."),
            ("Network Effect",
             "Giá trị sản phẩm tăng theo số lượng người dùng — người dùng càng nhiều, sản phẩm càng hữu ích. "
             "Ví dụ: Visa/Mastercard, Meta, LinkedIn, sàn giao dịch chứng khoán."),
            ("Intangible Assets (Patents)",
             "Bằng sáng chế, giấy phép, thương hiệu độc quyền bảo vệ khỏi cạnh tranh trong thời gian dài. "
             "Ví dụ: dược phẩm (Eli Lilly, AbbVie), bán dẫn (Qualcomm, ARM)."),
            ("Brand",
             "Thương hiệu mạnh cho phép định giá cao hơn đối thủ mà người tiêu dùng vẫn sẵn sàng trả. "
             "Ví dụ: Apple, Nike, Coca-Cola, LVMH."),
            ("Cost Advantage",
             "Chi phí sản xuất hoặc vận hành thấp hơn đối thủ nhờ quy mô, công nghệ, hoặc địa lý. "
             "Cho phép cạnh tranh giá hoặc giữ biên lợi nhuận cao hơn ngành. "
             "Ví dụ: Costco, ExxonMobil, Rio Tinto."),
            ("Scale",
             "Quy mô lớn giúp phân bổ chi phí cố định hiệu quả hơn — mỗi đơn vị sản phẩm rẻ hơn khi sản lượng tăng. "
             "Ví dụ: Amazon, Walmart, JPMorgan."),
            ("Efficient Scale",
             "Thị trường nhỏ chỉ đủ chỗ cho 1–2 công ty hoạt động hiệu quả — gia nhập thêm sẽ phá vỡ lợi nhuận cả ngành. "
             "Ví dụ: tiện ích điện/nước (regulated utilities), cảng biển, đường cao tốc thu phí."),
            ("Distribution Network",
             "Mạng lưới phân phối rộng khắp tạo rào cản khó xây dựng lại từ đầu. "
             "Ví dụ: hệ thống bán lẻ toàn quốc, chuỗi cung ứng độc quyền, logistics last-mile."),
            ("Location",
             "Vị trí địa lý độc quyền hoặc khó nhân rộng tạo lợi thế bền vững. "
             "Ví dụ: bất động sản trung tâm thương mại, trung tâm dữ liệu gần nguồn điện rẻ."),
        ]
        proxy_type_rows = ""
        for i, (ptype, pdesc) in enumerate(proxy_types):
            bg = H_ROW_E if i % 2 == 0 else H_ROW_O
            proxy_type_rows += (
                f'<tr style="background-color:{bg};">'
                f'<td style="color:{H_BLUE};font-weight:600;font-size:12px;padding:8px 13px;width:200px;">{ptype}</td>'
                f'<td style="color:{H_TEXT};font-size:12px;padding:8px 13px;">{pdesc}</td>'
                f'</tr>'
            )

        # ── Quality Compounder criteria rows ──────────────────────────────────
        qc_full_desc = {
            "ROIC": (f"<b>Chỉ số:</b> ROIC% = Net Income / (Debt + Equity) — lợi nhuận trên toàn bộ vốn đầu tư.<br>"
                     f"<b>Ngưỡng:</b> &gt;15% (biên an toàn rõ ràng so với chi phí vốn bình quân ~8–10%).<br>"
                     f"<b>Lý do:</b> Nếu ROIC &gt; WACC, công ty đang tạo ra giá trị thực — mỗi đồng tái đầu tư sinh thêm lợi nhuận. "
                     f"Đây là tiêu chí cốt lõi nhất của compounder: phân biệt 'công ty vĩ đại' (ROIC cao bền vững) "
                     f"với 'công ty bình thường' (ROIC tăng nhờ đòn bẩy hoặc may mắn).<br>"
                     f"<span style='color:{H_WARN};'>⚠ Financial Services:</span> ROIC bị méo bởi leverage cấu trúc → "
                     f"thay bằng <b>ROE &gt;12%</b> cho ngân hàng/bảo hiểm."),
            "OPGM": (f"<b>Chỉ số:</b> Operating Margin% = EBIT / Revenue — biên lợi nhuận vận hành.<br>"
                     f"<b>Ngưỡng:</b> &gt;15%.<br>"
                     f"<b>Lý do:</b> Loại bỏ ảnh hưởng cấu trúc vốn (lãi vay) và thuế, đo thuần hiệu quả core business. "
                     f"Margin cao + ổn định = pricing power thực sự: tăng giá mà không mất khách. "
                     f"OpMgn giảm dần qua nhiều năm là cảnh báo sớm áp lực cạnh tranh tăng."),
            "GM":   (f"<b>Chỉ số:</b> Gross Margin% = (Revenue − COGS) / Revenue — biên lợi nhuận gộp.<br>"
                     f"<b>Ngưỡng:</b> &gt;40%.<br>"
                     f"<b>Lý do:</b> GM cao là nền tảng của mọi thứ tốt khác. Công ty GM 70% (software) có thể "
                     f"chi nhiều cho R&amp;D và sales mà vẫn còn lãi; GM 15% (retail) phải vận hành cực kỳ hiệu quả mới tồn tại. "
                     f"Ngưỡng 40% phân biệt business tự tạo ra lợi thế với business cần scale mới sống."),
            "FCF":  (f"<b>Chỉ số:</b> FCF/share — Free Cash Flow trên mỗi cổ phiếu.<br>"
                     f"<b>Ngưỡng:</b> &gt;0 (tiền mặt thực về tay sau capex).<br>"
                     f"<b>Lý do:</b> Sau khi trả chi phí vận hành và đầu tư tài sản cố định, vẫn còn tiền mặt thực. "
                     f"Đây là điều kiện tối thiểu để compounder hoạt động: buyback, dividend, M&amp;A — tất cả cần FCF dương. "
                     f"EPS dương nhưng FCF âm = lợi nhuận kế toán, không phải tiền thật. "
                     f"FCF &gt; EPS nhiều năm liên tiếp = earnings quality rất cao."),
            "DE":   (f"<b>Chỉ số:</b> D/E Ratio = Tổng nợ / Vốn chủ sở hữu.<br>"
                     f"<b>Ngưỡng:</b> &lt;1.0 — chặt hơn tab Catalyst (1.5) vì compounder cần bền vững qua nhiều chu kỳ.<br>"
                     f"<b>Lý do:</b> Compounder thực sự tăng trưởng bằng FCF tái đầu tư, không cần vay nợ nhiều. "
                     f"Balance sheet sạch = linh hoạt khi khủng hoảng: khi đối thủ phải bán tài sản trả nợ, "
                     f"compounder có thể mua lại đối thủ hoặc buyback cổ phiếu.<br>"
                     f"<span style='color:{H_WARN};'>⚠ Financial Services:</span> D/E 8–15x là cấu trúc bình thường → "
                     f"tiêu chí này bị <b>bỏ qua hoàn toàn</b> cho sector này."),
            "MOAT": (f"<b>Chỉ số:</b> Moat Score từ ROE 5yr avg + GM 5yr avg (yfinance).<br>"
                     f"<b>Ngưỡng:</b> Wide ★★★ hoặc Narrow ★★ mới đủ điều kiện.<br>"
                     f"<b>Lý do:</b> Moat là lý do duy nhất ROIC cao có thể duy trì dài hạn. "
                     f"Không có moat: đối thủ thấy ROIC cao → đầu tư vào ngành → cạnh tranh → ROIC giảm về mức trung bình trong 3–7 năm. "
                     f"Có moat: rào cản cạnh tranh (switching cost, network effect, brand, patent) bảo vệ ROIC. "
                     f"Ổn định ROE+GM qua 5 năm = bằng chứng moat thực sự, không phải may mắn."),
        }
        qc_criteria_rows = ""
        for i, key in enumerate(QC_KEYS):
            cfg  = QC_CRITERIA[key]
            bg   = H_ROW_E if i % 2 == 0 else H_ROW_O
            if cfg["op"] == "in":
                thr_str = "Wide / Narrow"
            elif cfg["op"] == "gt":
                thr_str = f"&gt;&nbsp;{cfg['thr']}{'%' if cfg['thr'] >= 15 else ''}"
            else:
                thr_str = f"&lt;&nbsp;{cfg['thr']}"
            qc_criteria_rows += f"""
              <tr style="background-color:{bg};">
                <td style="color:{H_GREEN};font-weight:700;font-family:Consolas,monospace;
                           font-size:13px;padding:9px 13px;width:56px;">{key}</td>
                <td style="color:{H_TEXT};padding:9px 13px;width:130px;">{cfg['label']}</td>
                <td style="color:{H_GREEN};font-weight:700;padding:9px 13px;width:110px;">{thr_str}</td>
                <td style="color:{H_TEXT};font-size:12px;padding:9px 13px;line-height:1.6;">{qc_full_desc[key]}</td>
              </tr>"""

        qc_signal_rows = "".join([
            f"""<tr style="background-color:{bg};">
                  <td style="color:{fg};font-weight:700;font-size:14px;padding:10px 14px;width:160px;">{icon}</td>
                  <td style="color:{fg};font-weight:700;font-size:13px;padding:10px 14px;width:90px;">{score}</td>
                  <td style="color:{fg};font-size:12px;padding:10px 14px;">{desc}</td>
                </tr>"""
            for icon, score, fg, bg, desc in [
                ("🏆 COMPOUNDER", f"5 – {N_QC} / {N_QC}", "#1A5C2B", "#C6EFCE",
                 "Đạt ≥5/6 tiêu chí. Doanh nghiệp chất lượng cao — tăng trưởng bền vững, biên lợi nhuận tốt, bảng cân đối sạch và có moat. Ưu tiên theo dõi dài hạn."),
                ("⭐ QUALITY",     f"3 – 4 / {N_QC}",      "#1B3A5C", "#DDEEFF",
                 "Đạt 3–4 tiêu chí. Chất lượng tốt nhưng còn một vài điểm yếu. Cần kiểm tra kỹ tiêu chí chưa đạt trước khi đầu tư."),
                ("○ AVERAGE",     f"1 – 2 / {N_QC}",      "#7D6608", "#FFF2CC",
                 "Đạt 1–2 tiêu chí. Chất lượng trung bình, chưa đủ tiêu chuẩn compounder. Cần thêm phân tích định tính."),
                ("✗ WEAK",        f"0 / {N_QC}",          "#9C0006", "#FFC7CE",
                 "Không đạt tiêu chí nào. Không phải quality compounder. Bỏ qua hoặc chờ cải thiện fundamental."),
            ]
        ])

        _sb = round(N_CS * 0.80); _buy = round(N_CS * 0.625)
        conviction_html = f"""<h2>③ CONVICTION — ĐIỂM TỰ TIN TỔNG HỢP</h2>
<p style="color:{H_DIM};font-size:12px;margin:0 0 10px 0;">
  Cột <b style="color:{H_BLUE};">Conviction</b> xuất hiện trong cả hai tab (Catalyst và Quality Compounder),
  đứng ngay sau cột Score. Đây là <b>CS Score đã điều chỉnh theo chất lượng Moat</b>
  — phản ánh mức độ tự tin của hệ thống. Cùng CS Score nhưng Moat khác nhau → Conviction khác nhau → mức ưu tiên khác nhau.
</p>
<p style="color:{H_TEXT};font-size:12px;margin:0 0 4px 0;font-weight:700;">Công thức:</p>
<p style="color:{H_BLUE};font-size:13px;font-family:Consolas,monospace;
          background:{H_CODE_BG};padding:8px 14px;border-radius:4px;margin:0 0 12px 0;
          border-left:3px solid #2563EB;">
  Conviction&nbsp; = &nbsp;CS_Score &nbsp;×&nbsp; Moat_Multiplier
</p>
<table>
  <tr>
    <th style="width:155px;">Moat Score</th>
    <th style="width:110px;">Hệ số nhân</th>
    <th style="width:155px;">Conviction tối đa</th>
    <th>Ý nghĩa</th>
  </tr>
  <tr style="background-color:#C6EFCE;">
    <td style="color:#1A5C2B;font-weight:700;font-size:12px;padding:9px 13px;">WIDE ★★★</td>
    <td style="color:#1A5C2B;font-weight:700;padding:9px 13px;text-align:center;">× 1.2</td>
    <td style="color:#1A5C2B;font-weight:700;padding:9px 13px;text-align:center;">{round(N_CS*1.2, 1)}&nbsp; (CS = {N_CS})</td>
    <td style="color:#0D2010;font-size:12px;padding:9px 13px;">Hệ thống rất tự tin: momentum mạnh + lợi thế cạnh tranh bền vững. Ưu tiên cao nhất — nghiên cứu nhóm này đầu tiên.</td>
  </tr>
  <tr style="background-color:#DDEEFF;">
    <td style="color:#1B3A5C;font-weight:700;font-size:12px;padding:9px 13px;">NARROW ★★</td>
    <td style="color:#1B3A5C;font-weight:700;padding:9px 13px;text-align:center;">× 1.1</td>
    <td style="color:#1B3A5C;font-weight:700;padding:9px 13px;text-align:center;">{round(N_CS*1.1, 1)}&nbsp; (CS = {N_CS})</td>
    <td style="color:#0D1C30;font-size:12px;padding:9px 13px;">Tự tin cao: moat tốt nhưng hẹp hơn WIDE. Duy trì được 5–10 năm. Cần theo dõi áp lực cạnh tranh định kỳ.</td>
  </tr>
  <tr style="background-color:#FFF2CC;">
    <td style="color:#7D6608;font-weight:700;font-size:12px;padding:9px 13px;">UNCERTAIN ★</td>
    <td style="color:#7D6608;font-weight:700;padding:9px 13px;text-align:center;">× 1.0</td>
    <td style="color:#7D6608;font-weight:700;padding:9px 13px;text-align:center;">{float(N_CS):.1f}&nbsp; (CS = {N_CS})</td>
    <td style="color:#3D2E00;font-size:12px;padding:9px 13px;">Bằng CS Score thuần: chưa xác định được moat. Cần phân tích định tính thêm về mô hình kinh doanh.</td>
  </tr>
  <tr style="background-color:#FFC7CE;">
    <td style="color:#9C0006;font-weight:700;font-size:12px;padding:9px 13px;">WEAK</td>
    <td style="color:#9C0006;font-weight:700;padding:9px 13px;text-align:center;">× 0.85</td>
    <td style="color:#9C0006;font-weight:700;padding:9px 13px;text-align:center;">{round(N_CS*0.85, 1)}&nbsp; (CS = {N_CS})</td>
    <td style="color:#5C0000;font-size:12px;padding:9px 13px;">Giảm điểm: không có lợi thế cạnh tranh. CS cao nhưng Conviction thấp = momentum ngắn hạn, dễ đảo chiều. Thận trọng khi vào tiền.</td>
  </tr>
</table>
<p style="color:{H_TEXT};font-size:12px;margin:10px 0 5px 0;font-weight:700;">Ví dụ thực tế — hai cổ phiếu cùng CS Score = 8:</p>
<table>
  <tr>
    <th style="width:80px;">Cổ phiếu</th>
    <th style="width:90px;">CS Score</th>
    <th style="width:140px;">Moat</th>
    <th style="width:120px;">Conviction</th>
    <th>Nhận xét</th>
  </tr>
  <tr style="background-color:{H_ROW_E};">
    <td style="color:{H_BLUE};font-weight:700;padding:8px 13px;">Cổ A</td>
    <td style="color:{H_TEXT};padding:8px 13px;text-align:center;">8&nbsp;/&nbsp;{N_CS}</td>
    <td style="color:#1A5C2B;font-weight:700;padding:8px 13px;">WIDE ★★★</td>
    <td style="color:#16A34A;font-weight:700;font-size:15px;padding:8px 13px;text-align:center;">{round(8*1.2, 1)}</td>
    <td style="color:{H_TEXT};font-size:12px;padding:8px 13px;">Mạnh cả ngắn hạn (CS) lẫn dài hạn (Moat) → ưu tiên nghiên cứu sâu</td>
  </tr>
  <tr style="background-color:{H_ROW_O};">
    <td style="color:{H_BLUE};font-weight:700;padding:8px 13px;">Cổ B</td>
    <td style="color:{H_TEXT};padding:8px 13px;text-align:center;">8&nbsp;/&nbsp;{N_CS}</td>
    <td style="color:#9C0006;font-weight:700;padding:8px 13px;">WEAK</td>
    <td style="color:#9C0006;font-weight:700;font-size:15px;padding:8px 13px;text-align:center;">{round(8*0.85, 1)}</td>
    <td style="color:{H_TEXT};font-size:12px;padding:8px 13px;">Momentum tốt nhưng không có nền tảng bền vững → dễ bị xói mòn, theo dõi sát hơn</td>
  </tr>
</table>
<p class="tip">💡 Khi hai mã có CS Score gần bằng nhau, dùng <b>Conviction</b> để phân biệt: mã nào Conviction cao hơn thì hệ thống tự tin hơn cả ngắn lẫn dài hạn.
Conviction chính xác nhất khi bật <b>yfinance Moat</b> — khi tắt, Moat ước tính từ TTM ROE&amp;GM nên Conviction kém chính xác hơn.</p>"""

        return f"""<html><head><style>
  body  {{ background-color:#F8FAFC; color:#1A202C;
           font-family:'Segoe UI',sans-serif; font-size:13px;
           margin:0; padding:0; }}
  h1    {{ color:#1E40AF; font-size:20px; letter-spacing:4px;
           font-weight:700; margin:0 0 5px 0; }}
  .sub  {{ color:#64748B; font-size:11px; letter-spacing:2px;
           margin:0 0 20px 0; }}
  h2    {{ color:#374151; font-size:12px; letter-spacing:3px;
           font-weight:700; margin:26px 0 9px 0;
           border-bottom:2px solid #CBD5E0; padding-bottom:5px; }}
  table {{ border-collapse:collapse; width:100%; margin-bottom:10px; }}
  th    {{ background-color:#E2E8F0; color:#2D3748; text-align:left;
           padding:8px 13px; font-size:11px; letter-spacing:1px;
           font-weight:700; border-bottom:2px solid #CBD5E0; }}
  .tip  {{ color:#64748B; font-size:11px; font-style:italic;
           margin:5px 0 0 0; }}
  li    {{ margin:7px 0; color:#1A202C; font-size:12px; }}
  b     {{ color:#0F172A; font-weight:700; }}
</style></head><body>
<h1>FUNDAMENTAL SCREENER</h1>
<p class="sub">HƯỚNG DẪN SỬ DỤNG  ·  CATALYST · QUALITY COMPOUNDER · EQ BADGE · MOAT · SIGNAL</p>

<h2>① CATALYST — TIÊU CHÍ LỌC CỔ PHIẾU (dựa trên phương pháp CAN SLIM)</h2>
<p style="color:#64748B;font-size:12px;margin:0 0 10px 0;">
  Mỗi tiêu chí cho 1 điểm nếu đạt. Cột <b>Chỉ số</b> là tên dữ liệu từ TradingView/yfinance.
  Cột <b>Ngưỡng</b> là điều kiện cụ thể để đạt điểm. Cột <b>Lý do chọn</b> giải thích tại sao ngưỡng đó được dùng.
</p>
<table>
  <tr>
    <th style="width:36px;">Key</th>
    <th style="width:110px;">Tên</th>
    <th style="width:110px;">Chỉ số (nguồn)</th>
    <th style="width:70px;">Ngưỡng</th>
    <th>Lý do chọn giá trị này</th>
  </tr>
  {canslim_rows}
</table>
<p class="tip">⚠  Dữ liệu từ TradingView (TTM/FY). EPS &amp; Revenue là YoY growth %. ROE &amp; Gross Margin là trailing twelve months.</p>

<h2>② CATALYST SIGNAL — KẾT QUẢ TỔNG HỢP</h2>
<table>
  <tr>
    <th style="width:140px;">Tín hiệu</th>
    <th style="width:85px;">Điểm</th>
    <th>Ý nghĩa &amp; Hành động gợi ý</th>
  </tr>
  {signal_rows}
</table>
<p class="tip">Score = số tiêu chí đạt được (tối đa {N_CS}). ✓ = đạt (+1 điểm) · ✗ = không đạt · — = thiếu dữ liệu (không tính vào score). Signal dựa trên Score: STRONG BUY ≥{round(N_CS*0.80)}, BUY ≥{round(N_CS*0.625)}, WATCH ≥{round(N_CS*0.375)}, SKIP &lt;{round(N_CS*0.375)}.</p>

{conviction_html}

<h2>④ QUALITY COMPOUNDER — 6 TIÊU CHÍ CHẤT LƯỢNG BỀN VỮNG</h2>
<p style="color:#64748B;font-size:12px;margin:0 0 10px 0;">Tab <b style="color:#16A34A;">Quality Compounder</b> tìm doanh nghiệp có khả năng tăng trưởng kép bền vững dài hạn — không chỉ tốt về momentum mà còn mạnh về chất lượng nền tảng thực sự.</p>
<table>
  <tr>
    <th style="width:52px;">Key</th>
    <th style="width:130px;">Chỉ số</th>
    <th style="width:110px;">Ngưỡng</th>
    <th>Lý do chọn giá trị này</th>
  </tr>
  {qc_criteria_rows}
</table>
<p class="tip">⚠  Tiêu chí MOAT yêu cầu bật "yfinance Moat" để có dữ liệu 5 năm chính xác. Khi tắt, Moat được ước tính từ TTM ROE &amp; GM.<br>
⚠  <b>Financial Services</b> (ngân hàng, bảo hiểm): ROIC được thay bằng <b>ROE &gt; 12%</b>; D/E bị <b>bỏ qua</b> hoàn toàn. Điểm tối đa cho Financial Services vẫn là 6 (5 tiêu chí còn lại + Moat).</p>

<h2>⑤ QUALITY COMPOUNDER SIGNAL</h2>
<table>
  <tr>
    <th style="width:160px;">Tín hiệu</th>
    <th style="width:90px;">Điểm</th>
    <th>Ý nghĩa &amp; Hành động gợi ý</th>
  </tr>
  {qc_signal_rows}
</table>
<p class="tip">Score = số trong 6 tiêu chí đạt. Quick Filter "🟢 STRONG BUY" trên tab Quality Compounder sẽ lọc cột 🏆 COMPOUNDER.</p>

<h2>⑥ EQ BADGE — CHẤT LƯỢNG LỢI NHUẬN (Earnings Quality)</h2>
<p style="color:#64748B;font-size:12px;margin:0 0 10px 0;">
  Cột <b style="color:#0D9488;">EQ</b> trong bảng Quality Compounder trả lời câu hỏi:
  <b>Lợi nhuận công ty báo cáo có phải là tiền mặt thực hay chỉ là con số kế toán?</b>
</p>

<p style="color:#1A202C;font-size:12px;margin:0 0 6px 0;font-weight:700;">Tại sao cần chỉ số này?</p>
<p style="color:#374151;font-size:12px;margin:0 0 12px 0;line-height:1.7;">
  Net Income (lợi nhuận ròng) là con số kế toán — có thể bị <b>tô vẽ</b> bằng cách ghi nhận doanh thu sớm,
  trì hoãn chi phí, hoặc thay đổi ước tính kế toán. Trong khi đó,
  <b>Free Cash Flow (FCF)</b> là tiền mặt thực sự chảy vào tài khoản công ty sau khi đã trừ chi phí đầu tư —
  không thể làm giả. Nếu Net Income cao nhưng FCF thấp, đó là dấu hiệu lợi nhuận phụ thuộc vào kế toán,
  không phải kinh doanh thực chất.
</p>

<p style="color:#1A202C;font-size:12px;margin:0 0 6px 0;font-weight:700;">Công thức tính:</p>
<p style="color:#0D6E6E;font-size:13px;font-family:Consolas,monospace;
          background:#E6FFFA;padding:8px 14px;border-radius:4px;margin:0 0 12px 0;
          border-left:3px solid #0D9488;">
  FCF / Net Income Ratio &nbsp;=&nbsp; Free Cash Flow &nbsp;÷&nbsp; Net Income &nbsp;(TTM, từ yfinance)
</p>

<table>
  <tr>
    <th style="width:160px;">Badge</th>
    <th style="width:130px;">Điều kiện</th>
    <th>Ý nghĩa &amp; Hành động</th>
  </tr>
  <tr style="background-color:#F0FFF4;">
    <td style="color:#1A5C2B;font-weight:700;font-size:13px;padding:10px 14px;background:#C6EFCE;">💚 Cash Backed</td>
    <td style="color:#1A5C2B;font-weight:700;padding:10px 14px;background:#C6EFCE;">Ratio ≥ 0.80</td>
    <td style="color:#1A202C;font-size:12px;padding:10px 14px;">
      Mỗi 1 đồng lợi nhuận kế toán, công ty thực thu ≥ 0.8 đồng tiền mặt.
      Lợi nhuận <b>đáng tin cậy cao</b> — tài chính minh bạch, ít rủi ro kế toán.
      Compounder thực sự thường có chỉ số này ≥ 1.0 (FCF &gt; Net Income).
    </td>
  </tr>
  <tr style="background-color:#FFFBEB;">
    <td style="color:#7D6608;font-weight:700;font-size:13px;padding:10px 14px;background:#FFF2CC;">🟡 Mixed</td>
    <td style="color:#7D6608;font-weight:700;padding:10px 14px;background:#FFF2CC;">0.30 ≤ Ratio &lt; 0.80</td>
    <td style="color:#1A202C;font-size:12px;padding:10px 14px;">
      Tiền mặt thu về chỉ bằng 30–80% lợi nhuận báo cáo. Có thể do đầu tư mở rộng nặng (capex cao),
      hoặc tốc độ ghi nhận doanh thu nhanh hơn thu tiền. <b>Cần xem thêm</b> bối cảnh ngành
      — một số công ty tăng trưởng nhanh có FCF thấp tạm thời là bình thường.
    </td>
  </tr>
  <tr style="background-color:#FFF5F5;">
    <td style="color:#9C0006;font-weight:700;font-size:13px;padding:10px 14px;background:#FFC7CE;">🔴 Accrual Heavy</td>
    <td style="color:#9C0006;font-weight:700;padding:10px 14px;background:#FFC7CE;">Ratio &lt; 0.30</td>
    <td style="color:#1A202C;font-size:12px;padding:10px 14px;">
      Lợi nhuận chủ yếu là con số kế toán, tiền mặt thực thu rất ít.
      <b>Cảnh báo đỏ</b> — đặc biệt nguy hiểm nếu kết hợp D/E cao.
      Rủi ro: công ty có thể phải huy động vốn mới (dilution), hoặc lợi nhuận sẽ bị điều chỉnh giảm.
    </td>
  </tr>
  <tr style="background-color:#F8FAFC;">
    <td style="color:#64748B;font-weight:700;font-size:13px;padding:10px 14px;">— (không có)</td>
    <td style="color:#64748B;padding:10px 14px;">yfinance tắt</td>
    <td style="color:#64748B;font-size:12px;padding:10px 14px;">
      EQ Badge chỉ tính được khi bật checkbox "yfinance Moat" vì cần cashflow statement từ yfinance.
      Khi tắt yfinance, cột EQ hiển thị "—" cho tất cả.
    </td>
  </tr>
</table>
<p class="tip">⚠  EQ Badge dùng số liệu <b>năm tài chính gần nhất</b> (annual, không phải TTM quarterly).
Một số ngành có FCF tự nhiên thấp hơn Net Income do capex lớn (bán dẫn, pharma R&amp;D) — cần so sánh trong ngành, không áp dụng ngưỡng cứng cho mọi sector.</p>

<h2>⑦ MOAT SCORE — LỢI THẾ CẠNH TRANH (Economic Moat)</h2>
<p style="color:#64748B;font-size:12px;margin:0 0 8px 0;">
  Moat Score được tính từ ROE 5yr avg + GM 5yr avg (từ yfinance). Ổn định cao qua 5 năm = bằng chứng moat thực sự.
</p>
<table>
  <tr>
    <th style="width:160px;">Moat</th>
    <th style="width:240px;">Điều kiện (5yr avg từ yfinance)</th>
    <th>Ý nghĩa</th>
  </tr>
  {moat_rows}
</table>
<p class="tip">Khi tắt "yfinance Moat", dùng TTM ROE &amp; GM từ TradingView thay thế (nhanh hơn nhưng kém chính xác — chỉ 1 điểm thay vì trung bình 5 năm).</p>

<h2>⑧ MOAT PROXY — LOẠI LỢI THẾ CẠNH TRANH THEO NGÀNH</h2>
<p style="color:#64748B;font-size:12px;margin:0 0 9px 0;">Moat Proxy được gán tự động theo <b>ngành (Sector)</b> — giúp bạn hiểu loại lợi thế cạnh tranh điển hình của từng ngành để phân tích sâu hơn.</p>
<table>
  <tr>
    <th style="width:200px;">Sector</th>
    <th>Moat Proxy (loại lợi thế điển hình)</th>
  </tr>
  {proxy_sector_rows}
</table>

<p style="color:#374151;font-size:12px;margin:16px 0 8px 0;font-weight:700;letter-spacing:1px;">GIẢI THÍCH CÁC LOẠI LỢI THẾ CẠNH TRANH</p>
<table>
  <tr>
    <th style="width:200px;">Loại lợi thế</th>
    <th>Ý nghĩa &amp; Ví dụ thực tế</th>
  </tr>
  {proxy_type_rows}
</table>
<p class="tip">Moat Proxy chỉ là ước tính định tính theo ngành — không thay thế phân tích sâu từng công ty. Hai công ty cùng ngành có thể có moat type khác nhau hoàn toàn.</p>

<h2>⑨ CÁCH SỬ DỤNG APP</h2>
<ul>
  <li><b>SCAN</b> — Lấy top N cổ phiếu theo Market Cap từ TradingView. Tự động tính Catalyst Score, Conviction, Quality Compounder Score cho tất cả. Sau khi scan xong, tự chuyển sang Dashboard.</li>
  <li><b>Tab Catalyst</b> — Hiển thị tất cả cổ phiếu với 10 tiêu chí (C/A/S/L/Q/R/M/D/N/MKT), CS Score, Conviction và Signal. Click tiêu đề cột để sort. Cột Conviction nằm sau Score — dùng cột này để ưu tiên khi nhiều mã có cùng Signal.</li>
  <li><b>Tab Quality Compounder</b> — Góc nhìn dài hạn: ROIC, Op Margin, Gross Margin, FCF/sh, D/E, Moat. Cũng có cột Conviction. Hai tab dùng chung chart panel bên phải.</li>
  <li><b>yfinance Moat</b> — Bật checkbox này trước khi SCAN để:
    (1) Tính Moat chính xác từ ROE/GM trung bình 5 năm (thay vì TTM),
    (2) Tính <b>Conviction</b> chính xác hơn (WIDE/NARROW/WEAK có cơ sở thực tế),
    (3) Tính <b>EQ Badge</b> (FCF/Net Income — chất lượng lợi nhuận).
    Chậm hơn ~0.3 giây/cổ phiếu do gọi API yfinance.
  </li>
  <li><b>TICKER lookup</b> — Nhập mã (ví dụ: NVDA, AAPL) vào ô Ticker rồi Enter hoặc nhấn LOOKUP. Hữu ích khi muốn tra nhanh 1–2 mã mà không cần scan toàn bộ.</li>
  <li><b>FILTER</b> — Gõ vào ô filter để lọc bảng theo Ticker hoặc tên công ty theo thời gian thực. Ví dụ: gõ "tech" để chỉ hiện các công ty có "tech" trong tên.</li>
  <li><b>Quick Filter chips</b> — Lọc nhanh theo nhóm: STRONG BUY / COMPOUNDER (tab QC), Moat Wide/Narrow, 1Y%&gt;50, 52W High&gt;90%. Dùng nút ✕ để bỏ filter.</li>
  <li><b>Click vào hàng</b> — Xem breakdown chi tiết: ✓/✗ từng tiêu chí, giải thích từng chỉ số, và chart tự động tải bên phải. Đây là cách nhanh nhất để hiểu tại sao một mã đạt/không đạt.</li>
  <li><b>Sort</b> — Click tiêu đề cột để sort. Cột số (CS Score, Conviction, 1Y%...) sort theo giá trị số, không phải chữ. Click 2 lần để đảo chiều.</li>
  <li><b>Export</b> — Xuất Excel với các sheet: <b>Data</b> (toàn bộ), <b>Quality Compounder</b>, <b>Financials</b> (nếu bật yfinance — biểu đồ EPS/Revenue từng công ty), <b>Dashboard</b> (tóm tắt). Thanh tiến trình xanh ở status bar — UI không bị đơ trong lúc xuất. Xem hướng dẫn đọc Dashboard ở phần ⑩.</li>
  <li><b>F1</b> — Mở màn hình hướng dẫn này bất cứ lúc nào.</li>
</ul>

<h2>⑩ DASHBOARD EXCEL — HƯỚNG DẪN ĐỌC ĐẦY ĐỦ</h2>

<p style="color:#64748B;font-size:12px;margin:0 0 14px 0;">
  Tab <b style="color:#1E40AF;">Dashboard</b> là bản tóm tắt toàn bộ kết quả scan, được thiết kế để đọc
  <b>từ trên xuống</b> theo mức độ ưu tiên.
</p>

<!-- ══ GIẢI THÍCH KHÁI NIỆM CƠ BẢN ══ -->
<p style="color:#374151;font-size:12px;margin:10px 0 6px 0;font-weight:700;letter-spacing:1px;">① HIỂU CÁC KHÁI NIỆM CƠ BẢN</p>
<table>
  <tr style="background-color:#EFF6FF;">
    <td style="color:#1E40AF;font-weight:700;font-size:12px;padding:8px 12px;width:160px;">CS Score (0–{N_CS})</td>
    <td style="color:#1A202C;font-size:12px;padding:8px 12px;">
      <b>Catalyst Score</b> — số tiêu chí đạt được trong {N_CS} tiêu chí tăng trưởng + momentum.<br>
      Điểm càng cao, cổ phiếu càng mạnh cả về cơ bản lẫn kỹ thuật.<br>
      <span style="color:#6B7280;font-size:11px;">Score ≥ 7 = tốt · Score 5–6 = trung bình · Score &lt; 5 = yếu</span>
    </td>
  </tr>
  <tr style="background-color:#F0FFF4;">
    <td style="color:#16A34A;font-weight:700;font-size:12px;padding:8px 12px;">QC Score (0–6)</td>
    <td style="color:#1A202C;font-size:12px;padding:8px 12px;">
      <b>Quality Compounder Score</b> — số tiêu chí chất lượng dài hạn đạt được (ROIC, margin, FCF, D/E, Moat).<br>
      Cần bật <b>yfinance Moat</b> để tính đầy đủ.<br>
      <span style="color:#6B7280;font-size:11px;">Score ≥ 4 = chất lượng cao · Score 2–3 = ổn · Score &lt; 2 = yếu</span>
    </td>
  </tr>
  <tr style="background-color:#FFFBEB;">
    <td style="color:#B45309;font-weight:700;font-size:12px;padding:8px 12px;">ROIC%</td>
    <td style="color:#1A202C;font-size:12px;padding:8px 12px;">
      <b>Return on Invested Capital</b> — tỷ suất sinh lời trên toàn bộ vốn đầu tư.<br>
      ROIC 20% = mỗi 100đ đầu tư tạo ra 20đ lợi nhuận/năm. Thước đo "công ty kiếm tiền hiệu quả" tốt nhất.<br>
      <span style="color:#6B7280;font-size:11px;">ROIC &gt;15% = xuất sắc · 10–15% = tốt · &lt;10% = trung bình</span>
    </td>
  </tr>
  <tr style="background-color:#FFF5F5;">
    <td style="color:#B45309;font-weight:700;font-size:12px;padding:8px 12px;">P/E Ratio</td>
    <td style="color:#1A202C;font-size:12px;padding:8px 12px;">
      <b>Price-to-Earnings</b> — giá cổ phiếu / lợi nhuận mỗi cổ phiếu. Thước đo "giá đắt hay rẻ".<br>
      P/E = 20 nghĩa là bạn trả 20đ để mua 1đ lợi nhuận/năm.<br>
      <span style="color:#6B7280;font-size:11px;">P/E &lt; 15 = rẻ · 15–35 = hợp lý · &gt;50 = đắt/rủi ro cao · Âm = đang lỗ</span>
    </td>
  </tr>
  <tr style="background-color:#F0F4FA;">
    <td style="color:#B45309;font-weight:700;font-size:12px;padding:8px 12px;">EPS YoY%</td>
    <td style="color:#1A202C;font-size:12px;padding:8px 12px;">
      <b>Earnings Per Share growth</b> — tốc độ tăng trưởng lợi nhuận mỗi cổ phiếu so cùng kỳ năm ngoái.<br>
      EPS Annual +30% = lợi nhuận tăng 30% — dấu hiệu công ty đang tăng trưởng mạnh.<br>
      <span style="color:#6B7280;font-size:11px;">EPS &gt;25% YoY = rất tốt · 10–25% = tốt · &lt;10% = yếu · Âm = lỗ</span>
    </td>
  </tr>
  <tr style="background-color:#FFFFFF;">
    <td style="color:#B45309;font-weight:700;font-size:12px;padding:8px 12px;">D/E Ratio</td>
    <td style="color:#1A202C;font-size:12px;padding:8px 12px;">
      <b>Debt-to-Equity</b> — tỷ lệ nợ / vốn chủ sở hữu. D/E = 0.5 nghĩa là cứ 1đ vốn tự có, vay thêm 0.5đ.<br>
      D/E cao = rủi ro cao khi lãi suất tăng.<br>
      <span style="color:#6B7280;font-size:11px;">D/E &lt; 0.5 = an toàn · 0.5–1 = chấp nhận được · &gt;1 = ⚠ cảnh báo</span>
    </td>
  </tr>
  <tr style="background-color:#EFF6FF;">
    <td style="color:#1E40AF;font-weight:700;font-size:12px;padding:8px 12px;">1Y% / 52W High%</td>
    <td style="color:#1A202C;font-size:12px;padding:8px 12px;">
      <b>1Y%</b> = hiệu suất giá trong 1 năm qua. +60% = giá tăng 60% trong 12 tháng.<br>
      <b>52W High%</b> = giá hiện tại / đỉnh 52 tuần. 95% = đang rất gần đỉnh = cổ phiếu đang rất mạnh.<br>
      <span style="color:#6B7280;font-size:11px;">52W High &gt;90% = gần đỉnh (tích cực) · &lt;50% = đang thấp (cẩn thận)</span>
    </td>
  </tr>
  <tr style="background-color:#F5F0FF;">
    <td style="color:#6D28D9;font-weight:700;font-size:12px;padding:8px 12px;">Signal (tín hiệu)</td>
    <td style="color:#1A202C;font-size:12px;padding:8px 12px;">
      Kết quả phân loại tự động dựa trên tổng hợp tất cả tiêu chí:<br>
      🟢 <b>STRONG BUY</b> — CS Score cao ≥80%, đủ điều kiện mạnh nhất<br>
      🔵 <b>BUY</b> — CS Score khá ≥62.5%, đủ điều kiện cơ bản<br>
      🏆 <b>COMPOUNDER</b> — QC Score ≥5/6, công ty chất lượng xuất sắc dài hạn<br>
      ⭐ <b>QUALITY</b> — QC Score 3–4/6, nền tảng vững chắc<br>
      🟡 <b>WATCH</b> — CS Score trung bình, cần theo dõi thêm<br>
      🔴 <b>SKIP</b> — CS Score thấp, không khuyến nghị lúc này
    </td>
  </tr>
</table>

<!-- ══ TỪNG VÙNG DASHBOARD ══ -->
<p style="color:#374151;font-size:12px;margin:14px 0 6px 0;font-weight:700;letter-spacing:1px;">② TỪNG VÙNG TRONG DASHBOARD</p>
<table>
  <tr>
    <th style="width:180px;">Vùng</th>
    <th>Nội dung &amp; Cách đọc</th>
  </tr>

  <tr style="background-color:#EFF6FF;">
    <td style="color:#1E40AF;font-weight:700;font-size:13px;padding:10px 14px;">❶ Catalyst KPIs</td>
    <td style="color:#1A202C;font-size:12px;padding:10px 14px;">
      <b>Hàng thẻ tổng quan — đọc đầu tiên để biết thị trường đang như thế nào.</b><br><br>
      • <b>Total Stocks</b>: tổng số mã được scan<br>
      • <b># Strong Buy</b>: số mã đạt STRONG BUY. &gt;10% tổng → thị trường đang tốt<br>
      • <b># Buy</b>: số mã tín hiệu BUY — chất lượng khá<br>
      • <b>Avg CS Score</b>: điểm Catalyst trung bình. &gt;6 = thị trường đang chất lượng cao<br>
      • <b>Avg 1Y%</b>: hiệu suất trung bình 1 năm của toàn bộ cổ phiếu scan<br>
      • <b># Near 52W High</b>: số mã gần đỉnh (&gt;90%) — nhiều = thị trường đang tăng mạnh
    </td>
  </tr>

  <tr style="background-color:#F0FFF4;">
    <td style="color:#16A34A;font-weight:700;font-size:13px;padding:10px 14px;">❷ QC KPIs</td>
    <td style="color:#1A202C;font-size:12px;padding:10px 14px;">
      <b>Hàng thẻ chất lượng dài hạn — cần bật yfinance Moat để hiện đầy đủ.</b><br><br>
      • <b># Compounder</b>: số mã QC Signal "COMPOUNDER" — giữ 5–10 năm<br>
      • <b># Quality</b>: số mã QC tốt nhưng chưa đến mức Compounder<br>
      • <b>Avg QC Score</b>: điểm QC trung bình — &gt;3 là thị trường có nhiều công ty chất lượng<br>
      • <b>Avg ROIC%</b>: ROIC trung bình. &gt;15% = doanh nghiệp đang rất hiệu quả<br>
      • <b># Dual Leaders</b>: số mã đạt cả CS≥7 và QC≥4 — "tinh hoa" của danh sách<br>
      • <b># Cash Backed</b>: trong nhóm QC, số mã có FCF/NI ≥ 0.8
    </td>
  </tr>

  <tr style="background-color:#FFFBEB;">
    <td style="color:#B45309;font-weight:700;font-size:13px;padding:10px 14px;">⚠ Risk Flags</td>
    <td style="color:#1A202C;font-size:12px;padding:10px 14px;">
      <b>Đèn cảnh báo — đọc dòng này trước khi đưa ra quyết định mua.</b><br><br>
      • <b># D/E &gt; 1</b>: số mã có nợ &gt; vốn chủ sở hữu — rủi ro lãi suất tăng<br>
      • <b># P/E &gt; 50</b>: số mã định giá rất cao — nếu kỳ vọng sai, giảm mạnh<br>
      • <b># 1M% &lt; −10%</b>: số mã giảm &gt;10% trong tháng vừa rồi — thị trường đang yếu<br><br>
      <b>Cách đọc:</b> Risk Flags nhiều → tăng tiêu chuẩn lọc (CS ≥ 8) hoặc đợi thêm tín hiệu.
    </td>
  </tr>

  <tr style="background-color:#F0F4FA;">
    <td style="color:#1E40AF;font-weight:700;font-size:13px;padding:10px 14px;">📊 Top CS / Top QC</td>
    <td style="color:#1A202C;font-size:12px;padding:10px 14px;">
      <b>Hai bảng đặt cạnh nhau — danh sách đầy đủ các mã đủ tiêu chuẩn.</b><br><br>
      <b style="color:#1E40AF;">Trái — Top Catalyst:</b> tất cả mã STRONG BUY hoặc BUY, sắp xếp theo CS Score cao nhất.<br>
      Cột xanh lá = CS Score cao. Cột 1Y% màu xanh/đỏ theo chiều tăng/giảm.<br><br>
      <b style="color:#16A34A;">Phải — Top Quality Compounder:</b> tất cả mã COMPOUNDER hoặc QUALITY, sắp xếp theo QC Score.<br>
      Cột ROIC% xanh đậm nếu &gt;15% — mức ROIC của các công ty vĩ đại.<br><br>
      <b>Mẹo:</b> Mã xuất hiện trong <u>cả hai bảng</u> = đáng chú ý nhất — vừa mạnh ngắn hạn, vừa chất lượng dài hạn.
    </td>
  </tr>

  <tr style="background-color:#F5F0FF;">
    <td style="color:#6D28D9;font-weight:700;font-size:13px;padding:10px 14px;">⭐ Dual Leaders</td>
    <td style="color:#1A202C;font-size:12px;padding:10px 14px;">
      <b>Bảng "tinh hoa" — shortlist đáng nghiên cứu kỹ nhất.</b><br><br>
      Điều kiện: <b>CS Score ≥ 7</b> VÀ <b>QC Score ≥ 4</b> — giao điểm "đang chạy mạnh" và "công ty thực sự tốt".<br>
      Tránh được cả hai bẫy phổ biến: mua cổ phiếu kém chất lượng chỉ vì đang hot,
      hoặc mua công ty tốt nhưng không ai quan tâm.<br><br>
      • <b>1Y%</b> — data bar: thanh dài hơn = hiệu suất tốt hơn<br>
      • <b>ROE%</b> — lợi nhuận trên vốn cổ đông, &gt;17% là tốt<br>
      • <b>ROIC%</b> — xanh &gt;15%, vàng 10–15%, cam thấp hơn<br>
      • <b>D/E</b> — đỏ nếu &gt;1, xanh nếu &lt;0.5<br>
      • <b>⚠</b> — cảnh báo D/E &gt;1 hoặc P/E &gt;50 — cần nghiên cứu thêm
    </td>
  </tr>

  <tr style="background-color:#F0F4FA;">
    <td style="color:#6D28D9;font-weight:700;font-size:13px;padding:10px 14px;">🎯 Top Picks</td>
    <td style="color:#1A202C;font-size:12px;padding:10px 14px;">
      <b>4 bảng chiến lược — mỗi bảng phù hợp một mục tiêu đầu tư khác nhau.</b><br><br>

      <b style="color:#1E40AF;">🚀 Momentum — ngắn hạn 3–12 tháng</b><br>
      <span style="color:#374151;font-size:11px;">
        STRONG BUY sắp xếp theo 1Y% cao nhất. Nghiên cứu (Jegadeesh &amp; Titman 1993) chứng minh:
        cổ phiếu outperform 12 tháng qua có xu hướng tiếp tục outperform 3–12 tháng tới.<br>
        ⚠ Nếu thị trường đảo chiều, những mã tăng mạnh nhất sẽ giảm mạnh nhất.
      </span><br><br>

      <b style="color:#0D9488;">💎 Quality — dài hạn 5–10 năm</b><br>
      <span style="color:#374151;font-size:11px;">
        COMPOUNDER sắp xếp theo ROIC% cao nhất. ROIC &gt;15% bền vững = công ty có moat thực sự,
        đối thủ khó copy được. Triết lý của Warren Buffett và Charlie Munger.<br>
        ⚠ P/E thường cao, nhạy cảm khi lãi suất tăng.
      </span><br><br>

      <b style="color:#B45309;">📉 Value Growth — trung hạn 1–3 năm</b><br>
      <span style="color:#374151;font-size:11px;">
        P/E &lt;35, EPS &gt;10%/năm, CS ≥6 — sắp xếp theo P/E thấp nhất. Chiến lược GARP của Peter Lynch:
        tăng trưởng tốt nhưng giá chưa phản ánh hết tiềm năng.<br>
        ⚠ "Bẫy giá rẻ" — P/E thấp đôi khi vì công ty đang gặp vấn đề.
      </span><br><br>

      <b style="color:#6D28D9;">💥 Breakout — ngắn hạn, cần theo dõi sát</b><br>
      <span style="color:#374151;font-size:11px;">
        52W High ≥90% và CS ≥7. O'Neil và Minervini chỉ mua cổ phiếu đang phá đỉnh lịch sử:
        khi giá gần đỉnh 52W, không có áp lực bán cắt lỗ → giá dễ tiếp tục tăng.<br>
        ⚠ Nếu thị trường yếu, breakout thất bại và giá quay về rất nhanh.
      </span>
    </td>
  </tr>

  <tr style="background-color:#FFFFFF;">
    <td style="color:#374151;font-weight:700;font-size:13px;padding:10px 14px;">❸ Chart Analysis</td>
    <td style="color:#1A202C;font-size:12px;padding:10px 14px;">
      <b>Bảng 2×2 phân vùng — danh sách mã theo vị trí CS/QC.</b><br><br>
      Trục ngang: CS Score (trái &lt;7, phải ≥7) — Trục dọc: QC Score (trên ≥4, dưới &lt;4)<br><br>
      <table style="border-collapse:collapse;font-size:11px;width:100%;">
        <tr>
          <td style="padding:6px 10px;background:#1A5C2B;color:#FFF;font-weight:700;border:1px solid #CBD5E0;width:50%;">
            💎 Quality (CS&lt;7, QC≥4)<br>
            <span style="font-weight:400;font-size:10px;color:#C8E6C9;">Cơ bản tốt, chưa được thị trường chú ý. Phù hợp giữ dài hạn.</span>
          </td>
          <td style="padding:6px 10px;background:#4A235A;color:#FFF;font-weight:700;border:1px solid #CBD5E0;width:50%;">
            ⭐ Dual Leaders (CS≥7, QC≥4)<br>
            <span style="font-weight:400;font-size:10px;color:#E1BEE7;">Tốt nhất cả hai chiều. Ưu tiên nghiên cứu nhóm này đầu tiên.</span>
          </td>
        </tr>
        <tr>
          <td style="padding:6px 10px;background:#E2E8F0;color:#374151;font-weight:700;border:1px solid #CBD5E0;">
            👀 Watchlist (CS&lt;7, QC&lt;4)<br>
            <span style="font-weight:400;font-size:10px;">Chưa đủ điều kiện. Hiển thị tên mã — theo dõi và scan lại.</span>
          </td>
          <td style="padding:6px 10px;background:#1A5276;color:#FFF;font-weight:700;border:1px solid #CBD5E0;">
            🚀 Momentum (CS≥7, QC&lt;4)<br>
            <span style="font-weight:400;font-size:10px;color:#BBDEFB;">Đang tăng mạnh ngắn hạn, chưa đủ chất lượng dài hạn.</span>
          </td>
        </tr>
      </table><br>
      Mỗi ô: <b>Ticker · CS/QC · 1Y%</b> — đọc nhanh không cần mở sheet chính.
    </td>
  </tr>

  <tr style="background-color:#F0F4FA;">
    <td style="color:#374151;font-weight:700;font-size:13px;padding:10px 14px;">❹ Sector Breakdown</td>
    <td style="color:#1A202C;font-size:12px;padding:10px 14px;">
      <b>Phân tích theo ngành — tìm ngành đang dẫn đầu thị trường.</b><br><br>
      Thống kê theo sector: # Stocks · # Strong Buy · # Compounder · Avg CS · Avg QC · Avg 1Y%<br>
      Sắp xếp theo # Strong Buy giảm dần. Top 3 ngành được tô màu nổi bật.<br><br>
      <b>Tại sao sector quan trọng?</b> Thị trường di chuyển theo nhóm ngành (sector rotation).
      Khi Technology dẫn đầu, hầu hết cổ phiếu công nghệ đều tăng — ngay cả mã bình thường cũng được kéo theo.<br><br>
      <b>Mẹo:</b> Ngành có Avg 1Y% cao + nhiều Strong Buy = sector đang trending.
      Tìm Dual Leaders trong sector đó để có xác suất tốt cao nhất.
    </td>
  </tr>
</table>

<!-- ══ QUY TRÌNH ĐỌC ══ -->
<p style="color:#374151;font-size:12px;margin:14px 0 6px 0;font-weight:700;letter-spacing:1px;">③ QUY TRÌNH ĐỌC DASHBOARD — 5 BƯỚC</p>
<ol style="padding-left:20px;margin:0;">
  <li style="margin:8px 0;color:#1A202C;font-size:12px;">
    <b>Bắt đầu từ ❶+❷ KPIs</b> — "Thị trường hôm nay tốt hay xấu?"
    # Strong Buy &lt;5% tổng → thận trọng. Avg CS &gt;6 và # Dual Leaders &gt;5 → thị trường đang chất lượng.
  </li>
  <li style="margin:8px 0;color:#1A202C;font-size:12px;">
    <b>Kiểm tra ⚠ Risk Flags</b> — Nhiều cảnh báo (D/E cao, 1M% âm) → đừng vội vào tiền.
    Kiên nhẫn đợi thêm 1–2 tuần rồi scan lại.
  </li>
  <li style="margin:8px 0;color:#1A202C;font-size:12px;">
    <b>Xem ⭐ Dual Leaders</b> — Danh sách ngắn, chất lượng cao nhất.
    Chỉ có thời gian nghiên cứu 2–3 mã → lấy từ bảng này. Chú ý cột ⚠.
  </li>
  <li style="margin:8px 0;color:#1A202C;font-size:12px;">
    <b>Dùng 🎯 Top Picks theo chiến lược của bạn</b> —
    Ngắn hạn → 🚀 Momentum hoặc 💥 Breakout.
    Dài hạn, ít theo dõi → 💎 Quality hoặc 📉 Value Growth.
  </li>
  <li style="margin:8px 0;color:#1A202C;font-size:12px;">
    <b>Xác nhận bằng ❸+❹</b> — Mã bạn chọn đang ở vùng nào trong Chart Analysis?
    Ngành của mã đó có trong top 3 Sector Breakdown không? Có → xác suất tốt cao hơn.
  </li>
</ol>

<p style="color:#374151;font-size:12px;margin:14px 0 6px 0;font-weight:700;letter-spacing:1px;">④ LỖI PHỔ BIẾN CẦN TRÁNH</p>
<ul style="padding-left:20px;margin:0;">
  <li style="margin:6px 0;color:#1A202C;font-size:12px;"><b>❌ Chỉ nhìn 1Y% cao mà mua ngay</b> — Cổ phiếu tăng 200%/năm không có nghĩa là tiếp tục tăng. Phải kết hợp CS Score và Risk Flags.</li>
  <li style="margin:6px 0;color:#1A202C;font-size:12px;"><b>❌ Bỏ qua D/E và cột ⚠</b> — Công ty nợ nhiều, khi lãi suất tăng hoặc doanh thu giảm, có thể sụp đổ rất nhanh.</li>
  <li style="margin:6px 0;color:#1A202C;font-size:12px;"><b>❌ Mua tất cả mã trong Dual Leaders</b> — Dual Leaders là shortlist để nghiên cứu sâu hơn, không phải lệnh mua tự động.</li>
  <li style="margin:6px 0;color:#1A202C;font-size:12px;"><b>❌ Bỏ qua Sector Breakdown</b> — Mã tốt trong ngành đang yếu vẫn có thể bị kéo xuống. Luôn kiểm tra sector trend.</li>
  <li style="margin:6px 0;color:#1A202C;font-size:12px;"><b>❌ Dùng dữ liệu cũ</b> — Dashboard phản ánh thời điểm scan gần nhất. Scan lại ít nhất mỗi tuần 1 lần.</li>
</ul>

<p class="tip">⚠ Dashboard chỉ xuất hiện khi nhấn <b>Export Excel</b>. Thanh tiến trình xanh ở status bar cho biết quá trình xuất — UI không bị đơ. Dữ liệu phản ánh thời điểm scan gần nhất — nên scan lại ít nhất mỗi tuần 1 lần để cập nhật thị trường.</p>
</body></html>"""


# ─────────────────────────────────────────────────────────────────────────────
# ─────────────────────────────────────────────────────────────────────────────
# Export worker — chạy write_excel trên background thread
# ─────────────────────────────────────────────────────────────────────────────
class ExportWorker(QThread):
    progress = Signal(int, str)   # (0-100, message)
    done     = Signal(str)        # output path
    failed   = Signal(str)        # error message

    def __init__(self, df, path, market, top, use_yf):
        super().__init__()
        self._df     = df
        self._path   = path
        self._market = market
        self._top    = top
        self._use_yf = use_yf

    def run(self):
        try:
            def cb(pct, msg=""):
                self.progress.emit(pct, msg)
            write_excel(self._df, self._path, self._market, self._top,
                        use_yf=self._use_yf, progress_cb=cb)
            self.done.emit(self._path)
        except Exception as e:
            self.failed.emit(str(e))


# Main window
# ─────────────────────────────────────────────────────────────────────────────
class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Fundamental Screener")
        self.resize(1700, 960)
        self._is_dark     = False
        globals().update(LIGHT_THEME)
        self._df          = None
        self._worker      = None
        self._ticker_wkr  = None
        self._spin_idx    = 0
        self._spin_msg    = ""
        self._spin_tmr    = QTimer(self)
        self._spin_tmr.timeout.connect(self._tick)
        self._lookup_tmr  = QTimer(self)
        self._lookup_tmr.setSingleShot(True)
        self._lookup_tmr.timeout.connect(self._do_lookup)

        QShortcut(QKeySequence(Qt.Key_F1), self).activated.connect(self._show_help)

        self.setStyleSheet(f"""
            QMainWindow {{ background:{BG}; }}
            * {{ background:transparent; color:{TEXT1}; }}
            QScrollBar:vertical {{ background:{SURFACE}; width:8px; border-radius:4px; }}
            QScrollBar::handle:vertical {{ background:{BORDER2}; border-radius:4px; }}
            QScrollBar:horizontal {{ background:{SURFACE}; height:8px; border-radius:4px; }}
            QScrollBar::handle:horizontal {{ background:{BORDER2}; border-radius:4px; }}
            QTableWidget {{ gridline-color:{BORDER}; outline:none; }}
            QTableWidget::item:selected {{ background:{BORDER2}; }}
            QToolTip {{
                background:{SURFACE}; color:{TEXT1};
                border:1px solid {BORDER2}; font-size:10px; padding:4px;
            }}
        """)

        root = QWidget()
        self.setCentralWidget(root)
        vbox = QVBoxLayout(root)
        vbox.setContentsMargins(0, 0, 0, 0)
        vbox.setSpacing(0)

        vbox.addWidget(self._build_header())
        vbox.addWidget(self._sep())
        vbox.addWidget(self._build_controls())
        vbox.addWidget(self._sep())

        # Right: chart panel (always visible)
        self._chart_panel = ChartPanel()

        # Left: tab widget (CAN SLIM | Quality Compounder)
        self._tabs_w = self._build_tabs()
        self._tabs_w.currentChanged.connect(self._on_tab_changed)

        # Horizontal splitter: tabs | chart
        h_split = QSplitter(Qt.Horizontal)
        h_split.setStyleSheet(
            f"QSplitter::handle {{ background:{BORDER}; width:2px; }}")
        h_split.addWidget(self._tabs_w)
        h_split.addWidget(self._chart_panel)
        h_split.setSizes([920, 580])
        # Dashboard là tab mặc định (index 0) → ẩn chart panel ngay từ đầu
        self._chart_panel.setVisible(False)
        vbox.addWidget(h_split, stretch=1)

        vbox.addWidget(self._sep())
        vbox.addWidget(self._build_status())

    # ── Section builders ──────────────────────────────────────────────────────

    def _build_header(self):
        self._header_w = QWidget(); self._header_w.setFixedHeight(50)
        self._header_w.setStyleSheet(f"background:{SURFACE};")
        h = QHBoxLayout(self._header_w); h.setContentsMargins(22, 0, 16, 0); h.setSpacing(0)
        lbl = QLabel()
        lbl.setTextFormat(Qt.RichText)
        lbl.setText(
            f'<span style="color:{TEXT1};font-size:14px;font-weight:700;letter-spacing:4px">FUNDAMENTAL</span>'
            f'<span style="color:{BLUE};font-size:14px;font-weight:700;letter-spacing:4px"> SCREENER</span>'
        )
        self._title_lbl = lbl
        h.addWidget(lbl); h.addStretch()
        self._tag_lbl = QLabel("Catalyst · TradingView · yfinance")
        self._tag_lbl.setStyleSheet(f"color:{TEXT3}; font-size:9px; letter-spacing:2px;")
        h.addWidget(self._tag_lbl)
        h.addSpacing(10)

        self._btn_theme = QPushButton("☀")
        self._btn_theme.setFixedSize(28, 28)
        self._btn_theme.setToolTip("Switch Light / Dark")
        self._btn_theme.setCursor(Qt.PointingHandCursor)
        self._btn_theme.clicked.connect(self._toggle_theme)
        h.addWidget(self._btn_theme)
        h.addSpacing(6)

        self._btn_help = QPushButton("?")
        self._btn_help.setFixedSize(28, 28)
        self._btn_help.setToolTip("Help  (F1)")
        self._btn_help.setCursor(Qt.PointingHandCursor)
        self._btn_help.clicked.connect(self._show_help)
        h.addWidget(self._btn_help)
        self._refresh_icon_btns()
        return self._header_w

    def _build_controls(self):
        w = QWidget(); w.setFixedHeight(118)
        self._controls_w = w
        w.setStyleSheet(f"background:{BG};")
        v = QVBoxLayout(w); v.setContentsMargins(22, 6, 22, 6); v.setSpacing(5)

        # Row 1: scan controls
        row1 = QHBoxLayout(); row1.setSpacing(10)

        lbl_market = QLabel("MARKET")
        self._lbl_market = lbl_market
        lbl_market.setStyleSheet(
            f"color:{TEXT1}; font-size:9px; font-weight:600; letter-spacing:1.5px;")
        self._market = QComboBox()
        self._market.addItems(["AMERICA", "NASDAQ", "NYSE", "EURONEXT", "HONG_KONG", "VIETNAM"])
        self._market.setFixedHeight(32)
        self._market.setStyleSheet(self._combo_style())

        lbl_top = QLabel("TOP")
        self._lbl_top = lbl_top
        lbl_top.setStyleSheet(
            f"color:{TEXT1}; font-size:9px; font-weight:600; letter-spacing:1.5px;")
        self._top_spin = QSpinBox()
        self._top_spin.setRange(50, 1000); self._top_spin.setValue(300); self._top_spin.setSingleStep(50)
        self._top_spin.setFixedSize(60, 32)
        self._top_spin.setButtonSymbols(QSpinBox.ButtonSymbols.NoButtons)
        self._top_spin.setStyleSheet(self._spinbox_style())

        self._spin_minus = QPushButton("−")
        self._spin_plus  = QPushButton("+")
        for btn in (self._spin_minus, self._spin_plus):
            btn.setFixedSize(24, 32)
            btn.setCursor(Qt.PointingHandCursor)
            btn.setStyleSheet(self._spin_btn_style())
        self._spin_minus.clicked.connect(lambda: self._top_spin.setValue(self._top_spin.value() - self._top_spin.singleStep()))
        self._spin_plus.clicked.connect( lambda: self._top_spin.setValue(self._top_spin.value() + self._top_spin.singleStep()))

        self._use_yf = QCheckBox("yfinance Moat")
        self._use_yf.setChecked(True)
        self._use_yf.setStyleSheet(
            f"color:{TEXT1}; font-size:10px; font-family:'Segoe UI',sans-serif;")

        self._btn_scan = QPushButton("▶  SCAN")
        self._btn_scan.setFixedSize(110, 32)
        self._btn_scan.setCursor(Qt.PointingHandCursor)
        self._btn_scan.setStyleSheet(self._btn_style(BLUE, BLUE_HV))
        self._btn_scan.clicked.connect(self._on_scan)

        self._btn_export = QPushButton("💾  Export")
        self._btn_export.setFixedSize(95, 32)
        self._btn_export.setCursor(Qt.PointingHandCursor)
        self._btn_export.setEnabled(False)
        self._btn_export.setStyleSheet(self._btn_style("#2E5C2E", "#1A4C1A"))
        self._btn_export.clicked.connect(self._on_export)

        row1.addWidget(lbl_market); row1.addWidget(self._market)
        row1.addWidget(lbl_top)
        row1.addWidget(self._spin_minus)
        row1.addWidget(self._top_spin)
        row1.addWidget(self._spin_plus)
        row1.addWidget(self._use_yf)
        row1.addStretch()
        row1.addWidget(self._btn_scan)
        row1.addWidget(self._btn_export)
        v.addLayout(row1)

        # Row 2: ticker lookup + table filter
        row2 = QHBoxLayout(); row2.setSpacing(10)

        lbl_tk = QLabel("TICKER")
        self._lbl_tk = lbl_tk
        lbl_tk.setStyleSheet(
            f"color:{TEXT1}; font-size:9px; font-weight:600; letter-spacing:1.5px;")
        self._ticker_inp = QLineEdit()
        self._ticker_inp.setPlaceholderText("e.g. AAPL  · NVDA · MSFT")
        self._ticker_inp.setFixedHeight(28); self._ticker_inp.setMaxLength(12)
        self._ticker_inp.setStyleSheet(self._input_style(font="Consolas,monospace", size="13px", ls="3px"))
        self._ticker_inp.returnPressed.connect(self._do_lookup)
        self._ticker_inp.textChanged.connect(self._on_ticker_typed)

        self._btn_lookup = QPushButton("LOOKUP")
        self._btn_lookup.setFixedSize(80, 28)
        self._btn_lookup.setCursor(Qt.PointingHandCursor)
        self._btn_lookup.setStyleSheet(self._btn_style(BLUE, BLUE_HV, size="9px"))
        self._btn_lookup.clicked.connect(self._do_lookup)

        lbl_filter = QLabel("FILTER")
        self._lbl_filter = lbl_filter
        lbl_filter.setStyleSheet(
            f"color:{TEXT1}; font-size:9px; font-weight:600; letter-spacing:1.5px;")
        self._filter_inp = QLineEdit()
        self._filter_inp.setPlaceholderText("Search ticker or company name…")
        self._filter_inp.setFixedHeight(28)
        self._filter_inp.setStyleSheet(self._input_style())
        self._filter_inp.textChanged.connect(self._apply_filter)

        self._lbl_count = QLabel("0 / 0  rows")
        self._lbl_count.setStyleSheet(
            f"color:{TEXT3}; font-size:9px; letter-spacing:1px;")

        row2.addWidget(lbl_tk); row2.addWidget(self._ticker_inp)
        row2.addWidget(self._btn_lookup)
        sep = QFrame(); sep.setFrameShape(QFrame.VLine)
        sep.setStyleSheet(f"background:{BORDER}; max-width:1px; margin:2px 4px;")
        row2.addWidget(sep)
        row2.addWidget(lbl_filter); row2.addWidget(self._filter_inp)
        row2.addStretch(); row2.addWidget(self._lbl_count)
        v.addLayout(row2)

        # Row 3: quick filter chips
        row3 = QHBoxLayout(); row3.setSpacing(10)
        lbl_qf = QLabel("QUICK FILTER")
        lbl_qf.setStyleSheet(
            f"color:{TEXT1}; font-size:9px; font-weight:600; letter-spacing:1.5px;")
        self._lbl_qf = lbl_qf

        self._chk_sb   = QCheckBox("🟢  STRONG BUY")
        self._chk_moat = QCheckBox("🏰  WIDE / NARROW Moat")
        self._chk_1y   = QCheckBox("↑  1Y% > 0")
        self._chk_52w  = QCheckBox("📈  52W High ≥ 90%")
        self._chip_checks = (self._chk_sb, self._chk_moat, self._chk_1y, self._chk_52w)

        for chk in self._chip_checks:
            chk.setStyleSheet(self._chip_style())
            chk.stateChanged.connect(self._apply_filter)

        self._btn_clr = QPushButton("✕")
        self._btn_clr.setFixedSize(24, 22)
        self._btn_clr.setToolTip("Clear all filters")
        self._btn_clr.setCursor(Qt.PointingHandCursor)
        self._btn_clr.setStyleSheet(self._btn_style(PANEL, BORDER2, "9px"))
        self._btn_clr.clicked.connect(self._clear_chips)

        sep2 = QFrame(); sep2.setFrameShape(QFrame.VLine)
        sep2.setStyleSheet(f"background:{BORDER}; max-width:1px; margin:2px 4px;")
        row3.addWidget(lbl_qf); row3.addWidget(sep2)
        for chk in self._chip_checks:
            row3.addWidget(chk)
        row3.addStretch(); row3.addWidget(self._btn_clr)
        v.addLayout(row3)
        return w

    def _build_table(self):
        w = QWidget(); w.setStyleSheet(f"background:{BG};")
        v = QVBoxLayout(w); v.setContentsMargins(0, 0, 0, 0); v.setSpacing(0)
        self._table = QTableWidget(0, len(TABLE_COLS))
        self._table.setStyleSheet(f"""
            QTableWidget {{
                background:{BG}; alternate-background-color:{PANEL};
                font-size:11px; font-family:'Segoe UI',sans-serif;
                selection-background-color:{BORDER2};
                border:none;
            }}
        """)
        self._table.setAlternatingRowColors(True)
        self._table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self._table.setSelectionMode(QAbstractItemView.SingleSelection)
        self._table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self._table.setSortingEnabled(True)
        self._table.verticalHeader().setVisible(False)
        self._table.verticalHeader().setDefaultSectionSize(22)
        self._table.setHorizontalScrollMode(QAbstractItemView.ScrollPerPixel)
        self._table.setVerticalScrollMode(QAbstractItemView.ScrollPerPixel)
        _hdr = ColoredHeader(self._table)
        self._table.setHorizontalHeader(_hdr)

        headers = [c[0] for c in TABLE_COLS]
        self._table.setHorizontalHeaderLabels(headers)
        for i, (_, _, w_px, _) in enumerate(TABLE_COLS):
            self._table.setColumnWidth(i, w_px)

        self._color_header_cols()

        self._table.itemSelectionChanged.connect(self._on_row_selected)
        v.addWidget(self._table)
        return w

    def _build_detail(self):
        self._detail = DetailCard()
        return self._detail

    # ── Tab widget ────────────────────────────────────────────────────────────

    def _build_tabs(self):
        tabs = QTabWidget()
        tabs.setDocumentMode(True)
        tabs.setStyleSheet(self._tab_style())

        # Tab 0 — Dashboard
        tabs.addTab(self._build_dashboard_panel(), "  📊 Dashboard  ")

        # Tab 1 — CAN SLIM
        self._cs_w = QWidget(); self._cs_w.setStyleSheet(f"background:{BG};")
        cs_v = QVBoxLayout(self._cs_w)
        cs_v.setContentsMargins(0, 0, 0, 0); cs_v.setSpacing(0)
        self._cs_vsplit = QSplitter(Qt.Vertical)
        self._cs_vsplit.setStyleSheet(f"QSplitter::handle {{ background:{BORDER}; height:2px; }}")
        self._cs_vsplit.addWidget(self._build_table())
        self._cs_vsplit.addWidget(self._build_detail())
        self._cs_vsplit.setSizes([580, 180])
        cs_v.addWidget(self._cs_vsplit)
        tabs.addTab(self._cs_w, "  Catalyst  ")

        # Tab 2 — Quality Compounder
        self._qc_w = QWidget(); self._qc_w.setStyleSheet(f"background:{BG};")
        qc_v = QVBoxLayout(self._qc_w)
        qc_v.setContentsMargins(0, 0, 0, 0); qc_v.setSpacing(0)
        self._qc_vsplit = QSplitter(Qt.Vertical)
        self._qc_vsplit.setStyleSheet(f"QSplitter::handle {{ background:{BORDER}; height:2px; }}")
        self._qc_vsplit.addWidget(self._build_qc_table())
        self._qc_vsplit.addWidget(self._build_qc_detail())
        self._qc_vsplit.setSizes([580, 180])
        qc_v.addWidget(self._qc_vsplit)
        tabs.addTab(self._qc_w, "  Quality Compounder  ")

        return tabs

    def _build_dashboard_panel(self):
        w = QWidget(); w.setStyleSheet(f"background:{BG};")
        lay = QVBoxLayout(w); lay.setContentsMargins(0,0,0,0); lay.setSpacing(0)
        self._dash_browser = QTextBrowser()
        self._dash_browser.setOpenExternalLinks(False)
        self._dash_browser.setStyleSheet(f"""
            QTextBrowser {{
                background:{BG}; color:#DCE4EE;
                border:none; padding:16px 20px;
                font-family:'Segoe UI',sans-serif;
            }}
        """)
        self._dash_browser.setHtml(self._dash_placeholder())
        lay.addWidget(self._dash_browser)
        return w

    def _dash_placeholder(self):
        return f"""<html><body style="background:{BG};color:#4A6080;
            font-family:'Segoe UI',sans-serif;text-align:center;padding-top:120px;">
            <p style="font-size:48px;margin:0;">📊</p>
            <p style="font-size:18px;font-weight:700;color:#3D5070;margin:12px 0 6px 0;">
              Dashboard chưa có dữ liệu</p>
            <p style="font-size:13px;color:#2A3A50;">
              Nhấn <b style="color:#3D8EF0;">▶ SCAN</b> để tải dữ liệu và xem Dashboard</p>
        </body></html>"""

    def _tab_style(self):
        return f"""
            QTabWidget::pane {{
                border:none;
                background:{BG};
            }}
            QTabBar {{
                background:{SURFACE};
            }}
            QTabBar::tab {{
                background:{SURFACE}; color:{TEXT3};
                border:none;
                border-right:1px solid {BORDER};
                border-bottom:2px solid transparent;
                padding:9px 22px; font-size:9px; font-weight:700;
                letter-spacing:2px; font-family:'Segoe UI',sans-serif;
                min-width:80px;
            }}
            QTabBar::tab:selected {{
                background:{BG}; color:{TEXT1};
                border-bottom:2px solid {BLUE};
            }}
            QTabBar::tab:hover:!selected {{
                color:{TEXT2};
            }}
        """

    # ── QC table ──────────────────────────────────────────────────────────────

    def _build_qc_table(self):
        w = QWidget(); w.setStyleSheet(f"background:{BG};")
        v = QVBoxLayout(w); v.setContentsMargins(0, 0, 0, 0); v.setSpacing(0)
        self._qc_table = QTableWidget(0, len(QC_COLS))
        self._qc_table.setStyleSheet(f"""
            QTableWidget {{
                background:{BG}; alternate-background-color:{PANEL};
                font-size:11px; font-family:'Segoe UI',sans-serif;
                selection-background-color:{BORDER2};
                border:none;
            }}
        """)
        self._qc_table.setAlternatingRowColors(True)
        self._qc_table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self._qc_table.setSelectionMode(QAbstractItemView.SingleSelection)
        self._qc_table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self._qc_table.setSortingEnabled(True)
        self._qc_table.verticalHeader().setVisible(False)
        self._qc_table.verticalHeader().setDefaultSectionSize(22)
        self._qc_table.setHorizontalScrollMode(QAbstractItemView.ScrollPerPixel)
        self._qc_table.setVerticalScrollMode(QAbstractItemView.ScrollPerPixel)
        _hdr = ColoredHeader(self._qc_table)
        self._qc_table.setHorizontalHeader(_hdr)
        self._qc_table.setHorizontalHeaderLabels([c[0] for c in QC_COLS])
        for i, (_, _, w_px, _) in enumerate(QC_COLS):
            self._qc_table.setColumnWidth(i, w_px)
        self._color_qc_header_cols()
        self._qc_table.itemSelectionChanged.connect(self._on_qc_row_selected)
        v.addWidget(self._qc_table)
        return w

    def _build_qc_detail(self):
        self._qc_detail = QualityDetailCard()
        return self._qc_detail

    def _color_qc_header_cols(self):
        dark = self._is_dark
        _qc_bg = "#0D1F10" if dark else "#E8F5E9"
        _qc_fg = "#5EC472" if dark else "#276749"
        _mo_bg = "#1A0D2E" if dark else "#F3E8FF"
        _mo_fg = "#B07ED4" if dark else "#7C3AED"
        _fu_bg = "#0D1828" if dark else "#EEF2FB"
        _fu_fg = "#7EB8D4" if dark else "#1E40AF"
        for i, (_, key, _, _) in enumerate(QC_COLS):
            item = self._qc_table.horizontalHeaderItem(i)
            if not item:
                continue
            if key.startswith("QC_"):
                item.setBackground(QColor(_qc_bg)); item.setForeground(QColor(_qc_fg))
            elif key == "Moat Score":
                item.setBackground(QColor(_mo_bg)); item.setForeground(QColor(_mo_fg))
            elif key in ("Gross Margin%", "Op Margin%", "ROIC%", "FCF_Margin%", "D/E",
                         "Net Cash ($B)", "FCF/sh", "Current Ratio", "EV/EBITDA", "P/E", "1Y%"):
                item.setBackground(QColor(_fu_bg)); item.setForeground(QColor(_fu_fg))
            elif key == "EQ_Badge":
                _eq_hbg = "#0D2020" if dark else "#E0F7FA"
                _eq_hfg = "#4DB6AC" if dark else "#00695C"
                item.setBackground(QColor(_eq_hbg)); item.setForeground(QColor(_eq_hfg))
        self._qc_table.horizontalHeader().update()

    def _populate_qc_table(self, df: pd.DataFrame):
        self._qc_data = {}
        self._qc_table.setSortingEnabled(False)
        self._qc_table.setRowCount(0)
        self._qc_table.setRowCount(len(df))

        # Columns that are %-type and get green/red coloring
        PCT_GREEN_RED = {"Gross Margin%", "Op Margin%", "ROIC%", "1Y%", "FCF_Margin%"}
        is_vn = self._market.currentText().lower() == "vietnam"

        for ri, (_, row) in enumerate(df.iterrows()):
            rd = row.to_dict()
            qc = compute_qc_score(rd)
            rd.update(qc)
            ticker = rd.get("Ticker", "")
            if ticker:
                self._qc_data[ticker] = rd

            for ci, (_, key, _, align) in enumerate(QC_COLS):
                val  = rd.get(key)
                cell = None

                if key == "_no_":
                    cell = NumItem(ri + 1, "int")
                    cell.setForeground(QColor(TEXT1))
                elif key == "Ticker":
                    cell = QTableWidgetItem(str(val) if val else "")
                    cell.setFont(QFont("Consolas", 9, QFont.Bold))
                    cell.setForeground(QColor("#00BFFF"))
                elif key in ("Tên Công Ty", "Sector"):
                    cell = QTableWidgetItem(str(val) if val else "")
                    cell.setForeground(QColor(TEXT1))
                elif key == "Moat Score":
                    cell = QTableWidgetItem(str(val) if val else "—")
                    if val in MOAT_SCORE_STYLE:
                        fg, bg = MOAT_SCORE_STYLE[val]
                        cell.setForeground(QColor(f"#{fg}"))
                        cell.setBackground(QColor(f"#{bg}"))
                elif key == "EQ_Badge":
                    _eq_pal = {
                        "💚 Cash Backed":   ("#1A5C2B", "#C6EFCE"),
                        "🟡 Mixed":         ("#7D6608", "#FFF2CC"),
                        "🔴 Accrual Heavy": ("#9C0006", "#FFC7CE"),
                    }
                    cell = QTableWidgetItem(str(val) if val else "—")
                    if val in _eq_pal:
                        fg, bg = _eq_pal[val]
                        cell.setForeground(QColor(fg))
                        cell.setBackground(QColor(bg))
                    else:
                        cell.setForeground(QColor(TEXT3))
                elif key == "QC_Signal":
                    cell = QTableWidgetItem(str(val) if val else "")
                    if val in QC_SIGNAL_COLORS:
                        fg, bg = QC_SIGNAL_COLORS[val]
                        cell.setForeground(QColor(fg))
                        cell.setBackground(QColor(bg))
                elif key == "QC_Score":
                    score = int(val) if val is not None else 0
                    cell  = NumItem(score, "int")
                    pct   = score / N_QC
                    c  = "#1A5C2B" if pct >= 5/6 else "#1B3A5C" if pct >= 3/6 \
                         else "#7D6608" if pct >= 1/6 else "#9C0006"
                    bg = "#C6EFCE" if pct >= 5/6 else "#DDEEFF" if pct >= 3/6 \
                         else "#FFF2CC" if pct >= 1/6 else "#FFC7CE"
                    cell.setForeground(QColor(c)); cell.setBackground(QColor(bg))
                elif key.startswith("QC_"):
                    if val is True:
                        cell = QTableWidgetItem("✓")
                        cell.setForeground(QColor("#1A5C2B"))
                        cell.setBackground(QColor("#C6EFCE"))
                    elif val is False:
                        cell = QTableWidgetItem("✗")
                        cell.setForeground(QColor("#9C0006"))
                        cell.setBackground(QColor("#FFC7CE"))
                    else:
                        cell = QTableWidgetItem("—")
                        cell.setForeground(QColor(TEXT3))
                elif key == "Price ($)":
                    cell = NumItem(val, "vnd" if is_vn else "price")
                elif key == "MCap ($B)":
                    cell = NumItem(val, "mcap_vnd" if is_vn else "mcap")
                elif key == "Net Cash ($B)":
                    cell = NumItem(val)
                    if isinstance(val, float) and val == val:
                        if val > 5:
                            cell.setForeground(QColor("#1A5C2B"))
                            cell.setBackground(QColor("#C6EFCE"))
                        elif val > 0:
                            cell.setForeground(QColor("#1A5C2B"))
                        elif val > -5:
                            cell.setForeground(QColor("#7D6608"))
                        else:
                            cell.setForeground(QColor(RED))
                    else:
                        cell.setForeground(QColor(TEXT3))

                elif key == "FCF_Margin%":
                    cell = NumItem(val, "pct")
                    if isinstance(val, float):
                        if val >= 25:
                            cell.setForeground(QColor("#1A5C2B"))
                            cell.setBackground(QColor("#D5F5E3"))
                        elif val >= 15:
                            cell.setForeground(QColor("#1A5C2B"))
                        elif val < 0:
                            cell.setForeground(QColor(RED))
                        else:
                            cell.setForeground(QColor(TEXT3))

                elif key in PCT_GREEN_RED:
                    cell = NumItem(val, "pct")
                    if isinstance(val, float):
                        cell.setForeground(QColor(GREEN) if val > 0 else
                                           QColor(RED)   if val < 0 else
                                           QColor(TEXT3))
                else:
                    cell = NumItem(val)

                if cell:
                    cell.setTextAlignment(int(align))
                    self._qc_table.setItem(ri, ci, cell)

        self._qc_table.setSortingEnabled(True)
        self._apply_qc_filter()

    def _recolor_qc_table(self):
        PCT_GREEN_RED = {"Gross Margin%", "Op Margin%", "ROIC%", "1Y%", "FCF_Margin%"}
        for ri in range(self._qc_table.rowCount()):
            for ci, (_, key, _, _) in enumerate(QC_COLS):
                item = self._qc_table.item(ri, ci)
                if not item:
                    continue
                if key in ("_no_", "Tên Công Ty", "Sector"):
                    item.setForeground(QColor(TEXT1))
                elif key in PCT_GREEN_RED and hasattr(item, "_num"):
                    v = item._num
                    if v != float("-inf"):
                        item.setForeground(QColor(GREEN) if v > 0 else
                                           QColor(RED)   if v < 0 else
                                           QColor(TEXT3))

    def _on_qc_row_selected(self):
        rows = self._qc_table.selectedItems()
        if not rows:
            return
        ri = self._qc_table.currentRow()
        ticker_item = self._qc_table.item(ri, 1)
        if not ticker_item:
            return
        ticker = ticker_item.text()
        rd = getattr(self, "_qc_data", {}).get(ticker)
        if rd:
            self._qc_detail.show_row(rd)
            self._ticker_inp.setText(ticker)
            self._chart_panel.load_ticker(ticker)

    def _on_tab_changed(self, idx: int):
        # Tab 0 = Dashboard → full width, ẩn chart panel
        is_dash = (idx == 0)
        if hasattr(self, "_chart_panel"):
            self._chart_panel.setVisible(not is_dash)
        self._apply_filter()

    def _apply_qc_filter(self):
        text      = self._filter_inp.text().strip().lower()
        only_sb   = self._chk_sb.isChecked()    # maps to COMPOUNDER
        only_moat = self._chk_moat.isChecked()
        only_1y   = self._chk_1y.isChecked()
        only_52w  = self._chk_52w.isChecked()
        any_chip  = only_sb or only_moat or only_1y or only_52w
        qd        = getattr(self, "_qc_data", {})

        visible = 0
        total   = self._qc_table.rowCount()
        for ri in range(total):
            show = True
            if text:
                t_item = self._qc_table.item(ri, 1)
                n_item = self._qc_table.item(ri, 2)
                t_txt  = (t_item.text() if t_item else "").lower()
                n_txt  = (n_item.text() if n_item else "").lower()
                show   = (text in t_txt) or (text in n_txt)
            if show and any_chip:
                ticker = (self._qc_table.item(ri, 1).text()
                          if self._qc_table.item(ri, 1) else "")
                rd = qd.get(ticker, {})
                if only_sb and rd.get("QC_Signal") != "🏆 COMPOUNDER":
                    show = False
                if show and only_moat and rd.get("Moat Score") not in (
                        "WIDE  ★★★", "NARROW ★★"):
                    show = False
                if show and only_1y:
                    v = rd.get("1Y%")
                    if v is None or (isinstance(v, float) and pd.isna(v)) or v <= 0:
                        show = False
                if show and only_52w:
                    v = rd.get("52W_High%")
                    if v is None or (isinstance(v, float) and pd.isna(v)) or v < 90:
                        show = False
            self._qc_table.setRowHidden(ri, not show)
            if show:
                visible += 1
        self._lbl_count.setText(f"{visible} / {total}  rows")

    def _build_status(self):
        w = QWidget(); w.setFixedHeight(28)
        self._status_w = w
        w.setStyleSheet(f"background:{SURFACE};")
        h = QHBoxLayout(w); h.setContentsMargins(22, 0, 22, 0)
        self._dot = QLabel("●"); self._dot.setFixedWidth(14)
        self._dot.setStyleSheet(f"color:{TEXT3}; font-size:7px;")
        self.status_lbl = QLabel("Ready  ·  Enter market & top N, then click SCAN")
        self.status_lbl.setStyleSheet(
            f"color:{TEXT3}; font-size:10px; font-family:'Segoe UI',sans-serif;")
        self._export_bar = QProgressBar()
        self._export_bar.setFixedSize(180, 10)
        self._export_bar.setTextVisible(False)
        self._export_bar.setRange(0, 100)
        self._export_bar.setValue(0)
        self._export_bar.setVisible(False)
        self._export_bar.setStyleSheet("""
            QProgressBar { background:#333; border-radius:4px; }
            QProgressBar::chunk { background:#4AE06A; border-radius:4px; }
        """)
        h.addWidget(self._dot); h.addWidget(self.status_lbl)
        h.addWidget(self._export_bar)
        h.addStretch()

        self._ver_lbl = QLabel("TradingView  ·  yfinance  ·  PySide6")
        self._ver_lbl.setStyleSheet(f"color:{TEXT3}; font-size:9px; letter-spacing:1px;")
        h.addWidget(self._ver_lbl)
        return w

    # ── Theme ─────────────────────────────────────────────────────────────────

    def _refresh_icon_btns(self):
        _s = f"""
            QPushButton {{
                background:{BORDER2}; color:{TEXT1};
                border:none; border-radius:14px;
                font-size:12px; font-weight:700;
                font-family:'Segoe UI',sans-serif;
            }}
            QPushButton:hover {{ background:{BLUE}; color:#FFFFFF; }}
        """
        self._btn_theme.setStyleSheet(_s)
        self._btn_help.setStyleSheet(_s)

    def _color_header_cols(self):
        dark = self._is_dark
        _cs_bg  = "#0D1F10" if dark else "#E8F5E9"
        _cs_fg  = "#5EC472" if dark else "#276749"
        _mo_bg  = "#1A0D2E" if dark else "#F3E8FF"
        _mo_fg  = "#B07ED4" if dark else "#7C3AED"
        _fu_bg  = "#0D1828" if dark else "#EEF2FB"
        _fu_fg  = "#7EB8D4" if dark else "#1E40AF"
        _keys = [c[1] for c in TABLE_COLS]
        for i, key in enumerate(_keys):
            item = self._table.horizontalHeaderItem(i)
            if not item:
                continue
            if key.startswith("CS_") or key in ("CS_Score", "CS_Signal"):
                item.setBackground(QColor(_cs_bg)); item.setForeground(QColor(_cs_fg))
            elif key in ("Moat Score", "Moat Proxy"):
                item.setBackground(QColor(_mo_bg)); item.setForeground(QColor(_mo_fg))
            elif key in ("EPS Qtr%", "EPS Annual%", "Rev Qtr%", "Gross Margin%",
                         "ROE%", "D/E", "P/E", "3M%", "1Y%"):
                item.setBackground(QColor(_fu_bg)); item.setForeground(QColor(_fu_fg))
        self._table.horizontalHeader().update()

    def _toggle_theme(self):
        self._is_dark = not self._is_dark
        theme = DARK_THEME if self._is_dark else LIGHT_THEME
        globals().update(theme)
        self._btn_theme.setText("☀" if self._is_dark else "🌙")
        self._apply_theme()

    def _apply_theme(self):
        dark = self._is_dark
        # ── ColoredHeader ──
        ColoredHeader._DEF_BG   = QColor("#0A1628" if dark else "#E2E8F0")
        ColoredHeader._DEF_FG   = QColor("#7EB8D4" if dark else "#2D3748")
        ColoredHeader._BORDER_B = QColor(BLUE)
        ColoredHeader._BORDER_R = QColor(BORDER)
        self._color_header_cols()

        # ── Main window stylesheet (scrollbars, tooltip) ──
        self.setStyleSheet(f"""
            QMainWindow {{ background:{BG}; }}
            * {{ background:transparent; color:{TEXT1}; }}
            QScrollBar:vertical {{ background:{SURFACE}; width:8px; border-radius:4px; }}
            QScrollBar::handle:vertical {{ background:{BORDER2}; border-radius:4px; }}
            QScrollBar:horizontal {{ background:{SURFACE}; height:8px; border-radius:4px; }}
            QScrollBar::handle:horizontal {{ background:{BORDER2}; border-radius:4px; }}
            QTableWidget {{ gridline-color:{BORDER}; outline:none; }}
            QTableWidget::item:selected {{ background:{BORDER2}; }}
            QToolTip {{
                background:{SURFACE}; color:{TEXT1};
                border:1px solid {BORDER2}; font-size:10px; padding:4px;
            }}
        """)
        # ── Header bar ──
        self._header_w.setStyleSheet(f"background:{SURFACE};")
        self._tag_lbl.setStyleSheet(f"color:{TEXT3}; font-size:9px; letter-spacing:2px;")
        self._title_lbl.setText(
            f'<span style="color:{TEXT1};font-size:14px;font-weight:700;letter-spacing:4px">FUNDAMENTAL</span>'
            f'<span style="color:{BLUE};font-size:14px;font-weight:700;letter-spacing:4px"> SCREENER</span>'
        )
        self._refresh_icon_btns()

        # ── Controls bar ──
        _lbl_s = f"color:{TEXT1}; font-size:9px; font-weight:600; letter-spacing:1.5px;"
        self._controls_w.setStyleSheet(f"background:{BG};")
        self._lbl_market.setStyleSheet(_lbl_s)
        self._lbl_top.setStyleSheet(_lbl_s)
        self._lbl_tk.setStyleSheet(_lbl_s)
        self._lbl_filter.setStyleSheet(_lbl_s)
        self._market.setStyleSheet(self._combo_style())
        self._top_spin.setStyleSheet(self._spinbox_style())
        self._spin_minus.setStyleSheet(self._spin_btn_style())
        self._spin_plus.setStyleSheet(self._spin_btn_style())
        self._use_yf.setStyleSheet(f"color:{TEXT1}; font-size:10px; font-family:'Segoe UI',sans-serif;")
        self._btn_scan.setStyleSheet(self._btn_style(BLUE, BLUE_HV))
        self._btn_export.setStyleSheet(self._btn_style("#2E5C2E", "#1A4C1A"))
        self._ticker_inp.setStyleSheet(self._input_style(font="Consolas,monospace", size="13px", ls="3px"))
        self._btn_lookup.setStyleSheet(self._btn_style(BLUE, BLUE_HV, size="9px"))
        self._filter_inp.setStyleSheet(self._input_style())
        self._lbl_count.setStyleSheet(f"color:{TEXT3}; font-size:9px; letter-spacing:1px;")
        self._lbl_qf.setStyleSheet(f"color:{TEXT1}; font-size:9px; font-weight:600; letter-spacing:1.5px;")
        for chk in self._chip_checks:
            chk.setStyleSheet(self._chip_style())
        self._btn_clr.setStyleSheet(self._btn_style(PANEL, BORDER2, "9px"))

        # ── Table ──
        self._table.setStyleSheet(f"""
            QTableWidget {{
                background:{BG}; alternate-background-color:{PANEL};
                font-size:11px; font-family:'Segoe UI',sans-serif;
                selection-background-color:{BORDER2};
                border:none;
            }}
        """)
        self._recolor_table()
        # ── Detail card ──
        self._detail.apply_theme()
        # ── QC table ──
        self._qc_table.setStyleSheet(f"""
            QTableWidget {{
                background:{BG}; alternate-background-color:{PANEL};
                font-size:11px; font-family:'Segoe UI',sans-serif;
                selection-background-color:{BORDER2};
                border:none;
            }}
        """)
        self._color_qc_header_cols()
        self._recolor_qc_table()
        # ── QC detail card ──
        self._qc_detail.apply_theme()
        # ── Tab widget ──
        self._tabs_w.setStyleSheet(self._tab_style())
        # ── Tab containers + splitters ──
        if hasattr(self, "_cs_w"):
            self._cs_w.setStyleSheet(f"background:{BG};")
        if hasattr(self, "_qc_w"):
            self._qc_w.setStyleSheet(f"background:{BG};")
        if hasattr(self, "_cs_vsplit"):
            self._cs_vsplit.setStyleSheet(f"QSplitter::handle {{ background:{BORDER}; height:2px; }}")
        if hasattr(self, "_qc_vsplit"):
            self._qc_vsplit.setStyleSheet(f"QSplitter::handle {{ background:{BORDER}; height:2px; }}")
        # ── Chart panel ──
        self._chart_panel.set_theme(dark)
        # ── Dashboard — regenerate HTML theo theme mới ──
        self._dash_browser.setStyleSheet(f"""
            QTextBrowser {{
                background:{BG}; color:{TEXT1};
                border:none; padding:16px 20px;
                font-family:'Segoe UI',sans-serif;
            }}
        """)
        if hasattr(self, "_df") and self._df is not None and len(self._df) > 0:
            self._update_dashboard(self._df)
        # ── Status bar ──
        self._status_w.setStyleSheet(f"background:{SURFACE};")
        self._ver_lbl.setStyleSheet(f"color:{TEXT3}; font-size:9px; letter-spacing:1px;")

    # ── Style helpers ─────────────────────────────────────────────────────────

    def _sep(self):
        f = QFrame(); f.setFrameShape(QFrame.HLine); f.setFixedHeight(1)
        f.setStyleSheet(f"background:{BORDER}; border:none;")
        return f

    def _input_style(self, font="'Segoe UI',sans-serif", size="11px", ls="0"):
        return f"""
            QLineEdit {{
                background:{INPUT_BG}; color:{TEXT1};
                border:1px solid {BORDER}; border-radius:3px;
                padding:0 10px; font-size:{size};
                font-family:{font}; letter-spacing:{ls};
                selection-background-color:{BLUE};
            }}
            QLineEdit:focus {{ border-color:{BLUE}; }}
        """

    def _spin_btn_style(self):
        return f"""
            QPushButton {{
                background:{SURFACE}; color:{TEXT1};
                border:1px solid {BORDER}; border-radius:3px;
                font-size:15px; font-weight:600;
                padding:0;
            }}
            QPushButton:hover   {{ background:{BORDER};  color:{TEXT1}; }}
            QPushButton:pressed {{ background:{BORDER2}; color:{TEXT1}; }}
        """

    def _spinbox_style(self):
        return f"""
            QSpinBox {{
                background:{INPUT_BG}; color:{TEXT1};
                border:1px solid {BORDER}; border-radius:3px;
                padding:0 4px 0 6px; font-size:11px;
                font-family:'Segoe UI',sans-serif;
            }}
            QSpinBox:focus {{ border-color:{BLUE}; }}
            QSpinBox::up-button {{
                subcontrol-origin:border; subcontrol-position:top right;
                background:{SURFACE}; border-left:1px solid {BORDER};
                border-top-right-radius:3px;
                width:18px; height:16px;
            }}
            QSpinBox::up-button:hover   {{ background:{BORDER}; }}
            QSpinBox::up-button:pressed {{ background:{BORDER2}; }}
            QSpinBox::down-button {{
                subcontrol-origin:border; subcontrol-position:bottom right;
                background:{SURFACE}; border-left:1px solid {BORDER};
                border-bottom-right-radius:3px;
                width:18px; height:16px;
            }}
            QSpinBox::down-button:hover   {{ background:{BORDER}; }}
            QSpinBox::down-button:pressed {{ background:{BORDER2}; }}
            QSpinBox::up-arrow   {{ color:{TEXT1}; width:10px; height:10px; }}
            QSpinBox::down-arrow {{ color:{TEXT1}; width:10px; height:10px; }}
        """

    def _combo_style(self):
        return f"""
            QComboBox {{
                background:{INPUT_BG}; color:{TEXT1};
                border:1px solid {BORDER}; border-radius:3px;
                padding:0 10px; font-size:11px; min-width:100px;
            }}
            QComboBox:focus {{ border-color:{BLUE}; }}
            QComboBox::drop-down {{ border:none; width:20px; }}
            QComboBox QAbstractItemView {{
                background:{SURFACE}; color:{TEXT1};
                selection-background-color:{BORDER2};
                border:1px solid {BORDER};
            }}
        """

    def _chip_style(self, checked=False):
        fg = BLUE if checked else TEXT2
        border = BLUE if checked else BORDER
        return (f"QCheckBox {{ color:{fg}; font-size:9px; font-weight:600;"
                f" font-family:'Segoe UI',sans-serif; spacing:4px;"
                f" padding:2px 8px; border:1px solid {border}; border-radius:10px; }}"
                f"QCheckBox:checked {{ color:{BLUE}; border-color:{BLUE}; }}"
                f"QCheckBox:unchecked {{ color:{TEXT2}; border-color:{BORDER}; }}")

    def _btn_style(self, bg, hv, size="10px"):
        return f"""
            QPushButton {{
                background:{bg}; color:#FFFFFF;
                border:none; border-radius:3px;
                font-size:{size}; font-weight:700; letter-spacing:1.5px;
                font-family:'Segoe UI',sans-serif;
            }}
            QPushButton:hover   {{ background:{hv}; }}
            QPushButton:pressed {{ background:{hv}; }}
            QPushButton:disabled {{
                background:{PANEL}; color:{TEXT3};
            }}
        """

    # ── Spinner ───────────────────────────────────────────────────────────────

    def _set_status(self, text, color=TEXT3):
        self.status_lbl.setText(text)
        self.status_lbl.setStyleSheet(
            f"color:{color}; font-size:10px; font-family:'Segoe UI',sans-serif;")
        self._dot.setStyleSheet(f"color:{color}; font-size:7px;")

    def _tick(self):
        self._spin_idx = (self._spin_idx + 1) % len(SPINNER)
        self._set_status(f"{SPINNER[self._spin_idx]}  {self._spin_msg}", AMBER)

    # ── Scan logic ────────────────────────────────────────────────────────────

    def _on_scan(self):
        if self._worker and self._worker.isRunning():
            self._worker.cancel()
            self._btn_scan.setText("▶  SCAN")
            return

        market = self._market.currentText()
        top    = self._top_spin.value()
        use_yf = self._use_yf.isChecked()

        self._btn_scan.setText("⏹  Cancel")
        self._btn_export.setEnabled(False)
        self._spin_msg = f"Scanning {market.upper()} top {top}…"
        self._spin_tmr.start(80)

        self._worker = ScanWorker(market, top, use_yf)
        self._worker.progress.connect(lambda t: setattr(self, "_spin_msg", t))
        self._worker.ticker_update.connect(self._on_ticker_progress)
        self._worker.done.connect(self._on_scan_done)
        self._worker.failed.connect(self._on_scan_failed)
        self._worker.start()

    def _on_ticker_progress(self, idx, total, ticker, score):
        self._spin_msg = f"Moat [{idx}/{total}] {ticker} → {score}"

    def _on_scan_done(self, df: pd.DataFrame):
        self._spin_tmr.stop()
        self._btn_scan.setText("▶  SCAN")
        self._btn_export.setEnabled(True)
        self._populate_table(df)
        self._populate_qc_table(df)

        # Áp QC scores vào df để Dashboard và Export đọc được HAS_QC
        try:
            qc_results = df.apply(lambda r: compute_qc_score(r.to_dict()), axis=1)
            qc_df = pd.DataFrame(list(qc_results))
            # Drop QC cols cũ (nếu có) trước khi concat — tránh duplicate column labels
            qc_cols = [c for c in qc_df.columns if c in df.columns]
            df = df.drop(columns=qc_cols, errors="ignore")
            df = pd.concat([df.reset_index(drop=True),
                            qc_df.reset_index(drop=True)], axis=1)
        except Exception:
            pass

        self._df = df
        self._update_dashboard(df)
        self._tabs_w.setCurrentIndex(0)   # tự chuyển sang Dashboard
        self._set_status(
            f"✓  {len(df)} stocks  ·  "
            f"STRONG BUY: {(df['CS_Signal']=='🟢 STRONG BUY').sum()}  "
            f"BUY: {(df['CS_Signal']=='🔵 BUY').sum()}  "
            f"WATCH: {(df['CS_Signal']=='🟡 WATCH').sum()}  "
            f"SKIP: {(df['CS_Signal']=='🔴 SKIP').sum()}",
            GREEN
        )

    def _on_scan_failed(self, msg):
        self._spin_tmr.stop()
        self._btn_scan.setText("▶  SCAN")
        self._set_status(f"✕  {msg}", RED)

    # ── Table population ──────────────────────────────────────────────────────

    def _populate_table(self, df: pd.DataFrame):
        self._ticker_data = {row.get("Ticker", ""): row.to_dict()
                             for _, row in df.iterrows() if row.get("Ticker")}
        self._table.setSortingEnabled(False)
        self._table.setRowCount(0)
        self._table.setRowCount(len(df))

        PCT_KEYS = {"EPS Qtr%","EPS Annual%","Rev Annual%","Rev Qtr%",
                    "Gross Margin%","Net Margin%","ROE%",
                    "1W%","1M%","3M%","6M%","1Y%"}
        is_vn = self._market.currentText().lower() == "vietnam"

        for ri, (_, row) in enumerate(df.iterrows()):
            for ci, (hdr, key, _, align) in enumerate(TABLE_COLS):
                val = row.get(key)
                cell = None

                if key == "_no_":
                    cell = NumItem(ri + 1, "int")
                    cell.setForeground(QColor(TEXT1))

                elif key == "Ticker":
                    cell = QTableWidgetItem(str(val) if val else "")
                    cell.setFont(QFont("Consolas", 9, QFont.Bold))
                    cell.setForeground(QColor("#00BFFF"))

                elif key == "Tên Công Ty":
                    cell = QTableWidgetItem(str(val) if val else "")
                    cell.setForeground(QColor(TEXT1))

                elif key == "Sector":
                    cell = QTableWidgetItem(str(val) if val else "")
                    cell.setForeground(QColor(TEXT1))

                elif key == "Moat Proxy":
                    cell = QTableWidgetItem(str(val) if val else "—")
                    cell.setForeground(QColor(TEXT1))

                elif key == "Moat Score":
                    cell = QTableWidgetItem(str(val) if val else "—")
                    if val in MOAT_SCORE_STYLE:
                        fg, bg = MOAT_SCORE_STYLE[val]
                        cell.setForeground(QColor(f"#{fg}"))
                        cell.setBackground(QColor(f"#{bg}"))

                elif key == "CS_Signal":
                    cell = QTableWidgetItem(str(val) if val else "")
                    if val in SIGNAL_COLORS:
                        fg, bg = SIGNAL_COLORS[val]
                        cell.setForeground(QColor(fg))
                        cell.setBackground(QColor(bg))

                elif key == "CS_Score":
                    score = int(val) if val is not None else 0
                    cell  = NumItem(score, "int")
                    pct   = score / N_CS
                    c = "#1A5C2B" if pct >= 0.875 else "#1B3A5C" if pct >= 0.625 \
                        else "#7D6608" if pct >= 0.375 else "#9C0006"
                    bg = "#C6EFCE" if pct >= 0.875 else "#DDEEFF" if pct >= 0.625 \
                         else "#FFF2CC" if pct >= 0.375 else "#FFC7CE"
                    cell.setForeground(QColor(c)); cell.setBackground(QColor(bg))

                elif key.startswith("CS_"):
                    if val is True:
                        cell = QTableWidgetItem("✓")
                        cell.setForeground(QColor("#1A5C2B")); cell.setBackground(QColor("#C6EFCE"))
                    elif val is False:
                        cell = QTableWidgetItem("✗")
                        cell.setForeground(QColor("#9C0006")); cell.setBackground(QColor("#FFC7CE"))
                    else:
                        cell = QTableWidgetItem("—"); cell.setForeground(QColor(TEXT3))

                elif key == "PEG":
                    cell = NumItem(val)
                    if isinstance(val, float) and val == val:
                        if val <= 1.0:
                            cell.setForeground(QColor("#1A5C2B"))
                            cell.setBackground(QColor("#C6EFCE"))
                        elif val <= 1.5:
                            cell.setForeground(QColor("#1A3A5C"))
                            cell.setBackground(QColor("#DDEEFF"))
                        elif val <= 2.5:
                            cell.setForeground(QColor("#7D6608"))
                        else:
                            cell.setForeground(QColor(RED))
                    else:
                        cell.setForeground(QColor(TEXT3))

                elif key == "EPS_Acc":
                    text = str(val) if val else "—"
                    cell = QTableWidgetItem(text)
                    _acc_colors = {
                        "↑3Q": ("#1A5C2B", "#C6EFCE"),
                        "↑2Q": ("#1A5C2B", "#D8F5E5"),
                        "↑1Q": ("#7D6608", "#FFF2CC"),
                        "↓":   ("#9C0006", "#FFC7CE"),
                    }
                    if val in _acc_colors:
                        fg_a, bg_a = _acc_colors[val]
                        cell.setForeground(QColor(fg_a))
                        cell.setBackground(QColor(bg_a))
                    else:
                        cell.setForeground(QColor(TEXT3))

                elif key == "Price ($)":
                    cell = NumItem(val, "vnd" if is_vn else "price")
                elif key == "MCap ($B)":
                    cell = NumItem(val, "mcap_vnd" if is_vn else "mcap")
                elif key in PCT_KEYS:
                    cell = NumItem(val, "pct")
                    if isinstance(val, float):
                        cell.setForeground(QColor(GREEN) if val > 0 else
                                           QColor(RED)   if val < 0 else
                                           QColor(TEXT3))
                else:
                    cell = NumItem(val)

                if cell:
                    cell.setTextAlignment(int(align))
                    self._table.setItem(ri, ci, cell)

        self._table.setSortingEnabled(True)
        # Đổi header label theo market
        _keys = [c[1] for c in TABLE_COLS]
        if is_vn:
            if "Price ($)" in _keys:
                self._table.horizontalHeaderItem(_keys.index("Price ($)")).setText("Giá (₫)")
            if "MCap ($B)" in _keys:
                self._table.horizontalHeaderItem(_keys.index("MCap ($B)")).setText("MCap (tỷ₫)")
        else:
            if "Price ($)" in _keys:
                self._table.horizontalHeaderItem(_keys.index("Price ($)")).setText("Price($)")
            if "MCap ($B)" in _keys:
                self._table.horizontalHeaderItem(_keys.index("MCap ($B)")).setText("MCap($B)")
        self._apply_filter()

    # ── Dashboard ─────────────────────────────────────────────────────────────

    def _update_dashboard(self, df):
        self._dash_browser.setHtml(self._generate_dashboard_html(df))

    def _generate_dashboard_html(self, df):  # noqa: C901
        import pandas as pd
        from datetime import datetime

        HAS_QC = "QC_Score" in df.columns and "QC_Signal" in df.columns
        market = self._market.currentText()
        is_vn  = market.lower() == "vietnam"
        scan_t = datetime.now().strftime("%Y-%m-%d  %H:%M")
        dark   = self._is_dark

        # ── Theme color palette ───────────────────────────────────
        if dark:
            T_ROW_E   = "0E1820";  T_ROW_O   = "0B1218"
            T_HDR_BG  = "0A1422";  T_HDR_FG  = "5A7A9A"
            T_TEXT    = "DCE4EE";  T_DIM     = "7A8899"
            T_TICKER  = "00BFFF";  T_CO      = "B8C4D0"
            T_SECTOR  = "9AAFCC";  T_BDR     = "1px solid #1A2A3A"
            T_NUM     = "888888"
            # Value colors — bright for dark backgrounds
            T_GREEN   = "34C472";  T_BLUE    = "5D9EFF"
            T_RED     = "FF6B6B";  T_ORANGE  = "FFB347"
            T_PURPLE  = "CE93D8";  T_TEAL    = "4DB6AC"
            TOP3_BG  = ["1A2800","0A1F10","0A1530"]
            TOP3_FG  = ["FFC000","34C472","3D8EF0"]
            T_SEC_OE  = "0E1820";  T_SEC_OO  = "0B1218"
            T_SEC_TXT = "DCE4EE"
            CHART_Q   = "1A3A20";  CHART_DL  = "3A1A50"
            CHART_W   = "202020";  CHART_MOM = "1A3050"
            CHART_Q_F = "34C472";  CHART_DL_F = "CE93D8"
            CHART_W_F = "888888";  CHART_MOM_F = "4FC3F7"
            CHART_SUB = "555555"
        else:
            T_ROW_E   = "F5F8FC";  T_ROW_O   = "FFFFFF"
            T_HDR_BG  = "E8EDF5";  T_HDR_FG  = "4A5568"
            T_TEXT    = "1A202C";  T_DIM     = "64748B"
            T_TICKER  = "0055CC";  T_CO      = "374151"
            T_SECTOR  = "4A5568";  T_BDR     = "1px solid #D1D9E6"
            T_NUM     = "6B7280"
            # Value colors — dark for light backgrounds
            T_GREEN   = "276221";  T_BLUE    = "1B3A5C"
            T_RED     = "9C0006";  T_ORANGE  = "9C6500"
            T_PURPLE  = "6B21A8";  T_TEAL    = "0E6655"
            TOP3_BG  = ["FFF8E1","E8F5E9","E3F2FD"]
            TOP3_FG  = ["7B5900","1B5E20","0D3B7A"]
            T_SEC_OE  = "F5F8FC";  T_SEC_OO  = "FFFFFF"
            T_SEC_TXT = "1A202C"
            CHART_Q   = "E8F5E8";  CHART_DL  = "F3E8FF"
            CHART_W   = "F0F0F0";  CHART_MOM = "E8F0FF"
            CHART_Q_F = "1B5E20";  CHART_DL_F = "6B21A8"
            CHART_W_F = "6B7280";  CHART_MOM_F = "1E3A8A"
            CHART_SUB = "9CA3AF"

        def _v(row, key):
            v = row.get(key)
            return None if v is None or (isinstance(v, float) and pd.isna(v)) else v

        def _pct(v):
            if v is None: return "—"
            return f"{v:+.1f}%"

        # ── ❶ CAN SLIM KPIs ──────────────────────────────────────
        total   = len(df)
        sb_cnt  = int((df["CS_Signal"] == "🟢 STRONG BUY").sum())
        b_cnt   = int((df["CS_Signal"] == "🔵 BUY").sum())
        avg_cs  = f"{df['CS_Score'].mean():.1f}" if "CS_Score" in df.columns else "—"
        _1y     = df["1Y%"].dropna() if "1Y%" in df.columns else pd.Series(dtype=float)
        avg_1y  = f"{_1y.mean():.1f}%" if len(_1y) else "—"
        _roe    = df["ROE%"].dropna() if "ROE%" in df.columns else pd.Series(dtype=float)
        avg_roe = f"{_roe.mean():.1f}%" if len(_roe) else "—"

        # ── ❷ QC KPIs ────────────────────────────────────────────
        if HAS_QC:
            comp_cnt  = int((df["QC_Signal"] == "🏆 COMPOUNDER").sum())
            qual_cnt  = int((df["QC_Signal"] == "⭐ QUALITY").sum())
            avg_qc_s  = f"{df['QC_Score'].mean():.1f}"
            _roic     = df["ROIC%"].dropna() if "ROIC%" in df.columns else pd.Series(dtype=float)
            avg_roic  = f"{_roic.mean():.1f}%" if len(_roic) else "—"
            dual_cnt  = int(((df["CS_Score"] >= 7) & (df["QC_Score"] >= 4)).sum())
            _qc_mask  = df["QC_Signal"].isin(["🏆 COMPOUNDER", "⭐ QUALITY"])
            eq_cb_cnt = (int((df[_qc_mask]["EQ_Badge"] == "💚 Cash Backed").sum())
                         if "EQ_Badge" in df.columns else 0)
        else:
            comp_cnt = qual_cnt = dual_cnt = eq_cb_cnt = 0
            avg_qc_s = avg_roic = "—"

        # ── Risk Flags ────────────────────────────────────────────
        risk_de = int((df["D/E"].fillna(0) > 1.5).sum()) if "D/E"  in df.columns else 0
        risk_pe = int((df["P/E"].fillna(0) > 50).sum())  if "P/E"  in df.columns else 0
        risk_1m = int((df["1M%"].fillna(0) < -10).sum()) if "1M%"  in df.columns else 0

        # ── KPI card helper ────────────────────────────────────────
        def _kpi(val, lbl, bg, fg="FFFFFF"):
            return (f'<td width="16.66%" style="padding:3px;">'
                    f'<table width="100%" cellspacing="0" cellpadding="0">'
                    f'<tr><td align="center" style="background:#{bg};padding:10px 6px 3px;">'
                    f'<span style="font-size:20px;font-weight:700;color:#{fg};">{val}</span></td></tr>'
                    f'<tr><td align="center" style="background:#{bg};padding:3px 6px 9px;">'
                    f'<span style="font-size:9px;color:#{fg};letter-spacing:1px;">{lbl}</span>'
                    f'</td></tr></table></td>')

        sb_pct = f"{sb_cnt/total*100:.1f}%" if total else "0%"
        b_pct  = f"{b_cnt/total*100:.1f}%"  if total else "0%"

        cs_kpis = (
            _kpi(total,   "TOTAL STOCKS",              "1B3A5C") +
            _kpi(sb_pct,  f"STRONG BUY  ({sb_cnt})",  "00703A") +
            _kpi(b_pct,   f"BUY  ({b_cnt})",           "0070C0") +
            _kpi(avg_cs,  f"AVG CS / {N_CS}",          "5B2C6F") +
            _kpi(avg_1y,  "AVG 1Y RETURN",              "7D3C0A") +
            _kpi(avg_roe, "AVG ROE%",                   "0E6655")
        )
        qc_kpis = (
            _kpi(comp_cnt,   "# COMPOUNDER",               "1A5C2B") +
            _kpi(qual_cnt,   "# QUALITY",                   "1A3A5C") +
            _kpi(avg_roic,   "AVG ROIC%",                   "0E6655") +
            _kpi(avg_qc_s,   "AVG QC / 6",                  "5B2C6F") +
            _kpi(dual_cnt,   "DUAL LEADERS (CS≥7 &amp; QC≥4)","4A235A") +
            _kpi(eq_cb_cnt,  "💚 CASH BACKED (QC stocks)",  "155A28")
        ) if HAS_QC else ""

        # ── Section header ─────────────────────────────────────────
        def _sec(text, bg="2C4F7C"):
            return (f'<tr><td colspan="100%" style="background:#{bg};color:#FFFFFF;'
                    f'font-size:11px;font-weight:700;padding:7px 10px;letter-spacing:1px;'
                    f'margin-top:10px;">{text}</td></tr>')

        # ── Top CS | Top QC side by side ──────────────────────────
        _cs_buy = ["🟢 STRONG BUY", "🔵 BUY"]
        cs_top = (df[df["CS_Signal"].isin(_cs_buy)].sort_values("CS_Score", ascending=False)
                  if "CS_Score" in df.columns else pd.DataFrame())
        _qc_buy = ["🏆 COMPOUNDER", "⭐ QUALITY"]
        qc_top = (df[df["QC_Signal"].isin(_qc_buy)].sort_values("QC_Score", ascending=False)
                  if HAS_QC else pd.DataFrame())
        n_top = max(len(cs_top), len(qc_top), 1)

        SROW = ("#", "Ticker", "Company", "Sector", "Score", "Signal", "Key%")
        def _th(text, bg):
            return f'<td align="center" style="background:#{bg};color:#FFFFFF;font-size:9px;font-weight:700;padding:5px 6px;border:1px solid #1A2A3A;">{text}</td>'

        cs_hdr = "".join(_th(h,"1B3A5C") for h in SROW)
        qc_hdr = "".join(_th(h,"0E6655") for h in SROW)

        cs_rows_html = ""
        for ri in range(1, n_top + 1):
            bg = T_ROW_E if ri%2==0 else T_ROW_O
            if ri <= len(cs_top):
                r = cs_top.iloc[ri-1]
                sig = r.get("CS_Signal","") or ""
                sfg = {"🟢 STRONG BUY": T_GREEN, "🔵 BUY": T_BLUE}.get(sig, T_DIM)
                _1y_v = _v(r,"1Y%"); _1y_s = _pct(_1y_v)
                _1y_c = T_GREEN if _1y_v and _1y_v>0 else T_RED if _1y_v and _1y_v<0 else T_NUM
                cs_rows_html += (f'<tr style="background:#{bg};">'
                    f'<td align="center" style="padding:4px 6px;background:#1B3A5C;color:#FFFFFF;font-weight:700;font-size:9px;{T_BDR}">{ri}</td>'
                    f'<td align="center" style="padding:4px 6px;color:#{T_TICKER};font-weight:700;font-family:Consolas;{T_BDR}">{r.get("Ticker","")}</td>'
                    f'<td style="padding:4px 6px;color:#{T_CO};font-size:11px;{T_BDR}">{str(r.get("Tên Công Ty",""))[:22]}</td>'
                    f'<td style="padding:4px 6px;color:#{T_SECTOR};font-size:10px;font-style:italic;{T_BDR}">{r.get("Sector","")}</td>'
                    f'<td align="center" style="padding:4px 6px;background:#C6EFCE;color:#1A5C2B;font-weight:700;{T_BDR}">{int(r.get("CS_Score",0))}</td>'
                    f'<td align="center" style="padding:4px 6px;color:#{sfg};font-size:10px;{T_BDR}">{sig}</td>'
                    f'<td align="center" style="padding:4px 6px;color:#{_1y_c};font-weight:700;{T_BDR}">{_1y_s}</td>'
                    f'</tr>')

        qc_rows_html = ""
        for ri in range(1, n_top + 1):
            bg = T_ROW_E if ri%2==0 else T_ROW_O
            if ri <= len(qc_top):
                r = qc_top.iloc[ri-1]
                sig = r.get("QC_Signal","") or ""
                sfg = {"🏆 COMPOUNDER":"1A5C2B","⭐ QUALITY":"1A3A5C"}.get(sig,"000000")
                roic = _v(r,"ROIC%")
                roic_s = f"{roic:.1f}%" if roic is not None else "—"
                roic_c = T_GREEN if roic and roic>=15 else T_BLUE
                qc_rows_html += (f'<tr style="background:#{bg};">'
                    f'<td align="center" style="padding:4px 6px;background:#0E6655;color:#FFFFFF;font-weight:700;font-size:9px;{T_BDR}">{ri}</td>'
                    f'<td align="center" style="padding:4px 6px;color:#{T_TICKER};font-weight:700;font-family:Consolas;{T_BDR}">{r.get("Ticker","")}</td>'
                    f'<td style="padding:4px 6px;color:#{T_CO};font-size:11px;{T_BDR}">{str(r.get("Tên Công Ty",""))[:22]}</td>'
                    f'<td style="padding:4px 6px;color:#{T_SECTOR};font-size:10px;font-style:italic;{T_BDR}">{r.get("Sector","")}</td>'
                    f'<td align="center" style="padding:4px 6px;background:#DDEEFF;color:#1A3A5C;font-weight:700;{T_BDR}">{int(r.get("QC_Score",0))}</td>'
                    f'<td align="center" style="padding:4px 6px;color:#{sfg};font-size:10px;{T_BDR}">{sig}</td>'
                    f'<td align="center" style="padding:4px 6px;color:#{roic_c};font-weight:700;{T_BDR}">{roic_s}</td>'
                    f'</tr>')
            elif not HAS_QC and ri == 1:
                qc_rows_html += f'<tr><td colspan="7" align="center" style="padding:8px;color:#888;font-style:italic;">— Yêu cầu bật yfinance Moat —</td></tr>'

        # ── ⭐ Dual Leaders ───────────────────────────────────────
        if HAS_QC:
            dl_df = (df[(df["CS_Score"]>=7)&(df["QC_Score"]>=4)]
                     .sort_values(["CS_Score","QC_Score"],ascending=[False,False]))
        else:
            dl_df = pd.DataFrame()

        DL_H = ["#","Ticker","Company","Sector","Price","CS","QC","1Y%","ROE%","ROIC%","D/E","Net Cash($B)","Signal","⚠"]
        dl_hdr_html = "".join(
            f'<td align="center" style="background:#4A235A;color:#FFFFFF;font-size:9px;'
            f'font-weight:700;padding:5px 6px;border:1px solid #3A1A4A;">{h}</td>'
            for h in DL_H)

        dl_rows_html = ""
        if dl_df.empty:
            dl_rows_html = '<tr><td colspan="14" align="center" style="padding:10px;color:#888;font-style:italic;">— Không có mã nào đạt CS≥7 và QC≥4 —</td></tr>'
        else:
            for ri,(_, r) in enumerate(dl_df.iterrows(),1):
                rbg = T_ROW_E if ri%2==0 else T_ROW_O
                sig = r.get("CS_Signal","") or ""
                sfg,sbg = {"🟢 STRONG BUY":("276221","C6EFCE"),"🔵 BUY":("1B3A5C","DDEEFF")}.get(sig,("000000","FFFFFF"))
                _1y_v = _v(r,"1Y%"); _1y_s = _pct(_1y_v); _1y_c = T_GREEN if _1y_v and _1y_v>0 else T_RED if _1y_v and _1y_v<0 else T_NUM
                roe_v = _v(r,"ROE%"); roe_s = f"{roe_v:.1f}%" if roe_v is not None else "—"; roe_c = T_GREEN if roe_v and roe_v>17 else T_TEXT
                roic_v = _v(r,"ROIC%"); roic_s = f"{roic_v:.1f}%" if roic_v is not None else "—"
                roic_c = T_GREEN if roic_v and roic_v>=15 else T_BLUE if roic_v and roic_v>=10 else T_ORANGE
                de_v = _v(r,"D/E"); de_s = f"{de_v:.1f}" if de_v is not None else "—"; de_c = T_RED if de_v and de_v>2 else T_ORANGE if de_v and de_v>1 else T_GREEN
                nc_v = _v(r,"Net Cash ($B)"); nc_s = f"{nc_v:+.1f}" if nc_v is not None else "—"
                nc_c = T_GREEN if nc_v and nc_v>5 else T_GREEN if nc_v and nc_v>0 else T_ORANGE if nc_v and nc_v>-5 else T_RED
                nc_bg = "C6EFCE" if nc_v and nc_v>5 else rbg
                px_v = _v(r,"Price ($)"); px_s = f"₫{px_v:,.0f}" if is_vn and px_v else (f"${px_v:,.2f}" if px_v else "—")
                flags = []
                if de_v and de_v>1: flags.append("D/E")
                if _v(r,"P/E") and _v(r,"P/E")>50: flags.append("P/E")
                warn_s = "⚠ "+" · ".join(flags) if flags else ""
                td = lambda val,fg=T_TEXT,bg=None,fw="normal",fs="12px": (
                    f'<td align="center" style="padding:4px 6px;color:#{fg};font-weight:{fw};'
                    f'font-size:{fs};background:{"#"+bg if bg else "#"+rbg};{T_BDR}">{val}</td>')
                dl_rows_html += (f'<tr>'
                    + f'<td align="center" style="padding:4px 6px;background:#4A235A;color:#FFFFFF;font-weight:700;font-size:9px;{T_BDR}">{ri}</td>'
                    + td(r.get("Ticker",""), T_TICKER,"","700","11px")
                    + f'<td style="padding:4px 6px;color:#{T_CO};font-size:11px;{T_BDR}">{str(r.get("Tên Công Ty",""))[:24]}</td>'
                    + f'<td style="padding:4px 6px;color:#{T_SECTOR};font-size:10px;font-style:italic;{T_BDR}">{r.get("Sector","")}</td>'
                    + td(px_s)
                    + td(int(r.get("CS_Score",0)), "276221", "C6EFCE", "700")
                    + td(int(r.get("QC_Score",0)), "1A3A5C", "DDEEFF", "700")
                    + td(_1y_s, _1y_c)
                    + td(roe_s, roe_c)
                    + td(roic_s, roic_c, None, "700")
                    + td(de_s, de_c)
                    + f'<td align="center" style="padding:4px 6px;color:#{nc_c};font-weight:700;background:#{nc_bg};border:1px solid #E0D8F0;">{nc_s}</td>'
                    + f'<td align="center" style="padding:4px 6px;color:#{sfg};background:#{sbg};font-size:10px;font-weight:700;border:1px solid #E0D8F0;">{sig}</td>'
                    + f'<td align="center" style="padding:4px 6px;color:#9C6500;background:{"#FFF2CC" if flags else "#"+rbg};font-size:10px;border:1px solid #E0D8F0;">{warn_s}</td>'
                    + '</tr>')

        # ── 🎯 Top Picks ─────────────────────────────────────────
        # Momentum
        mom = []
        if "1Y%" in df.columns and "CS_Signal" in df.columns:
            for sig_f in ["🟢 STRONG BUY", ["🟢 STRONG BUY","🔵 BUY"]]:
                mask = df["CS_Signal"]==sig_f if isinstance(sig_f,str) else df["CS_Signal"].isin(sig_f)
                d = df[mask][["Ticker","1Y%","Sector","CS_Signal"]].dropna(subset=["1Y%"]).sort_values("1Y%",ascending=False)
                if not d.empty:
                    mom = [(r["Ticker"], f"{r['1Y%']:+.1f}%  ({r['CS_Signal'].split()[1]})", r.get("Sector","")) for _,r in d.iterrows()]
                    break
        # Quality
        qua = []
        if HAS_QC and "ROIC%" in df.columns:
            d = df[df["QC_Signal"]=="🏆 COMPOUNDER"][["Ticker","ROIC%","Sector"]].dropna(subset=["ROIC%"]).sort_values("ROIC%",ascending=False)
            qua = [(r["Ticker"],f"ROIC {r['ROIC%']:.1f}%",r.get("Sector","")) for _,r in d.iterrows()]
        # Value Growth
        val = []
        if "P/E" in df.columns and "EPS Annual%" in df.columns:
            _w52_ok = (df["52W_High%"] < 80) if ("52W_High%" in df.columns and not df["52W_High%"].isna().all()) else True
            tmp = df[(df["CS_Score"]>=6)&(df["P/E"]>0)&(df["P/E"]<35)&
                     (df["EPS Annual%"]>10)&(df["EPS Annual%"]<200)&_w52_ok].dropna(subset=["P/E"]).sort_values("P/E")
            val = [(r["Ticker"],f"P/E {r['P/E']:.1f}× | EPS +{r['EPS Annual%']:.0f}%",r.get("Sector","")) for _,r in tmp.iterrows()]
        # Breakout
        brk = []
        if "52W_High%" in df.columns and not df["52W_High%"].isna().all():
            tmp = df[(df["52W_High%"]>=90)&(df["CS_Score"]>=7)].sort_values(["CS_Score","52W_High%"],ascending=[False,False])
            brk = [(r["Ticker"],f"52W {r['52W_High%']:.1f}% | CS {int(r['CS_Score'])}",r.get("Sector","")) for _,r in tmp.iterrows()]
        elif "52W_High%" in df.columns:
            brk = [("—","52W High% N/A","yfinance tắt")]

        TP_GROUPS = [
            ("🚀 Momentum  (STRONG BUY, 1Y%↓)",    "1A5276", mom),
            ("💎 Quality  (Compounder, ROIC%↓)",    "1A5C2B", qua),
            ("📉 Value  (P/E<35, EPS>10%, CS≥6)",   "784212", val),
            ("💥 Breakout  (52W High≥90%, CS≥7)",   "4A235A", brk),
        ]
        COL_H = ["Ticker","Metric","Sector"]

        tp_cols_html = ""
        for title, bg, rows in TP_GROUPS:
            hdr_html = "".join(
                f'<td align="center" style="background:#{bg};color:#FFFFFF;font-size:9px;'
                f'font-weight:700;padding:4px 6px;border:1px solid #1A2A3A;">{h}</td>'
                for h in COL_H)
            body_html = ""
            for ri,(ticker,metric,sector) in enumerate(rows):
                rbg2 = T_ROW_E if ri%2==0 else T_ROW_O
                body_html += (f'<tr style="background:#{rbg2};">'
                    f'<td align="center" style="padding:4px 6px;color:#{T_TICKER};font-weight:700;font-family:Consolas;font-size:11px;{T_BDR}">{ticker}</td>'
                    f'<td style="padding:4px 6px;color:#{T_TEXT};font-size:11px;{T_BDR}">{metric}</td>'
                    f'<td style="padding:4px 6px;color:#{T_SECTOR};font-size:10px;font-style:italic;{T_BDR}">{sector}</td>'
                    f'</tr>')
            if not body_html:
                body_html = '<tr><td colspan="3" align="center" style="padding:6px;color:#888;font-style:italic;">— No data —</td></tr>'
            tp_cols_html += (f'<td width="25%" style="padding:4px;vertical-align:top;">'
                f'<table width="100%" cellspacing="0" cellpadding="0">'
                f'<tr><td colspan="3" style="background:#{bg};color:#FFFFFF;font-size:10px;font-weight:700;padding:6px 8px;">{title}</td></tr>'
                f'<tr>{hdr_html}</tr>{body_html}</table></td>')

        # ── ❸ Chart Analysis (2×2) ────────────────────────────────
        def _zone_tickers(zone_df, limit=10):
            if zone_df.empty: return "<i style='color:#888;'>—</i>"
            parts = []
            for _, r in zone_df.head(limit).iterrows():
                cs = int(r.get("CS_Score",0)); qc = int(r.get("QC_Score",0)) if HAS_QC else "—"
                _1y_v2 = _v(r,"1Y%"); _1y_s2 = f"{_1y_v2:+.0f}%" if _1y_v2 is not None else ""
                parts.append(f'<b style="color:#{T_TICKER};font-family:Consolas;">{r.get("Ticker","")}</b>'
                             f'<span style="color:#{T_DIM};font-size:10px;"> {cs}/{qc} {_1y_s2}</span>')
            txt = "  ·  ".join(parts)
            if len(zone_df)>limit: txt += f'<span style="color:#555;"> +{len(zone_df)-limit}</span>'
            return txt

        if HAS_QC:
            _cs=df["CS_Score"]; _qc=df["QC_Score"]
            dual_z=df[(_cs>=7)&(_qc>=4)].sort_values(["CS_Score","QC_Score"],ascending=[False,False])
            mom_z =df[(_cs>=7)&(_qc< 4)].sort_values(["CS_Score","1Y%"],ascending=[False,False])
            qual_z=df[(_cs< 7)&(_qc>=4)].sort_values("QC_Score",ascending=False)
            watch_df=df[(_cs<7)&(_qc<4)].sort_values("CS_Score",ascending=False)
            watch_n=len(watch_df)
            chart_html = f"""<table width="100%" cellspacing="3" cellpadding="0">
              <tr>
                <td width="50%" style="background:#{CHART_Q};padding:10px 12px;vertical-align:top;">
                  <p style="margin:0 0 5px;font-size:11px;font-weight:700;color:#{CHART_Q_F};">💎 Quality
                    <span style="font-weight:400;color:#{CHART_SUB};"> (CS&lt;7, QC≥4) · {len(qual_z)} mã</span></p>
                  <p style="margin:0;font-size:11px;line-height:1.9;">{_zone_tickers(qual_z)}</p>
                </td>
                <td width="50%" style="background:#{CHART_DL};padding:10px 12px;vertical-align:top;">
                  <p style="margin:0 0 5px;font-size:11px;font-weight:700;color:#{CHART_DL_F};">⭐ Dual Leaders
                    <span style="font-weight:400;color:#{CHART_SUB};"> (CS≥7, QC≥4) · {len(dual_z)} mã</span></p>
                  <p style="margin:0;font-size:11px;line-height:1.9;">{_zone_tickers(dual_z)}</p>
                </td>
              </tr>
              <tr>
                <td style="background:#{CHART_W};padding:10px 12px;vertical-align:top;">
                  <p style="margin:0 0 5px;font-size:11px;font-weight:700;color:#{CHART_W_F};">👀 Watchlist
                    <span style="font-weight:400;color:#{CHART_SUB};"> (CS&lt;7, QC&lt;4) · {watch_n} mã</span></p>
                  <p style="margin:0;font-size:11px;line-height:1.9;">{_zone_tickers(watch_df)}</p>
                </td>
                <td style="background:#{CHART_MOM};padding:10px 12px;vertical-align:top;">
                  <p style="margin:0 0 5px;font-size:11px;font-weight:700;color:#{CHART_MOM_F};">🚀 Momentum
                    <span style="font-weight:400;color:#{CHART_SUB};"> (CS≥7, QC&lt;4) · {len(mom_z)} mã</span></p>
                  <p style="margin:0;font-size:11px;line-height:1.9;">{_zone_tickers(mom_z)}</p>
                </td>
              </tr>
            </table>"""
        else:
            chart_html = '<p style="color:#4A6080;font-style:italic;padding:10px;">Bật yfinance Moat để xem Chart Analysis đầy đủ.</p>'

        # ── ❹ Sector Breakdown ────────────────────────────────────
        S4_H = ["Sector","# Stocks","# Str.Buy","# Buy","# Comp.","Avg CS","Avg QC","Avg 1Y%"]
        s4_hdr = "".join(
            f'<td style="background:#2C4F7C;color:#FFFFFF;font-size:9px;font-weight:700;'
            f'padding:5px 8px;{T_BDR};text-align:{"left" if i==0 else "center"};">{h}</td>'
            for i,h in enumerate(S4_H))
        sec_pivot = []
        if "Sector" in df.columns:
            for sec,g in df.assign(Sector=df["Sector"].fillna("Unknown")).groupby("Sector"):
                sec_pivot.append({
                    "s": sec, "n": len(g),
                    "sb": int((g["CS_Signal"]=="🟢 STRONG BUY").sum()),
                    "b":  int((g["CS_Signal"]=="🔵 BUY").sum()),
                    "cp": int((g["QC_Signal"]=="🏆 COMPOUNDER").sum()) if HAS_QC else 0,
                    "acs": g["CS_Score"].mean() if "CS_Score" in g.columns else None,
                    "aqc": g["QC_Score"].mean() if HAS_QC else None,
                    "a1y": g["1Y%"].dropna().mean() if "1Y%" in g.columns else None,
                })
            sec_pivot.sort(key=lambda x: x["sb"], reverse=True)
        s4_rows_html = ""
        for ri, row in enumerate(sec_pivot):
            if ri < 3:
                rbg = TOP3_BG[ri]; sec_fg = TOP3_FG[ri]; fw = "font-weight:700;"
            else:
                rbg = T_SEC_OE if ri%2==0 else T_SEC_OO; sec_fg = T_SEC_TXT; fw = ""
            a1y = _pct(row["a1y"]) if row["a1y"] is not None and not pd.isna(row["a1y"]) else "—"
            a1y_c = T_GREEN if row["a1y"] and row["a1y"]>0 else T_RED if row["a1y"] and row["a1y"]<0 else T_NUM
            aqc   = f"{row['aqc']:.1f}" if row["aqc"] is not None and not pd.isna(row["aqc"]) else "—"
            acs_s = "%.1f" % row["acs"] if row["acs"] is not None else "—"
            s4_rows_html += (f'<tr style="background:#{rbg};">'
                f'<td style="padding:5px 10px;color:#{sec_fg};{fw}{T_BDR}">{row["s"]}</td>'
                f'<td align="center" style="padding:5px 8px;color:#{T_DIM};{T_BDR}">{row["n"]}</td>'
                f'<td align="center" style="padding:5px 8px;color:#{T_GREEN};font-weight:700;{T_BDR}">{row["sb"]}</td>'
                f'<td align="center" style="padding:5px 8px;color:#{T_BLUE};{T_BDR}">{row["b"]}</td>'
                f'<td align="center" style="padding:5px 8px;color:#{T_PURPLE};{T_BDR}">{row["cp"]}</td>'
                f'<td align="center" style="padding:5px 8px;color:#{T_TEXT};{T_BDR}">{acs_s}</td>'
                f'<td align="center" style="padding:5px 8px;color:#{T_TEXT};{T_BDR}">{aqc}</td>'
                f'<td align="center" style="padding:5px 8px;color:#{a1y_c};font-weight:700;{T_BDR}">{a1y}</td>'
                f'</tr>')

        # ── Assemble ─────────────────────────────────────────────
        qc_section = (
            f'<tr><td colspan="100%" style="padding:8px 0 0 0;">'
            f'<table width="100%" cellspacing="0" cellpadding="0" style="table-layout:fixed;">'
            f'{_sec("❷  QUALITY COMPOUNDER  —  KEY METRICS")}'
            f'<tr>{qc_kpis}</tr>'
            f'</table></td></tr>'
        ) if HAS_QC else ""

        hdr_bg  = "0D2137" if dark else "1B3A5C"
        hdr_sub = "88AACC" if dark else "B8CCDD"
        return f"""<html><head><style>
          body{{background:{BG};color:{TEXT1};
               font-family:'Segoe UI',sans-serif;font-size:12px;margin:0;padding:10px 14px;}}
          table{{border-collapse:collapse;width:100%;}}
          p{{margin:0;}}
        </style></head><body>

        <table width="100%" cellspacing="0" cellpadding="0" style="margin-bottom:10px;">
          <tr>
            <td bgcolor="#{hdr_bg}" style="background:#{hdr_bg};padding:10px 14px;">
              <span style="font-size:15px;font-weight:700;color:#FFFFFF;letter-spacing:2px;">
                STOCK SCREENER  ·  DECISION DASHBOARD</span>
            </td>
            <td bgcolor="#{hdr_bg}" align="right" style="background:#{hdr_bg};padding:10px 14px;">
              <span style="font-size:10px;color:#{hdr_sub};font-style:italic;">
                {market}  ·  Last updated: {scan_t}</span>
            </td>
          </tr>
        </table>

        <table width="100%" cellspacing="0" cellpadding="0" style="table-layout:fixed;">
          {_sec("❶  CAN SLIM  —  KEY METRICS")}
          <tr>{cs_kpis}</tr>
          {qc_section}
          <tr><td colspan="100%" style="background:#FFF2CC;padding:6px 10px;
              color:#9C6500;font-size:11px;font-weight:700;">
            ⚠️  RISK FLAGS: &nbsp;&nbsp;
            D/E &gt; 1.5: <b>{risk_de} mã</b> &nbsp;&nbsp;|&nbsp;&nbsp;
            P/E &gt; 50: <b>{risk_pe} mã</b> &nbsp;&nbsp;|&nbsp;&nbsp;
            1M% &lt; −10%: <b>{risk_1m} mã</b>
          </td></tr>
        </table>

        <p style="height:8px;"></p>

        <table width="100%" cellspacing="0" cellpadding="0">
          <tr>
            <td width="49%" style="vertical-align:top;">
              <table width="100%" cellspacing="0" cellpadding="0">
                {_sec("📊  TOP CAN SLIM SCORE","1B3A5C")}
                <tr>{cs_hdr}</tr>
                {cs_rows_html}
              </table>
            </td>
            <td width="2%" style="background:{BG};"></td>
            <td width="49%" style="vertical-align:top;">
              <table width="100%" cellspacing="0" cellpadding="0">
                {_sec("📊  TOP QUALITY COMPOUNDER SCORE","0E6655")}
                <tr>{qc_hdr}</tr>
                {qc_rows_html}
              </table>
            </td>
          </tr>
        </table>

        <p style="height:8px;"></p>

        <table width="100%" cellspacing="0" cellpadding="0">
          {_sec("⭐  DUAL LEADERS  —  CS Score ≥ 7  AND  QC Score ≥ 4","4A235A")}
          <tr>{dl_hdr_html}</tr>
          {dl_rows_html}
        </table>

        <p style="height:8px;"></p>

        <table width="100%" cellspacing="0" cellpadding="0">
          {_sec("🎯  TOP PICKS","1A3A5C")}
          <tr>{tp_cols_html}</tr>
        </table>

        <p style="height:8px;"></p>

        <table width="100%" cellspacing="0" cellpadding="0">
          {_sec("❸  CHART ANALYSIS")}
          <tr><td style="padding:6px 0;">{chart_html}</td></tr>
        </table>

        <p style="height:8px;"></p>

        <table width="100%" cellspacing="0" cellpadding="0">
          {_sec("❹  SECTOR BREAKDOWN  —  Sort: # Strong Buy ↓")}
          <tr>{s4_hdr}</tr>
          {s4_rows_html}
        </table>

        <p style="color:#2A3A50;font-size:10px;margin-top:12px;text-align:right;">
          Dashboard  ·  {scan_t}</p>
        </body></html>"""

        # ── KPIs ──────────────────────────────────────────────────
        total   = len(df)
        sb_cnt  = int((df["CS_Signal"] == "🟢 STRONG BUY").sum())
        buy_cnt = int((df["CS_Signal"] == "🔵 BUY").sum())
        avg_cs  = f"{df['CS_Score'].mean():.1f}" if "CS_Score" in df.columns else "—"
        _1y     = df["1Y%"].dropna() if "1Y%" in df.columns else pd.Series(dtype=float)
        avg_1y  = f"{_1y.mean():+.1f}%" if len(_1y) else "—"
        w52s    = df["52W_High%"].dropna() if "52W_High%" in df.columns else pd.Series(dtype=float)
        near_hi = str(int((w52s >= 90).sum())) if len(w52s) else "—"

        if HAS_QC:
            comp_cnt = int((df["QC_Signal"] == "🏆 COMPOUNDER").sum())
            qual_cnt = int((df["QC_Signal"] == "⭐ QUALITY").sum())
            avg_qc   = f"{df['QC_Score'].mean():.1f}"
            _roic    = df["ROIC%"].dropna() if "ROIC%" in df.columns else pd.Series(dtype=float)
            avg_roic = f"{_roic.mean():.1f}%" if len(_roic) else "—"
            dual_cnt = int(((df["CS_Score"] >= 7) & (df["QC_Score"] >= 4)).sum())
            eq_cb    = int((df["EQ_Badge"] == "💚 Cash Backed").sum()) if "EQ_Badge" in df.columns else "—"
        else:
            comp_cnt = qual_cnt = dual_cnt = eq_cb = "—"
            avg_qc = avg_roic = "—"

        _de  = df["D/E"].dropna()  if "D/E"  in df.columns else pd.Series(dtype=float)
        _pe  = df["P/E"].dropna()  if "P/E"  in df.columns else pd.Series(dtype=float)
        _1m  = df["1M%"].dropna()  if "1M%"  in df.columns else pd.Series(dtype=float)
        de_r = str(int((_de > 1).sum()))    if len(_de) else "—"
        pe_r = str(int((_pe > 50).sum()))   if len(_pe) else "—"
        m1_r = str(int((_1m < -10).sum()))  if len(_1m) else "—"

        # ── KPI card helper ────────────────────────────────────────
        def _kpi(val, label, val_col, bg):
            return (f'<td style="padding:3px;" width="16%">'
                    f'<table width="100%" cellspacing="0" cellpadding="0">'
                    f'<tr><td align="center" style="background:{bg};padding:10px 4px 2px 4px;'
                    f'border-radius:3px 3px 0 0;">'
                    f'<span style="font-size:22px;font-weight:700;color:{val_col};">{val}</span></td></tr>'
                    f'<tr><td align="center" style="background:{bg};padding:2px 4px 8px 4px;'
                    f'border-radius:0 0 3px 3px;filter:brightness(0.85);">'
                    f'<span style="font-size:9px;color:{val_col};letter-spacing:1px;">{label}</span>'
                    f'</td></tr></table></td>')

        cs_kpi_row = (
            _kpi(total,   "TOTAL STOCKS",    "#DCE4EE", "#1C2A3A") +
            _kpi(sb_cnt,  "STRONG BUY 🟢",   "#1A5C2B", "#C6EFCE") +
            _kpi(buy_cnt, "BUY 🔵",           "#1A3A5C", "#DDEEFF") +
            _kpi(avg_cs,  f"AVG CS / {N_CS}", "#B8C4D0", "#1A2435") +
            _kpi(avg_1y,  "AVG 1Y%",          "#3D8EF0", "#0D1830") +
            _kpi(near_hi, "NEAR 52W HIGH",    "#7D6608", "#2A2010")
        )
        qc_kpi_row = (
            _kpi(comp_cnt, "COMPOUNDER 🏆",   "#1A5C2B", "#C6EFCE") +
            _kpi(qual_cnt, "QUALITY ⭐",        "#1A3A5C", "#DDEEFF") +
            _kpi(avg_qc,   "AVG QC / 6",       "#34C472", "#0D2010") +
            _kpi(avg_roic, "AVG ROIC%",         "#FFD700", "#1A1500") +
            _kpi(dual_cnt, "DUAL LEADERS",      "#9B59B6", "#2A0A3A") +
            _kpi(eq_cb,    "CASH BACKED 💚",    "#4DB6AC", "#0A2020")
        ) if HAS_QC else ""
        risk_row = (
            _kpi(de_r, "D/E > 1  ⚠",    "#9C0006", "#FFC7CE") +
            _kpi(pe_r, "P/E > 50  ⚠",   "#9C6500", "#FFF2CC") +
            _kpi(m1_r, "1M% < −10%  ⚠", "#9C0006", "#FFE0E0") +
            "<td></td><td></td><td></td>"
        )
        # [old implementation removed]

    def _recolor_table(self):
        """Re-apply theme-sensitive foreground colors to existing cells."""
        PCT_KEYS = {"EPS Qtr%","EPS Annual%","Rev Annual%","Rev Qtr%",
                    "Gross Margin%","Net Margin%","ROE%",
                    "1W%","1M%","3M%","6M%","1Y%"}
        _keys = [c[1] for c in TABLE_COLS]
        for ri in range(self._table.rowCount()):
            for ci, key in enumerate(_keys):
                item = self._table.item(ri, ci)
                if not item:
                    continue
                if key in ("_no_", "Tên Công Ty", "Sector", "Moat Proxy"):
                    item.setForeground(QColor(TEXT1))
                elif key in PCT_KEYS and hasattr(item, "_num"):
                    v = item._num
                    if v != float("-inf"):
                        item.setForeground(QColor(GREEN) if v > 0 else
                                           QColor(RED)   if v < 0 else
                                           QColor(TEXT3))

    def _clear_chips(self):
        for chk in self._chip_checks:
            chk.blockSignals(True)
            chk.setChecked(False)
            chk.blockSignals(False)
        self._apply_filter()

    def _apply_filter(self):
        if hasattr(self, "_tabs_w") and self._tabs_w.currentIndex() == 2:
            self._apply_qc_filter()
            return
        text      = self._filter_inp.text().strip().lower()
        only_sb   = self._chk_sb.isChecked()
        only_moat = self._chk_moat.isChecked()
        only_1y   = self._chk_1y.isChecked()
        only_52w  = self._chk_52w.isChecked()
        any_chip  = only_sb or only_moat or only_1y or only_52w
        td        = getattr(self, "_ticker_data", {})

        visible = 0
        total = self._table.rowCount()
        for ri in range(total):
            show = True
            if text:
                t_item = self._table.item(ri, 1)
                n_item = self._table.item(ri, 2)
                t_txt  = (t_item.text() if t_item else "").lower()
                n_txt  = (n_item.text() if n_item else "").lower()
                show   = (text in t_txt) or (text in n_txt)
            if show and any_chip:
                ticker = (self._table.item(ri, 1).text()
                          if self._table.item(ri, 1) else "")
                rd = td.get(ticker, {})
                if only_sb and rd.get("CS_Signal") != "🟢 STRONG BUY":
                    show = False
                if show and only_moat and rd.get("Moat Score") not in (
                        "WIDE  ★★★", "NARROW ★★"):
                    show = False
                if show and only_1y:
                    v = rd.get("1Y%")
                    if v is None or (isinstance(v, float) and pd.isna(v)) or v <= 0:
                        show = False
                if show and only_52w:
                    v = rd.get("52W_High%")
                    if v is None or (isinstance(v, float) and pd.isna(v)) or v < 90:
                        show = False
            self._table.setRowHidden(ri, not show)
            if show:
                visible += 1
        self._lbl_count.setText(f"{visible} / {total}  rows")

    # ── Row selection → detail card ───────────────────────────────────────────

    def _on_row_selected(self):
        rows = self._table.selectedItems()
        if not rows or self._df is None: return
        ri = self._table.currentRow()
        ticker_item = self._table.item(ri, 1)
        if not ticker_item: return
        ticker = ticker_item.text()
        match = self._df[self._df["Ticker"] == ticker]
        if not match.empty:
            self._detail.show_row(match.iloc[0].to_dict())
            self._ticker_inp.setText(ticker)
            self._chart_panel.load_ticker(ticker)

    # ── Ticker lookup ─────────────────────────────────────────────────────────

    def _on_ticker_typed(self, text):
        if len(text) >= 1:
            self._lookup_tmr.start(700)

    def _do_lookup(self):
        self._lookup_tmr.stop()
        ticker = self._ticker_inp.text().strip().upper()
        if not ticker: return

        # Check if already in scan results
        if self._df is not None:
            match = self._df[self._df["Ticker"] == ticker]
            if not match.empty:
                self._detail.show_row(match.iloc[0].to_dict())
                # Highlight the row in table
                for ri in range(self._table.rowCount()):
                    item = self._table.item(ri, 1)
                    if item and item.text() == ticker:
                        self._table.selectRow(ri)
                        self._table.scrollToItem(item)
                        return
                return

        # Not in df — fetch from TradingView + yfinance
        if self._ticker_wkr and self._ticker_wkr.isRunning():
            return
        self._set_status(f"Looking up {ticker}…", AMBER)
        market = self._market.currentText()
        self._ticker_wkr = TickerWorker(ticker, market)
        self._ticker_wkr.done.connect(self._on_ticker_done)
        self._ticker_wkr.failed.connect(lambda m: self._set_status(f"✕  {m}", RED))
        self._ticker_wkr.start()

    def _on_ticker_done(self, df: pd.DataFrame):
        if df.empty: return
        row = df.iloc[0].to_dict()
        self._detail.show_row(row)
        qc = compute_qc_score(row)
        row_qc = {**row, **qc}
        self._qc_detail.show_row(row_qc)
        ticker = row.get("Ticker", "")
        self._set_status(
            f"✓  {ticker}  ·  {row.get('Tên Công Ty','')}  "
            f"·  {row.get('CS_Signal','')}  (Score {row.get('CS_Score',0)}/{N_CS})",
            GREEN
        )
        self._chart_panel.load_ticker(ticker)

    # ── Help ──────────────────────────────────────────────────────────────────

    def _show_help(self):
        dlg = HelpDialog(self)
        dlg.exec()

    # ── Export ────────────────────────────────────────────────────────────────

    def _on_export(self):
        if self._df is None: return
        from datetime import datetime
        market  = self._market.currentText()
        top     = self._top_spin.value()
        default = f"tv_top{top}_{market}_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
        path, _ = QFileDialog.getSaveFileName(
            self, "Export Excel", default, "Excel (*.xlsx)")
        if not path: return
        if not path.endswith(".xlsx"): path += ".xlsx"

        # Tính QC score — drop cột cũ trước để tránh duplicate sau concat
        df_export = self._df.copy()
        qc_df = pd.DataFrame(
            [compute_qc_score(r.to_dict()) for _, r in df_export.iterrows()])
        dup_cols = [c for c in qc_df.columns if c in df_export.columns]
        df_export = df_export.drop(columns=dup_cols, errors="ignore")
        df_export = pd.concat(
            [df_export.reset_index(drop=True), qc_df.reset_index(drop=True)], axis=1)

        # Hiện progress bar, disable nút
        self._export_bar.setValue(0)
        self._export_bar.setVisible(True)
        self._btn_export.setEnabled(False)
        self._set_status("Exporting…", AMBER)

        self._export_wkr = ExportWorker(
            df_export, path, market, top, self._use_yf.isChecked())
        self._export_wkr.progress.connect(self._on_export_progress)
        self._export_wkr.done.connect(self._on_export_done)
        self._export_wkr.failed.connect(self._on_export_failed)
        self._export_wkr.start()

    def _on_export_progress(self, pct, msg):
        self._export_bar.setValue(pct)
        if msg:
            self._set_status(f"💾  {msg}", AMBER)

    def _on_export_done(self, path):
        self._export_bar.setValue(100)
        self._set_status(f"✓  Exported → {path}", GREEN)
        self._btn_export.setEnabled(True)
        QTimer.singleShot(2000, lambda: self._export_bar.setVisible(False))

    def _on_export_failed(self, msg):
        self._export_bar.setVisible(False)
        self._set_status(f"✕  Export failed: {msg}", RED)
        self._btn_export.setEnabled(True)


# ─────────────────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    win = MainWindow()
    win.showMaximized()
    sys.exit(app.exec())
