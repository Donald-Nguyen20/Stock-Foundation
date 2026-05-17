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
    QDialog, QTextBrowser, QTabWidget,
)
from PySide6.QtCore import Qt, QThread, Signal, QTimer, QSortFilterProxyModel
from PySide6.QtGui import QFont, QColor, QKeySequence, QShortcut, QPixmap, QPainter
from PySide6.QtWidgets import QHeaderView
#
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
    ("EPS A%",      "EPS Annual%",  62,  Qt.AlignRight  | Qt.AlignVCenter),
    ("Rev Q%",      "Rev Qtr%",     62,  Qt.AlignRight  | Qt.AlignVCenter),
    ("GM%",         "Gross Margin%",60,  Qt.AlignRight  | Qt.AlignVCenter),
    ("ROE%",        "ROE%",         60,  Qt.AlignRight  | Qt.AlignVCenter),
    ("D/E",         "D/E",          52,  Qt.AlignRight  | Qt.AlignVCenter),
    ("P/E",         "P/E",          52,  Qt.AlignRight  | Qt.AlignVCenter),
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
    ("FCF/sh",    "FCF/sh",        65,  Qt.AlignRight  | Qt.AlignVCenter),
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
    ("Quality",   "QC_Signal",    138,  Qt.AlignCenter | Qt.AlignVCenter),
    ("EQ",        "EQ_Badge",     118,  Qt.AlignCenter | Qt.AlignVCenter),
]


def compute_qc_score(row: dict) -> dict:
    """Return QC pass/fail flags + Score + Signal for one stock row dict.

    Criteria (all sourced from TradingView):
      ROIC > 15%   — earns above cost of capital (core compounder trait)
      Op Mgn > 15% — operational leverage / pricing power
      GM > 40%     — structural margin advantage
      FCF/sh > 0   — actually generating free cash (not just accounting profit)
      D/E < 1.0    — clean balance sheet
      Moat Wide/Narrow — confirmed competitive advantage
    """
    result = {}
    score  = 0
    moat_good = {"WIDE  ★★★", "NARROW ★★"}

    for key, field, op, thr in [
        ("QC_ROIC", "ROIC%",         "gt", 15),
        ("QC_OPGM", "Op Margin%",    "gt", 15),
        ("QC_GM",   "Gross Margin%", "gt", 40),
        ("QC_FCF",  "FCF/sh",        "gt", 0),
        ("QC_DE",   "D/E",           "lt", 1.0),
    ]:
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
    if score >= 5:
        result["QC_Signal"] = "🏆 COMPOUNDER"
    elif score >= 3:
        result["QC_Signal"] = "⭐ QUALITY"
    elif score >= 1:
        result["QC_Signal"] = "○ AVERAGE"
    else:
        result["QC_Signal"] = "✗ WEAK"
    return result


from fundamental_chart import ChartWorkerPng, _RenderWorkerPng


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
    def __init__(self, parent=None):
        super().__init__(parent)
        self._dark = True
        self.setStyleSheet(f"background:{BG};")
        self._pix_q        = None   # QPixmap (quarterly)
        self._pix_a        = None   # QPixmap (annual)
        self._mode         = "quarterly"
        self._worker       = None
        self._render_worker = None
        self._chart_data   = None   # cached raw data dict from yfinance
        self._ticker       = ""
        self._spin_idx = 0
        self._spin_msg = ""
        self._spin_tmr = QTimer(self)
        self._spin_tmr.timeout.connect(self._tick)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)
        self._bar_w = self._build_bar()
        self._sep_line = self._sep()
        layout.addWidget(self._bar_w)
        layout.addWidget(self._sep_line)

        self._img = QLabel()
        self._img.setAlignment(Qt.AlignCenter)
        self._img.setStyleSheet(f"background:{BG}; color:{TEXT3};"
                                f" font-size:11px; letter-spacing:1px;")
        self._img.setText("Click a row to load chart")
        layout.addWidget(self._img, stretch=1)

    def _build_bar(self):
        w = QWidget(); w.setFixedHeight(36)
        w.setStyleSheet(f"background:{SURFACE};")
        h = QHBoxLayout(w); h.setContentsMargins(14, 0, 14, 0); h.setSpacing(8)

        self._chart_lbl = QLabel("—")
        self._chart_lbl.setStyleSheet(
            f"color:{TEXT1}; font-size:13px; font-weight:700;"
            f" font-family:'Consolas',monospace; letter-spacing:2px;")
        h.addWidget(self._chart_lbl)
        h.addStretch()

        self._chart_status = QLabel("Click a row to load chart")
        self._chart_status.setStyleSheet(
            f"color:{TEXT3}; font-size:9px; letter-spacing:1px;")
        h.addWidget(self._chart_status)
        h.addSpacing(10)

        self._btn_q = QPushButton("Quarterly")
        self._btn_a = QPushButton("Annual")
        for btn in (self._btn_q, self._btn_a):
            btn.setFixedSize(84, 22)
            btn.setCursor(Qt.PointingHandCursor)
        self._btn_q.clicked.connect(lambda: self._switch_mode("quarterly"))
        self._btn_a.clicked.connect(lambda: self._switch_mode("annual"))
        h.addWidget(self._btn_q); h.addWidget(self._btn_a)
        self._refresh_btns()
        return w

    def _sep(self):
        f = QFrame(); f.setFrameShape(QFrame.HLine); f.setFixedHeight(1)
        f.setStyleSheet(f"background:{BORDER}; border:none;"); return f

    # ── pixmap display ────────────────────────────────────────────────────────
    def _show_pixmap(self, pix):
        if pix and not pix.isNull():
            scaled = pix.scaled(self._img.size(),
                                Qt.KeepAspectRatio, Qt.SmoothTransformation)
            self._img.setPixmap(scaled)
            self._img.setText("")
        else:
            self._img.clear()

    def resizeEvent(self, event):
        super().resizeEvent(event)
        self._show_pixmap(self._pix_a if self._mode == "annual" else self._pix_q)

    # ── Public API ────────────────────────────────────────────────────────────
    def load_ticker(self, ticker: str):
        ticker = ticker.upper()
        if ticker == self._ticker and (self._pix_q or self._pix_a):
            return
        self._ticker = ticker
        self._pix_q = None; self._pix_a = None
        self._mode  = "quarterly"
        self._chart_lbl.setText(ticker)
        self._img.clear()
        self._img.setText(f"Loading {ticker}…")
        self._refresh_btns()

        if self._worker and self._worker.isRunning():
            self._worker.terminate(); self._worker.wait()

        self._chart_data = None
        self._spin_msg = f"Loading {ticker}…"
        self._spin_tmr.start(80)
        self._worker = ChartWorkerPng(ticker, dark_mode=self._dark)
        self._worker.data_ready.connect(self._cache_data)
        self._worker.done.connect(self._on_done)
        self._worker.failed.connect(self._on_failed)
        self._worker.msg.connect(lambda m: setattr(self, "_spin_msg", m))
        self._worker.start()

    # ── Slots ─────────────────────────────────────────────────────────────────
    def _on_done(self, png_a, png_q, summary):
        self._spin_tmr.stop()
        self._pix_q = QPixmap(); self._pix_q.loadFromData(png_q)
        if png_a:
            self._pix_a = QPixmap(); self._pix_a.loadFromData(png_a)
        self._mode = "quarterly"
        self._show_pixmap(self._pix_q)
        parts = summary.split("·")
        self._set_status(f"✓  {parts[-1].strip()}" if len(parts) > 1 else "✓", GREEN)
        self._refresh_btns()

    def _on_failed(self, msg):
        self._spin_tmr.stop()
        self._img.clear()
        self._img.setText(f"Error: {msg}")
        self._img.setStyleSheet(f"background:{BG}; color:{RED};"
                                f" font-size:11px; letter-spacing:1px;")
        self._set_status(f"✕  {msg}", RED)

    def _cache_data(self, d):
        self._chart_data = d

    def _fast_rerender(self):
        if not self._chart_data or not self._ticker:
            return
        if self._render_worker and self._render_worker.isRunning():
            self._render_worker.terminate()
            self._render_worker.wait()
        self._pix_q = None
        self._pix_a = None
        self._spin_msg = "Re-rendering…"
        self._spin_tmr.start(80)
        self._render_worker = _RenderWorkerPng(
            self._ticker, self._chart_data, dark_mode=self._dark)
        self._render_worker.done.connect(self._on_done)
        self._render_worker.failed.connect(self._on_failed)
        self._render_worker.msg.connect(lambda m: setattr(self, "_spin_msg", m))
        self._render_worker.start()

    def _switch_mode(self, mode):
        if mode == self._mode: return
        pix = self._pix_a if mode == "annual" else self._pix_q
        if not pix: return
        self._mode = mode
        self._show_pixmap(pix)
        self._refresh_btns()

    def _set_status(self, text, color=TEXT3):
        self._chart_status.setText(text)
        self._chart_status.setStyleSheet(
            f"color:{color}; font-size:9px; letter-spacing:1px;")

    def _tick(self):
        self._spin_idx = (self._spin_idx + 1) % len(SPINNER)
        self._set_status(f"{SPINNER[self._spin_idx]}  {self._spin_msg}", AMBER)

    def _refresh_btns(self):
        has_q = self._pix_q is not None and not self._pix_q.isNull()
        has_a = self._pix_a is not None and not self._pix_a.isNull()
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
        self._img.setStyleSheet(f"background:{BG}; color:{TEXT3}; font-size:11px; letter-spacing:1px;")
        self._refresh_btns()
        if self._ticker:
            if self._chart_data:
                self._fast_rerender()
            else:
                prev = self._ticker
                self._ticker = ""
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
        elif fmt == "mcap":
            text = f"{val:,.1f}"
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
            _, raw = (Query().set_markets(self.market)
                      .select(*FETCH_COLS)
                      .order_by("market_cap_basic", ascending=False)
                      .limit(self.top)
                      .get_scanner_data())
            df = clean_df(raw)
            self.progress.emit(f"Got {len(df)} stocks.")

            moat_cache = None
            if self.use_yf:
                tickers = df["Ticker"].tolist()
                sectors = dict(zip(df["Ticker"], df["Sector"].fillna("")))
                moat_cache = {}
                total = len(tickers)
                for i, ticker in enumerate(tickers, 1):
                    if self._cancel:
                        self.failed.emit("Cancelled.")
                        return
                    sector = sectors.get(ticker, "")
                    proxy, score, w52_pct, eq_badge = fetch_moat_yfinance(ticker, sector)
                    moat_cache[ticker] = (proxy, score, w52_pct, eq_badge)
                    self.ticker_update.emit(i, total, ticker, score)
                    time.sleep(0.3)

            self.progress.emit("Fetching market direction (^GSPC)…")
            market_ok = fetch_market_direction()
            self.progress.emit("Applying CAN SLIM scoring…")
            df = score_canslim(df, moat_cache, market_ok=market_ok)
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
            _, raw = (Query().set_markets(self.market)
                      .select(*FETCH_COLS)
                      .where(Column("name") == self.ticker)
                      .get_scanner_data())
            if raw.empty:
                self.failed.emit(f"'{self.ticker}' not found on TradingView.")
                return
            df = clean_df(raw)
            proxy, score, w52_pct, eq_badge = fetch_moat_yfinance(
                self.ticker, df["Sector"].iloc[0] if "Sector" in df.columns else "")
            moat_cache = {self.ticker: (proxy, score, w52_pct, eq_badge)}
            market_ok = fetch_market_direction()
            df = score_canslim(df, moat_cache, market_ok=market_ok)
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
            if lbl.text() == "—":
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
            if lbl.text() == "—":
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

    def _html(self):
        ops = {"gt": ">", "lt": "<"}
        full_desc = {
            "C": "EPS quý gần nhất tăng ≥25% so cùng kỳ năm trước. Tiêu chí quan trọng nhất — tăng trưởng lợi nhuận hiện tại phải mạnh và tăng tốc.",
            "A": "EPS năm (FY) tăng ≥20% YoY. Xác nhận tăng trưởng bền vững qua nhiều năm, không phải chỉ một quý tốt đột biến.",
            "S": "Doanh thu quý tăng ≥20% YoY. Đảm bảo EPS growth đến từ top-line thực sự, không phải từ cắt chi phí hay buyback.",
            "L": "1-Year Performance >20% là proxy cho Relative Strength. Cổ phiếu dẫn đầu thường outperform index 12 tháng trước breakout.",
            "Q": "Gross Margin >40% — dấu hiệu pricing power và lợi thế cạnh tranh. Công ty có thể bảo vệ biên lợi nhuận khi chi phí tăng.",
            "R": "ROE >17% — hiệu quả sử dụng vốn chủ sở hữu. Bao nhiêu lợi nhuận tạo ra trên mỗi đồng equity của cổ đông.",
            "M": "3-Month Performance >0% — giá đang trong xu hướng tăng ngắn hạn. Không mua cổ phiếu đang giảm (catching a falling knife).",
            "D": "D/E <2.0 — kiểm soát đòn bẩy tài chính. Tránh công ty overleveraged, đặc biệt nguy hiểm khi lãi suất tăng cao.",
            "N": "Price ≥90% of 52-Week High — cổ phiếu đang ở vùng sức mạnh dài hạn, gần đỉnh 52 tuần. O'Neil: mua cổ phiếu phá đỉnh, không phải đang trong downtrend.",
            "MKT": "S&P 500 trên MA50 & MA200 — thị trường chung đang uptrend. Ngay cả cổ phiếu tốt cũng khó tăng khi thị trường correction.",
        }
        canslim_rows = ""
        for i, key in enumerate(CS_KEYS):
            cfg  = CANSLIM[key]
            op   = ops[cfg["op"]]
            thr  = f"{op}&nbsp;{cfg['thr']}{'%' if cfg['thr'] != 2.0 else ''}"
            name = cfg["label"].split("—")[1].strip() if "—" in cfg["label"] else cfg["label"]
            bg   = "#0E1220" if i % 2 == 0 else "#0B0E18"
            canslim_rows += f"""
              <tr style="background-color:{bg};">
                <td style="color:#3D8EF0;font-weight:700;font-family:Consolas,monospace;
                           font-size:15px;padding:9px 13px;width:36px;">{key}</td>
                <td style="color:#DCE4EE;padding:9px 13px;width:120px;">{name}</td>
                <td style="color:#B8C4D0;font-family:Consolas,monospace;font-size:11px;
                           padding:9px 13px;width:115px;">{cfg['field']}</td>
                <td style="color:#34C472;font-weight:700;padding:9px 13px;width:70px;">{thr}</td>
                <td style="color:#DCE4EE;font-size:12px;padding:9px 13px;">{full_desc[key]}</td>
              </tr>"""

        signal_rows = "".join([
            f"""<tr style="background-color:{bg};">
                  <td style="color:{fg};font-weight:700;font-size:13px;padding:10px 14px;width:140px;">{icon}&nbsp;{label}</td>
                  <td style="color:{fg};font-weight:700;font-size:13px;padding:10px 14px;width:85px;">{score}</td>
                  <td style="color:{fg};font-size:12px;padding:10px 14px;">{desc}</td>
                </tr>"""
            for icon, label, score, fg, bg, desc in [
                ("🟢", "STRONG BUY", f"≥{round(N_CS*0.875)} / {N_CS}", "#1A5C2B", "#C6EFCE",
                 "Đạt ≥87.5% tiêu chí. Nền tảng cơ bản + kỹ thuật đều mạnh. Ưu tiên theo dõi và cân nhắc mua."),
                ("🔵", "BUY",        f"≥{round(N_CS*0.625)} / {N_CS}", "#1B3A5C", "#DDEEFF",
                 "Đạt ≥62.5% tiêu chí. Nền tảng tốt, đáng xem xét nhưng kiểm tra thêm tiêu chí còn thiếu."),
                ("🟡", "WATCH",      f"≥{round(N_CS*0.375)} / {N_CS}", "#7D6608", "#FFF2CC",
                 "Đạt ≥37.5% tiêu chí. Tiềm năng nhưng chưa đủ điều kiện. Đưa vào watchlist, chờ cải thiện."),
                ("🔴", "SKIP",       f"< {round(N_CS*0.375)} / {N_CS}", "#9C0006", "#FFC7CE",
                 "Không đủ điều kiện theo CAN SLIM. Bỏ qua hoặc chờ fundamental xoay chiều."),
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
            bg = "#0E1220" if i % 2 == 0 else "#0B0E18"
            proxy_sector_rows += (
                f'<tr style="background-color:{bg};">'
                f'<td style="color:#DCE4EE;font-size:12px;padding:7px 13px;">{sector}</td>'
                f'<td style="color:#DCE4EE;font-size:12px;padding:7px 13px;">{proxy}</td>'
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
            bg = "#0E1220" if i % 2 == 0 else "#0B0E18"
            proxy_type_rows += (
                f'<tr style="background-color:{bg};">'
                f'<td style="color:#3D8EF0;font-weight:600;font-size:12px;padding:8px 13px;">{ptype}</td>'
                f'<td style="color:#DCE4EE;font-size:12px;padding:8px 13px;">{pdesc}</td>'
                f'</tr>'
            )

        # ── Quality Compounder criteria rows ──────────────────────────────────
        qc_full_desc = {
            "ROIC": "ROIC >15% — Return on Invested Capital: lợi nhuận tạo ra trên mỗi đồng vốn đầu tư (cả debt lẫn equity). ĐÂY là tiêu chí cốt lõi nhất của compounder — công ty kiếm được nhiều hơn chi phí vốn, tạo ra giá trị thực sự cho cổ đông dài hạn.",
            "OPGM": "Operating Margin >15% — biên lợi nhuận hoạt động từ business core, trước lãi vay và thuế. Cao và ổn định qua nhiều năm là dấu hiệu rõ ràng của pricing power và cấu trúc chi phí tối ưu.",
            "GM":   "Gross Margin >40% — biên lợi nhuận gộp. Mức nền tảng quyết định công ty có thể tái đầu tư vào R&D, marketing, và nhân sự hay không. Compounder thực sự cần GM cao để duy trì lợi thế cạnh tranh.",
            "FCF":  "FCF/share >0 — Free Cash Flow trên mỗi cổ phiếu dương, nghĩa là công ty thực sự tạo ra tiền mặt, không chỉ là lợi nhuận kế toán. Compounder dùng FCF để buyback, dividend, hoặc tái đầu tư — đây là nguồn gốc của compound return.",
            "DE":   "D/E <1.0 — đòn bẩy tài chính thấp. Tiêu chuẩn chặt hơn CAN SLIM (D/E <2). Compounder chất lượng tăng trưởng bằng FCF tái đầu tư, không cần đòn bẩy nợ. Balance sheet sạch giúp vượt qua khủng hoảng và tận dụng cơ hội M&A.",
            "MOAT": "Moat Wide ★★★ hoặc Narrow ★★ — có lợi thế cạnh tranh được xác nhận bởi ROE 5yr avg + GM 5yr avg. Moat là điều kiện bắt buộc để ROIC cao được duy trì dài hạn — không có moat, lợi nhuận sẽ bị cạnh tranh xói mòn.",
        }
        qc_criteria_rows = ""
        for i, key in enumerate(QC_KEYS):
            cfg  = QC_CRITERIA[key]
            bg   = "#0E1220" if i % 2 == 0 else "#0B0E18"
            if cfg["op"] == "in":
                thr_str = "Wide / Narrow"
            elif cfg["op"] == "gt":
                thr_str = f"&gt;&nbsp;{cfg['thr']}{'%' if cfg['thr'] >= 15 else ''}"
            else:
                thr_str = f"&lt;&nbsp;{cfg['thr']}"
            qc_criteria_rows += f"""
              <tr style="background-color:{bg};">
                <td style="color:#34C472;font-weight:700;font-family:Consolas,monospace;
                           font-size:13px;padding:9px 13px;width:56px;">{key}</td>
                <td style="color:#DCE4EE;padding:9px 13px;width:130px;">{cfg['label']}</td>
                <td style="color:#34C472;font-weight:700;padding:9px 13px;width:120px;">{thr_str}</td>
                <td style="color:#DCE4EE;font-size:12px;padding:9px 13px;">{qc_full_desc[key]}</td>
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

        return f"""<html><head><style>
  body  {{ background-color:#0B0E18; color:#DCE4EE;
           font-family:'Segoe UI',sans-serif; font-size:13px;
           margin:0; padding:0; }}
  h1    {{ color:#3D8EF0; font-size:20px; letter-spacing:4px;
           font-weight:700; margin:0 0 5px 0; }}
  .sub  {{ color:#7A8899; font-size:11px; letter-spacing:2px;
           margin:0 0 20px 0; }}
  h2    {{ color:#B8C4D0; font-size:12px; letter-spacing:3px;
           font-weight:700; margin:26px 0 9px 0;
           border-bottom:1px solid #1A2133; padding-bottom:5px; }}
  table {{ border-collapse:collapse; width:100%; margin-bottom:10px; }}
  th    {{ background-color:#0E1220; color:#B8C4D0; text-align:left;
           padding:8px 13px; font-size:11px; letter-spacing:1px;
           font-weight:600; border-bottom:1px solid #1A2133; }}
  .tip  {{ color:#7A8899; font-size:11px; font-style:italic;
           margin:5px 0 0 0; }}
  li    {{ margin:7px 0; color:#DCE4EE; font-size:12px; }}
  b     {{ color:#FFFFFF; font-weight:600; }}
</style></head><body>
<h1>FUNDAMENTAL SCREENER</h1>
<p class="sub">HƯỚNG DẪN SỬ DỤNG  ·  CAN SLIM · QUALITY COMPOUNDER · EQ BADGE · MOAT · SIGNAL</p>

<h2>① CAN SLIM — 8 TIÊU CHÍ LỌC CỔ PHIẾU</h2>
<table>
  <tr>
    <th style="width:36px;">Key</th>
    <th style="width:120px;">Tên</th>
    <th style="width:115px;">Chỉ số</th>
    <th style="width:70px;">Ngưỡng</th>
    <th>Ý nghĩa</th>
  </tr>
  {canslim_rows}
</table>
<p class="tip">⚠  Dữ liệu từ TradingView (TTM/FY). EPS &amp; Revenue là YoY growth %. ROE &amp; Gross Margin là trailing twelve months.</p>

<h2>② CAN SLIM SIGNAL — KẾT QUẢ TỔNG HỢP</h2>
<table>
  <tr>
    <th style="width:140px;">Tín hiệu</th>
    <th style="width:85px;">Điểm</th>
    <th>Ý nghĩa &amp; Hành động gợi ý</th>
  </tr>
  {signal_rows}
</table>
<p class="tip">Score = số tiêu chí đạt được (tối đa {N_CS}). ✓ = đạt (+1 điểm) · ✗ = không đạt · — = thiếu dữ liệu (không tính vào score).</p>

<h2>③ QUALITY COMPOUNDER — 6 TIÊU CHÍ CHẤT LƯỢNG BỀN VỮNG</h2>
<p style="color:#7A8899;font-size:12px;margin:0 0 10px 0;">Tab <b style="color:#34C472;">Quality Compounder</b> tìm doanh nghiệp có khả năng tăng trưởng kép bền vững dài hạn — không chỉ tốt về momentum mà còn mạnh về chất lượng nền tảng.</p>
<table>
  <tr>
    <th style="width:52px;">Key</th>
    <th style="width:130px;">Chỉ số</th>
    <th style="width:110px;">Ngưỡng</th>
    <th>Ý nghĩa</th>
  </tr>
  {qc_criteria_rows}
</table>
<p class="tip">⚠  Tiêu chí MOAT yêu cầu bật "yfinance Moat" để có dữ liệu 5 năm chính xác. Khi tắt, Moat được ước tính từ TTM ROE &amp; GM.</p>

<h2>④ QUALITY COMPOUNDER SIGNAL</h2>
<table>
  <tr>
    <th style="width:160px;">Tín hiệu</th>
    <th style="width:90px;">Điểm</th>
    <th>Ý nghĩa &amp; Hành động gợi ý</th>
  </tr>
  {qc_signal_rows}
</table>
<p class="tip">Score = số trong 6 tiêu chí đạt. Quick Filter "🟢 STRONG BUY" trên tab Quality Compounder sẽ lọc cột 🏆 COMPOUNDER.</p>

<h2>⑤ EQ BADGE — CHẤT LƯỢNG LỢI NHUẬN (Earnings Quality)</h2>
<p style="color:#7A8899;font-size:12px;margin:0 0 10px 0;">
  Cột <b style="color:#4DB6AC;">EQ</b> trong bảng Quality Compounder trả lời câu hỏi:
  <b style="color:#FFFFFF;">Lợi nhuận công ty báo cáo có phải là tiền mặt thực hay chỉ là con số kế toán?</b>
</p>

<p style="color:#DCE4EE;font-size:12px;margin:0 0 6px 0;font-weight:700;">Tại sao cần chỉ số này?</p>
<p style="color:#B8C4D0;font-size:12px;margin:0 0 12px 0;line-height:1.7;">
  Net Income (lợi nhuận ròng) là con số kế toán — có thể bị <b>tô vẽ</b> bằng cách ghi nhận doanh thu sớm,
  trì hoãn chi phí, hoặc thay đổi ước tính kế toán. Trong khi đó,
  <b>Free Cash Flow (FCF)</b> là tiền mặt thực sự chảy vào tài khoản công ty sau khi đã trừ chi phí đầu tư —
  không thể làm giả. Nếu Net Income cao nhưng FCF thấp, đó là dấu hiệu lợi nhuận phụ thuộc vào kế toán,
  không phải kinh doanh thực chất.
</p>

<p style="color:#DCE4EE;font-size:12px;margin:0 0 6px 0;font-weight:700;">Công thức tính:</p>
<p style="color:#4DB6AC;font-size:13px;font-family:Consolas,monospace;
          background:#0D2020;padding:8px 14px;border-radius:4px;margin:0 0 12px 0;">
  FCF / Net Income Ratio &nbsp;=&nbsp; Free Cash Flow &nbsp;÷&nbsp; Net Income &nbsp;(TTM, từ yfinance)
</p>

<table>
  <tr>
    <th style="width:160px;">Badge</th>
    <th style="width:130px;">Điều kiện</th>
    <th>Ý nghĩa &amp; Hành động</th>
  </tr>
  <tr style="background-color:#0E1220;">
    <td style="color:#1A5C2B;font-weight:700;font-size:13px;padding:10px 14px;background:#C6EFCE;">💚 Cash Backed</td>
    <td style="color:#1A5C2B;font-weight:700;padding:10px 14px;background:#C6EFCE;">Ratio ≥ 0.80</td>
    <td style="color:#DCE4EE;font-size:12px;padding:10px 14px;">
      Mỗi 1 đồng lợi nhuận kế toán, công ty thực thu ≥ 0.8 đồng tiền mặt.
      Lợi nhuận <b>đáng tin cậy cao</b> — tài chính minh bạch, ít rủi ro kế toán.
      Compounder thực sự thường có chỉ số này ≥ 1.0 (FCF &gt; Net Income).
    </td>
  </tr>
  <tr style="background-color:#0B0E18;">
    <td style="color:#7D6608;font-weight:700;font-size:13px;padding:10px 14px;background:#FFF2CC;">🟡 Mixed</td>
    <td style="color:#7D6608;font-weight:700;padding:10px 14px;background:#FFF2CC;">0.30 ≤ Ratio &lt; 0.80</td>
    <td style="color:#DCE4EE;font-size:12px;padding:10px 14px;">
      Tiền mặt thu về chỉ bằng 30–80% lợi nhuận báo cáo. Có thể do đầu tư mở rộng nặng (capex cao),
      hoặc tốc độ ghi nhận doanh thu nhanh hơn thu tiền. <b>Cần xem thêm</b> bối cảnh ngành
      và xu hướng nhiều năm — một số công ty tăng trưởng nhanh có FCF thấp tạm thời là bình thường.
    </td>
  </tr>
  <tr style="background-color:#0E1220;">
    <td style="color:#9C0006;font-weight:700;font-size:13px;padding:10px 14px;background:#FFC7CE;">🔴 Accrual Heavy</td>
    <td style="color:#9C0006;font-weight:700;padding:10px 14px;background:#FFC7CE;">Ratio &lt; 0.30</td>
    <td style="color:#DCE4EE;font-size:12px;padding:10px 14px;">
      Lợi nhuận chủ yếu là con số kế toán, tiền mặt thực thu rất ít.
      <b>Cảnh báo đỏ</b> — đặc biệt nguy hiểm nếu kết hợp D/E cao.
      Rủi ro: công ty có thể phải huy động vốn mới (dilution), hoặc lợi nhuận sẽ bị điều chỉnh giảm trong tương lai.
    </td>
  </tr>
  <tr style="background-color:#0B0E18;">
    <td style="color:#7A8899;font-weight:700;font-size:13px;padding:10px 14px;">— (không có)</td>
    <td style="color:#7A8899;padding:10px 14px;">yfinance tắt</td>
    <td style="color:#7A8899;font-size:12px;padding:10px 14px;">
      EQ Badge chỉ tính được khi bật checkbox "yfinance Moat" vì cần cashflow statement từ yfinance.
      Khi tắt yfinance, cột EQ hiển thị "—" cho tất cả.
    </td>
  </tr>
</table>
<p class="tip">⚠  EQ Badge dùng số liệu <b>năm tài chính gần nhất</b> (annual, không phải TTM quarterly).
Một số ngành có FCF tự nhiên thấp hơn Net Income do capex lớn (bán dẫn, pharma R&amp;D) — cần so sánh trong ngành, không áp dụng ngưỡng cứng cho mọi sector.</p>

<h2>⑥ MOAT SCORE — LỢI THẾ CẠNH TRANH (Economic Moat)</h2>
<table>
  <tr>
    <th style="width:160px;">Moat</th>
    <th style="width:240px;">Điều kiện (5yr avg từ yfinance)</th>
    <th>Ý nghĩa</th>
  </tr>
  {moat_rows}
</table>
<p class="tip">Moat Score dựa trên ROE 5yr avg + GM 5yr avg từ yfinance (hoặc TTM từ TradingView nếu tắt checkbox).
Khi tắt "yfinance Moat", dùng TV TTM ROE &amp; GM thay thế (nhanh hơn nhưng kém chính xác hơn).</p>

<h2>⑦ MOAT PROXY — LOẠI LỢI THẾ CẠNH TRANH</h2>
<p style="color:#DCE4EE;font-size:12px;margin:0 0 9px 0;">Moat Proxy được gán tự động theo <b>ngành (Sector)</b> của cổ phiếu.</p>
<table>
  <tr>
    <th style="width:200px;">Sector</th>
    <th>Moat Proxy</th>
  </tr>
  {proxy_sector_rows}
</table>

<p style="color:#B8C4D0;font-size:12px;margin:16px 0 8px 0;font-weight:700;letter-spacing:1px;">GIẢI THÍCH CÁC LOẠI LỢI THẾ</p>
<table>
  <tr>
    <th style="width:220px;">Loại lợi thế</th>
    <th>Ý nghĩa &amp; Ví dụ</th>
  </tr>
  {proxy_type_rows}
</table>
<p class="tip">Moat Proxy chỉ là ước tính định tính theo ngành — không thay thế phân tích sâu từng công ty.
Hai công ty cùng ngành có thể có moat type khác nhau hoàn toàn.</p>

<h2>⑧ CÁCH SỬ DỤNG</h2>
<ul>
  <li><b>SCAN</b> — Lấy top N cổ phiếu theo Market Cap từ TradingView. Tự động tính CAN SLIM score và Quality Compounder score cho tất cả.</li>
  <li><b>Tab CAN SLIM / Quality Compounder</b> — Chuyển tab để xem hai góc nhìn khác nhau trên cùng một bộ dữ liệu. Chart panel bên phải dùng chung.</li>
  <li><b>yfinance Moat</b> — Bật để lấy ROE/GM 5 năm từ yfinance cho Moat chính xác hơn, đồng thời tính <b>EQ Badge</b> (chất lượng lợi nhuận FCF/NI) cho tab Quality Compounder. Chậm hơn ~0.3 giây/cổ phiếu.</li>
  <li><b>TICKER lookup</b> — Nhập mã (ví dụ: NVDA) rồi Enter hoặc nhấn LOOKUP để tra cứu ngay mà không cần scan toàn bộ.</li>
  <li><b>FILTER</b> — Gõ vào ô filter để lọc bảng đang hiển thị theo ticker hoặc tên công ty theo thời gian thực.</li>
  <li><b>Quick Filter chips</b> — Lọc nhanh STRONG BUY (hoặc COMPOUNDER ở tab QC), Moat, 1Y%, 52W High. Dùng nút ✕ để xóa tất cả.</li>
  <li><b>Click vào hàng</b> — Xem breakdown chi tiết (✓/✗ từng tiêu chí) ở panel phía dưới bảng. Chart tự động tải.</li>
  <li><b>Sort</b> — Click vào tiêu đề cột để sắp xếp. Cột số sắp xếp đúng theo giá trị số, không phải theo chữ.</li>
  <li><b>Export</b> — Xuất Excel với 3 sheet: <b>Data</b> (toàn bộ cổ phiếu), <b>Quality Compounder</b> (nếu có), <b>Dashboard</b> (tóm tắt trực quan). Xem hướng dẫn đọc Dashboard ở phần ⑨.</li>
  <li><b>F1</b> — Mở màn hình hướng dẫn này bất cứ lúc nào.</li>
</ul>

<h2>⑨ DASHBOARD EXCEL — CÁCH ĐỌC</h2>
<p style="color:#7A8899;font-size:12px;margin:0 0 12px 0;">
  Sheet <b style="color:#3D8EF0;">Dashboard</b> trong file Excel là tóm tắt toàn bộ kết quả scan — đọc từ trên xuống theo thứ tự ưu tiên.
</p>

<table>
  <tr>
    <th style="width:200px;">Vùng</th>
    <th>Nội dung &amp; Cách đọc</th>
  </tr>
  <tr style="background-color:#0E1220;">
    <td style="color:#3D8EF0;font-weight:700;font-size:13px;padding:10px 14px;">❶ CAN SLIM KPIs</td>
    <td style="color:#DCE4EE;font-size:12px;padding:10px 14px;">
      6 thẻ tóm tắt kết quả CAN SLIM: <b>Total Stocks</b> · <b># Strong Buy</b> · <b># Buy</b> · <b>Avg Score</b> · <b>Avg 1Y%</b> · <b># Near 52W High</b>.
      Đây là bức tranh tổng quan — thị trường đang có bao nhiêu cổ phiếu đủ tiêu chuẩn.
    </td>
  </tr>
  <tr style="background-color:#0B0E18;">
    <td style="color:#34C472;font-weight:700;font-size:13px;padding:10px 14px;">❷ QC KPIs</td>
    <td style="color:#DCE4EE;font-size:12px;padding:10px 14px;">
      6 thẻ Quality Compounder: <b># Compounder</b> · <b># Quality</b> · <b>Avg QC Score</b> · <b>Avg ROIC%</b> · <b># Dual Leaders</b> · <b># Cash Backed</b>.
      <b style="color:#FFC000;">Cash Backed</b> chỉ đếm trong số cổ phiếu đạt QC signal — không tính tất cả 300 mã.
      Hiển thị <b>—</b> nếu chưa bật yfinance Moat.
    </td>
  </tr>
  <tr style="background-color:#0E1220;">
    <td style="color:#FFC000;font-weight:700;font-size:13px;padding:10px 14px;">⚠ Risk Flags</td>
    <td style="color:#DCE4EE;font-size:12px;padding:10px 14px;">
      3 cảnh báo rủi ro thị trường: <b>D/E &gt; 1</b> (đòn bẩy cao) · <b>P/E &gt; 50</b> (định giá cao) · <b>1M% &lt; −10%</b> (đang giảm mạnh).
      Đọc dòng này trước khi quyết định — nếu có nhiều mã trong cảnh báo, thị trường đang rủi ro cao.
    </td>
  </tr>
  <tr style="background-color:#0B0E18;">
    <td style="color:#FFD700;font-weight:700;font-size:13px;padding:10px 14px;">⭐ Dual Leaders</td>
    <td style="color:#DCE4EE;font-size:12px;padding:10px 14px;">
      Bảng vàng: cổ phiếu đạt <b>cả hai</b> tiêu chuẩn — <b style="color:#34C472;">CS Score ≥ 7</b> và <b style="color:#3D8EF0;">QC Score ≥ 4</b>.
      Đây là danh sách ưu tiên cao nhất: vừa có momentum (CAN SLIM) vừa có chất lượng bền vững (QC).
      Cột <b>1Y%</b> có thanh màu xanh — dài hơn = hiệu suất tốt hơn.
      Sắp xếp: CS Score cao nhất trước, sau đó QC Score.
    </td>
  </tr>
  <tr style="background-color:#0E1220;">
    <td style="color:#E040FB;font-weight:700;font-size:13px;padding:10px 14px;">🎯 Top Picks</td>
    <td style="color:#DCE4EE;font-size:12px;padding:10px 14px;">
      3 bảng nhỏ đặt cạnh nhau, mỗi bảng top 5:<br>
      &nbsp;• <b style="color:#FF7043;">🚀 Momentum</b> — Top 5 theo 1Y% cao nhất (tất cả cổ phiếu).<br>
      &nbsp;• <b style="color:#4DB6AC;">💎 Quality</b> — Top 5 theo ROIC% cao nhất, chỉ lọc COMPOUNDER.<br>
      &nbsp;• <b style="color:#FFD54F;">💰 Value</b> — Top 5 P/E thấp nhất với CS Score ≥ 6 (tăng trưởng tốt, định giá hợp lý).
    </td>
  </tr>
  <tr style="background-color:#0B0E18;">
    <td style="color:#B8C4D0;font-weight:700;font-size:13px;padding:10px 14px;">❸ Chart Analysis</td>
    <td style="color:#DCE4EE;font-size:12px;padding:10px 14px;">
      <b>Trái — Scatter CS × QC:</b> Mỗi chấm = 1 cổ phiếu. Góc trên phải (CS≥7, QC≥4) là vùng Dual Leaders lý tưởng.
      Đường cam dọc x=7 và ngang y=4 là ngưỡng phân chia. Màu chấm = tín hiệu CAN SLIM.<br><br>
      <b>Phải — Bar Top 10 1Y%:</b> 10 cổ phiếu có hiệu suất 1 năm tốt nhất. Đọc cùng Scatter để tìm cổ phiếu vừa momentum vừa chất lượng.
    </td>
  </tr>
  <tr style="background-color:#0E1220;">
    <td style="color:#B8C4D0;font-weight:700;font-size:13px;padding:10px 14px;">❹ Sector Breakdown</td>
    <td style="color:#DCE4EE;font-size:12px;padding:10px 14px;">
      Pivot theo ngành: <b># Stocks</b> · <b># Strong Buy</b> · <b># Buy</b> · <b># Compounder</b> · <b>Avg CS</b> · <b>Avg QC</b> · <b>Avg 1Y%</b>.
      Sắp xếp theo # Strong Buy giảm dần — <b style="color:#FFD700;">Top 3 ngành</b> được highlight vàng/xanh lá/xanh dương.
      Dùng bảng này để tìm <b>sector rotation</b>: ngành nào đang có nhiều cổ phiếu mạnh nhất.
    </td>
  </tr>
</table>

<p style="color:#B8C4D0;font-size:12px;margin:16px 0 6px 0;font-weight:700;letter-spacing:1px;">QUY TRÌNH ĐỌC DASHBOARD NHANH (30 giây)</p>
<ol style="padding-left:20px;margin:0;">
  <li style="margin:6px 0;color:#DCE4EE;font-size:12px;"><b>❶+❷ KPIs</b> — Thị trường đang tốt hay xấu? # Strong Buy &gt; 10% tổng là dấu hiệu tốt.</li>
  <li style="margin:6px 0;color:#DCE4EE;font-size:12px;"><b>⚠ Risk Flags</b> — Có nhiều cảnh báo không? Nếu có, tăng tiêu chuẩn lọc.</li>
  <li style="margin:6px 0;color:#DCE4EE;font-size:12px;"><b>⭐ Dual Leaders</b> — Đây là shortlist thực sự. Nghiên cứu kỹ từng mã trong bảng này.</li>
  <li style="margin:6px 0;color:#DCE4EE;font-size:12px;"><b>🎯 Top Picks</b> — Bổ sung từ 3 góc nhìn: momentum / chất lượng / định giá.</li>
  <li style="margin:6px 0;color:#DCE4EE;font-size:12px;"><b>❸ Scatter</b> — Xác nhận trực quan: mã nào ở góc trên phải là ứng viên tốt nhất.</li>
  <li style="margin:6px 0;color:#DCE4EE;font-size:12px;"><b>❹ Sector</b> — Sector nào dẫn đầu? Ưu tiên mã trong sector mạnh.</li>
</ol>
<p class="tip">⚠  Dashboard chỉ xuất hiện khi nhấn <b>Export Excel</b> từ app. Dữ liệu phản ánh thời điểm scan gần nhất — xem timestamp ở góc phải hàng đầu của sheet.</p>
</body></html>"""


# ─────────────────────────────────────────────────────────────────────────────
# Main window
# ─────────────────────────────────────────────────────────────────────────────
class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Fundamental Screener")
        self.resize(1700, 960)
        self._is_dark     = True
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
        self._tag_lbl = QLabel("CAN SLIM · TradingView · yfinance")
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
        self._market.addItems(["america", "nasdaq", "nyse", "euronext", "hong_kong", "vietnam"])
        self._market.setFixedHeight(32)
        self._market.setStyleSheet(self._combo_style())

        lbl_top = QLabel("TOP")
        self._lbl_top = lbl_top
        lbl_top.setStyleSheet(
            f"color:{TEXT1}; font-size:9px; font-weight:600; letter-spacing:1.5px;")
        self._top_spin = QSpinBox()
        self._top_spin.setRange(50, 1000); self._top_spin.setValue(300); self._top_spin.setSingleStep(50)
        self._top_spin.setFixedSize(80, 32)
        self._top_spin.setStyleSheet(self._spinbox_style())

        self._use_yf = QCheckBox("yfinance Moat")
        self._use_yf.setChecked(False)
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
        row1.addWidget(lbl_top);    row1.addWidget(self._top_spin)
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

        # Tab 0 — CAN SLIM
        cs_w = QWidget(); cs_w.setStyleSheet(f"background:{BG};")
        cs_v = QVBoxLayout(cs_w)
        cs_v.setContentsMargins(0, 0, 0, 0); cs_v.setSpacing(0)
        v_split = QSplitter(Qt.Vertical)
        v_split.setStyleSheet(f"QSplitter::handle {{ background:{BORDER}; height:2px; }}")
        v_split.addWidget(self._build_table())
        v_split.addWidget(self._build_detail())
        v_split.setSizes([580, 180])
        cs_v.addWidget(v_split)
        tabs.addTab(cs_w, "  CAN SLIM  ")

        # Tab 1 — Quality Compounder
        qc_w = QWidget(); qc_w.setStyleSheet(f"background:{BG};")
        qc_v = QVBoxLayout(qc_w)
        qc_v.setContentsMargins(0, 0, 0, 0); qc_v.setSpacing(0)
        v_split2 = QSplitter(Qt.Vertical)
        v_split2.setStyleSheet(f"QSplitter::handle {{ background:{BORDER}; height:2px; }}")
        v_split2.addWidget(self._build_qc_table())
        v_split2.addWidget(self._build_qc_detail())
        v_split2.setSizes([580, 180])
        qc_v.addWidget(v_split2)
        tabs.addTab(qc_w, "  Quality Compounder  ")

        return tabs

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
            elif key in ("Gross Margin%", "Op Margin%", "ROIC%", "D/E",
                         "FCF/sh", "Current Ratio", "EV/EBITDA", "P/E", "1Y%"):
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
        PCT_GREEN_RED = {"Gross Margin%", "Op Margin%", "ROIC%", "1Y%"}

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
                    cell = NumItem(val, "price")
                elif key == "MCap ($B)":
                    cell = NumItem(val, "mcap")
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
        PCT_GREEN_RED = {"Gross Margin%", "Op Margin%", "ROIC%", "1Y%"}
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

    def _on_tab_changed(self, _idx: int):
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
        h.addWidget(self._dot); h.addWidget(self.status_lbl); h.addStretch()
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
        # ── Chart panel ──
        self._chart_panel.set_theme(dark)
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

    def _spinbox_style(self):
        return f"""
            QSpinBox {{
                background:{INPUT_BG}; color:{TEXT1};
                border:1px solid {BORDER}; border-radius:3px;
                padding:0 4px; font-size:11px;
                font-family:'Segoe UI',sans-serif;
            }}
            QSpinBox:focus {{ border-color:{BLUE}; }}
            QSpinBox::up-button, QSpinBox::down-button {{
                background:{SURFACE}; border:none; width:16px;
            }}
            QSpinBox::up-arrow  {{ color:{TEXT2}; }}
            QSpinBox::down-arrow {{ color:{TEXT2}; }}
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
        self._df = df
        self._btn_scan.setText("▶  SCAN")
        self._btn_export.setEnabled(True)
        self._populate_table(df)
        self._populate_qc_table(df)
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

                elif key == "Price ($)":
                    cell = NumItem(val, "price")
                elif key == "MCap ($B)":
                    cell = NumItem(val, "mcap")
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
        self._apply_filter()

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
        if hasattr(self, "_tabs_w") and self._tabs_w.currentIndex() == 1:
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
        try:
            df_export = self._df.copy()
            qc_df = pd.DataFrame(
                [compute_qc_score(r.to_dict()) for _, r in df_export.iterrows()])
            df_export = pd.concat(
                [df_export.reset_index(drop=True), qc_df.reset_index(drop=True)], axis=1)
            write_excel(df_export, path, market, top)
            self._set_status(f"✓  Exported → {path}", GREEN)
        except Exception as e:
            self._set_status(f"✕  Export failed: {e}", RED)


# ─────────────────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    win = MainWindow()
    win.showMaximized()
    sys.exit(app.exec())
