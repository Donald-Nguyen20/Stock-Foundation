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
    QDialog, QTextBrowser,
)
from PySide6.QtCore import Qt, QThread, Signal, QTimer, QSortFilterProxyModel
from PySide6.QtGui import QFont, QColor, QKeySequence, QShortcut, QPixmap

# ─────────────────────────────────────────────────────────────────────────────
# Load python screen_top100.py via importlib (filename has a space)
# ─────────────────────────────────────────────────────────────────────────────
_dir  = os.path.dirname(os.path.abspath(__file__))
_spec = importlib.util.spec_from_file_location(
    "screen_top100", os.path.join(_dir, "python screen_top100.py"))
_m = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(_m)

fetch_data          = _m.fetch_data
clean_df            = _m.clean_df
score_canslim       = _m.score_canslim
fetch_moat_yfinance = _m.fetch_moat_yfinance
write_excel         = _m.write_excel
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

# Signal badge colors  (text, background)
SIGNAL_COLORS = {
    "🟢 STRONG BUY": ("#1A5C2B", "#C6EFCE"),
    "🔵 BUY":        ("#1B3A5C", "#DDEEFF"),
    "🟡 WATCH":      ("#7D6608", "#FFF2CC"),
    "🔴 SKIP":       ("#9C0006", "#FFC7CE"),
}

# Table columns: (header, data-key, width, align)
TABLE_COLS = [
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
    ("Score",       "CS_Score",     52,  Qt.AlignCenter | Qt.AlignVCenter),
    ("Signal",      "CS_Signal",   118,  Qt.AlignCenter | Qt.AlignVCenter),
]


from fundamental_chart import ChartWorkerPng


# ─────────────────────────────────────────────────────────────────────────────
# Chart panel — Plotly rendered to PNG, displayed as a native QLabel pixmap
# ─────────────────────────────────────────────────────────────────────────────
class ChartPanel(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setStyleSheet(f"background:{BG};")
        self._pix_q  = None   # QPixmap (quarterly)
        self._pix_a  = None   # QPixmap (annual)
        self._mode   = "quarterly"
        self._worker = None
        self._ticker = ""
        self._spin_idx = 0
        self._spin_msg = ""
        self._spin_tmr = QTimer(self)
        self._spin_tmr.timeout.connect(self._tick)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)
        layout.addWidget(self._build_bar())
        layout.addWidget(self._sep())

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

        self._spin_msg = f"Loading {ticker}…"
        self._spin_tmr.start(80)
        self._worker = ChartWorkerPng(ticker)
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
                    proxy, score = fetch_moat_yfinance(ticker, sector)
                    moat_cache[ticker] = (proxy, score)
                    self.ticker_update.emit(i, total, ticker, score)
                    time.sleep(0.3)

            self.progress.emit("Applying CAN SLIM scoring…")
            df = score_canslim(df, moat_cache)
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
            proxy, score = fetch_moat_yfinance(self.ticker,
                           df["Sector"].iloc[0] if "Sector" in df.columns else "")
            moat_cache = {self.ticker: (proxy, score)}
            df = score_canslim(df, moat_cache)
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
        }
        canslim_rows = ""
        for i, key in enumerate(CS_KEYS):
            cfg = CANSLIM[key]
            op  = ops[cfg["op"]]
            thr = f"{op}&nbsp;{cfg['thr']}{'%' if cfg['thr'] != 2.0 else ''}"
            name = cfg["label"].split("—")[1].strip() if "—" in cfg["label"] else cfg["label"]
            bg  = "#0E1220" if i % 2 == 0 else "#0B0E18"
            canslim_rows += f"""
              <tr style="background-color:{bg};">
                <td style="color:#3D8EF0;font-weight:700;font-family:Consolas,monospace;
                           font-size:14px;padding:8px 12px;width:28px;">{key}</td>
                <td style="color:#DCE4EE;padding:8px 12px;width:110px;">{name}</td>
                <td style="color:#B8C4D0;font-family:Consolas,monospace;font-size:10px;
                           padding:8px 12px;width:105px;">{cfg['field']}</td>
                <td style="color:#34C472;font-weight:600;padding:8px 12px;width:60px;">{thr}</td>
                <td style="color:#DCE4EE;font-size:10px;padding:8px 12px;">{full_desc[key]}</td>
              </tr>"""

        signal_rows = "".join([
            f"""<tr style="background-color:{bg};">
                  <td style="color:{fg};font-weight:700;padding:9px 14px;width:130px;">{icon}&nbsp;{label}</td>
                  <td style="color:{fg};font-weight:700;padding:9px 14px;width:80px;">{score}</td>
                  <td style="color:{fg};padding:9px 14px;">{desc}</td>
                </tr>"""
            for icon, label, score, fg, bg, desc in [
                ("🟢", "STRONG BUY", "7 – 8 / 8", "#1A5C2B", "#C6EFCE",
                 "Đạt ≥7 tiêu chí. Nền tảng cơ bản + kỹ thuật đều mạnh. Ưu tiên theo dõi và cân nhắc mua."),
                ("🔵", "BUY",        "5 – 6 / 8", "#1B3A5C", "#DDEEFF",
                 "Đạt 5–6 tiêu chí. Nền tảng tốt, đáng xem xét nhưng kiểm tra thêm tiêu chí còn thiếu."),
                ("🟡", "WATCH",      "3 – 4 / 8", "#7D6608", "#FFF2CC",
                 "Đạt 3–4 tiêu chí. Tiềm năng nhưng chưa đủ điều kiện. Đưa vào watchlist, chờ cải thiện."),
                ("🔴", "SKIP",       "0 – 2 / 8", "#9C0006", "#FFC7CE",
                 "Đạt ≤2 tiêu chí. Không đủ điều kiện theo CAN SLIM. Bỏ qua hoặc chờ fundamental xoay chiều."),
            ]
        ])

        moat_rows = "".join([
            f"""<tr style="background-color:{bg};">
                  <td style="color:{fg};font-weight:700;padding:9px 14px;width:150px;">{label}</td>
                  <td style="color:{fg};padding:9px 14px;width:220px;">{cond}</td>
                  <td style="color:{fg};padding:9px 14px;">{desc}</td>
                </tr>"""
            for label, cond, fg, bg, desc in [
                ("🏰 WIDE  ★★★",    "ROE 5yr ≥20%  &amp;  GM 5yr ≥50%",          "#1A5C2B", "#C6EFCE",
                 "Lợi thế cạnh tranh bền vững rộng. Rất khó bị xói mòn trong 10+ năm. Ví dụ: Apple, NVIDIA, Visa."),
                ("🏰 NARROW ★★",   "ROE ≥15%  &amp;  GM ≥35%  (hoặc Utilities/RE)", "#1A3A5C", "#DDEEFF",
                 "Lợi thế hẹp hơn, duy trì được 5–10 năm. Cần giám sát cạnh tranh định kỳ."),
                ("🏰 UNCERTAIN ★", "ROE ≥10%  (GM bất kỳ)",                        "#7D6608", "#FFF2CC",
                 "Chưa rõ có moat bền vững. Cần phân tích định tính thêm: sản phẩm, thị trường, ban lãnh đạo."),
                ("🏰 WEAK",        "ROE &lt;10%",                                  "#9C0006", "#FFC7CE",
                 "Không có dấu hiệu lợi thế cạnh tranh. Dễ bị cạnh tranh hoặc biên lợi nhuận xói mòn."),
            ]
        ])

        proxy_sector_rows = ""
        for i, (sector, proxy) in enumerate(MOAT_PROXY_MAP.items()):
            bg = "#0E1220" if i % 2 == 0 else "#0B0E18"
            proxy_sector_rows += (
                f'<tr style="background-color:{bg};">'
                f'<td style="color:#DCE4EE;padding:6px 12px;">{sector}</td>'
                f'<td style="color:#DCE4EE;font-size:10px;padding:6px 12px;">{proxy}</td>'
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
                f'<td style="color:#3D8EF0;font-weight:600;padding:7px 12px;">{ptype}</td>'
                f'<td style="color:#DCE4EE;font-size:10px;padding:7px 12px;">{pdesc}</td>'
                f'</tr>'
            )

        return f"""<html><head><style>
  body  {{ background-color:#0B0E18; color:#DCE4EE;
           font-family:'Segoe UI',sans-serif; font-size:11px;
           margin:0; padding:0; }}
  h1    {{ color:#3D8EF0; font-size:16px; letter-spacing:4px;
           font-weight:700; margin:0 0 3px 0; }}
  .sub  {{ color:#7A8899; font-size:9px; letter-spacing:2px;
           margin:0 0 18px 0; }}
  h2    {{ color:#B8C4D0; font-size:9px; letter-spacing:3px;
           font-weight:700; margin:22px 0 7px 0;
           border-bottom:1px solid #1A2133; padding-bottom:4px; }}
  table {{ border-collapse:collapse; width:100%; margin-bottom:8px; }}
  th    {{ background-color:#0E1220; color:#B8C4D0; text-align:left;
           padding:6px 12px; font-size:9px; letter-spacing:1px;
           font-weight:600; border-bottom:1px solid #1A2133; }}
  .tip  {{ color:#7A8899; font-size:9px; font-style:italic;
           margin:4px 0 0 0; }}
  li    {{ margin:5px 0; color:#DCE4EE; font-size:10px; }}
  b     {{ color:#FFFFFF; font-weight:600; }}
</style></head><body>
<h1>FUNDAMENTAL SCREENER</h1>
<p class="sub">HƯỚNG DẪN ĐỌC TÍN HIỆU  ·  CAN SLIM · MOAT · SIGNAL</p>

<h2>① CAN SLIM — 8 TIÊU CHÍ LỌC CỔ PHIẾU</h2>
<table>
  <tr>
    <th style="width:28px;">Key</th>
    <th style="width:110px;">Tên</th>
    <th style="width:105px;">Chỉ số</th>
    <th style="width:60px;">Ngưỡng</th>
    <th>Ý nghĩa</th>
  </tr>
  {canslim_rows}
</table>
<p class="tip">⚠  Dữ liệu từ TradingView (TTM/FY). EPS &amp; Revenue là YoY growth %. ROE &amp; Gross Margin là trailing twelve months.</p>

<h2>② SIGNAL — KẾT QUẢ TỔNG HỢP CAN SLIM</h2>
<table>
  <tr>
    <th style="width:130px;">Tín hiệu</th>
    <th style="width:80px;">Điểm</th>
    <th>Ý nghĩa &amp; Hành động gợi ý</th>
  </tr>
  {signal_rows}
</table>
<p class="tip">Score = số tiêu chí đạt được (tối đa {N_CS}). ✓ = đạt (+1 điểm) · ✗ = không đạt · — = thiếu dữ liệu (không tính vào score).</p>

<h2>③ MOAT SCORE — LỢI THẾ CẠNH TRANH (Economic Moat)</h2>
<table>
  <tr>
    <th style="width:150px;">Moat</th>
    <th style="width:220px;">Điều kiện (5yr avg từ yfinance)</th>
    <th>Ý nghĩa</th>
  </tr>
  {moat_rows}
</table>
<p class="tip">Moat Score dựa trên ROE 5yr avg + GM 5yr avg từ yfinance (hoặc TTM từ TradingView nếu tắt checkbox).
Khi tắt "yfinance Moat", dùng TV TTM ROE &amp; GM thay thế (nhanh hơn nhưng kém chính xác hơn).</p>

<h2>④ MOAT PROXY — LOẠI LỢI THẾ CẠNH TRANH</h2>
<p style="color:#DCE4EE;font-size:10px;margin:0 0 8px 0;">Moat Proxy được gán tự động theo <b>ngành (Sector)</b> của cổ phiếu.</p>
<table>
  <tr>
    <th style="width:190px;">Sector</th>
    <th>Moat Proxy</th>
  </tr>
  {proxy_sector_rows}
</table>

<p style="color:#B8C4D0;font-size:10px;margin:14px 0 7px 0;font-weight:700;letter-spacing:1px;">GIẢI THÍCH CÁC LOẠI LỢI THẾ</p>
<table>
  <tr>
    <th style="width:210px;">Loại lợi thế</th>
    <th>Ý nghĩa &amp; Ví dụ</th>
  </tr>
  {proxy_type_rows}
</table>
<p class="tip">Moat Proxy chỉ là ước tính định tính theo ngành — không thay thế phân tích sâu từng công ty.
Hai công ty cùng ngành có thể có moat type khác nhau hoàn toàn.</p>

<h2>⑤ CÁCH SỬ DỤNG</h2>
<ul>
  <li><b>SCAN</b> — Lấy top N cổ phiếu theo Market Cap từ TradingView, tự động tính CAN SLIM score cho tất cả.</li>
  <li><b>yfinance Moat</b> — Bật để lấy ROE/GM 5 năm từ yfinance cho Moat chính xác hơn (chậm hơn ~0.3 giây/cổ phiếu).</li>
  <li><b>TICKER lookup</b> — Nhập mã (ví dụ: NVDA) rồi Enter hoặc nhấn LOOKUP để tra cứu ngay mà không cần scan toàn bộ.</li>
  <li><b>FILTER</b> — Gõ vào ô filter để lọc bảng theo ticker hoặc tên công ty theo thời gian thực.</li>
  <li><b>Click vào hàng</b> — Xem CAN SLIM breakdown chi tiết (✓/✗ từng tiêu chí) ở panel phía dưới.</li>
  <li><b>Sort</b> — Click vào tiêu đề cột để sắp xếp. Các cột số sắp xếp đúng theo giá trị (không phải theo chữ).</li>
  <li><b>Export</b> — Xuất Excel đầy đủ với màu sắc, 3 sheet: Data · CAN SLIM Legend · Summary.</li>
  <li><b>F1</b> — Mở màn hình hướng dẫn này bất cứ lúc nào.</li>
</ul>
</body></html>"""


# ─────────────────────────────────────────────────────────────────────────────
# Main window
# ─────────────────────────────────────────────────────────────────────────────
class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Fundamental Screener")
        self.resize(1700, 960)
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
            QHeaderView::section {{
                background:{SURFACE}; color:{TEXT2}; font-size:9px;
                font-weight:600; letter-spacing:1px;
                border:none; border-right:1px solid {BORDER};
                border-bottom:1px solid {BORDER};
                padding:4px 6px;
            }}
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

        # Left: table (top) + detail card (bottom)
        v_split = QSplitter(Qt.Vertical)
        v_split.setStyleSheet(
            f"QSplitter::handle {{ background:{BORDER}; height:2px; }}")
        v_split.addWidget(self._build_table())
        v_split.addWidget(self._build_detail())
        v_split.setSizes([580, 180])

        # Right: fundamental chart panel
        self._chart_panel = ChartPanel()

        # Horizontal splitter: table side | chart side
        h_split = QSplitter(Qt.Horizontal)
        h_split.setStyleSheet(
            f"QSplitter::handle {{ background:{BORDER}; width:2px; }}")
        h_split.addWidget(v_split)
        h_split.addWidget(self._chart_panel)
        h_split.setSizes([920, 580])
        vbox.addWidget(h_split, stretch=1)

        vbox.addWidget(self._sep())
        vbox.addWidget(self._build_status())

    # ── Section builders ──────────────────────────────────────────────────────

    def _build_header(self):
        w = QWidget(); w.setFixedHeight(50)
        w.setStyleSheet(f"background:{SURFACE};")
        h = QHBoxLayout(w); h.setContentsMargins(22, 0, 16, 0); h.setSpacing(0)
        lbl = QLabel()
        lbl.setTextFormat(Qt.RichText)
        lbl.setText(
            f'<span style="color:{TEXT1};font-size:14px;font-weight:700;letter-spacing:4px">FUNDAMENTAL</span>'
            f'<span style="color:{BLUE};font-size:14px;font-weight:700;letter-spacing:4px"> SCREENER</span>'
        )
        h.addWidget(lbl); h.addStretch()
        tag = QLabel("CAN SLIM · TradingView · yfinance")
        tag.setStyleSheet(f"color:{TEXT3}; font-size:9px; letter-spacing:2px;")
        h.addWidget(tag)
        h.addSpacing(16)
        btn_help = QPushButton("?")
        btn_help.setFixedSize(28, 28)
        btn_help.setToolTip("Help  (F1)")
        btn_help.setCursor(Qt.PointingHandCursor)
        btn_help.setStyleSheet(f"""
            QPushButton {{
                background:{BORDER2}; color:{TEXT2};
                border:none; border-radius:14px;
                font-size:13px; font-weight:700;
                font-family:'Segoe UI',sans-serif;
            }}
            QPushButton:hover {{ background:{BLUE}; color:#FFFFFF; }}
        """)
        btn_help.clicked.connect(self._show_help)
        h.addWidget(btn_help)
        return w

    def _build_controls(self):
        w = QWidget(); w.setFixedHeight(86)
        w.setStyleSheet(f"background:{BG};")
        v = QVBoxLayout(w); v.setContentsMargins(22, 8, 22, 8); v.setSpacing(6)

        # Row 1: scan controls
        row1 = QHBoxLayout(); row1.setSpacing(10)

        lbl_market = QLabel("MARKET")
        lbl_market.setStyleSheet(
            f"color:{TEXT3}; font-size:9px; font-weight:600; letter-spacing:1.5px;")
        self._market = QComboBox()
        self._market.addItems(["america", "nasdaq", "nyse", "euronext", "hong_kong", "vietnam"])
        self._market.setFixedHeight(32)
        self._market.setStyleSheet(self._combo_style())

        lbl_top = QLabel("TOP")
        lbl_top.setStyleSheet(
            f"color:{TEXT3}; font-size:9px; font-weight:600; letter-spacing:1.5px;")
        self._top_spin = QSpinBox()
        self._top_spin.setRange(50, 1000); self._top_spin.setValue(300); self._top_spin.setSingleStep(50)
        self._top_spin.setFixedSize(80, 32)
        self._top_spin.setStyleSheet(self._spinbox_style())

        self._use_yf = QCheckBox("yfinance Moat")
        self._use_yf.setChecked(False)
        self._use_yf.setStyleSheet(
            f"color:{TEXT2}; font-size:10px; font-family:'Segoe UI',sans-serif;")

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
        lbl_tk.setStyleSheet(
            f"color:{TEXT3}; font-size:9px; font-weight:600; letter-spacing:1.5px;")
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
        lbl_filter.setStyleSheet(
            f"color:{TEXT3}; font-size:9px; font-weight:600; letter-spacing:1.5px;")
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
        self._table.horizontalHeader().setHighlightSections(False)

        headers = [c[0] for c in TABLE_COLS]
        self._table.setHorizontalHeaderLabels(headers)
        for i, (_, _, w_px, _) in enumerate(TABLE_COLS):
            self._table.setColumnWidth(i, w_px)

        self._table.itemSelectionChanged.connect(self._on_row_selected)
        v.addWidget(self._table)
        return w

    def _build_detail(self):
        self._detail = DetailCard()
        return self._detail

    def _build_status(self):
        w = QWidget(); w.setFixedHeight(28)
        w.setStyleSheet(f"background:{SURFACE};")
        h = QHBoxLayout(w); h.setContentsMargins(22, 0, 22, 0)
        self._dot = QLabel("●"); self._dot.setFixedWidth(14)
        self._dot.setStyleSheet(f"color:{TEXT3}; font-size:7px;")
        self.status_lbl = QLabel("Ready  ·  Enter market & top N, then click SCAN")
        self.status_lbl.setStyleSheet(
            f"color:{TEXT3}; font-size:10px; font-family:'Segoe UI',sans-serif;")
        h.addWidget(self._dot); h.addWidget(self.status_lbl); h.addStretch()
        ver = QLabel("TradingView  ·  yfinance  ·  PySide6")
        ver.setStyleSheet(f"color:{TEXT3}; font-size:9px; letter-spacing:1px;")
        h.addWidget(ver)
        return w

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

                if key == "Ticker":
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

    def _apply_filter(self):
        text = self._filter_inp.text().strip().lower()
        visible = 0
        total = self._table.rowCount()
        for ri in range(total):
            show = True
            if text:
                t_item = self._table.item(ri, 0)   # Ticker col
                n_item = self._table.item(ri, 1)   # Name col
                t_txt  = (t_item.text() if t_item else "").lower()
                n_txt  = (n_item.text() if n_item else "").lower()
                show   = (text in t_txt) or (text in n_txt)
            self._table.setRowHidden(ri, not show)
            if show: visible += 1
        self._lbl_count.setText(f"{visible} / {total}  rows")

    # ── Row selection → detail card ───────────────────────────────────────────

    def _on_row_selected(self):
        rows = self._table.selectedItems()
        if not rows or self._df is None: return
        ri = self._table.currentRow()
        ticker_item = self._table.item(ri, 0)
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
                    item = self._table.item(ri, 0)
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
        path, _ = QFileDialog.getSaveFileName(
            self, "Export Excel", "", "Excel (*.xlsx)")
        if not path: return
        if not path.endswith(".xlsx"): path += ".xlsx"
        try:
            market = self._market.currentText()
            top    = self._top_spin.value()
            write_excel(self._df, path, market, top)
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
