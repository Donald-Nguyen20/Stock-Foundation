"""
TV Top 100 MCap Fundamental Screener — xuất Excel + CAN SLIM Score
===================================================================
pip install tradingview-screener pandas openpyxl yfinance
python tv_top100_screener.py
python tv_top100_screener.py --market nasdaq --top 200
"""

import argparse, sys, time, warnings
from datetime import datetime
import numpy as np
warnings.filterwarnings("ignore")

try:
    from tradingview_screener import Query, Column
    import pandas as pd
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
    from openpyxl.formatting.rule import ColorScaleRule
except ImportError as e:
    print(f"❌ Thiếu thư viện: {e}\n   pip install tradingview-screener pandas openpyxl yfinance")
    sys.exit(1)


# ═══════════════════════════════════════════════════════════════
# CAN SLIM CRITERIA
# ═══════════════════════════════════════════════════════════════
CANSLIM = {
    "C":   dict(label="C — EPS Qtr",   field="EPS Qtr%",      thr=25,  op="gt", desc="EPS Quarterly YoY > 25%"),
    "A":   dict(label="A — EPS Ann",   field="EPS Annual%",   thr=20,  op="gt", desc="EPS Annual YoY > 20%"),
    "S":   dict(label="S — Sales",     field="Rev Qtr%",      thr=20,  op="gt", desc="Revenue Quarterly YoY > 20%"),
    "L":   dict(label="L — RS (1Y)",   field="1Y%",           thr=20,  op="gt", desc="1-Year Perf > 20% (RS proxy)"),
    "Q":   dict(label="Quality GM",    field="Gross Margin%", thr=40,  op="gt", desc="Gross Margin > 40%"),
    "R":   dict(label="ROE",           field="ROE%",          thr=17,  op="gt", desc="ROE > 17%"),
    "M":   dict(label="Momentum 3M",   field="3M%",           thr=0,   op="gt", desc="3-Month Perf > 0%"),
    "D":   dict(label="Debt OK",       field="D/E",           thr=2.0, op="lt", desc="D/E < 2.0"),
    "N":   dict(label="N — 52W High",  field="52W_High%",     thr=90,  op="gt", desc="Price ≥ 90% of 52-Week High"),
    "MKT": dict(label="Market",        field="Market_OK",     thr=0.5, op="gt", desc="S&P 500 above MA50 & MA200"),
}
CS_KEYS = ["C", "A", "S", "L", "Q", "R", "M", "D", "N", "MKT"]
N_CS    = len(CS_KEYS)

FETCH_COLS = [
    "name", "description", "close", "volume", "market_cap_basic",
    "sector",
    "earnings_per_share_diluted_yoy_growth_fy",
    "earnings_per_share_diluted_yoy_growth_fq",
    "total_revenue_yoy_growth_fy",
    "total_revenue_yoy_growth_fq",
    "gross_margin", "net_margin", "return_on_equity", "debt_to_equity",
    "price_earnings_ttm", "price_sales_current",
    "relative_volume_10d_calc",
    "Perf.W", "Perf.1M", "Perf.3M", "Perf.6M", "Perf.Y",
    # Quality Compounder fields
    "return_on_invested_capital", "operating_margin",
    "enterprise_value_ebitda_ttm", "free_cash_flow_per_share_ttm",
    "current_ratio",
]

RENAME = {
    "name": "Ticker", "description": "Tên Công Ty", "close": "Price ($)", "market_cap_basic": "MCap ($B)",
    "sector": "Sector",
    "earnings_per_share_diluted_yoy_growth_fy": "EPS Annual%",
    "earnings_per_share_diluted_yoy_growth_fq": "EPS Qtr%",
    "total_revenue_yoy_growth_fy": "Rev Annual%",
    "total_revenue_yoy_growth_fq": "Rev Qtr%",
    "gross_margin": "Gross Margin%", "net_margin": "Net Margin%",
    "return_on_equity": "ROE%", "debt_to_equity": "D/E",
    "price_earnings_ttm": "P/E", "price_sales_current": "P/S",
    "relative_volume_10d_calc": "Rel Vol",
    "Perf.W": "1W%", "Perf.1M": "1M%", "Perf.3M": "3M%",
    "Perf.6M": "6M%", "Perf.Y": "1Y%",
    "return_on_invested_capital": "ROIC%",
    "operating_margin": "Op Margin%",
    "enterprise_value_ebitda_ttm": "EV/EBITDA",
    "free_cash_flow_per_share_ttm": "FCF/sh",
    "current_ratio": "Current Ratio",
}

DATA_HEADERS = [
    "Ticker", "Tên Công Ty", "Sector", "Price ($)", "MCap ($B)",
    "Moat Proxy", "Moat Score",
    "EPS Annual%", "EPS Qtr%", "Rev Annual%", "Rev Qtr%",
    "Gross Margin%", "Net Margin%", "ROE%", "D/E",
    "P/E", "P/S", "Rel Vol",
    "1W%", "1M%", "3M%", "6M%", "1Y%",
]
# 1=Ticker,2=TênCT,3=Sector,4=Price,5=MCap
# 6=MoatProxy,7=MoatScore
# 8=EPS_A,9=EPS_Q,10=Rev_A,11=Rev_Q
# 12=GM,13=NM,14=ROE,15=DE,16=PE,17=PS,18=RVol
# 19=1W,20=1M,21=3M,22=6M,23=1Y
PCT_COLS = {8,9,10,11,12,13,14,19,20,21,22,23}


# ═══════════════════════════════════════════════════════════════
# MOAT — lấy từ yfinance (5yr ROE avg + 5yr GM avg)
# Logic giống moat_score() trong screen_accumulate.py
# ═══════════════════════════════════════════════════════════════
MOAT_PROXY_MAP = {
    "Technology":            "Switching Cost / Network Effect",
    "Healthcare":            "Intangible Assets (Patents)",
    "Consumer Defensive":    "Brand / Distribution Network",
    "Financial Services":    "Network Effect / Scale",
    "Industrials":           "Cost Advantage / Scale",
    "Communication Services":"Network Effect",
    "Consumer Cyclical":     "Brand / Cost Advantage",
    "Basic Materials":       "Cost Advantage / Efficient Scale",
    "Energy":                "Cost Advantage",
    "Utilities":             "Efficient Scale (Regulated)",
    "Real Estate":           "Efficient Scale / Location",
}

MOAT_SCORE_STYLE = {
    "WIDE  ★★★": ("1A5C2B", "C6EFCE"),
    "NARROW ★★":  ("1A3A5C", "DDEEFF"),
    "UNCERTAIN ★": ("7D6608", "FFF2CC"),
    "WEAK":        ("9C0006", "FFC7CE"),
}

QC_SIGNAL_STYLE = {
    "🏆 COMPOUNDER": ("1A5C2B", "C6EFCE"),
    "⭐ QUALITY":     ("1A3A5C", "DDEEFF"),
    "○ AVERAGE":     ("7D6608", "FFF2CC"),
    "✗ WEAK":        ("9C0006", "FFC7CE"),
}

EQ_BADGE_STYLE = {
    "💚 Cash Backed":   ("1A5C2B", "C6EFCE"),
    "🟡 Mixed":         ("7D6608", "FFF2CC"),
    "🔴 Accrual Heavy": ("9C0006", "FFC7CE"),
}

def _sv(s, i=0):
    try:
        v = s.iloc[i] if hasattr(s, "iloc") else s
        return float(v) if pd.notna(v) else None
    except: return None

def _get_row(df_fin, *keys):
    for k in keys:
        if k in df_fin.index:
            s = df_fin.loc[k]
            return s if hasattr(s, "iloc") else pd.Series([s])
    return pd.Series(dtype=float)

def _safe_mean(lst):
    v = [x for x in lst if x is not None]
    return round(float(np.mean(v)), 1) if v else None

def _moat_score_from_metrics(roe_avg, gm_avg, sector=""):
    """Copy từ moat_score() trong screen_accumulate.py"""
    r = roe_avg or 0;  g = gm_avg or 0
    if sector in ("Utilities", "Real Estate"):
        return "NARROW ★★" if r >= 10 and g >= 20 else "UNCERTAIN ★"
    if   r >= 20 and g >= 50: return "WIDE  ★★★"
    elif r >= 15 and g >= 35: return "NARROW ★★"
    elif r >= 10:             return "UNCERTAIN ★"
    return "WEAK"

def fetch_moat_yfinance(ticker: str, sector: str = "") -> tuple:
    """
    Lấy 5yr ROE avg + 5yr GM avg từ yfinance.
    Trả về (moat_proxy, moat_score).
    Giống fetch() trong screen_accumulate.py.
    """
    try:
        import yfinance as yf
        tk   = yf.Ticker(ticker)
        info = tk.info or {}

        # Sector từ yfinance (chính xác hơn TV)
        yf_sector = info.get("sector", "") or sector
        proxy = MOAT_PROXY_MAP.get(yf_sector, MOAT_PROXY_MAP.get(sector, "Cost Advantage"))

        # TTM fallback
        roe_ttm = (info.get("returnOnEquity") or 0) * 100 or None
        gm_ttm  = (info.get("grossMargins")   or 0) * 100 or None

        # 5yr ROE từ income + balance sheet
        roe_avg = None
        try:
            inc = tk.income_stmt.sort_index(axis=1, ascending=False)
            bal = tk.balance_sheet.sort_index(axis=1, ascending=False)
            ni  = _get_row(inc, "Net Income", "NetIncome")
            eq  = _get_row(bal, "Stockholders Equity", "StockholdersEquity",
                           "Total Stockholders Equity", "Common Stock Equity")
            roe_list = []
            for i in range(min(5, min(len(ni), len(eq)))):
                n = _sv(ni, i); e = _sv(eq, i)
                roe_list.append(round(n/e*100,1) if n and e and e != 0 else None)
            roe_avg = _safe_mean(roe_list)
        except: pass

        # 5yr GM từ income stmt
        gm_avg = None
        try:
            inc2 = tk.income_stmt.sort_index(axis=1, ascending=False)
            gp   = _get_row(inc2, "Gross Profit", "GrossProfit")
            rev  = _get_row(inc2, "Total Revenue", "TotalRevenue")
            gm_list = []
            for i in range(min(5, min(len(gp), len(rev)))):
                g = _sv(gp, i); r = _sv(rev, i)
                gm_list.append(round(g/r*100,1) if g and r and r != 0 else None)
            gm_avg = _safe_mean(gm_list)
        except: pass

        # Dùng 5yr avg nếu có, fallback TTM
        score = _moat_score_from_metrics(
            roe_avg or roe_ttm,
            gm_avg  or gm_ttm,
            yf_sector
        )

        # 52-week high proximity
        w52_high  = info.get("fiftyTwoWeekHigh")
        cur_price = (info.get("currentPrice") or info.get("regularMarketPrice")
                     or info.get("previousClose"))
        w52_pct = (round(cur_price / w52_high * 100, 1)
                   if w52_high and cur_price and w52_high > 0 else None)

        # Earnings quality — FCF / Net Income ratio (TTM)
        eq_badge = None
        try:
            cf   = tk.cashflow.sort_index(axis=1, ascending=False)
            inc3 = tk.income_stmt.sort_index(axis=1, ascending=False)
            fcf_row = _get_row(cf,   "Free Cash Flow",     "FreeCashFlow")
            ni_row  = _get_row(inc3, "Net Income",          "NetIncome")
            if fcf_row.empty:
                ocf = _get_row(cf, "Operating Cash Flow",  "OperatingCashFlow")
                cex = _get_row(cf, "Capital Expenditure",  "CapitalExpenditure")
                ocf_v = _sv(ocf, 0); cex_v = _sv(cex, 0)
                fcf_v = (ocf_v + (cex_v or 0)) if ocf_v is not None else None
            else:
                fcf_v = _sv(fcf_row, 0)
            ni_v = _sv(ni_row, 0)
            if fcf_v is not None and ni_v is not None and ni_v != 0:
                ratio = fcf_v / ni_v
                eq_badge = ("💚 Cash Backed"  if ratio >= 0.8
                            else "🟡 Mixed"       if ratio >= 0.3
                            else "🔴 Accrual Heavy")
        except:
            pass

        return proxy, score, w52_pct, eq_badge

    except Exception as e:
        # Fallback: sector hardcode + UNCERTAIN
        proxy = MOAT_PROXY_MAP.get(sector, "Cost Advantage")
        return proxy, "UNCERTAIN ★", None, None


def build_moat_cache(tickers: list, sectors: dict) -> dict:
    """
    Fetch moat cho toàn bộ danh sách tickers từ yfinance.
    Trả về dict: {ticker: (proxy, score)}
    """
    cache = {}
    total = len(tickers)
    print(f"\n  🏰 Fetch Moat từ yfinance ({total} tickers)...")
    for i, ticker in enumerate(tickers, 1):
        sector = sectors.get(ticker, "")
        print(f"  [{i:>3}/{total}] {ticker:<8}", end="", flush=True)
        proxy, score, w52_pct, eq_badge = fetch_moat_yfinance(ticker, sector)
        cache[ticker] = (proxy, score, w52_pct, eq_badge)
        eq_str = f" [{eq_badge}]" if eq_badge else ""
        print(f" {score}{eq_str}")
        time.sleep(0.4)   # tránh rate limit
    print()
    return cache


def fetch_market_direction():
    """True nếu S&P 500 đang trên MA50 và MA200 (uptrend), False nếu không, None nếu lỗi."""
    try:
        import yfinance as yf
        spx = yf.Ticker("^GSPC").history(period="1y")["Close"]
        if len(spx) < 200:
            return None
        price = float(spx.iloc[-1])
        ma50  = float(spx.rolling(50).mean().iloc[-1])
        ma200 = float(spx.rolling(200).mean().iloc[-1])
        return price > ma50 and price > ma200
    except Exception:
        return None


# ─── Colors ─────────────────────────────────────────────────────────────────
C_NAVY="1B3A5C"; C_DARK="0D2137"; C_WHITE="FFFFFF"; C_ALT="F0F4FA"; C_BRD="C5D5E8"
CG_BG="C6EFCE"; CG_FG="276221"
CR_BG="FFC7CE"; CR_FG="9C0006"
CY_BG="FFEB9C"; CY_FG="9C6500"
CB_BG="DDEEFF"; CB_FG="1B3A5C"

SIGNAL_STYLE = {
    "🟢 STRONG BUY": (CG_FG, CG_BG),
    "🔵 BUY":        (CB_FG, CB_BG),
    "🟡 WATCH":      (CY_FG, CY_BG),
    "🔴 SKIP":       (CR_FG, CR_BG),
}

def F(bold=False, size=9, color="000000", italic=False):
    return Font(name="Calibri", bold=bold, size=size, color=color, italic=italic)
def BG(c): return PatternFill("solid", fgColor=c)
def AL(h="center", v="center", wrap=False): return Alignment(horizontal=h, vertical=v, wrap_text=wrap)
def BD(c=C_BRD):
    s = Side(style="thin", color=c); return Border(left=s, right=s, top=s, bottom=s)
def CL(i): return get_column_letter(i)


# ═══════════════════════════════════════════════════════════════
# SCORING
# ═══════════════════════════════════════════════════════════════
def score_canslim(df, moat_cache: dict = None, market_ok=None):
    # Populate 52W_High% from moat_cache (3rd element of tuple)
    if moat_cache:
        def _get_w52(ticker):
            t = moat_cache.get(ticker)
            return t[2] if t and len(t) >= 3 else None
        df["52W_High%"] = df["Ticker"].apply(_get_w52)
    else:
        df["52W_High%"] = None

    # Market direction — same value for all rows
    df["Market_OK"] = market_ok

    def chk(row, key):
        cfg = CANSLIM[key]
        v = row.get(cfg["field"])
        if v is None or (isinstance(v, float) and pd.isna(v)): return None
        return v > cfg["thr"] if cfg["op"] == "gt" else v < cfg["thr"]

    for key in CS_KEYS:
        df[f"CS_{key}"] = df.apply(lambda r: chk(r, key), axis=1)

    df["CS_Score"] = df[[f"CS_{k}" for k in CS_KEYS]].apply(
        lambda r: sum(1 for v in r if v is True), axis=1)

    def sig(s):
        if s >= round(N_CS * 0.875): return "🟢 STRONG BUY"
        if s >= round(N_CS * 0.625): return "🔵 BUY"
        if s >= round(N_CS * 0.375): return "🟡 WATCH"
        return "🔴 SKIP"
    df["CS_Signal"] = df["CS_Score"].apply(sig)

    # ── MOAT: dùng yfinance cache nếu có, fallback sector ──────
    def apply_moat(row):
        ticker = row.get("Ticker", "")
        sector = row.get("Sector", "")
        if moat_cache and ticker in moat_cache:
            t = moat_cache[ticker]
            return t[0], t[1]   # proxy, score (bỏ qua w52_pct)
        # Fallback: tính từ TV TTM data (giống cũ)
        gm  = row.get("Gross Margin%"); gm  = gm  if (gm  is not None and not (isinstance(gm,  float) and pd.isna(gm)))  else None
        roe = row.get("ROE%");          roe = roe if (roe is not None and not (isinstance(roe, float) and pd.isna(roe))) else None
        proxy = MOAT_PROXY_MAP.get(sector, "Cost Advantage")
        score = _moat_score_from_metrics(roe, gm, sector)
        return proxy, score

    moat = df.apply(apply_moat, axis=1)
    df["Moat Proxy"] = moat.apply(lambda x: x[0])
    df["Moat Score"] = moat.apply(lambda x: x[1])

    # Earnings quality badge (only when moat_cache / yfinance is available)
    if moat_cache:
        def _get_eq(ticker):
            t = moat_cache.get(ticker)
            return t[3] if t and len(t) >= 4 else None
        df["EQ_Badge"] = df["Ticker"].apply(_get_eq)
    else:
        df["EQ_Badge"] = None

    _MW = {"WIDE  ★★★": 1.2, "NARROW ★★": 1.1, "UNCERTAIN ★": 1.0, "WEAK": 0.85}
    df["Conviction"] = df.apply(
        lambda r: round(r["CS_Score"] * _MW.get(r.get("Moat Score", ""), 1.0), 1),
        axis=1)
    return df


# ═══════════════════════════════════════════════════════════════
# FETCH
# ═══════════════════════════════════════════════════════════════
def fetch_data(market, top):
    print(f"  ⏳ Query TradingView ({market.upper()}, top {top} MCap)...", end="", flush=True)
    try:
        _, df = (Query().set_markets(market).select(*FETCH_COLS)
                 .order_by("market_cap_basic", ascending=False).limit(top).get_scanner_data())
        print(f" ✅ {len(df)} mã"); return df
    except Exception as e:
        print(f"\n❌ {e}"); sys.exit(1)

def clean_df(df):
    df = df.rename(columns=RENAME)
    df["MCap ($B)"] = df["MCap ($B)"] / 1e9
    return df


# ═══════════════════════════════════════════════════════════════
# EXCEL — giữ nguyên như cũ
# ═══════════════════════════════════════════════════════════════
def write_excel(df, path, market, top):
    wb = Workbook()
    _main_sheet(wb, df, market, top)
    if "QC_Score" in df.columns:
        _qc_sheet(wb, df)
    _dashboard_sheet(wb, df)
    wb.save(path)
    print(f"  💾 {path}")


def _dashboard_sheet(wb, df):
    from openpyxl.chart import PieChart, BarChart, ScatterChart, Reference, Series
    from openpyxl.chart.series import DataPoint

    ws = wb.create_sheet("Dashboard")
    ws.sheet_view.showGridLines = False
    ws.sheet_tab_color = "1F3864"

    SIGNALS     = ["\U0001f7e2 STRONG BUY", "\U0001f535 BUY", "\U0001f7e1 WATCH", "\U0001f534 SKIP"]
    SIG_COLS    = ["00AA44", "0070C0", "FFC000", "FF0000"]
    SIG_LBLS    = ["STRONG BUY", "BUY", "WATCH", "SKIP"]
    EQ_LBLS     = ["\U0001f49a Cash Backed", "\U0001f7e1 Mixed", "\U0001f534 Accrual Heavy"]
    EQ_PIE_COLS = ["1A5C2B", "FFC000", "FF0000"]
    HAS_QC      = "QC_Score" in df.columns and "QC_Signal" in df.columns
    N_QC_LOCAL  = 6  # ROIC, OpMgn, GM, FCF, D/E, Moat

    # -- Column widths (24 cols, A-X)
    for ci in range(1, 25):
        ws.column_dimensions[CL(ci)].width = 10

    # -- Helpers
    def sec_header(row, text, end_col=24):
        ws.merge_cells(f"A{row}:{CL(end_col)}{row}")
        c = ws.cell(row, 1, text)
        c.font = F(True, 10, "FFFFFF"); c.fill = BG("2C4F7C"); c.alignment = AL()
        ws.row_dimensions[row].height = 18

    def kpi_card(sc, val_text, label, bg_hex, fg_hex="FFFFFF"):
        ec = sc + 3
        ws.merge_cells(f"{CL(sc)}4:{CL(ec)}4")
        c = ws.cell(4, sc, val_text)
        c.font = F(True, 18, fg_hex); c.fill = BG(bg_hex); c.alignment = AL()
        ws.merge_cells(f"{CL(sc)}5:{CL(ec)}5")
        c = ws.cell(5, sc, label)
        c.font = F(False, 9, fg_hex); c.fill = BG(bg_hex); c.alignment = AL()
        ws.merge_cells(f"{CL(sc)}6:{CL(ec)}6")
        ws.cell(6, sc).fill = BG("1B3A5C")

    def kpi_card_qc(sc, val_text, label, bg_hex, fg_hex="FFFFFF"):
        ec = sc + 3
        ws.merge_cells(f"{CL(sc)}9:{CL(ec)}9")
        c = ws.cell(9, sc, val_text)
        c.font = F(True, 18, fg_hex); c.fill = BG(bg_hex); c.alignment = AL()
        ws.merge_cells(f"{CL(sc)}10:{CL(ec)}10")
        c = ws.cell(10, sc, label)
        c.font = F(False, 9, fg_hex); c.fill = BG(bg_hex); c.alignment = AL()
        ws.merge_cells(f"{CL(sc)}11:{CL(ec)}11")
        ws.cell(11, sc).fill = BG("0D2020")

    # -- Title (rows 1-2)
    ws.merge_cells("A1:X2")
    c = ws["A1"]
    c.value = "STOCK SCREENER  ·  DASHBOARD"
    c.font = Font(name="Calibri", bold=True, size=16, color="FFFFFF")
    c.fill = BG("0D2137"); c.alignment = AL()
    ws.row_dimensions[1].height = 36
    ws.row_dimensions[2].height = 4

    # -- [1] CAN SLIM KPI CARDS (rows 3-7)
    sec_header(3, "❶  CAN SLIM  —  KEY METRICS")
    ws.row_dimensions[4].height = 38
    ws.row_dimensions[5].height = 18
    ws.row_dimensions[6].height = 6
    ws.row_dimensions[7].height = 8

    total        = len(df)
    sb_cnt       = int((df["CS_Signal"] == "\U0001f7e2 STRONG BUY").sum())
    b_cnt        = int((df["CS_Signal"] == "\U0001f535 BUY").sum())
    sb_pct       = f"{sb_cnt/total*100:.1f}%" if total else "—"
    b_pct        = f"{b_cnt/total*100:.1f}%"  if total else "—"
    avg_cs_score = f"{df['CS_Score'].mean():.1f}" if "CS_Score" in df.columns else "—"
    avg_1y_ser   = df["1Y%"].dropna()
    avg_1y_val   = f"{avg_1y_ser.mean():.1f}%" if len(avg_1y_ser) else "—"
    avg_roe_ser  = df["ROE%"].dropna()
    avg_roe_val  = f"{avg_roe_ser.mean():.1f}%" if len(avg_roe_ser) else "—"

    kpi_card(1,  str(total),     "Total Stocks",            "1B3A5C")
    kpi_card(5,  sb_pct,         f"Strong Buy  ({sb_cnt})", "00703A")
    kpi_card(9,  b_pct,          f"Buy  ({b_cnt})",         "0070C0")
    kpi_card(13, avg_cs_score,   f"Avg CS Score / {N_CS}",  "5B2C6F")
    kpi_card(17, avg_1y_val,     "Avg 1Y Return",           "7D3C0A")
    kpi_card(21, avg_roe_val,    "Avg ROE%",                "0E6655")

    # -- [2] QUALITY COMPOUNDER KPI CARDS (rows 8-13)
    sec_header(8, "❷  QUALITY COMPOUNDER  —  KEY METRICS")
    ws.row_dimensions[9].height  = 38
    ws.row_dimensions[10].height = 18
    ws.row_dimensions[11].height = 6
    ws.row_dimensions[12].height = 8
    ws.row_dimensions[13].height = 8

    if HAS_QC:
        comp_cnt     = int((df["QC_Signal"] == "\U0001f3c6 COMPOUNDER").sum())
        qual_cnt     = int((df["QC_Signal"] == "⭐ QUALITY").sum())
        avg_qc_v     = df["QC_Score"].mean()
        avg_qc_s     = f"{avg_qc_v:.1f}" if pd.notna(avg_qc_v) else "—"
        roic_ser     = df["ROIC%"].dropna() if "ROIC%" in df.columns else pd.Series(dtype=float)
        avg_roic_val = f"{roic_ser.mean():.1f}%" if len(roic_ser) else "—"
        dual_cnt     = int(((df["CS_Score"] >= 7) & (df["QC_Score"] >= 4)).sum())
        if "EQ_Badge" in df.columns:
            eq_cb_txt = str(int((df["EQ_Badge"] == "\U0001f49a Cash Backed").sum()))
            eq_cb_lbl = "\U0001f49a Cash Backed"
        else:
            eq_cb_txt = "—"; eq_cb_lbl = "EQ Badge (yfinance)"
    else:
        comp_cnt = 0; qual_cnt = 0; avg_qc_s = "—"; avg_roic_val = "—"
        dual_cnt = 0; eq_cb_txt = "—"; eq_cb_lbl = "EQ Badge (yfinance)"

    kpi_card_qc(1,  str(comp_cnt), "# COMPOUNDER",                 "1A5C2B")
    kpi_card_qc(5,  str(qual_cnt), "# QUALITY",                    "1A3A5C")
    kpi_card_qc(9,  avg_roic_val,  "Avg ROIC%",                    "0E6655")
    kpi_card_qc(13, avg_qc_s,      f"Avg QC Score / {N_QC_LOCAL}", "5B2C6F")
    kpi_card_qc(17, str(dual_cnt), "Dual Leaders (CS≧7 & QC≧4)",  "4A235A")
    kpi_card_qc(21, eq_cb_txt,     eq_cb_lbl,                      "155A28")

    # -- [3] CHART ANALYSIS (rows 14-60)
    # Row 1 of charts: Signal Pie (A15) | CS x QC Scatter (I15) | EQ Badge Pie (Q15)
    # Row 2 of charts: Top 20 Bar (A38) | Score x 1Y% Scatter (I38)
    sec_header(14, "❸  CHART ANALYSIS")
    for r in range(15, 61):
        ws.row_dimensions[r].height = 16
    ws.row_dimensions[37].height = 8   # spacer between chart rows
    ws.row_dimensions[60].height = 10  # gap before [4]

    # -- [4] SECTOR BREAKDOWN (rows 61-85)
    sec_header(61, "❹  SECTOR BREAKDOWN")
    for r in range(62, 86):
        ws.row_dimensions[r].height = 17
    ws.row_dimensions[85].height = 10  # gap before [5]

    # -- [5] TOP 10 PERFORMERS (rows 90-101)
    ws.row_dimensions[89].height = 6
    sec_header(90, "❺  TOP 10 PERFORMERS")

    T10_HDRS   = ["#", "Ticker", "Tên Công Ty", "Sector", "Signal",
                  f"Score/{N_CS}", "Conv.", "1Y%", "52W Hi%", "Moat", "Price ($)"]
    T10_WIDTHS = [4, 9, 22, 14, 14, 7, 7, 8, 9, 13, 9]
    for ci, (h, w) in enumerate(zip(T10_HDRS, T10_WIDTHS), 1):
        ws.column_dimensions[CL(ci)].width = w
        c = ws.cell(91, ci, h)
        c.font = F(True, 9, "FFFFFF"); c.fill = BG("1B3A5C")
        c.alignment = AL(); c.border = BD()
    ws.row_dimensions[91].height = 20

    _t10_cols = ["Ticker", "Tên Công Ty", "Sector", "CS_Signal", "CS_Score",
                 "Conviction", "1Y%", "52W_High%", "Moat Score", "Price ($)"]
    _t10_safe = [col for col in _t10_cols if col in df.columns]
    top10 = (df[_t10_safe]
             .sort_values(["Conviction", "1Y%"] if "Conviction" in df.columns
                          else ["CS_Score", "1Y%"], ascending=[False, False])
             .head(10))

    def _fmt_t10_cell(c, ci, val, fg_s, bg_s, fg_m, bg_m):
        c.border = BD(); c.font = F(size=9)
        c.alignment = AL("left" if ci in (3, 4) else "center")
        if ci == 2:
            c.font = F(True, 9, "1B3A5C")
        elif ci == 5:
            c.font = F(True, 9, fg_s); c.fill = BG(bg_s)
        elif ci == 7 and val is not None:
            fv = float(val); c.value = round(fv, 1)
            c.font = F(True, 9, "276221" if fv >= 9 else "0070C0" if fv >= 7 else "7D6608" if fv >= 5 else "9C0006")
        elif ci == 8 and val is not None and not (isinstance(val, float) and pd.isna(val)):
            fv = float(val); c.value = fv / 100; c.number_format = '+0.0%;(0.0%);"-"'
            c.font = F(size=9, color=("276221" if fv > 0 else ("9C0006" if fv < 0 else "000000")))
        elif ci == 9 and val is not None and not (isinstance(val, float) and pd.isna(val)):
            fv = float(val); c.value = fv / 100; c.number_format = '0.0%'
            c.font = F(size=9, color=("276221" if fv >= 90 else ("9C0006" if fv < 70 else "000000")))
        elif ci == 10:
            c.font = F(True, 8, fg_m); c.fill = BG(bg_m)
        elif ci == 11 and val is not None:
            c.number_format = '"$"#,##0.00'

    for ri, (_, row) in enumerate(top10.iterrows(), 1):
        er  = 91 + ri
        rbg = "F0F4FA" if ri % 2 == 0 else "FFFFFF"
        ws.row_dimensions[er].height = 18
        sig  = row.get("CS_Signal", "") or ""
        moat = row.get("Moat Score", "") or ""
        fg_s, bg_s = SIGNAL_STYLE.get(sig,  ("000000", "FFFFFF"))
        fg_m, bg_m = MOAT_SCORE_STYLE.get(moat, ("000000", "FFFFFF"))
        vals = [ri, row.get("Ticker",""), row.get("Tên Công Ty",""), row.get("Sector",""),
                sig, int(row.get("CS_Score", 0) or 0), row.get("Conviction"),
                row.get("1Y%"), row.get("52W_High%"), moat, row.get("Price ($)")]
        for ci, val in enumerate(vals, 1):
            c = ws.cell(er, ci, val); c.fill = BG(rbg)
            _fmt_t10_cell(c, ci, val, fg_s, bg_s, fg_m, bg_m)

    t10_end = 91 + len(top10)

    # -- [6] CONVICTION SHORTLIST (dynamic rows after [5])
    ws.row_dimensions[t10_end + 1].height = 10
    sec_header(t10_end + 2, "❻  CONVICTION SHORTLIST  —  STRONG BUY  ×  WIDE / NARROW Moat  ×  1Y% > 0")
    conv_hdr_row = t10_end + 3

    for ci, h in enumerate(T10_HDRS, 1):
        c = ws.cell(conv_hdr_row, ci, h)
        c.font = F(True, 9, "FFFFFF"); c.fill = BG("1A5C2B")
        c.alignment = AL(); c.border = BD()
    ws.row_dimensions[conv_hdr_row].height = 20

    _conv_filter = (
        (df["CS_Signal"] == "\U0001f7e2 STRONG BUY") &
        (df["Moat Score"].isin(["WIDE  ★★★", "NARROW ★★"])) &
        (df["1Y%"].fillna(0) > 0)
    )
    conv_df = (df[_conv_filter][_t10_safe]
               .sort_values(["Conviction", "1Y%"] if "Conviction" in df.columns
                             else ["CS_Score", "1Y%"], ascending=[False, False]))

    if conv_df.empty:
        ws.merge_cells(f"A{conv_hdr_row+1}:{CL(len(T10_HDRS))}{conv_hdr_row+1}")
        c = ws.cell(conv_hdr_row+1, 1, "— Không có mã nào đạt đồng thời cả 3 điều kiện —")
        c.font = F(False, 9, "888888", italic=True); c.alignment = AL()
        ws.row_dimensions[conv_hdr_row+1].height = 18
        conv_end = conv_hdr_row + 1
    else:
        for ri, (_, row) in enumerate(conv_df.iterrows(), 1):
            er  = conv_hdr_row + ri
            rbg = "EFFFEE" if ri % 2 == 0 else "F5FFF5"
            ws.row_dimensions[er].height = 18
            sig  = row.get("CS_Signal", "") or ""
            moat = row.get("Moat Score", "") or ""
            fg_s, bg_s = SIGNAL_STYLE.get(sig,  ("000000", "FFFFFF"))
            fg_m, bg_m = MOAT_SCORE_STYLE.get(moat, ("000000", "FFFFFF"))
            vals = [ri, row.get("Ticker",""), row.get("Tên Công Ty",""), row.get("Sector",""),
                    sig, int(row.get("CS_Score", 0) or 0), row.get("Conviction"),
                    row.get("1Y%"), row.get("52W_High%"), moat, row.get("Price ($)")]
            for ci, val in enumerate(vals, 1):
                c = ws.cell(er, ci, val); c.fill = BG(rbg)
                _fmt_t10_cell(c, ci, val, fg_s, bg_s, fg_m, bg_m)
        conv_end = conv_hdr_row + len(conv_df)

    # -- [7] DUAL STRATEGY SHORTLIST (dynamic rows after [6])
    ws.row_dimensions[conv_end + 1].height = 10
    ds_sec_row = conv_end + 2
    ds_hdr_row = conv_end + 3
    sec_header(ds_sec_row, "❼  DUAL STRATEGY SHORTLIST  —  CS Score ≥ 7  AND  QC Score ≥ 4")
    ws.row_dimensions[ds_sec_row].height = 18

    DS_HDRS   = ["#", "Ticker", "Tên Công Ty", "Sector",
                 "CS Signal", f"CS/{N_CS}", "QC Signal", "QC/6",
                 "ROIC%", "EQ Badge", "1Y%"]
    DS_WIDTHS = [4, 9, 22, 14, 14, 6, 16, 6, 9, 16, 8]
    for ci, (h, w) in enumerate(zip(DS_HDRS, DS_WIDTHS), 1):
        ws.column_dimensions[CL(ci)].width = w
        c = ws.cell(ds_hdr_row, ci, h)
        c.font = F(True, 9, "FFFFFF"); c.fill = BG("4A235A")
        c.alignment = AL(); c.border = BD()
    ws.row_dimensions[ds_hdr_row].height = 20

    if HAS_QC:
        _ds_filter = (df["CS_Score"] >= 7) & (df["QC_Score"] >= 4)
        ds_df = df[_ds_filter].sort_values(
            ["QC_Score", "CS_Score", "1Y%"], ascending=[False, False, False])
    else:
        ds_df = pd.DataFrame()

    if ds_df.empty:
        ws.merge_cells(f"A{ds_hdr_row+1}:{CL(len(DS_HDRS))}{ds_hdr_row+1}")
        c = ws.cell(ds_hdr_row+1, 1, "— Không có mã nào đạt cả 2 chiến lược (cần chạy QC scan) —")
        c.font = F(False, 9, "888888", italic=True); c.alignment = AL()
        ws.row_dimensions[ds_hdr_row+1].height = 18
    else:
        for ri, (_, row) in enumerate(ds_df.iterrows(), 1):
            er  = ds_hdr_row + ri
            rbg = "F5F0FF" if ri % 2 == 0 else "FBF7FF"
            ws.row_dimensions[er].height = 18
            sig_cs = row.get("CS_Signal", "") or ""
            sig_qc = row.get("QC_Signal", "") or ""
            eq_b   = row.get("EQ_Badge",  "") or ""
            roic_v = row.get("ROIC%")
            perf_v = row.get("1Y%")
            fg_s, bg_s = SIGNAL_STYLE.get(sig_cs,    ("000000", "FFFFFF"))
            fg_q, bg_q = QC_SIGNAL_STYLE.get(sig_qc,  ("000000", "FFFFFF"))
            fg_e, bg_e = EQ_BADGE_STYLE.get(eq_b,     ("444444", "EEEEEE"))

            col_defs = [
                (1, ri,                               True,  "FFFFFF", "4A235A"),
                (2, row.get("Ticker",""),              True,  "1B3A5C", None),
                (3, row.get("Tên Công Ty",""), False, None,  None, "left"),
                (4, row.get("Sector",""),              False, None,    None, "left"),
                (5, sig_cs,                           True,  fg_s,    bg_s),
                (6, int(row.get("CS_Score", 0) or 0), False, None,   None),
                (7, sig_qc,                           True,  fg_q,    bg_q),
                (8, int(row.get("QC_Score", 0) or 0), False, None,   None),
            ]
            for cd in col_defs:
                ci_   = cd[0]; val_  = cd[1]; bold_ = cd[2]
                fg_   = cd[3]; bg_c_ = cd[4]
                align_ = cd[5] if len(cd) > 5 else "center"
                c2 = ws.cell(er, ci_, val_)
                c2.border = BD(); c2.alignment = AL(align_)
                c2.font = F(bold_, 9, fg_ if fg_ else "000000")
                c2.fill = BG(bg_c_) if bg_c_ else BG(rbg)

            c2 = ws.cell(er, 9); c2.border = BD(); c2.alignment = AL(); c2.fill = BG(rbg)
            if roic_v is not None and not (isinstance(roic_v, float) and pd.isna(roic_v)):
                fv = float(roic_v); c2.value = fv / 100; c2.number_format = '0.0%'
                c2.font = F(True, 9, "276221" if fv >= 15 else "0070C0" if fv >= 10 else "7D6608" if fv >= 0 else "9C0006")
            else:
                c2.value = "—"; c2.font = F(size=9, color="BBBBBB")

            c2 = ws.cell(er, 10, str(eq_b) if eq_b else "—")
            c2.border = BD(); c2.alignment = AL()
            c2.font = F(True, 8, fg_e)
            c2.fill = BG(bg_e) if eq_b else BG(rbg)

            c2 = ws.cell(er, 11); c2.border = BD(); c2.alignment = AL(); c2.fill = BG(rbg)
            if perf_v is not None and not (isinstance(perf_v, float) and pd.isna(perf_v)):
                fv2 = float(perf_v); c2.value = fv2 / 100; c2.number_format = '+0.0%;(0.0%);"-"'
                c2.font = F(size=9, color=("276221" if fv2 > 0 else "9C0006" if fv2 < 0 else "000000"))
            else:
                c2.value = "—"; c2.font = F(size=9, color="BBBBBB")

    # ===================================================================
    # DATA TABLES (row 200+, safely below all visible sections)
    # ===================================================================
    DR = 200

    def _hdr(r, c_idx, text):
        ws.cell(r, c_idx, text).font = F(True, 8, "FFFFFF")
        ws.cell(r, c_idx).fill = BG("1F3864")

    # Table 1 -- Signal Pie (cols 1:2)
    _hdr(DR, 1, "Signal"); _hdr(DR, 2, "Count")
    sig_counts = df["CS_Signal"].value_counts()
    for i, sig in enumerate(SIGNALS, 1):
        ws.cell(DR+i, 1, sig)
        ws.cell(DR+i, 2, int(sig_counts.get(sig, 0)))

    # Table 2 -- CS x QC Scatter (cols 3:7): col3=CS_Score_X, col4..7=QC_Score_Y split by signal
    _hdr(DR, 3, "CS_Score_X")
    for j, lbl in enumerate(SIG_LBLS, 1):
        _hdr(DR, 3+j, lbl)
    if HAS_QC:
        cxq_df = df[["CS_Score", "QC_Score", "CS_Signal"]].dropna(subset=["CS_Score", "QC_Score"])
        for i, (_, row) in enumerate(cxq_df.iterrows(), 1):
            x_val = int(row["CS_Score"]); y_val = round(float(row["QC_Score"]), 2)
            sig   = row["CS_Signal"]
            ws.cell(DR+i, 3, x_val)
            for j, s in enumerate(SIGNALS, 1):
                ws.cell(DR+i, 3+j, y_val if sig == s else None)
        n_cxq = len(cxq_df)
    else:
        n_cxq = 0

    # Table 3 -- Top 20 Bar (cols 8:9)
    _hdr(DR, 8, "Ticker"); _hdr(DR, 9, f"Score/{N_CS}")
    top20 = (df[["Ticker", "CS_Score", "1Y%"]]
             .sort_values(["CS_Score", "1Y%"], ascending=[False, False])
             .head(20)
             .sort_values(["CS_Score", "1Y%"], ascending=[True, True]))
    for i, (_, row) in enumerate(top20.iterrows(), 1):
        ws.cell(DR+i, 8, str(row.get("Ticker", "")))
        ws.cell(DR+i, 9, int(row.get("CS_Score", 0)))

    # Table 4 -- Score x 1Y% Scatter (cols 10:14): col10=Score_X, col11..14=1Y%_Y split by signal
    _hdr(DR, 10, "Score_X")
    for j, lbl in enumerate(SIG_LBLS, 1):
        _hdr(DR, 10+j, lbl)
    sc_df = df[["CS_Score", "1Y%", "CS_Signal"]].dropna(subset=["1Y%"])
    for i, (_, row) in enumerate(sc_df.iterrows(), 1):
        x_val = int(row["CS_Score"]); y_val = round(float(row["1Y%"]), 2)
        sig   = row["CS_Signal"]
        ws.cell(DR+i, 10, x_val)
        for j, s in enumerate(SIGNALS, 1):
            ws.cell(DR+i, 10+j, y_val if sig == s else None)
    n_sc = len(sc_df)

    # Table 5 -- Sector x Signal Stacked Bar (cols 15:19)
    _hdr(DR, 15, "Sector")
    for j, lbl in enumerate(SIG_LBLS, 1):
        _hdr(DR, 15+j, lbl)
    grp = (df.assign(Sector=df["Sector"].fillna("Unknown"))
             .groupby(["Sector", "CS_Signal"])
             .size()
             .unstack(fill_value=0)
             .reindex(columns=SIGNALS, fill_value=0))
    grp["_tot"] = grp.sum(axis=1)
    grp = grp.sort_values("_tot", ascending=True).drop(columns="_tot")
    n_sec = len(grp)
    for i, (sector, row) in enumerate(grp.iterrows(), 1):
        ws.cell(DR+i, 15, str(sector))
        for j, sig in enumerate(SIGNALS, 1):
            ws.cell(DR+i, 15+j, int(row.get(sig, 0)))

    # Table 6 -- ROIC% by Sector (cols 20:21, replaces ROE%/GM%)
    _hdr(DR, 20, "Sector"); _hdr(DR, 21, "Avg ROIC%")
    if "ROIC%" in df.columns:
        sec_roic = (df.assign(Sector=df["Sector"].fillna("Unknown"))
                      .groupby("Sector")["ROIC%"]
                      .mean().round(1)
                      .sort_values(ascending=True))
        n_roic = len(sec_roic)
        for i, (sector, val) in enumerate(sec_roic.items(), 1):
            ws.cell(DR+i, 20, str(sector))
            ws.cell(DR+i, 21, float(val) if not pd.isna(val) else None)
    else:
        n_roic = 0

    # Table 7 -- EQ Badge Pie (cols 22:23)
    _hdr(DR, 22, "EQ Badge"); _hdr(DR, 23, "Count")
    if "EQ_Badge" in df.columns:
        eq_counts = df["EQ_Badge"].value_counts()
        for i, lbl in enumerate(EQ_LBLS, 1):
            ws.cell(DR+i, 22, lbl)
            ws.cell(DR+i, 23, int(eq_counts.get(lbl, 0)))
        n_eq = len(EQ_LBLS)
    else:
        n_eq = 0

    # ===================================================================
    # CHARTS
    # ===================================================================

    # Chart 1: Pie -- Signal Distribution (A15, [3] row-1 left)
    pie1 = PieChart()
    pie1.title = "Signal Distribution"; pie1.style = 10
    pie1.add_data(Reference(ws, min_col=2, min_row=DR, max_row=DR+4), titles_from_data=True)
    pie1.set_categories(Reference(ws, min_col=1, min_row=DR+1, max_row=DR+4))
    for idx, color in enumerate(SIG_COLS):
        pt = DataPoint(idx=idx); pt.graphicalProperties.solidFill = color
        pie1.series[0].dPt.append(pt)
    pie1.width = 12; pie1.height = 12
    ws.add_chart(pie1, "A15")

    # Chart 2: Scatter -- CAN SLIM Score x QC Score (I15, [3] row-1 center)
    if HAS_QC and n_cxq > 0:
        cxq_sct = ScatterChart()
        cxq_sct.title = "CAN SLIM Score × QC Score"; cxq_sct.style = 10
        cxq_sct.x_axis.title = f"CAN SLIM Score (0–{N_CS})"
        cxq_sct.y_axis.title = "QC Score (0–6)"
        cxq_sct.x_axis.scaling.min = 0; cxq_sct.x_axis.scaling.max = N_CS
        cxq_sct.y_axis.scaling.min = 0; cxq_sct.y_axis.scaling.max = N_QC_LOCAL
        for j, (sig_lbl, color) in enumerate(zip(SIG_LBLS, SIG_COLS)):
            xv  = Reference(ws, min_col=3,   min_row=DR+1, max_row=DR+n_cxq)
            yv  = Reference(ws, min_col=4+j, min_row=DR+1, max_row=DR+n_cxq)
            ser = Series(yv, xv, title=sig_lbl)
            ser.marker.symbol = "circle"; ser.marker.size = 5
            ser.graphicalProperties.line.noFill           = True
            ser.marker.graphicalProperties.solidFill      = color
            ser.marker.graphicalProperties.line.solidFill = color
            cxq_sct.series.append(ser)
        cxq_sct.width = 12; cxq_sct.height = 12
        ws.add_chart(cxq_sct, "I15")

    # Chart 3: Pie -- EQ Badge Distribution (Q15, [3] row-1 right)
    if n_eq > 0:
        eq_pie = PieChart()
        eq_pie.title = "Earnings Quality Badge"; eq_pie.style = 10
        eq_pie.add_data(Reference(ws, min_col=23, min_row=DR, max_row=DR+n_eq), titles_from_data=True)
        eq_pie.set_categories(Reference(ws, min_col=22, min_row=DR+1, max_row=DR+n_eq))
        for idx, color in enumerate(EQ_PIE_COLS):
            pt = DataPoint(idx=idx); pt.graphicalProperties.solidFill = color
            eq_pie.series[0].dPt.append(pt)
        eq_pie.width = 12; eq_pie.height = 12
        ws.add_chart(eq_pie, "Q15")

    # Chart 4: Horizontal Bar -- Top 20 Score (A38, [3] row-2 left)
    bar = BarChart()
    bar.type = "bar"; bar.style = 10
    bar.title = f"Top 20  ·  CAN SLIM Score / {N_CS}"
    bar.x_axis.title = "Score"; bar.y_axis.title = "Ticker"
    bar.x_axis.scaling.min = 0; bar.x_axis.scaling.max = N_CS
    bar.add_data(Reference(ws, min_col=9, min_row=DR, max_row=DR+20), titles_from_data=True)
    bar.set_categories(Reference(ws, min_col=8, min_row=DR+1, max_row=DR+20))
    bar.series[0].graphicalProperties.solidFill = "0070C0"
    bar.width = 14; bar.height = 12
    ws.add_chart(bar, "A38")

    # Chart 5: Scatter -- Score vs 1Y% by signal (I38, [3] row-2 right)
    sct = ScatterChart()
    sct.title = "CAN SLIM Score  vs  1Y Return %"; sct.style = 10
    sct.x_axis.title = "CAN SLIM Score"; sct.y_axis.title = "1Y Return %"
    for j, (sig_lbl, color) in enumerate(zip(SIG_LBLS, SIG_COLS)):
        xv  = Reference(ws, min_col=10,   min_row=DR+1, max_row=DR+n_sc)
        yv  = Reference(ws, min_col=11+j, min_row=DR+1, max_row=DR+n_sc)
        ser = Series(yv, xv, title=sig_lbl)
        ser.marker.symbol = "circle"; ser.marker.size = 5
        ser.graphicalProperties.line.noFill           = True
        ser.marker.graphicalProperties.solidFill      = color
        ser.marker.graphicalProperties.line.solidFill = color
        sct.series.append(ser)
    sct.width = 14; sct.height = 12
    ws.add_chart(sct, "I38")

    # Chart 6: Stacked Bar -- Sector x Signal (A62, [4] left)
    stk = BarChart()
    stk.type = "bar"; stk.grouping = "stacked"; stk.style = 10
    stk.title = "Signal Distribution by Sector"
    stk.x_axis.title = "Count"; stk.y_axis.title = "Sector"
    for j, (lbl, color) in enumerate(zip(SIG_LBLS, SIG_COLS)):
        stk.add_data(Reference(ws, min_col=16+j, min_row=DR, max_row=DR+n_sec), titles_from_data=True)
        stk.series[-1].graphicalProperties.solidFill = color
    stk.set_categories(Reference(ws, min_col=15, min_row=DR+1, max_row=DR+n_sec))
    stk.legend.position = "b"
    stk.width = 14; stk.height = 14
    ws.add_chart(stk, "A62")

    # Chart 7: Bar -- Avg ROIC% by Sector (I62, [4] right, replaces ROE%/GM%)
    if n_roic > 0:
        roic_bar = BarChart()
        roic_bar.type = "bar"; roic_bar.style = 10
        roic_bar.title = "Avg ROIC% by Sector"
        roic_bar.x_axis.title = "Avg ROIC%"; roic_bar.y_axis.title = "Sector"
        roic_bar.add_data(Reference(ws, min_col=21, min_row=DR, max_row=DR+n_roic), titles_from_data=True)
        roic_bar.series[-1].graphicalProperties.solidFill = "1A5C2B"
        roic_bar.set_categories(Reference(ws, min_col=20, min_row=DR+1, max_row=DR+n_roic))
        roic_bar.legend.position = "b"
        roic_bar.width = 14; roic_bar.height = 14
        ws.add_chart(roic_bar, "I62")

    # -- Minimize data rows so they're invisible but charts still read them
    n_cxq_safe = n_cxq if HAS_QC else 1
    max_data_row = DR + max(4, 20, n_sc, n_sec, max(n_roic, 1), max(n_eq, 1), n_cxq_safe) + 2
    for r in range(DR, max_data_row + 1):
        ws.row_dimensions[r].height = 1


def _main_sheet(wb, df, market, top):
    ws = wb.active; ws.title = f"Top{top}"
    ws.sheet_view.showGridLines = False
    bdr = BD(); nd = len(DATA_HEADERS)
    CS_START = nd + 1
    CI_SCORE  = CS_START + N_CS
    CI_SIGNAL = CS_START + N_CS + 1
    TOTAL     = CI_SIGNAL
    DR = 5

    ws.merge_cells(f"A1:{CL(TOTAL)}1")
    c = ws["A1"]; c.value = f"📊  TOP {top} {market.upper()}  —  FUNDAMENTAL + CAN SLIM SCREENER"
    c.font = F(True,15,C_WHITE); c.fill = BG(C_DARK); c.alignment = AL(); ws.row_dimensions[1].height = 32

    ws.merge_cells(f"A2:{CL(TOTAL)}2")
    c = ws["A2"]; c.value = f"Source: TradingView + yfinance (Moat)   |   {datetime.now().strftime('%Y-%m-%d %H:%M')}   |   Sorted by Market Cap ↓"
    c.font = F(False,8,"88AACC",True); c.fill = BG(C_DARK); c.alignment = AL(); ws.row_dimensions[2].height = 16

    groups = [
        (1,1,"TICKER","213A5C"),(2,2,"TÊN CÔNG TY","1C3550"),
        (3,3,"SECTOR","2C4F7C"),
        (4,5,"PRICE","2E5E8E"),
        (6,7,"🏰  MOAT (yfinance 5yr)","4A235A"),
        (8,11,"GROWTH","1A5276"),
        (12,15,"QUALITY","154360"),(16,18,"VALUATION","1B2631"),
        (19,23,"PERFORMANCE","0E3251"),
        (CS_START, CI_SCORE-1, "✅  CAN SLIM  (✓ = đạt / ✗ = không đạt)", "1A5C2B"),
        (CI_SCORE, CI_SCORE, "SCORE", "0D3B1A"),
        (CI_SIGNAL, CI_SIGNAL, "SIGNAL", "0D3B1A"),
    ]
    ws.row_dimensions[3].height = 14
    for cs, ce, lbl, bg in groups:
        c = ws.cell(3, cs, lbl)
        c.font = F(True,8,"C8DDEF"); c.fill = BG(bg); c.alignment = AL(); c.border = bdr
        if cs < ce: ws.merge_cells(f"{CL(cs)}3:{CL(ce)}3")

    ws.row_dimensions[4].height = 30
    for ci, h in enumerate(DATA_HEADERS, 1):
        c = ws.cell(4,ci,h); c.font=F(True,8,C_WHITE); c.fill=BG(C_NAVY); c.alignment=AL(wrap=True); c.border=bdr

    for i, key in enumerate(CS_KEYS):
        ci = CS_START + i
        c = ws.cell(4,ci,CANSLIM[key]["label"])
        c.font=F(True,8,C_WHITE); c.fill=BG("1A5C2B"); c.alignment=AL(wrap=True); c.border=bdr
    for ci, lbl in [(CI_SCORE,f"Score\n/{N_CS}"), (CI_SIGNAL,"Signal")]:
        c = ws.cell(4,ci,lbl); c.font=F(True,9,C_WHITE); c.fill=BG("0D3B1A"); c.alignment=AL(wrap=True); c.border=bdr

    for ri, (_, row) in enumerate(df.iterrows(), 1):
        er = DR + ri - 1; alt = ri%2==0; rbg = C_ALT if alt else C_WHITE

        for ci, h in enumerate(DATA_HEADERS, 1):
            val = row.get(h); c = ws.cell(er,ci)
            c.border=bdr; c.fill=BG(rbg); c.alignment=AL("left" if ci==1 else "center")
            if ci==1:
                c.value=val; c.font=F(True,9,C_NAVY); continue
            if val is None or (isinstance(val,float) and pd.isna(val)):
                c.font=F(size=9,color="BBBBBB"); continue
            if h=="Tên Công Ty":
                c.value=val; c.font=F(size=8,color="1C3550"); c.alignment=AL("left")
            elif h=="Sector":
                c.value=val; c.font=F(size=8,italic=True,color="334466"); c.alignment=AL("left")
            elif h=="Moat Proxy":
                c.value=val; c.font=F(size=8,italic=True,color="4A235A")
                c.fill=BG("F5EEF8"); c.alignment=AL("left")
            elif h=="Moat Score":
                fg,bg = MOAT_SCORE_STYLE.get(val, ("000000",C_WHITE))
                c.value=val; c.font=F(True,9,fg); c.fill=BG(bg); c.alignment=AL()
            elif h=="Price ($)":
                c.value=val; c.number_format='"$"#,##0.00'; c.font=F(size=9)
            elif h=="MCap ($B)":
                c.value=val; c.number_format='#,##0.0'; c.font=F(size=9)
            elif ci in PCT_COLS:
                c.value=val/100; c.number_format='+0.0%;(0.0%);"-"'
                c.font = F(size=9,color=CG_FG,bold=(val>25)) if val>0 else (F(size=9,color=CR_FG) if val<0 else F(size=9))
            else:
                c.value=val; c.number_format='0.0'; c.font=F(size=9)

        for i, key in enumerate(CS_KEYS):
            ci = CS_START + i; v = row.get(f"CS_{key}"); c = ws.cell(er,ci)
            c.border=bdr; c.alignment=AL()
            if v is True:    c.value="✓"; c.font=F(True,11,CG_FG); c.fill=BG(CG_BG)
            elif v is False: c.value="✗"; c.font=F(True,10,CR_FG); c.fill=BG(CR_BG)
            else:            c.value="–"; c.font=F(size=9,color="AAAAAA"); c.fill=BG("F2F2F2")

        sc = row.get("CS_Score",0); c = ws.cell(er,CI_SCORE,sc)
        c.border=bdr; c.alignment=AL(); c.number_format="0"
        pct = sc/N_CS
        if pct>=0.875:  c.fill=BG(CG_BG); c.font=F(True,11,CG_FG)
        elif pct>=0.625: c.fill=BG(CB_BG); c.font=F(True,11,CB_FG)
        elif pct>=0.375: c.fill=BG(CY_BG); c.font=F(True,11,CY_FG)
        else:            c.fill=BG(CR_BG); c.font=F(True,11,CR_FG)

        sig = row.get("CS_Signal",""); c = ws.cell(er,CI_SIGNAL,sig)
        c.border=bdr; c.alignment=AL()
        fg,bg = SIGNAL_STYLE.get(sig,("000000",C_WHITE))
        c.font=F(True,9,fg); c.fill=BG(bg)

    for ci,w in {1:9,2:22,3:16,4:8,5:9,
                 6:28,7:13,
                 8:8,9:8,10:8,11:8,12:8,13:7,14:7,15:6,
                 16:6,17:6,18:6,
                 19:6,20:6,21:6,22:6,23:7}.items():
        ws.column_dimensions[CL(ci)].width=w
    for i in range(N_CS): ws.column_dimensions[CL(CS_START+i)].width=9
    ws.column_dimensions[CL(CI_SCORE)].width=7
    ws.column_dimensions[CL(CI_SIGNAL)].width=15
    for ri in range(DR, DR+len(df)): ws.row_dimensions[ri].height=16

    for col_i in [9,11,21]:
        rng=f"{CL(col_i)}{DR}:{CL(col_i)}{DR+len(df)-1}"
        ws.conditional_formatting.add(rng, ColorScaleRule(
            start_type="num",start_value=-0.3,start_color="FFC7CE",
            mid_type="num",mid_value=0,mid_color="FFFFFF",
            end_type="num",end_value=0.5,end_color="C6EFCE"))

    ws.freeze_panes = ws[f"B{DR}"]


def _qc_sheet(wb, df):
    """Quality Compounder sheet — metrics, QC checkmarks, score, signal, EQ badge."""
    ws = wb.create_sheet("Quality Compounder")
    ws.sheet_view.showGridLines = False
    ws.sheet_tab_color = "1A5C2B"
    bdr = BD()

    # Sort by QC_Score desc → 1Y% desc
    sort_cols = [c for c in ["QC_Score", "1Y%"] if c in df.columns]
    df_qc = (df.sort_values(sort_cols, ascending=[False] * len(sort_cols))
               .reset_index(drop=True) if sort_cols else df.reset_index(drop=True))

    DATA_HDRS = [
        ("Ticker",      "Ticker",        9),
        ("Tên Công Ty", "Tên Công Ty",  22),
        ("Sector",      "Sector",        14),
        ("Price ($)",   "Price ($)",      8),
        ("MCap ($B)",   "MCap ($B)",      9),
        ("Moat Score",  "Moat Score",    13),
        ("GM%",         "Gross Margin%",  7),
        ("Op Margin%",  "Op Margin%",     9),
        ("ROIC%",       "ROIC%",          7),
        ("D/E",         "D/E",            6),
        ("FCF/sh",      "FCF/sh",         7),
        ("Curr.R",      "Current Ratio",  7),
        ("EV/EBITDA",   "EV/EBITDA",      9),
        ("P/E",         "P/E",            6),
        ("1Y%",         "1Y%",            7),
    ]
    QC_CHECK_HDRS = [
        ("ROIC✓",  "QC_ROIC"),
        ("OpMgn✓", "QC_OPGM"),
        ("GM✓",    "QC_GM"),
        ("FCF✓",   "QC_FCF"),
        ("DE✓",    "QC_DE"),
        ("Moat✓",  "QC_MOAT"),
    ]
    RESULT_HDRS = [
        ("Score",    "QC_Score"),
        ("Quality",  "QC_Signal"),
        ("EQ Badge", "EQ_Badge"),
    ]

    ND   = len(DATA_HDRS)
    NQC  = len(QC_CHECK_HDRS)
    NR   = len(RESULT_HDRS)
    TOTAL = ND + NQC + NR
    PCT_COLS = {"Gross Margin%", "Op Margin%", "ROIC%", "1Y%"}
    DR = 5

    # ── Title ─────────────────────────────────────────────────────────────────
    ws.merge_cells(f"A1:{CL(TOTAL)}1")
    c = ws["A1"]
    c.value = "📊  QUALITY COMPOUNDER SCREENER"
    c.font = Font(name="Calibri", bold=True, size=15, color="FFFFFF")
    c.fill = BG("0D2137"); c.alignment = AL()
    ws.row_dimensions[1].height = 32

    ws.merge_cells(f"A2:{CL(TOTAL)}2")
    c = ws["A2"]
    c.value = (f"Source: TradingView + yfinance   |   "
               f"{datetime.now().strftime('%Y-%m-%d %H:%M')}   |   Sorted by QC Score ↓")
    c.font = F(False, 8, "88AACC", True); c.fill = BG("0D2137"); c.alignment = AL()
    ws.row_dimensions[2].height = 16

    # ── Group headers (row 3) ─────────────────────────────────────────────────
    groups = [
        (1,       3,        "IDENTITY",                                    "1C3550"),
        (4,       5,        "PRICE & SIZE",                                "2E5E8E"),
        (6,       6,        "🏰  MOAT",                                    "4A235A"),
        (7,       15,       "📊  QC METRICS",                              "0E3251"),
        (16,      21,       "✅  QC CRITERIA  (✓ = đạt / ✗ = không đạt)", "1A5C2B"),
        (22,      22,       "SCORE",                                       "0D3B1A"),
        (23,      23,       "QUALITY",                                     "0D3B1A"),
        (24,      24,       "EQ BADGE",                                    "0D2020"),
    ]
    ws.row_dimensions[3].height = 14
    for cs, ce, lbl, bg_hex in groups:
        c = ws.cell(3, cs, lbl)
        c.font = F(True, 8, "C8DDEF"); c.fill = BG(bg_hex)
        c.alignment = AL(); c.border = bdr
        if cs < ce:
            ws.merge_cells(f"{CL(cs)}3:{CL(ce)}3")

    # ── Column headers (row 4) ────────────────────────────────────────────────
    ws.row_dimensions[4].height = 28
    for ci, (hdr, _, w) in enumerate(DATA_HDRS, 1):
        c = ws.cell(4, ci, hdr)
        c.font = F(True, 8, C_WHITE); c.fill = BG(C_NAVY)
        c.alignment = AL(wrap=True); c.border = bdr
        ws.column_dimensions[CL(ci)].width = w

    for i, (hdr, _) in enumerate(QC_CHECK_HDRS):
        ci = ND + 1 + i
        c = ws.cell(4, ci, hdr)
        c.font = F(True, 8, "5EC472"); c.fill = BG("0D2A10")
        c.alignment = AL(wrap=True); c.border = bdr
        ws.column_dimensions[CL(ci)].width = 7

    for i, (hdr, _) in enumerate(RESULT_HDRS):
        ci = ND + NQC + 1 + i
        bg_h = "0D3B1A" if i < 2 else "0D2020"
        c = ws.cell(4, ci, hdr)
        c.font = F(True, 8, C_WHITE); c.fill = BG(bg_h)
        c.alignment = AL(wrap=True); c.border = bdr
        ws.column_dimensions[CL(ci)].width = 7 if i == 0 else 16

    # ── Data rows ─────────────────────────────────────────────────────────────
    for ri, (_, row) in enumerate(df_qc.iterrows(), 1):
        er  = DR + ri - 1
        alt = ri % 2 == 0
        rbg = C_ALT if alt else C_WHITE
        ws.row_dimensions[er].height = 16

        for ci, (_, key, _) in enumerate(DATA_HDRS, 1):
            val = row.get(key)
            c   = ws.cell(er, ci)
            c.border = bdr; c.fill = BG(rbg)
            c.alignment = AL("left" if ci <= 3 else "center")
            if val is None or (isinstance(val, float) and pd.isna(val)):
                c.value = "—"; c.font = F(size=9, color="BBBBBB"); continue
            if key == "Ticker":
                c.value = val; c.font = F(True, 9, C_NAVY); c.alignment = AL("left")
            elif key == "Tên Công Ty":
                c.value = val; c.font = F(size=8, color="1C3550"); c.alignment = AL("left")
            elif key == "Sector":
                c.value = val; c.font = F(size=8, italic=True, color="334466"); c.alignment = AL("left")
            elif key == "Moat Score":
                fg_m, bg_m = MOAT_SCORE_STYLE.get(val, ("000000", C_WHITE))
                c.value = val; c.font = F(True, 9, fg_m); c.fill = BG(bg_m); c.alignment = AL()
            elif key == "Price ($)":
                c.value = val; c.number_format = '"$"#,##0.00'; c.font = F(size=9)
            elif key == "MCap ($B)":
                c.value = val; c.number_format = "#,##0.0"; c.font = F(size=9)
            elif key in PCT_COLS:
                c.value = val / 100; c.number_format = '+0.0%;(0.0%);"-"'
                c.font = (F(size=9, color=CG_FG, bold=(val > 25)) if val > 0
                          else F(size=9, color=CR_FG) if val < 0 else F(size=9))
            else:
                c.value = val; c.number_format = "0.0"; c.font = F(size=9)

        for i, (_, qc_key) in enumerate(QC_CHECK_HDRS):
            ci = ND + 1 + i
            v  = row.get(qc_key)
            c  = ws.cell(er, ci)
            c.border = bdr; c.alignment = AL()
            if v is True:
                c.value = "✓"; c.font = F(True, 11, CG_FG); c.fill = BG(CG_BG)
            elif v is False:
                c.value = "✗"; c.font = F(True, 10, CR_FG); c.fill = BG(CR_BG)
            else:
                c.value = "–"; c.font = F(size=9, color="AAAAAA"); c.fill = BG("F2F2F2")

        # Score
        sc  = int(row.get("QC_Score", 0) or 0)
        pct = sc / NQC
        c   = ws.cell(er, ND + NQC + 1, sc)
        c.border = bdr; c.alignment = AL(); c.number_format = "0"
        if pct >= 5/6:   c.fill = BG(CG_BG); c.font = F(True, 11, CG_FG)
        elif pct >= 3/6: c.fill = BG(CB_BG); c.font = F(True, 11, CB_FG)
        elif pct >= 1/6: c.fill = BG(CY_BG); c.font = F(True, 11, CY_FG)
        else:            c.fill = BG(CR_BG); c.font = F(True, 11, CR_FG)

        # Quality signal
        sig = row.get("QC_Signal", "")
        fg_s, bg_s = QC_SIGNAL_STYLE.get(sig, ("000000", C_WHITE))
        c = ws.cell(er, ND + NQC + 2, sig or "")
        c.border = bdr; c.alignment = AL()
        c.font = F(True, 9, fg_s); c.fill = BG(bg_s)

        # EQ Badge
        eq = row.get("EQ_Badge")
        c  = ws.cell(er, ND + NQC + 3, eq if eq else "—")
        c.border = bdr; c.alignment = AL()
        if eq and eq in EQ_BADGE_STYLE:
            fg_e, bg_e = EQ_BADGE_STYLE[eq]
            c.font = F(True, 9, fg_e); c.fill = BG(bg_e)
        else:
            c.font = F(size=9, color="AAAAAA"); c.fill = BG(rbg)

    ws.freeze_panes = ws[f"B{DR}"]


# ═══════════════════════════════════════════════════════════════
# MAIN
# ═══════════════════════════════════════════════════════════════
def main():
    p=argparse.ArgumentParser()
    p.add_argument("--market", default="america")
    p.add_argument("--top",    type=int, default=300)
    p.add_argument("--output", default=None)
    p.add_argument("--no-yf",  action="store_true",
                   help="Bỏ qua yfinance, dùng TV TTM data cho Moat (nhanh hơn)")
    a=p.parse_args()
    out = a.output or f"tv_top{a.top}_{a.market}_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"

    print(f"\n{'='*60}")
    print(f"  TV SCREENER — TOP {a.top} {a.market.upper()} + CAN SLIM")
    print(f"{'='*60}\n")

    # 1. Lấy data từ TradingView
    df = clean_df(fetch_data(a.market, a.top))

    # 2. Fetch Moat từ yfinance (trừ khi --no-yf)
    moat_cache = None
    if not a.no_yf:
        tickers = df["Ticker"].tolist()
        sectors = dict(zip(df["Ticker"], df["Sector"].fillna("")))
        moat_cache = build_moat_cache(tickers, sectors)
    else:
        print("  ⚡ Bỏ qua yfinance — dùng TV TTM data cho Moat\n")

    # 3. Score + Moat
    df = score_canslim(df, moat_cache)

    # Preview
    prev = ["Ticker","Moat Proxy","Moat Score","EPS Qtr%","ROE%","CS_Score","CS_Signal"]
    print(df[[c for c in prev if c in df.columns]].head(10).to_string(index=False))
    print()

    print("  📝 Tạo Excel...", end="", flush=True)
    write_excel(df, out, a.market, a.top)
    print(f" ✅\n  ✨ Done → {out}\n")


if __name__ == "__main__":
    main()