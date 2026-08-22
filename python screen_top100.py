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
# NHÓM A — TÀI CHÍNH TỐT  (số liệu đẹp: sinh lời, biên, tăng trưởng, nợ)
# Toàn bộ lấy từ TradingView → có cho mọi thị trường.
# ═══════════════════════════════════════════════════════════════
FIN_CRITERIA = {
    "ROIC": dict(label="ROIC",       field="ROIC%",        thr=15, op="gt",
                 desc="ROIC > 15% — sinh lời cao hơn chi phí vốn (nhóm Tài chính: ROE > 12%)"),
    "OPM":  dict(label="Op Margin",  field="Op Margin%",   thr=15, op="gt",
                 desc="Biên lợi nhuận hoạt động > 15% — kỷ luật chi phí, đòn bẩy vận hành"),
    "GM":   dict(label="Gross Mgn",  field="Gross Margin%", thr=40, op="gt",
                 desc="Biên lợi nhuận gộp > 40% — lợi thế cấu trúc về giá bán"),
    "NIG":  dict(label="NI Growth",  field="NI Growth%",   thr=20, op="gt",
                 desc="Lợi nhuận ròng TTM tăng > 20% so cùng kỳ"),
    "REVG": dict(label="Rev Growth", field="Rev Growth%",  thr=15, op="gt",
                 desc="Doanh thu TTM tăng > 15% so cùng kỳ"),
    "DE":   dict(label="D/E",        field="D/E",          thr=1.0, op="lt",
                 desc="Nợ trên vốn chủ < 1.0 (không áp dụng cho nhóm Tài chính)"),
}
FIN_KEYS = list(FIN_CRITERIA.keys())
N_FIN    = len(FIN_KEYS)

# ═══════════════════════════════════════════════════════════════
# NHÓM B — CHỨNG MINH LÀ THỰC
# Mỗi tiêu chí bác bỏ một cách cụ thể mà số liệu nhóm A có thể bị làm đẹp.
# 5/6 lấy từ TradingView; riêng SUS cần yfinance (trung bình 5 năm).
# ═══════════════════════════════════════════════════════════════
REAL_CRITERIA = {
    "CASH": dict(label="FCF/NI",     field="FCF/NI",       thr=0.8, op="gt",
                 bac_bo="Lợi nhuận ghi sổ nhưng chưa thu được tiền (accrual)",
                 desc="Dòng tiền tự do ≥ 80% lợi nhuận ròng — lợi nhuận có tiền mặt thật đi kèm"),
    "FCFM": dict(label="FCF Mgn",    field="FCF_Margin%",  thr=8,  op="gt",
                 bac_bo="Biên lợi nhuận cao nhờ vốn hóa chi phí hoặc hoãn đầu tư",
                 desc="Biên dòng tiền tự do > 8% — tiền thực trên mỗi đồng doanh thu"),
    "BUYB": dict(label="Không buyback", field="EPS_NI_Gap", thr=-10, op="gt",
                 bac_bo="EPS tăng nhờ mua lại cổ phiếu chứ không phải lợi nhuận tăng",
                 desc="Tăng trưởng EPS không vượt tăng trưởng lợi nhuận quá 10 điểm %"),
    "LEV":  dict(label="ROE thật",   field="ROE_ROIC_Gap", thr=15, op="lt",
                 bac_bo="ROE cao nhờ đòn bẩy hoặc vốn chủ teo lại vì mua cổ phiếu quỹ",
                 desc="ROE không vượt ROIC quá 15 điểm % — sinh lời đến từ kinh doanh, không từ nợ"),
    "BAL":  dict(label="Bảng cân đối", field="Nợ ròng/FCF", thr=0, op="bal",
                 bac_bo="Bảng cân đối trông sạch nhưng thực tế không gánh nổi nghĩa vụ nợ",
                 desc="Có tiền ròng dương, hoặc thanh khoản tốt (Current Ratio > 1.5), "
                      "hoặc trả hết được nợ ròng trong dưới 3 năm bằng dòng tiền tự do"),
    "TREND": dict(label="Còn đúng",  field="NI Qtr%",      thr=10, op="trend",
                 bac_bo="Số TTM còn đẹp chỉ vì ba quý cũ đè lên, quý gần nhất đã xấu đi",
                 desc="Quý gần nhất vẫn tăng trưởng lợi nhuận trên 10%, hoặc không tụt quá 15 điểm % "
                      "so với mức TTM — số liệu tốt vẫn còn đúng ở kỳ mới nhất"),
    "SUS":  dict(label="Bền vững 5N", field="Moat Score",   thr=None, op="in",
                 bac_bo="Một kỳ may mắn được trình bày như năng lực dài hạn",
                 desc="ROE và biên gộp trung bình 5 năm ở mức cao — số đẹp lặp lại nhiều năm"),
}
REAL_KEYS = list(REAL_CRITERIA.keys())
N_REAL    = len(REAL_KEYS)

FETCH_COLS = [
    "name", "description", "close", "volume", "market_cap_basic",
    "sector",
    "gross_margin", "net_margin", "return_on_equity", "debt_to_equity",
    "price_earnings_ttm", "price_sales_current",
    "relative_volume_10d_calc",
    "Perf.W", "Perf.1M", "Perf.3M", "Perf.6M", "Perf.Y",
    "return_on_invested_capital", "operating_margin",
    "enterprise_value_ebitda_ttm", "free_cash_flow_per_share_ttm",
    "current_ratio",
    # ── Nhóm A: tăng trưởng TTM (mượt hơn số quý đơn lẻ) ──────────
    "net_income_yoy_growth_ttm",
    "total_revenue_yoy_growth_ttm",
    "earnings_per_share_diluted_yoy_growth_ttm",
    # ── Nhóm B: dữ liệu kiểm chứng (đều có sẵn trên TradingView) ──
    "free_cash_flow_ttm",
    "free_cash_flow_margin_ttm",
    "net_income",
    "total_revenue_ttm",
    # ── Số liệu quý gần nhất: TTM có độ trễ, quý cho biết hiện tại ra sao ──
    "total_revenue_fq",
    "net_income_fq",
    "total_revenue_yoy_growth_fq",
    "net_income_yoy_growth_fq",
    "cash_n_short_term_invest_fq",
    "total_debt_fq",
]

RENAME = {
    "name": "Ticker", "description": "Tên Công Ty", "close": "Price ($)", "market_cap_basic": "MCap ($B)",
    "sector": "Sector",
    "net_income_yoy_growth_ttm": "NI Growth%",
    "total_revenue_yoy_growth_ttm": "Rev Growth%",
    "earnings_per_share_diluted_yoy_growth_ttm": "EPS Growth%",
    "free_cash_flow_ttm": "FCF_abs",
    "free_cash_flow_margin_ttm": "FCF_Margin%",
    "net_income": "NI_abs",
    "total_revenue_ttm": "Rev_abs",
    "total_revenue_fq": "Rev_q_abs",
    "net_income_fq": "NI_q_abs",
    "total_revenue_yoy_growth_fq": "Rev Qtr%",
    "net_income_yoy_growth_fq": "NI Qtr%",
    "cash_n_short_term_invest_fq": "Cash_abs",
    "total_debt_fq": "Debt_abs",
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
    "Ticker", "Tên Công Ty", "Sector", "Price ($)", "MCap ($B)", "Moat Score",
    "ROIC%", "Op Margin%", "Gross Margin%", "Net Margin%", "NI Growth%", "Rev Growth%",
    "ROE%", "D/E", "FCF/NI", "FCF_Margin%", "EPS_NI_Gap", "ROE_ROIC_Gap",
    "Net Cash ($B)", "Current Ratio", "P/E", "PEG", "1Y%",
]
PCT_COLS = {7, 8, 9, 10, 11, 12, 13, 16, 17, 18, 23}


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
        import yfinance as yf, sys, io
        tk   = yf.Ticker(ticker)
        _stderr, sys.stderr = sys.stderr, io.StringIO()
        try:
            info = tk.info or {}
        finally:
            sys.stderr = _stderr

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

        # ── EQ Badge + FCF Margin% (shared data fetch) ───────────
        eq_badge   = None
        fcf_margin = None
        try:
            cf_s   = tk.cashflow.sort_index(axis=1, ascending=False)
            inc_s  = tk.income_stmt.sort_index(axis=1, ascending=False)

            # FCF (TTM) — prefer direct row, fallback OCF−CapEx
            fcf_row = _get_row(cf_s,  "Free Cash Flow",        "FreeCashFlow")
            if fcf_row.empty:
                ocf = _get_row(cf_s, "Operating Cash Flow",    "OperatingCashFlow")
                cex = _get_row(cf_s, "Capital Expenditure",    "CapitalExpenditure")
                ocf_v = _sv(ocf, 0); cex_v = _sv(cex, 0)
                fcf_v = (ocf_v + (cex_v or 0)) if ocf_v is not None else None
            else:
                fcf_v = _sv(fcf_row, 0)

            # Net Income (TTM) → EQ Badge
            ni_row = _get_row(inc_s, "Net Income", "NetIncome")
            ni_v   = _sv(ni_row, 0)
            if fcf_v is not None and ni_v is not None and ni_v != 0:
                ratio    = fcf_v / ni_v
                eq_badge = ("💚 Cash Backed"   if ratio >= 0.8
                            else "🟡 Mixed"     if ratio >= 0.3
                            else "🔴 Accrual Heavy")

            # Revenue (TTM) → FCF Margin%
            rev_row = _get_row(inc_s, "Total Revenue", "TotalRevenue")
            rev_v   = _sv(rev_row, 0)
            if fcf_v is not None and rev_v is not None and rev_v != 0:
                fcf_margin = round(fcf_v / rev_v * 100, 1)
        except Exception:
            pass

        # ── Net Cash = Cash + STI − Total Debt ───────────────────
        net_cash = None
        try:
            bal_nc = tk.balance_sheet.sort_index(axis=1, ascending=False)
            cash_r = _get_row(bal_nc, "Cash And Cash Equivalents",
                              "CashAndCashEquivalents",
                              "Cash Cash Equivalents And Short Term Investments", "Cash")
            sti_r  = _get_row(bal_nc, "Short Term Investments",
                              "OtherShortTermInvestments")
            debt_r = _get_row(bal_nc, "Total Debt", "TotalDebt")
            cash_v = _sv(cash_r, 0) or 0
            sti_v  = _sv(sti_r,  0) or 0
            debt_v = _sv(debt_r, 0)
            if debt_v is None:
                ltd = _sv(_get_row(bal_nc, "Long Term Debt", "LongTermDebt"), 0) or 0
                std = _sv(_get_row(bal_nc, "Current Debt", "CurrentDebt",
                                   "Short Long Term Debt"), 0) or 0
                debt_v = ltd + std
            net_cash = round((cash_v + sti_v - (debt_v or 0)) / 1e9, 2)
        except Exception:
            pass

        # ── EPS Quarterly YoY Acceleration ───────────────────────
        eps_acc = (0, [])   # (consecutive_acc_qtrs, yoy_list oldest→newest)
        try:
            # Thử quarterly_income_stmt trước (mới hơn), fallback quarterly_financials
            try:
                qfin_acc = tk.quarterly_income_stmt.sort_index(axis=1, ascending=True)
            except Exception:
                qfin_acc = tk.quarterly_financials.sort_index(axis=1, ascending=True)

            eps_key = next((k for k in ["Diluted EPS", "Basic EPS"]
                            if k in qfin_acc.index), None)
            if eps_key:
                eps_s = qfin_acc.loc[eps_key].dropna()
                n = len(eps_s)
                # Cần ≥5 quý để có ít nhất 1 điểm YoY (Q4 vs Q0)
                if n >= 5:
                    # Lấy tất cả điểm YoY có thể (tối đa 4 quý gần nhất)
                    yoy_list = []
                    for _i in range(4, n):
                        curr   = eps_s.iloc[_i]
                        yr_ago = eps_s.iloc[_i - 4]
                        if yr_ago != 0:
                            yoy_list.append(round((curr - yr_ago) / abs(yr_ago) * 100, 1))

                    if len(yoy_list) >= 2:
                        # Đủ dữ liệu: đếm streak acceleration liên tiếp từ quý mới nhất
                        consec = 0
                        for _j in range(len(yoy_list) - 1, 0, -1):
                            if yoy_list[_j] > yoy_list[_j - 1]:
                                consec += 1
                            else:
                                break
                    elif len(yoy_list) == 1:
                        # Chỉ 1 điểm YoY — dùng dấu hiệu positive/negative
                        consec = 1 if yoy_list[0] > 0 else 0
                    else:
                        consec = 0
                    eps_acc = (consec, yoy_list)
        except Exception:
            pass

        return proxy, score, w52_pct, eq_badge, eps_acc, fcf_margin, net_cash

    except Exception as e:
        # Fallback: sector hardcode + UNCERTAIN
        proxy = MOAT_PROXY_MAP.get(sector, "Cost Advantage")
        return proxy, "UNCERTAIN ★", None, None, (0, []), None, None


def build_moat_cache(tickers: list, sectors: dict) -> dict:
    """
    Fetch moat cho toàn bộ danh sách tickers từ yfinance.
    Trả về dict: {ticker: (proxy, score, w52_pct, eq_badge, eps_acc, fcf_margin, net_cash)}
    """
    cache = {}
    total = len(tickers)
    print(f"\n  🏰 Fetch Moat từ yfinance ({total} tickers)...")
    for i, ticker in enumerate(tickers, 1):
        sector = sectors.get(ticker, "")
        print(f"  [{i:>3}/{total}] {ticker:<8}", end="", flush=True)
        proxy, score, w52_pct, eq_badge, eps_acc, fcf_margin, net_cash = fetch_moat_yfinance(ticker, sector)
        cache[ticker] = (proxy, score, w52_pct, eq_badge, eps_acc, fcf_margin, net_cash)
        eq_str  = f" [{eq_badge}]" if eq_badge else ""
        acc_str = f" [ACC:{eps_acc[0]}Q]" if eps_acc and eps_acc[0] > 0 else ""
        fcf_str = f" [FCF:{fcf_margin:.0f}%]" if fcf_margin is not None else ""
        nc_str  = f" [NC:{net_cash:+.1f}B]" if net_cash is not None else ""
        print(f" {score}{eq_str}{acc_str}{fcf_str}{nc_str}")
        time.sleep(0.4)   # tránh rate limit
    print()
    return cache


_MKT_INDEX = {
    "america":   ("america",   "SP:SPX"),
    "nasdaq":    ("america",   "NASDAQ:NDX"),
    "nyse":      ("america",   "NYSE:NYA"),
    "euronext":  ("global",    "TVC:CAC40"),
    "hong_kong": ("hongkong",  "HSI:HSI"),
    "vietnam":   ("vietnam",   "HOSE:VNINDEX"),
}

def fetch_market_direction(market="america"):
    """Returns (market_ok, index_1y).
    market_ok: True nếu index trên MA50 & MA200; index_1y: % 1-year perf của index."""
    import requests
    tv_market, ticker = _MKT_INDEX.get(market.lower(), ("america", "SP:SPX"))
    try:
        r = requests.post(
            f"https://scanner.tradingview.com/{tv_market}/scan",
            json={"symbols": {"tickers": [ticker], "query": {"types": []}},
                  "columns": ["close", "SMA50", "SMA200", "Perf.Y"]},
            timeout=10)
        d = r.json()["data"][0]["d"]
        price, ma50, ma200, perf_y = d[0], d[1], d[2], d[3]
        if price is None or ma50 is None or ma200 is None:
            return None, None
        market_ok = float(price) > float(ma50) and float(price) > float(ma200)
        index_1y = float(perf_y) if perf_y is not None else None
        return market_ok, index_1y
    except Exception:
        return None, None


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
# SCORING — Nhóm A (Tài chính tốt) + Nhóm B (Chứng minh là thực)
# ═══════════════════════════════════════════════════════════════
_FIN_SECTORS = ("Financial Services", "Finance")
_MOAT_GOOD   = {"WIDE  ★★★", "NARROW ★★"}


def _num(v):
    """float nếu là số dùng được, ngược lại None."""
    if v is None:
        return None
    if isinstance(v, float) and pd.isna(v):
        return None
    try:
        return float(v)
    except (TypeError, ValueError):
        return None


def _derive_fields(df, moat_cache=None):
    """Các trường phái sinh mà nhóm B cần — đều tính từ dữ liệu TradingView,
    riêng Moat Score (trung bình 5 năm) cần yfinance."""

    # FCF / Net Income — lợi nhuận có kèm tiền mặt thật hay chỉ là bút toán
    def _fcf_ni(r):
        fcf = _num(r.get("FCF_abs")); ni = _num(r.get("NI_abs"))
        if fcf is None or ni is None or ni <= 0:
            return None          # đang lỗ thì tỷ lệ này vô nghĩa
        return round(fcf / ni, 2)
    df["FCF/NI"] = df.apply(_fcf_ni, axis=1)

    # EPS tăng nhanh hơn lợi nhuận ròng bao nhiêu điểm % → phần chênh đến từ
    # việc số cổ phiếu lưu hành giảm (mua cổ phiếu quỹ), không phải kinh doanh
    def _gap(r):
        eps = _num(r.get("EPS Growth%")); ni = _num(r.get("NI Growth%"))
        if eps is None or ni is None:
            return None
        return round(eps - ni, 1)
    df["EPS_NI_Gap"] = df.apply(_gap, axis=1)

    # ROE vượt ROIC bao nhiêu điểm % → phần chênh đến từ đòn bẩy tài chính
    def _lev_gap(r):
        roe = _num(r.get("ROE%")); roic = _num(r.get("ROIC%"))
        if roe is None or roic is None:
            return None
        return round(roe - roic, 1)
    df["ROE_ROIC_Gap"] = df.apply(_lev_gap, axis=1)

    # Đà tăng trưởng: quý gần nhất so với mức TTM.
    # TTM có độ trễ — ba quý cũ tốt vẫn kéo số TTM lên trong khi quý mới đã xấu đi.
    # Số âm lớn = đang giảm tốc; dương = đang tăng tốc.
    def _trend(r):
        q = _num(r.get("NI Qtr%")); t = _num(r.get("NI Growth%"))
        if q is None or t is None:
            return None
        return round(q - t, 1)
    df["NI_Trend_Gap"] = df.apply(_trend, axis=1)

    def _trend_rev(r):
        q = _num(r.get("Rev Qtr%")); t = _num(r.get("Rev Growth%"))
        if q is None or t is None:
            return None
        return round(q - t, 1)
    df["Rev_Trend_Gap"] = df.apply(_trend_rev, axis=1)

    # Số tuyệt đối quy về tỷ — chỉ để hiển thị, không chấm điểm.
    # Đặt Doanh thu · Lợi nhuận · FCF cạnh nhau thì đọc được bằng mắt
    # phép kiểm chứng FCF/NI thay vì chỉ thấy tỷ số.
    for src, dst in (("Rev_abs",   "Doanh thu ($B)"),
                     ("NI_abs",    "Lợi nhuận ($B)"),
                     ("FCF_abs",   "FCF ($B)"),
                     ("Rev_q_abs", "DThu quý ($B)"),
                     ("NI_q_abs",  "LN quý ($B)")):
        df[dst] = df[src].apply(
            lambda v: round(_num(v) / 1e9, 2) if _num(v) is not None else None) \
            if src in df.columns else None

    # Net Cash ($B) = tiền và tương đương + đầu tư ngắn hạn − tổng nợ
    def _net_cash(r):
        cash = _num(r.get("Cash_abs")); debt = _num(r.get("Debt_abs"))
        if cash is None and debt is None:
            return None
        return round(((cash or 0) - (debt or 0)) / 1e9, 2)
    df["Net Cash ($B)"] = df.apply(_net_cash, axis=1)

    # Bao nhiêu năm dòng tiền tự do thì trả hết nợ ròng.
    # Current Ratio vốn là thước đo cho doanh nghiệp sản xuất có tồn kho và khoản
    # phải thu; với công ty dịch vụ tạo nhiều tiền mặt thì khả năng trả nợ đo bằng
    # dòng tiền mới đúng bản chất. Rỗng khi đã có tiền ròng dương (không cần trả).
    def _debt_years(r):
        nc  = _num(r.get("Net Cash ($B)")); fcf = _num(r.get("FCF_abs"))
        if nc is None or nc >= 0 or fcf is None or fcf <= 0:
            return None
        return round(-nc / (fcf / 1e9), 1)
    df["Nợ ròng/FCF"] = df.apply(_debt_years, axis=1)

    # Moat Score = ROE & biên gộp trung bình 5 năm (chỉ từ yfinance).
    # Không có cache thì để trống — KHÔNG thay bằng số liệu một kỳ, vì tiêu chí
    # này đo tính BỀN VỮNG; lấy một kỳ mà gọi là bền vững thì chính là số ảo.
    if moat_cache:
        df["Moat Score"] = df["Ticker"].apply(
            lambda t: moat_cache[t][1] if t in moat_cache else None)
        df["52W_High%"] = df["Ticker"].apply(
            lambda t: (moat_cache[t][2] if t in moat_cache and len(moat_cache[t]) >= 3 else None))
    else:
        df["Moat Score"] = None
        df["52W_High%"]  = None

    # PEG — hiển thị tham khảo, không chấm điểm
    def _peg(r):
        pe = _num(r.get("P/E")); g = _num(r.get("NI Growth%"))
        if pe is None or g is None or pe <= 0 or g < 5:
            return None
        peg = pe / g
        return round(peg, 2) if peg <= 50 else None
    df["PEG"] = df.apply(_peg, axis=1)
    return df


# ── Nhóm A ──────────────────────────────────────────────────────
def _chk_fin(row, key):
    cfg    = FIN_CRITERIA[key]
    is_fin = row.get("Sector", "") in _FIN_SECTORS
    if key == "DE" and is_fin:
        return None                       # tiền gửi khách hàng là "nợ" — D/E vô nghĩa với ngân hàng
    if key == "ROIC" and is_fin:
        v = _num(row.get("ROE%"))         # nhóm tài chính: dùng ROE > 12% thay cho ROIC
        return None if v is None else v > 12
    v = _num(row.get(cfg["field"]))
    if v is None:
        return None
    return v > cfg["thr"] if cfg["op"] == "gt" else v < cfg["thr"]


# ── Nhóm B ──────────────────────────────────────────────────────
def _chk_real(row, key):
    is_fin = row.get("Sector", "") in _FIN_SECTORS

    if key == "CASH":                     # lợi nhuận có tiền mặt thật đi kèm
        v = _num(row.get("FCF/NI"))
        return None if v is None else v >= 0.8

    if key == "FCFM":                     # tiền thực trên mỗi đồng doanh thu
        v = _num(row.get("FCF_Margin%"))
        return None if v is None else v > 8

    if key == "BUYB":                     # EPS tăng không phải nhờ mua cổ phiếu quỹ
        v = _num(row.get("EPS_NI_Gap"))
        return None if v is None else v <= 10

    if key == "LEV":                      # ROE cao đến từ kinh doanh, không từ đòn bẩy
        if is_fin:
            return None                   # ngân hàng vốn dĩ vận hành bằng đòn bẩy
        v = _num(row.get("ROE_ROIC_Gap"))
        return None if v is None else v <= 15

    if key == "BAL":                      # bảng cân đối chịu được kiểm tra
        if is_fin:
            return None
        nc = _num(row.get("Net Cash ($B)"))
        cr = _num(row.get("Current Ratio"))
        yr = _num(row.get("Nợ ròng/FCF"))
        if nc is None and cr is None:
            return None
        if nc is not None and nc > 0:            # có nhiều tiền hơn nợ
            return True
        if cr is not None and cr > 1.5:          # thanh khoản ngắn hạn dư dả
            return True
        if yr is not None and yr < 3:            # trả hết nợ trong dưới 3 năm bằng FCF
            return True
        return False

    if key == "TREND":                    # số TTM đẹp có còn đúng ở quý gần nhất không
        q   = _num(row.get("NI Qtr%"))
        gap = _num(row.get("NI_Trend_Gap"))
        if q is None:
            return None
        # Đạt nếu quý gần nhất vẫn tăng trưởng tốt, HOẶC không tụt xa so với chính mức TTM.
        # Phải có vế đầu: so bằng điểm phần trăm sẽ vô nghĩa khi tăng trưởng ba chữ số
        # (TTM 295% mà quý 225% chênh 70 điểm, nhưng 225% nào có phải xấu đi).
        if q >= 10:
            return True
        return None if gap is None else gap >= -15

    if key == "SUS":                      # số đẹp lặp lại nhiều năm, không phải một kỳ
        m = row.get("Moat Score")
        if m is None or (isinstance(m, float) and pd.isna(m)):
            return None
        return m in _MOAT_GOOD
    return None


def _tally(vals):
    """(số tiêu chí đạt, số tiêu chí thực sự áp dụng được)."""
    got = sum(1 for v in vals if v is True)
    app = sum(1 for v in vals if v is not None)
    return got, app


def _signal(fin_pct, real_pct, fin_app, real_app):
    """Tín hiệu gộp. Điểm mấu chốt: số liệu đẹp mà không chứng minh được
    thì bị gắn cờ đỏ, chứ không được xếp ngang với mã trung bình đều."""
    if fin_app < 3 or real_app < 3:
        return "⚪ THIẾU DỮ LIỆU"
    if fin_pct >= 0.80 and real_pct >= 0.80: return "🟢 TỐT & THỰC"
    # Ngưỡng 0.60 tương đương "trượt từ 3 tiêu chí kiểm chứng trở lên", ổn định
    # với mọi mẫu số. Ngưỡng 0.50 cũ để lọt các mã rơi đúng 50% (3/6, 2/4).
    if fin_pct >= 0.80 and real_pct <  0.60: return "🔴 CỜ ĐỎ"
    if fin_pct >= 0.60 and real_pct >= 0.60: return "🔵 KHÁ"
    if real_pct >= 0.80 and fin_pct <  0.50: return "🟠 THẬT NHƯNG YẾU"
    if fin_pct >= 0.50:                      return "🟡 THEO DÕI"
    return "⚫ LOẠI"


def score_stocks(df, moat_cache: dict = None):
    """Chấm cả hai nhóm và sinh tín hiệu gộp."""
    df = _derive_fields(df, moat_cache)

    for k in FIN_KEYS:
        df[f"FIN_{k}"] = df.apply(lambda r, _k=k: _chk_fin(r, _k), axis=1)
    for k in REAL_KEYS:
        df[f"REAL_{k}"] = df.apply(lambda r, _k=k: _chk_real(r, _k), axis=1)

    fin  = df[[f"FIN_{k}"  for k in FIN_KEYS]].apply(lambda r: _tally(list(r)), axis=1)
    real = df[[f"REAL_{k}" for k in REAL_KEYS]].apply(lambda r: _tally(list(r)), axis=1)

    df["FIN_Score"]  = fin.apply(lambda x: x[0])
    df["FIN_App"]    = fin.apply(lambda x: x[1])
    df["REAL_Score"] = real.apply(lambda x: x[0])
    df["REAL_App"]   = real.apply(lambda x: x[1])

    # Chấm theo tỷ lệ trên số tiêu chí THỰC SỰ áp dụng được — mã thiếu dữ liệu
    # và nhóm Tài chính không bị phạt oan bởi mẫu số cố định.
    df["FIN_Pct"]  = df.apply(
        lambda r: round(r["FIN_Score"] / r["FIN_App"], 3) if r["FIN_App"] else 0.0, axis=1)
    df["REAL_Pct"] = df.apply(
        lambda r: round(r["REAL_Score"] / r["REAL_App"], 3) if r["REAL_App"] else 0.0, axis=1)

    df["Signal"] = df.apply(
        lambda r: _signal(r["FIN_Pct"], r["REAL_Pct"], r["FIN_App"], r["REAL_App"]), axis=1)
    return df


# ═══════════════════════════════════════════════════════════════
# FETCH
# ═══════════════════════════════════════════════════════════════
def fetch_data(market, top):
    print(f"  ⏳ Query TradingView ({market.upper()}, top {top} MCap)...", end="", flush=True)
    try:
        _, df = (Query().set_markets(market.lower()).select(*FETCH_COLS)
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
def write_excel(df, path, market, top, use_yf=False, progress_cb=None):
    def _pb(pct, msg=""):
        if progress_cb:
            progress_cb(pct, msg)

    wb = Workbook()
    _pb(5,  "Building main sheet…")
    _main_sheet(wb, df, market, top)

    _pb(20, "Sheet Cờ đỏ…")
    _qc_sheet(wb, df)

    _US_MARKETS = {"america", "nasdaq", "nyse"}
    can_yf = use_yf and market.lower() in _US_MARKETS
    if can_yf:
        n = len(df)
        def _fin_cb(idx):
            _pb(25 + int(idx / max(n, 1) * 45),
                f"Financials  {idx}/{n}…")
        _financials_sheet(wb, df, progress_cb=_fin_cb)
    else:
        if use_yf and market.lower() not in _US_MARKETS:
            print(f"  ⚠  yfinance không hỗ trợ {market.upper()} — bỏ qua sheet Financials")
        _pb(25, "")

    _pb(75, "Dashboard sheet…")
    _dashboard_sheet(wb, df)
    _pb(95, "Saving file…")
    wb.save(path)
    _pb(100, "")
    print(f"  💾 {path}")


def _financials_sheet(wb, df, progress_cb=None):
    import yfinance as yf
    from openpyxl.chart import BarChart, Reference

    ws = wb.create_sheet("Financials")
    ws.sheet_view.showGridLines = False
    ws.sheet_tab_color = "0E6655"

    N_QTR, N_YR = 8, 4
    C_LBL  = 1
    C_QTR0 = 2
    C_SEP  = C_QTR0 + N_QTR      # col 10
    C_YR0  = C_SEP + 1            # col 11
    LAST_C = C_YR0 + N_YR - 1    # col 14
    C_GAP  = LAST_C + 1           # col 15 — gap before chart area
    C_CH1  = C_GAP + 1            # col 16 — Revenue quarterly
    C_CH2  = C_CH1 + 7            # col 23 — Revenue annual
    C_CH3  = C_CH2 + 7            # col 30 — EPS quarterly
    C_CH4  = C_CH3 + 7            # col 37 — EPS annual
    C_CH5  = C_CH4 + 7            # col 44 — ROE quarterly
    C_CH6  = C_CH5 + 7            # col 51 — ROE annual
    C_CH7  = C_CH6 + 7            # col 58 — CF/Share quarterly
    C_CH8  = C_CH7 + 7            # col 65 — CF/Share annual
    CHART_SPACER = 6
    CHART_H, CHART_W = 4.5, 9.0

    ws.column_dimensions[CL(C_LBL)].width = 15
    for ci in range(C_QTR0, C_SEP):
        ws.column_dimensions[CL(ci)].width = 10
    ws.column_dimensions[CL(C_SEP)].width = 2
    for ci in range(C_YR0, LAST_C + 1):
        ws.column_dimensions[CL(ci)].width = 11
    ws.column_dimensions[CL(C_GAP)].width = 2

    # Title
    ws.merge_cells(f"A1:{CL(LAST_C)}1")
    c = ws.cell(1, 1, "FINANCIAL HISTORY  ·  REVENUE & EPS DILUTED  (nguồn: yfinance)")
    c.font = Font(name="Calibri", bold=True, size=13, color="FFFFFF")
    c.fill = BG("0D2137"); c.alignment = AL()
    ws.row_dimensions[1].height = 30

    # Sub-header: quarters vs years
    ws.merge_cells(f"{CL(C_QTR0)}2:{CL(C_SEP-1)}2")
    c2 = ws.cell(2, C_QTR0, f"◀  {N_QTR} QUÝ GẦN NHẤT  ▶")
    c2.font = F(True, 9, "FFFFFF"); c2.fill = BG("1B4F72"); c2.alignment = AL()
    ws.merge_cells(f"{CL(C_YR0)}2:{CL(LAST_C)}2")
    c3 = ws.cell(2, C_YR0, f"◀  {N_YR} NĂM GẦN NHẤT  ▶")
    c3.font = F(True, 9, "FFFFFF"); c3.fill = BG("0E6655"); c3.alignment = AL()
    ws.cell(2, C_LBL).fill = BG("0D2137")
    ws.cell(2, C_SEP).fill = BG("F0F4FA")
    ws.row_dimensions[2].height = 16

    cur = 3
    tickers = df["Ticker"].tolist()
    total   = len(tickers)

    for idx, ticker in enumerate(tickers, 1):
        if progress_cb: progress_cb(idx)
        print(f"  📊 Financials {idx}/{total}: {ticker}", end="\r", flush=True)
        name = str(df.loc[df["Ticker"] == ticker, "Tên Công Ty"].values[0]
                   if len(df.loc[df["Ticker"] == ticker]) else "")

        try:
            yft  = yf.Ticker(ticker)
            qfin = yft.quarterly_financials
            afin = yft.financials
            qbal = yft.quarterly_balance_sheet
            abal = yft.balance_sheet

            qtr_rev = (qfin.loc["Total Revenue"].sort_index().tail(N_QTR) / 1e9
                       if "Total Revenue" in qfin.index else pd.Series(dtype=float))
            qtr_eps = (qfin.loc["Diluted EPS"].sort_index().tail(N_QTR)
                       if "Diluted EPS" in qfin.index else pd.Series(dtype=float))
            ann_rev = (afin.loc["Total Revenue"].sort_index().dropna().tail(N_YR) / 1e9
                       if "Total Revenue" in afin.index else pd.Series(dtype=float))
            ann_eps = (afin.loc["Diluted EPS"].sort_index().dropna().tail(N_YR)
                       if "Diluted EPS" in afin.index else pd.Series(dtype=float))

            def _row(bs, *keys):
                for k in keys:
                    if k in bs.index: return bs.loc[k].sort_index()
                return pd.Series(dtype=float)

            qtr_ni  = _row(qfin, "Net Income", "Net Income Common Stockholders")
            ann_ni  = _row(afin, "Net Income", "Net Income Common Stockholders")
            qtr_eq  = _row(qbal, "Stockholders Equity", "Common Stock Equity",
                           "Total Equity Gross Minority Interest")
            ann_eq  = _row(abal, "Stockholders Equity", "Common Stock Equity",
                           "Total Equity Gross Minority Interest")

            def _roe_map(ni_s, eq_s, annualize=False):
                ni_sorted = ni_s.sort_index()
                eq_sorted = eq_s.sort_index()
                # TTM rolling: sum 4 quarters instead of multiplying 1 quarter × 4
                ni_use = ni_sorted.rolling(4, min_periods=1).sum() if annualize else ni_sorted
                common = ni_use.index.intersection(eq_sorted.index)
                out = {}
                for dt in common:
                    ni = float(ni_use.loc[dt]); eq = float(eq_sorted.loc[dt])
                    if pd.notna(ni) and pd.notna(eq) and eq != 0:
                        out[dt] = ni / abs(eq)
                return out

            qtr_roe = _roe_map(qtr_ni, qtr_eq.tail(N_QTR), annualize=True)  # full qtr_ni so rolling has enough history
            ann_roe = _roe_map(ann_ni.dropna().tail(N_YR), ann_eq.dropna().tail(N_YR))

            qcf = yft.quarterly_cashflow
            acf = yft.cashflow
            qtr_ocf = _row(qcf, "Operating Cash Flow",
                           "Cash Flow From Continuing Operating Activities")
            ann_ocf = _row(acf, "Operating Cash Flow",
                           "Cash Flow From Continuing Operating Activities")
            qtr_sh  = _row(qbal, "Ordinary Shares Number", "Share Issued",
                           "Common Stock Shares Outstanding")
            ann_sh  = _row(abal, "Ordinary Shares Number", "Share Issued",
                           "Common Stock Shares Outstanding")

            def _cfps_map(ocf_s, sh_s, tail_n):
                ocf_t = ocf_s.sort_index().tail(tail_n)
                common = ocf_t.index.intersection(sh_s.index)
                out = {}
                for dt in common:
                    ocf = float(ocf_t.loc[dt]); sh = float(sh_s.loc[dt])
                    if pd.notna(ocf) and pd.notna(sh) and sh != 0:
                        out[dt] = ocf / sh
                return out

            qtr_cfps = _cfps_map(qtr_ocf, qtr_sh, N_QTR)
            ann_cfps = _cfps_map(ann_ocf.dropna(), ann_sh, N_YR)
        except Exception:
            continue

        stock_row = cur

        # ── Stock header ──────────────────────────────────────────
        ws.merge_cells(f"A{cur}:{CL(LAST_C)}{cur}")
        c = ws.cell(cur, 1, f"  {ticker}  —  {name}")
        c.font = F(True, 10, "FFFFFF"); c.fill = BG("1B3A5C"); c.alignment = AL("left")
        ws.row_dimensions[cur].height = 18; cur += 1

        # ── Column headers: dates ─────────────────────────────────
        col_hdr_row = cur
        ws.cell(cur, C_LBL, "Chỉ số").font      = F(True, 8, "FFFFFF")
        ws.cell(cur, C_LBL).fill                 = BG("2C4F7C")
        ws.cell(cur, C_LBL).alignment            = AL()
        ws.cell(cur, C_LBL).border               = BD()
        ws.cell(cur, C_SEP).fill                 = BG("F0F4FA")

        qtr_idx = qtr_rev.index if len(qtr_rev) else qtr_eps.index
        for i, dt in enumerate(qtr_idx):
            lbl = dt.strftime("%b'%y") if hasattr(dt, "strftime") else str(dt)[:7]
            c = ws.cell(cur, C_QTR0 + i, lbl)
            c.font = F(True, 8, "FFFFFF"); c.fill = BG("1B4F72")
            c.alignment = AL(); c.border = BD()

        yr_idx = ann_rev.index if len(ann_rev) else ann_eps.index
        for i, dt in enumerate(yr_idx):
            lbl = f"FY{dt.year}" if hasattr(dt, "year") else str(dt)[:4]
            c = ws.cell(cur, C_YR0 + i, lbl)
            c.font = F(True, 8, "FFFFFF"); c.fill = BG("0E6655")
            c.alignment = AL(); c.border = BD()

        ws.row_dimensions[cur].height = 16; cur += 1

        # ── Revenue row ───────────────────────────────────────────
        rev_row = cur
        c = ws.cell(cur, C_LBL, "Revenue ($B)")
        c.font = F(True, 9, "1B4F72"); c.fill = BG("EBF5FB")
        c.alignment = AL("left"); c.border = BD()
        ws.cell(cur, C_SEP).fill = BG("F0F4FA")

        for i, v in enumerate(qtr_rev.values):
            cell = ws.cell(cur, C_QTR0 + i, round(float(v), 2) if pd.notna(v) else None)
            cell.font = F(size=9, color="1B4F72"); cell.fill = BG("EBF5FB")
            cell.alignment = AL(); cell.border = BD()
            if pd.notna(v): cell.number_format = "#,##0.00"
        for i, v in enumerate(ann_rev.values):
            cell = ws.cell(cur, C_YR0 + i, round(float(v), 2) if pd.notna(v) else None)
            cell.font = F(True, 9, "1B4F72"); cell.fill = BG("DDEEFF")
            cell.alignment = AL(); cell.border = BD()
            if pd.notna(v): cell.number_format = "#,##0.00"

        ws.row_dimensions[cur].height = 17; cur += 1

        # ── EPS row ───────────────────────────────────────────────
        eps_row = cur
        c = ws.cell(cur, C_LBL, "EPS Diluted ($)")
        c.font = F(True, 9, "276221"); c.fill = BG("E8F8F5")
        c.alignment = AL("left"); c.border = BD()
        ws.cell(cur, C_SEP).fill = BG("F0F4FA")

        for i, v in enumerate(qtr_eps.values):
            fv = float(v) if pd.notna(v) else None
            cell = ws.cell(cur, C_QTR0 + i, fv)
            cell.font = F(size=9, color="276221" if (fv or 0) >= 0 else "9C0006")
            cell.fill = BG("E8F8F5"); cell.alignment = AL(); cell.border = BD()
            if fv is not None: cell.number_format = '$#,##0.00'
        for i, v in enumerate(ann_eps.values):
            fv = float(v) if pd.notna(v) else None
            cell = ws.cell(cur, C_YR0 + i, fv)
            cell.font = F(True, 9, "276221" if (fv or 0) >= 0 else "9C0006")
            cell.fill = BG("C6EFCE"); cell.alignment = AL(); cell.border = BD()
            if fv is not None: cell.number_format = '$#,##0.00'

        ws.row_dimensions[cur].height = 17; cur += 1

        # ── ROE row ───────────────────────────────────────────────
        roe_row = cur
        c = ws.cell(cur, C_LBL, "ROE (annualized)")
        c.font = F(True, 9, "4A235A"); c.fill = BG("F4ECF7")
        c.alignment = AL("left"); c.border = BD()
        ws.cell(cur, C_SEP).fill = BG("F0F4FA")

        qtr_idx = qtr_rev.index if len(qtr_rev) else qtr_eps.index
        for i, dt in enumerate(qtr_idx):
            if dt in qtr_roe:
                v = qtr_roe[dt]
                cell = ws.cell(cur, C_QTR0 + i, round(v, 4))
                cell.font = F(size=9, color="4A235A" if v >= 0 else "9C0006")
                cell.fill = BG("F4ECF7"); cell.alignment = AL(); cell.border = BD()
                cell.number_format = '0.0%'

        yr_idx = ann_rev.index if len(ann_rev) else ann_eps.index
        for i, dt in enumerate(yr_idx):
            if dt in ann_roe:
                v = ann_roe[dt]
                cell = ws.cell(cur, C_YR0 + i, round(v, 4))
                cell.font = F(True, 9, "4A235A" if v >= 0 else "9C0006")
                cell.fill = BG("E8DAEF"); cell.alignment = AL(); cell.border = BD()
                cell.number_format = '0.0%'

        ws.row_dimensions[cur].height = 17; cur += 1

        # ── CF/Share row ──────────────────────────────────────────
        cfps_row = cur
        c = ws.cell(cur, C_LBL, "CF/Share ($)")
        c.font = F(True, 9, "0E6655"); c.fill = BG("E8F6F3")
        c.alignment = AL("left"); c.border = BD()
        ws.cell(cur, C_SEP).fill = BG("F0F4FA")

        qtr_idx = qtr_rev.index if len(qtr_rev) else qtr_eps.index
        for i, dt in enumerate(qtr_idx):
            if dt in qtr_cfps:
                v = qtr_cfps[dt]
                cell = ws.cell(cur, C_QTR0 + i, round(v, 2))
                cell.font = F(size=9, color="0E6655" if v >= 0 else "9C0006")
                cell.fill = BG("E8F6F3"); cell.alignment = AL(); cell.border = BD()
                cell.number_format = '$#,##0.00'

        yr_idx = ann_rev.index if len(ann_rev) else ann_eps.index
        for i, dt in enumerate(yr_idx):
            if dt in ann_cfps:
                v = ann_cfps[dt]
                cell = ws.cell(cur, C_YR0 + i, round(v, 2))
                cell.font = F(True, 9, "0E6655" if v >= 0 else "9C0006")
                cell.fill = BG("D0EDE8"); cell.alignment = AL(); cell.border = BD()
                cell.number_format = '$#,##0.00'

        ws.row_dimensions[cur].height = 17; cur += 1

        # ── Revenue QoQ % row ─────────────────────────────────────
        c = ws.cell(cur, C_LBL, "Rev QoQ %")
        c.font = F(True, 8, "155680"); c.fill = BG("D6EEF8")
        c.alignment = AL("left"); c.border = BD()
        ws.cell(cur, C_SEP).fill = BG("F0F4FA")

        rv = list(qtr_rev.values)
        for i, v in enumerate(rv):
            if i == 0 or not pd.notna(v) or not pd.notna(rv[i-1]) or rv[i-1] == 0:
                continue
            pct = (v - rv[i-1]) / abs(rv[i-1])
            cell = ws.cell(cur, C_QTR0 + i, round(pct, 4))
            cell.font = F(size=8, color="276221" if pct >= 0 else "9C0006")
            cell.fill = BG("D6EEF8"); cell.alignment = AL(); cell.border = BD()
            cell.number_format = '0.0%'
        arv = list(ann_rev.values)
        for i, v in enumerate(arv):
            if i == 0 or not pd.notna(v) or not pd.notna(arv[i-1]) or arv[i-1] == 0:
                continue
            pct = (v - arv[i-1]) / abs(arv[i-1])
            cell = ws.cell(cur, C_YR0 + i, round(pct, 4))
            cell.font = F(True, 8, "276221" if pct >= 0 else "9C0006")
            cell.fill = BG("C4E0F0"); cell.alignment = AL(); cell.border = BD()
            cell.number_format = '0.0%'

        ws.row_dimensions[cur].height = 16; cur += 1

        # ── EPS QoQ % row ─────────────────────────────────────────
        c = ws.cell(cur, C_LBL, "EPS QoQ %")
        c.font = F(True, 8, "276221"); c.fill = BG("D5F5E3")
        c.alignment = AL("left"); c.border = BD()
        ws.cell(cur, C_SEP).fill = BG("F0F4FA")

        ev = list(qtr_eps.values)
        for i, v in enumerate(ev):
            if i == 0 or not pd.notna(v) or not pd.notna(ev[i-1]) or ev[i-1] == 0:
                continue
            pct = (v - ev[i-1]) / abs(ev[i-1])
            cell = ws.cell(cur, C_QTR0 + i, round(pct, 4))
            cell.font = F(size=8, color="276221" if pct >= 0 else "9C0006")
            cell.fill = BG("D5F5E3"); cell.alignment = AL(); cell.border = BD()
            cell.number_format = '0.0%'
        aev = list(ann_eps.values)
        for i, v in enumerate(aev):
            if i == 0 or not pd.notna(v) or not pd.notna(aev[i-1]) or aev[i-1] == 0:
                continue
            pct = (v - aev[i-1]) / abs(aev[i-1])
            cell = ws.cell(cur, C_YR0 + i, round(pct, 4))
            cell.font = F(True, 8, "276221" if pct >= 0 else "9C0006")
            cell.fill = BG("A9DFBF"); cell.alignment = AL(); cell.border = BD()
            cell.number_format = '0.0%'

        ws.row_dimensions[cur].height = 16; cur += 1

        # ── Spacer (chart breathing room) ────────────────────────
        for _ in range(CHART_SPACER):
            ws.row_dimensions[cur].height = 8; cur += 1

        def _bar(title, data_row, cat_col0, cat_col1, anchor, y_fmt='#,##0.00'):
            ch = BarChart()
            ch.type = "col"; ch.grouping = "clustered"; ch.style = 10
            ch.title = title; ch.legend = None
            ch.y_axis.numFmt = y_fmt
            ch.add_data(Reference(ws, min_col=cat_col0, max_col=cat_col1,
                                  min_row=data_row, max_row=data_row), from_rows=True)
            ch.set_categories(Reference(ws, min_col=cat_col0, max_col=cat_col1,
                                        min_row=col_hdr_row))
            ch.width = CHART_W; ch.height = CHART_H
            ch.anchor = anchor
            ws.add_chart(ch)

        q_rev = len(qtr_rev); a_rev = len(ann_rev)
        q_eps = len(qtr_eps); a_eps = len(ann_eps)

        if q_rev > 0:
            _bar(f"{ticker}  |  Doanh thu quý ($B)",
                 rev_row, C_QTR0, C_QTR0+q_rev-1, f"{CL(C_CH1)}{stock_row}")
        if a_rev > 0:
            _bar(f"{ticker}  |  Doanh thu năm ($B)",
                 rev_row, C_YR0, C_YR0+a_rev-1, f"{CL(C_CH2)}{stock_row}")
        if q_eps > 0:
            _bar(f"{ticker}  |  EPS quý ($)",
                 eps_row, C_QTR0, C_QTR0+q_eps-1, f"{CL(C_CH3)}{stock_row}")
        if a_eps > 0:
            _bar(f"{ticker}  |  EPS năm ($)",
                 eps_row, C_YR0, C_YR0+a_eps-1, f"{CL(C_CH4)}{stock_row}")
        if qtr_roe:
            _bar(f"{ticker}  |  ROE quý (ann.)",
                 roe_row, C_QTR0, C_QTR0+N_QTR-1, f"{CL(C_CH5)}{stock_row}", y_fmt='0%')
        if ann_roe:
            _bar(f"{ticker}  |  ROE năm",
                 roe_row, C_YR0, C_YR0+N_YR-1, f"{CL(C_CH6)}{stock_row}", y_fmt='0%')
        if qtr_cfps:
            _bar(f"{ticker}  |  CF/Share quý ($)",
                 cfps_row, C_QTR0, C_QTR0+N_QTR-1, f"{CL(C_CH7)}{stock_row}")
        if ann_cfps:
            _bar(f"{ticker}  |  CF/Share năm ($)",
                 cfps_row, C_YR0, C_YR0+N_YR-1, f"{CL(C_CH8)}{stock_row}")

    print(f"  📊 Financials done ({total} stocks)          ")


def _sig_style(sig):
    return SIGNAL_STYLE.get(sig, (C_NAVY, C_WHITE))


def _score_fmt(got, app):
    try:
        got = int(got); app = int(app)
    except (TypeError, ValueError):
        return "—", 0.0
    if not app:
        return "—", 0.0
    return f"{got}/{app}", got / app


def _score_fill(pct):
    if pct >= 0.8:  return CG_FG, CG_BG
    if pct >= 0.6:  return CB_FG, CB_BG
    if pct >= 0.4:  return CY_FG, CY_BG
    return CR_FG, CR_BG


def _main_sheet(wb, df, market, top):
    """Bảng chính: số liệu gốc + 6 tiêu chí nhóm A + 6 tiêu chí nhóm B + hai điểm."""
    ws = wb.active
    ws.title = "Screener"
    ws.sheet_view.showGridLines = False

    BASE = [("Ticker", "Ticker", 10), ("Tên Công Ty", "Công ty", 30),
            ("Sector", "Ngành", 20), ("Price ($)", "Giá", 11),
            ("MCap ($B)", "Vốn hóa (tỷ$)", 14),
            ("Doanh thu ($B)", "Doanh thu (tỷ)", 14),
            ("Lợi nhuận ($B)", "Lợi nhuận (tỷ)", 14),
            ("FCF ($B)", "FCF (tỷ)", 12),
            ("DThu quý ($B)", "DThu quý (tỷ)", 13),
            ("LN quý ($B)", "LN quý (tỷ)", 12),
            ("ROIC%", "ROIC%", 9), ("Op Margin%", "Biên HĐ%", 10),
            ("Gross Margin%", "Biên gộp%", 10), ("NI Growth%", "LN tăng%", 10),
            ("Rev Growth%", "DT tăng%", 10), ("NI Qtr%", "LN quý%", 10),
            ("Rev Qtr%", "DT quý%", 10), ("NI_Trend_Gap", "Đà LN", 9),
            ("D/E", "D/E", 8),
            ("FCF/NI", "FCF/LN", 9), ("FCF_Margin%", "Biên FCF%", 10),
            ("EPS_NI_Gap", "EPS−LN", 9), ("ROE_ROIC_Gap", "ROE−ROIC", 10),
            ("Net Cash ($B)", "Tiền ròng", 11), ("Current Ratio", "T.khoản", 9),
            ("Nợ ròng/FCF", "Nợ/FCF", 9)]
    PCT = {"ROIC%", "Op Margin%", "Gross Margin%", "NI Growth%", "Rev Growth%",
           "FCF_Margin%", "EPS_NI_Gap", "ROE_ROIC_Gap",
           "NI Qtr%", "Rev Qtr%", "NI_Trend_Gap"}

    nb   = len(BASE)
    a_st = nb + 1
    b_st = a_st + len(FIN_KEYS)
    s_st = b_st + len(REAL_KEYS)
    total_cols = s_st + 2

    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=total_cols)
    t = ws.cell(1, 1, f"SCREENER TÀI CHÍNH  —  {market.upper()}  ·  TOP {top}  ·  "
                      f"{datetime.now():%d/%m/%Y %H:%M}")
    t.font = F(True, 13, C_WHITE); t.fill = BG(C_DARK); t.alignment = AL("left")
    ws.row_dimensions[1].height = 26

    for a, b, label, color in ((1, nb, "SỐ LIỆU", "1B3A5C"),
                               (a_st, b_st - 1, "A — TÀI CHÍNH TỐT", "1A5C2B"),
                               (b_st, s_st - 1, "B — CHỨNG MINH LÀ THỰC", "7A3E00"),
                               (s_st, total_cols, "KẾT QUẢ", "4A235A")):
        ws.merge_cells(start_row=2, start_column=a, end_row=2, end_column=b)
        c = ws.cell(2, a, label)
        c.font = F(True, 10, C_WHITE); c.fill = BG(color); c.alignment = AL()
    ws.row_dimensions[2].height = 19

    hdrs = ([h for _, h, _ in BASE]
            + [FIN_CRITERIA[k]["label"] for k in FIN_KEYS]
            + [REAL_CRITERIA[k]["label"] for k in REAL_KEYS]
            + ["Tài chính", "Thực", "Tín hiệu"])
    for i, h in enumerate(hdrs, 1):
        c = ws.cell(3, i, h)
        c.font = F(True, 9, C_WHITE); c.fill = BG(C_NAVY)
        c.alignment = AL(wrap=True); c.border = BD()
    ws.row_dimensions[3].height = 30

    sort_cols = [c for c in ("FIN_Pct", "REAL_Pct") if c in df.columns]
    d = df.sort_values(sort_cols, ascending=False) if sort_cols else df

    for ri, (_, row) in enumerate(d.iterrows()):
        r  = ri + 4
        bg = C_ALT if ri % 2 else C_WHITE
        for ci, (key, _, _) in enumerate(BASE, 1):
            v = row.get(key)
            c = ws.cell(r, ci)
            if v is None or (isinstance(v, float) and pd.isna(v)):
                c.value = "—"
            elif key in PCT:
                c.value = float(v) / 100; c.number_format = '+0.0%;(0.0%);"—"'
            elif key == "MCap ($B)":
                c.value = float(v); c.number_format = '#,##0.0'
            elif key == "Price ($)":
                c.value = float(v); c.number_format = '#,##0.00'
            elif isinstance(v, (int, float)):
                c.value = float(v); c.number_format = '0.00'
            else:
                c.value = str(v)
            c.font = F(size=9, bold=(ci == 1)); c.fill = BG(bg); c.border = BD()
            c.alignment = AL("left" if ci in (1, 2, 3) else "right")

        for gi, (keys, prefix) in enumerate(((FIN_KEYS, "FIN"), (REAL_KEYS, "REAL"))):
            start = a_st if gi == 0 else b_st
            for ki, k in enumerate(keys):
                v = row.get(f"{prefix}_{k}")
                c = ws.cell(r, start + ki)
                if v is True:
                    c.value = "✓"; c.font = F(True, 11, CG_FG); c.fill = BG(CG_BG)
                elif v is False:
                    c.value = "✗"; c.font = F(True, 11, CR_FG); c.fill = BG(CR_BG)
                else:
                    c.value = "—"; c.font = F(size=9, color="94A3B8"); c.fill = BG(bg)
                c.alignment = AL(); c.border = BD()

        for off, (g, a) in enumerate(((row.get("FIN_Score"), row.get("FIN_App")),
                                      (row.get("REAL_Score"), row.get("REAL_App")))):
            txt, pct = _score_fmt(g, a)
            fg, fbg  = _score_fill(pct)
            c = ws.cell(r, s_st + off, txt)
            c.font = F(True, 10, fg); c.fill = BG(fbg); c.alignment = AL(); c.border = BD()

        sig = row.get("Signal", "") or ""
        fg, fbg = _sig_style(sig)
        c = ws.cell(r, s_st + 2, sig)
        c.font = F(True, 9, fg); c.fill = BG(fbg); c.alignment = AL(); c.border = BD()

    for i, (_, _, w) in enumerate(BASE, 1):
        ws.column_dimensions[CL(i)].width = w
    for i in range(a_st, s_st):
        ws.column_dimensions[CL(i)].width = 11
    ws.column_dimensions[CL(s_st)].width     = 11
    ws.column_dimensions[CL(s_st + 1)].width = 11
    ws.column_dimensions[CL(s_st + 2)].width = 20
    ws.freeze_panes = "B4"
    ws.auto_filter.ref = f"A3:{CL(total_cols)}{len(d) + 3}"


def _qc_sheet(wb, df):
    """Sheet Cờ đỏ — số liệu đẹp nhưng không vượt được kiểm chứng."""
    ws = wb.create_sheet("Cờ đỏ")
    ws.sheet_view.showGridLines = False

    HDRS = ["Ticker", "Công ty", "Ngành", "Tài chính", "Thực", "Tín hiệu",
            "Chưa chứng minh được ở điểm nào"]
    WIDS = [10, 30, 20, 11, 11, 20, 62]

    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(HDRS))
    t = ws.cell(1, 1, "CỜ ĐỎ — SỐ LIỆU TÀI CHÍNH ĐẸP NHƯNG CHƯA CHỨNG MINH ĐƯỢC")
    t.font = F(True, 12, C_WHITE); t.fill = BG("9C0006"); t.alignment = AL("left")
    ws.row_dimensions[1].height = 24

    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=len(HDRS))
    s = ws.cell(2, 1, "Đạt ≥80% tiêu chí tài chính nhưng dưới 50% tiêu chí kiểm chứng. "
                      "Đây là nhóm mà bộ lọc được thiết kế để phát hiện.")
    s.font = F(False, 9, "5C0000", italic=True); s.fill = BG("FFF0F0"); s.alignment = AL("left")

    for i, (h, w) in enumerate(zip(HDRS, WIDS), 1):
        c = ws.cell(3, i, h)
        c.font = F(True, 9, C_WHITE); c.fill = BG(C_NAVY)
        c.alignment = AL(); c.border = BD()
        ws.column_dimensions[CL(i)].width = w
    ws.row_dimensions[3].height = 20

    _WHY = {
        "CASH": "dòng tiền tự do không theo kịp lợi nhuận ghi sổ",
        "FCFM": "biên dòng tiền tự do mỏng",
        "BUYB": "EPS tăng nhanh hơn lợi nhuận — do mua cổ phiếu quỹ",
        "LEV":  "ROE vượt xa ROIC — sinh lời đến từ đòn bẩy",
        "BAL":  "nợ ròng lớn mà dòng tiền không đủ trả trong 3 năm",
        "TREND": "quý gần nhất đã xấu đi — số TTM đẹp chủ yếu nhờ các quý cũ",
        "SUS":  "chưa có bề dày 5 năm để gọi là bền vững",
    }

    flag = df[df["Signal"] == "🔴 CỜ ĐỎ"] if "Signal" in df.columns else df.iloc[0:0]
    flag = flag.sort_values("FIN_Pct", ascending=False) if "FIN_Pct" in flag.columns else flag

    if flag.empty:
        ws.merge_cells(start_row=4, start_column=1, end_row=4, end_column=len(HDRS))
        c = ws.cell(4, 1, "Không có mã nào rơi vào nhóm này trong lần quét hiện tại.")
        c.font = F(False, 10, "276221", italic=True); c.alignment = AL("left")
        return

    for ri, (_, row) in enumerate(flag.iterrows()):
        r  = ri + 4
        bg = C_ALT if ri % 2 else C_WHITE
        why = "; ".join(_WHY[k] for k in REAL_KEYS if row.get(f"REAL_{k}") is False) or "—"
        vals = [row.get("Ticker", ""), row.get("Tên Công Ty", ""), row.get("Sector", ""),
                _score_fmt(row.get("FIN_Score"), row.get("FIN_App"))[0],
                _score_fmt(row.get("REAL_Score"), row.get("REAL_App"))[0],
                row.get("Signal", ""), why]
        for ci, v in enumerate(vals, 1):
            c = ws.cell(r, ci, v)
            c.border = BD(); c.fill = BG(bg)
            c.alignment = AL("left" if ci in (1, 2, 3, 7) else "center", wrap=(ci == 7))
            if ci == 4:
                c.font = F(True, 10, CG_FG); c.fill = BG(CG_BG)
            elif ci == 5:
                c.font = F(True, 10, CR_FG); c.fill = BG(CR_BG)
            elif ci == 6:
                c.font = F(True, 9, CR_FG); c.fill = BG(CR_BG)
            elif ci == 7:
                c.font = F(False, 9, "9C0006")
            else:
                c.font = F(ci == 1, 9)
    ws.freeze_panes = "A4"


def _dashboard_sheet(wb, df):
    """Sheet tổng hợp: phân bố tín hiệu, tỷ lệ đạt từng tiêu chí, theo ngành."""
    ws = wb.create_sheet("Tổng hợp")
    ws.sheet_view.showGridLines = False
    for i, w in enumerate([26, 13, 13, 13, 46], 1):
        ws.column_dimensions[CL(i)].width = w

    ws.merge_cells("A1:E1")
    t = ws.cell(1, 1, f"TỔNG HỢP  ·  {len(df)} mã  ·  {datetime.now():%d/%m/%Y %H:%M}")
    t.font = F(True, 13, C_WHITE); t.fill = BG(C_DARK); t.alignment = AL("left")
    ws.row_dimensions[1].height = 26

    r = 3

    def section(title, color="1B3A5C"):
        nonlocal r
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=5)
        c = ws.cell(r, 1, title)
        c.font = F(True, 10, C_WHITE); c.fill = BG(color); c.alignment = AL("left")
        r += 1

    def head(cols):
        nonlocal r
        for i, h in enumerate(cols, 1):
            c = ws.cell(r, i, h)
            c.font = F(True, 9, C_WHITE); c.fill = BG(C_NAVY)
            c.alignment = AL(); c.border = BD()
        r += 1

    section("PHÂN BỐ TÍN HIỆU")
    head(["Tín hiệu", "Số mã", "Tỷ lệ", "", ""])
    sig = df["Signal"] if "Signal" in df.columns else pd.Series(dtype=str)
    for name in ["🟢 TỐT & THỰC", "🔵 KHÁ", "🔴 CỜ ĐỎ", "🟠 THẬT NHƯNG YẾU",
                 "🟡 THEO DÕI", "⚫ LOẠI", "⚪ THIẾU DỮ LIỆU"]:
        n = int((sig == name).sum())
        fg, bg = _sig_style(name)
        ws.cell(r, 1, name).font = F(True, 9, fg)
        ws.cell(r, 1).fill = BG(bg); ws.cell(r, 1).border = BD()
        c = ws.cell(r, 2, n); c.font = F(True, 10); c.alignment = AL(); c.border = BD()
        p = ws.cell(r, 3, n / len(df) if len(df) else 0)
        p.number_format = '0.0%'; p.alignment = AL(); p.border = BD()
        r += 1
    r += 1

    section("TỶ LỆ ĐẠT CỦA TỪNG TIÊU CHÍ", "1A5C2B")
    head(["Tiêu chí", "Nhóm", "Đạt / Áp dụng", "Tỷ lệ", "Nhận xét"])
    for grp, keys, crit, pre in (("A", FIN_KEYS, FIN_CRITERIA, "FIN"),
                                 ("B", REAL_KEYS, REAL_CRITERIA, "REAL")):
        for k in keys:
            col = f"{pre}_{k}"
            if col not in df.columns:
                continue
            s   = df[col]
            app = int(s.notna().sum())
            got = int((s == True).sum())  # noqa: E712
            pct = got / app if app else 0
            if app == 0:
                note, nc = "không có dữ liệu — tiêu chí đang vô hiệu", CR_FG
            elif pct >= 0.9:
                note, nc = "quá lỏng — hầu như mã nào cũng đạt", CY_FG
            elif pct <= 0.05:
                note, nc = "quá chặt — gần như không mã nào đạt", CY_FG
            else:
                note, nc = "phân loại tốt", CG_FG
            ws.cell(r, 1, crit[k]["label"]).font = F(True, 9)
            ws.cell(r, 2, f"Nhóm {grp}").font = F(size=9)
            ws.cell(r, 3, f"{got} / {app}").font = F(size=9)
            pc = ws.cell(r, 4, pct); pc.number_format = '0.0%'; pc.font = F(True, 9)
            ws.cell(r, 5, note).font = F(size=9, color=nc)
            for i in range(1, 6):
                ws.cell(r, i).border = BD()
                ws.cell(r, i).alignment = AL("left" if i in (1, 5) else "center")
            r += 1
    r += 1

    section("THEO NGÀNH", "7A3E00")
    head(["Ngành", "Số mã", "Tốt & Thực", "Cờ đỏ", "TB Tài chính / TB Thực"])
    if "Sector" in df.columns:
        stats = []
        for name, g in df.groupby("Sector"):
            if not name:
                continue
            gs = g["Signal"] if "Signal" in g.columns else pd.Series(dtype=str)
            stats.append((str(name), len(g),
                          int((gs == "🟢 TỐT & THỰC").sum()),
                          int((gs == "🔴 CỜ ĐỎ").sum()),
                          g["FIN_Pct"].mean() if "FIN_Pct" in g.columns else 0,
                          g["REAL_Pct"].mean() if "REAL_Pct" in g.columns else 0))
        for name, n, good, bad, fp, rp in sorted(stats, key=lambda x: -x[2]):
            ws.cell(r, 1, name).font = F(True, 9)
            ws.cell(r, 2, n).font = F(size=9)
            ws.cell(r, 3, good).font = F(True, 9, CG_FG)
            ws.cell(r, 4, bad).font = F(True, 9, CR_FG)
            ws.cell(r, 5, f"{fp*100:.0f}%  /  {rp*100:.0f}%").font = F(size=9)
            for i in range(1, 6):
                ws.cell(r, i).border = BD()
                ws.cell(r, i).alignment = AL("left" if i in (1, 5) else "center")
            r += 1

    r += 1
    ws.merge_cells(start_row=r, start_column=1, end_row=r + 2, end_column=5)
    note = ws.cell(r, 1,
                   "Điểm chấm theo tỷ lệ trên số tiêu chí thực sự áp dụng được. "
                   "Nhóm A đo số liệu có đẹp không; nhóm B đo số liệu đó có thật không. "
                   "Điểm A cao mà điểm B thấp là cảnh báo, không phải trung bình.")
    note.font = F(False, 9, "64748B", italic=True)
    note.alignment = AL("left", wrap=True)


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

    # 3. Chấm điểm: nhóm A (tài chính tốt) + nhóm B (chứng minh là thực)
    print("  📊 Chấm điểm Tài chính + Kiểm chứng...", end="", flush=True)
    df = score_stocks(df, moat_cache)
    print(f" ✅  ({int((df['Signal'] == '🟢 TỐT & THỰC').sum())} TỐT & THỰC, "
          f"{int((df['Signal'] == '🔴 CỜ ĐỎ').sum())} CỜ ĐỎ)")

    # Hướng thị trường — chỉ in ra làm bối cảnh, không cộng vào điểm
    market_ok, index_1y = fetch_market_direction(a.market)
    if market_ok is not None:
        print(f"  🧭 Thị trường: {'THUẬN LỢI' if market_ok else 'BẤT LỢI'}"
              + (f"  ·  chỉ số 1 năm {index_1y:+.1f}%" if index_1y is not None else ""))

    # Preview
    prev = ["Ticker","Moat Score","ROIC%","FCF/NI","FIN_Score","REAL_Score","Signal"]
    print(df[[c for c in prev if c in df.columns]].head(10).to_string(index=False))
    print()

    print("  📝 Tạo Excel...", end="", flush=True)
    write_excel(df, out, a.market, a.top)
    print(f" ✅\n  ✨ Done → {out}\n")


if __name__ == "__main__":
    main()