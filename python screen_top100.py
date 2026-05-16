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
    "C": dict(label="C — EPS Qtr",   field="EPS Qtr%",      thr=25,  op="gt", desc="EPS Quarterly YoY > 25%"),
    "A": dict(label="A — EPS Ann",   field="EPS Annual%",   thr=20,  op="gt", desc="EPS Annual YoY > 20%"),
    "S": dict(label="S — Sales",     field="Rev Qtr%",      thr=20,  op="gt", desc="Revenue Quarterly YoY > 20%"),
    "L": dict(label="L — RS (1Y)",   field="1Y%",           thr=20,  op="gt", desc="1-Year Perf > 20% (RS proxy)"),
    "Q": dict(label="Quality GM",    field="Gross Margin%", thr=40,  op="gt", desc="Gross Margin > 40%"),
    "R": dict(label="ROE",           field="ROE%",          thr=17,  op="gt", desc="ROE > 17%"),
    "M": dict(label="Momentum 3M",   field="3M%",           thr=0,   op="gt", desc="3-Month Perf > 0%"),
    "D": dict(label="Debt OK",       field="D/E",           thr=2.0, op="lt", desc="D/E < 2.0"),
}
CS_KEYS = ["C", "A", "S", "L", "Q", "R", "M", "D"]
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
        return proxy, score

    except Exception as e:
        # Fallback: sector hardcode + UNCERTAIN
        proxy = MOAT_PROXY_MAP.get(sector, "Cost Advantage")
        return proxy, "UNCERTAIN ★"


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
        proxy, score = fetch_moat_yfinance(ticker, sector)
        cache[ticker] = (proxy, score)
        print(f" {score}")
        time.sleep(0.4)   # tránh rate limit
    print()
    return cache


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
def score_canslim(df, moat_cache: dict = None):
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
        if s >= 7: return "🟢 STRONG BUY"
        if s >= 5: return "🔵 BUY"
        if s >= 3: return "🟡 WATCH"
        return "🔴 SKIP"
    df["CS_Signal"] = df["CS_Score"].apply(sig)

    # ── MOAT: dùng yfinance cache nếu có, fallback sector ──────
    def apply_moat(row):
        ticker = row.get("Ticker", "")
        sector = row.get("Sector", "")
        if moat_cache and ticker in moat_cache:
            return moat_cache[ticker]   # (proxy, score) từ yfinance
        # Fallback: tính từ TV TTM data (giống cũ)
        gm  = row.get("Gross Margin%"); gm  = gm  if (gm  is not None and not (isinstance(gm,  float) and pd.isna(gm)))  else None
        roe = row.get("ROE%");          roe = roe if (roe is not None and not (isinstance(roe, float) and pd.isna(roe))) else None
        proxy = MOAT_PROXY_MAP.get(sector, "Cost Advantage")
        score = _moat_score_from_metrics(roe, gm, sector)
        return proxy, score

    moat = df.apply(apply_moat, axis=1)
    df["Moat Proxy"] = moat.apply(lambda x: x[0])
    df["Moat Score"] = moat.apply(lambda x: x[1])
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
    _legend_sheet(wb)
    _summary_sheet(wb, df)
    wb.save(path)
    print(f"  💾 {path}")


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


def _legend_sheet(wb):
    ws = wb.create_sheet("CAN SLIM Legend"); ws.sheet_view.showGridLines=False; bdr=BD()
    ws.merge_cells("A1:E1"); c=ws["A1"]
    c.value="📋  CAN SLIM — Tiêu Chí & Ngưỡng"
    c.font=F(True,13,C_WHITE); c.fill=BG(C_DARK); c.alignment=AL(); ws.row_dimensions[1].height=28
    for ci,h in enumerate(["Ký hiệu","Tên","Field","Điều kiện","Ý nghĩa"],1):
        c=ws.cell(2,ci,h); c.font=F(True,9,C_WHITE); c.fill=BG(C_NAVY); c.alignment=AL(); c.border=bdr
    ws.row_dimensions[2].height=20
    meanings={
        "C":"Current Earnings — lợi nhuận hiện tại phải tăng mạnh (≥25% YoY)",
        "A":"Annual Earnings — EPS tăng trưởng ổn định nhiều năm liên tiếp",
        "S":"Sales — doanh thu tăng, xác nhận EPS growth không phải từ cắt chi phí",
        "L":"Leader — cổ phiếu dẫn đầu ngành, 1Y Perf > 20% là RS proxy",
        "Q":"Quality — Gross Margin cao = lợi thế cạnh tranh / pricing power",
        "R":"ROE — hiệu quả sử dụng vốn, công ty tạo ra lợi nhuận tốt trên equity",
        "M":"Momentum — giá đang trong xu hướng tăng, không mua cổ phiếu đang giảm",
        "D":"Debt — nợ kiểm soát được, tránh công ty overleveraged",
    }
    ops={"gt":">","lt":"<"}
    for ri,key in enumerate(CS_KEYS,3):
        cfg=CANSLIM[key]; alt=ri%2==0
        for ci,val in enumerate([key,cfg["label"],cfg["field"],
                  f"{ops[cfg['op']]} {cfg['thr']}{'%' if cfg['thr']!=2.0 else ''}",
                  meanings[key]],1):
            c=ws.cell(ri,ci,val)
            c.font=F(size=9,bold=(ci==1),color=(CG_FG if ci==1 else "222222"))
            c.fill=BG(C_ALT if alt else C_WHITE); c.border=bdr
            c.alignment=AL("left" if ci>=3 else "center")
        ws.row_dimensions[ri].height=18
    ws.row_dimensions[11].height=6
    ws.merge_cells("A12:E12"); c=ws["A12"]
    c.value="SIGNAL LEGEND"; c.font=F(True,10,C_WHITE); c.fill=BG(C_NAVY); c.alignment=AL(); ws.row_dimensions[12].height=20
    for ri,(sig,desc,fg,bg) in enumerate([
        ("🟢 STRONG BUY","≥ 7/8 tiêu chí đạt",CG_FG,CG_BG),
        ("🔵 BUY","5–6/8 tiêu chí đạt",CB_FG,CB_BG),
        ("🟡 WATCH","3–4/8 tiêu chí đạt",CY_FG,CY_BG),
        ("🔴 SKIP","0–2/8 tiêu chí đạt",CR_FG,CR_BG)],13):
        ws.merge_cells(f"A{ri}:B{ri}"); c=ws.cell(ri,1,sig)
        c.font=F(True,10,fg); c.fill=BG(bg); c.alignment=AL(); c.border=bdr
        ws.merge_cells(f"C{ri}:E{ri}"); c=ws.cell(ri,3,desc)
        c.font=F(size=9); c.fill=BG(bg); c.alignment=AL("left"); c.border=bdr
        ws.row_dimensions[ri].height=18
    for ci,w in enumerate([8,16,16,14,45],1): ws.column_dimensions[CL(ci)].width=w
    ws.merge_cells("A18:E18"); c=ws["A18"]
    c.value="⚠️  Moat Proxy/Score lấy từ yfinance 5yr ROE avg + 5yr GM avg (giống screen_accumulate.py). N/I/M cần kiểm tra thủ công trên TradingView."
    c.font=F(False,8,italic=True,color="888888"); c.alignment=AL("left"); ws.row_dimensions[18].height=28


def _summary_sheet(wb, df):
    ws=wb.create_sheet("Summary"); ws.sheet_view.showGridLines=False; bdr=BD()
    ws.merge_cells("A1:F1"); c=ws["A1"]
    c.value="📊  SUMMARY"; c.font=F(True,13,C_WHITE); c.fill=BG(C_DARK); c.alignment=AL(); ws.row_dimensions[1].height=28
    ws["A3"]="Signal Distribution"; ws["A3"].font=F(True,10,C_NAVY); ws.row_dimensions[3].height=20
    for ci,h in enumerate(["Signal","Count","% of Total"],1):
        c=ws.cell(4,ci,h); c.font=F(True,9,C_WHITE); c.fill=BG(C_NAVY); c.border=bdr; c.alignment=AL()
    ws.row_dimensions[4].height=18
    total=len(df)
    for ri,(sig,fg,bg) in enumerate([
        ("🟢 STRONG BUY",CG_FG,CG_BG),("🔵 BUY",CB_FG,CB_BG),
        ("🟡 WATCH",CY_FG,CY_BG),("🔴 SKIP",CR_FG,CR_BG)],5):
        cnt=(df["CS_Signal"]==sig).sum()
        for ci,val in enumerate([sig,cnt,cnt/total if total else 0],1):
            c=ws.cell(ri,ci,val); c.font=F(size=9,color=fg,bold=(ci==1))
            c.fill=BG(bg); c.border=bdr; c.alignment=AL()
            if ci==3: c.number_format="0%"
        ws.row_dimensions[ri].height=18
    ws["A10"]="Metric Statistics"; ws["A10"].font=F(True,10,C_NAVY); ws.row_dimensions[10].height=20
    for ci,h in enumerate(["Metric","Median","Mean","Min","Max","% Pass"],1):
        c=ws.cell(11,ci,h); c.font=F(True,9,C_WHITE); c.fill=BG(C_NAVY); c.border=bdr; c.alignment=AL()
    ws.row_dimensions[11].height=18
    for ri,(metric,thr) in enumerate([
        ("EPS Annual%",20),("EPS Qtr%",25),("Rev Annual%",20),("Rev Qtr%",20),
        ("Gross Margin%",40),("Net Margin%",10),("ROE%",17),("3M%",0),("1Y%",20)],12):
        col=df.get(metric)
        if col is None: continue
        vals=col.dropna()
        if len(vals)==0: continue
        pp=(vals>thr).sum()/len(vals); alt=ri%2==0
        for ci,val in enumerate([metric,vals.median()/100,vals.mean()/100,
                                  vals.min()/100,vals.max()/100,pp],1):
            c=ws.cell(ri,ci,val); c.font=F(size=9)
            c.fill=BG(C_ALT if alt else C_WHITE); c.border=bdr
            c.alignment=AL("left" if ci==1 else "center")
            if ci>1: c.number_format="+0.0%;(0.0%)" if ci<=5 else "0%"
        ws.row_dimensions[ri].height=16
    for ci,w in enumerate([18,10,10,10,10,10],1): ws.column_dimensions[CL(ci)].width=w


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