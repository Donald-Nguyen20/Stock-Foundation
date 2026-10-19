# -*- coding: utf-8 -*-
"""Kiểm thử lùi thô cho bộ tiêu chí.

GIỚI HẠN — đọc trước khi tin bất kỳ con số nào ở đây:

1. Thiên lệch sống sót. Vũ trụ là các mã còn tồn tại và còn lớn tới hôm nay.
   Doanh nghiệp đã phá sản hoặc bị huỷ niêm yết không có mặt, nên mọi kết quả
   đều lạc quan hơn thực tế.
2. Mẫu nhỏ và cùng một chu kỳ. yfinance chỉ cho khoảng 4–5 năm báo cáo thường
   niên, tức mỗi mã vài điểm quan sát, tất cả nằm trong cùng một giai đoạn
   thị trường. Kết luận không suy rộng được sang chu kỳ khác.
3. Số liệu có thể đã bị trình bày lại. yfinance trả về con số hiện tại, không
   phải con số công bố lúc đó.
4. Độ trễ công bố lấy ước lượng 90 ngày sau ngày kết thúc năm tài chính.

Vì vậy đây là chỉ dấu để loại bỏ tiêu chí vô dụng, không phải bằng chứng
rằng tiêu chí còn lại là đúng.
"""
from __future__ import annotations

import pickle
import time
from pathlib import Path

import numpy as np
import pandas as pd

CACHE = Path(__file__).resolve().parent / "backtest_cache.pkl"
LAG_DAYS = 90          # báo cáo thường công bố trong vòng ~3 tháng
HOLD_DAYS = 365        # đo lợi suất 12 tháng kể từ ngày vào


def _col(df, *keys):
    for k in keys:
        if k in df.index:
            return df.loc[k]
    return pd.Series(dtype=float)


def fetch(tickers: list[str], pause: float = 0.25, progress=None) -> dict:
    """Tải báo cáo thường niên và giá cho từng mã. Có cache trên đĩa."""
    import yfinance as yf

    cache = {}
    if CACHE.is_file():
        try:
            cache = pickle.loads(CACHE.read_bytes())
        except Exception:
            cache = {}

    todo = [t for t in tickers if t not in cache]
    for i, t in enumerate(todo, 1):
        if progress:
            progress(f"[{i}/{len(todo)}] {t}")
        try:
            tk = yf.Ticker(t)
            rec = {
                "inc": tk.income_stmt,
                "bal": tk.balance_sheet,
                "cf": tk.cashflow,
                "px": tk.history(period="7y", interval="1d", auto_adjust=True)["Close"],
            }
            cache[t] = rec
        except Exception as e:
            cache[t] = {"error": str(e)[:120]}
        time.sleep(pause)
        if i % 25 == 0:
            CACHE.write_bytes(pickle.dumps(cache))
    CACHE.write_bytes(pickle.dumps(cache))
    return cache


def _metrics_at(inc, bal, cf, col) -> dict:
    """Các chỉ số của bộ lọc, tính từ báo cáo của một năm tài chính."""
    def g(df, *keys):
        s = _col(df, *keys)
        try:
            return float(s[col]) if col in s.index and pd.notna(s[col]) else None
        except Exception:
            return None

    rev = g(inc, "Total Revenue", "TotalRevenue", "Operating Revenue")
    gp = g(inc, "Gross Profit", "GrossProfit")
    op = g(inc, "Operating Income", "OperatingIncome", "EBIT")
    ni = g(inc, "Net Income", "NetIncome", "Net Income Common Stockholders")
    eq = g(bal, "Common Stock Equity", "Stockholders Equity", "Total Stockholders Equity")
    debt = g(bal, "Total Debt", "TotalDebt")
    ocf = g(cf, "Operating Cash Flow", "Total Cash From Operating Activities")
    capex = g(cf, "Capital Expenditure", "CapitalExpenditures")

    m = {}
    if rev and rev > 0:
        if gp is not None:
            m["gm"] = gp / rev * 100
        if op is not None:
            m["opm"] = op / rev * 100
        if ocf is not None:
            fcf = ocf + (capex or 0)
            m["fcfm"] = fcf / rev * 100
            if ni and ni > 0:
                m["fcfni"] = fcf / ni
    if ni is not None and eq and eq > 0:
        m["roe"] = ni / eq * 100
        cap = eq + (debt or 0)
        if cap > 0:
            m["roic"] = ni / cap * 100
    if eq and eq > 0 and debt is not None:
        m["de"] = debt / eq
    m["_rev"] = rev
    m["_ni"] = ni
    return m


def build_panel(cache: dict) -> pd.DataFrame:
    """Một dòng = một mã tại một năm tài chính, kèm lợi suất 12 tháng sau đó."""
    rows = []
    for tk, rec in cache.items():
        if not isinstance(rec, dict) or "error" in rec:
            continue
        inc, bal, cf, px = rec.get("inc"), rec.get("bal"), rec.get("cf"), rec.get("px")
        if inc is None or px is None or len(px) < 400:
            continue
        try:
            px = px.copy()
            px.index = pd.to_datetime(px.index).tz_localize(None)
        except Exception:
            continue

        cols = sorted([c for c in inc.columns], reverse=True)
        for idx, col in enumerate(cols):
            try:
                fy_end = pd.Timestamp(col).tz_localize(None)
            except Exception:
                continue
            entry = fy_end + pd.Timedelta(days=LAG_DAYS)
            exit_ = entry + pd.Timedelta(days=HOLD_DAYS)
            if exit_ > px.index.max():
                continue

            def px_at(d):
                s = px[px.index <= d]
                return float(s.iloc[-1]) if len(s) else None

            p0, p1 = px_at(entry), px_at(exit_)
            if not p0 or not p1 or p0 <= 0:
                continue

            m = _metrics_at(inc, bal, cf, col)
            if not m.get("roic") and not m.get("gm"):
                continue

            # tăng trưởng so với năm liền trước
            if idx + 1 < len(cols):
                prev = _metrics_at(inc, bal, cf, cols[idx + 1])
                if m.get("_rev") and prev.get("_rev"):
                    m["rev_g"] = (m["_rev"] / prev["_rev"] - 1) * 100
                if m.get("_ni") and prev.get("_ni") and prev["_ni"] > 0:
                    m["ni_g"] = (m["_ni"] / prev["_ni"] - 1) * 100
                if m.get("gm") is not None and prev.get("gm") is not None:
                    m["d_gm"] = m["gm"] - prev["gm"]

            # mức sụt so với đỉnh 52 tuần tại ngày vào
            win = px[(px.index > entry - pd.Timedelta(days=365)) & (px.index <= entry)]
            m["dd"] = (p0 / float(win.max()) - 1) * 100 if len(win) else None
            ma = px[(px.index > entry - pd.Timedelta(days=365)) & (px.index <= entry)]
            m["vs_ma"] = (p0 / float(ma.mean()) - 1) * 100 if len(ma) else None

            m["ticker"] = tk
            m["entry"] = entry
            m["year"] = entry.year
            m["fwd"] = (p1 / p0 - 1) * 100
            rows.append(m)

    df = pd.DataFrame(rows)
    if df.empty:
        return df
    # Lợi suất vượt trội so với trung vị cùng năm — tách ảnh hưởng chu kỳ
    df["excess"] = df["fwd"] - df.groupby("year")["fwd"].transform("median")
    return df
