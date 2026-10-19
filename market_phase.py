# -*- coding: utf-8 -*-
"""Đoán giai đoạn thị trường theo Wyckoff và Elliott.

ĐỘ TIN CẬY CỦA HAI PHƯƠNG PHÁP RẤT KHÁC NHAU — đọc trước khi dùng:

· Wyckoff cơ giới hoá được tương đối. Nó dựa vào quan hệ giữa giá và khối
  lượng, và những thứ như vùng đi ngang, cú rũ bỏ dưới đáy vùng, hay khối
  lượng lớn mà giá không đi được đều đo được bằng quy tắc. Kết quả ở đây là
  một cách đọc có căn cứ, không phải chân lý.

· Đếm sóng Elliott tự động thì KHÔNG đáng tin. Cùng một đồ thị luôn có nhiều
  cách đếm hợp lệ, và người hành nghề thật cũng bất đồng với nhau. Phần
  Elliott ở đây chỉ kiểm ba quy tắc cứng của sóng đẩy rồi nêu một giả thuyết
  kèm phương án thay thế. Hãy coi nó là gợi ý để tự nhìn lại đồ thị, không
  phải kết luận.

Cả hai đều mô tả hiện tại, không dự báo tương lai.
"""
from __future__ import annotations

import numpy as np
import pandas as pd

YF_INDEX = {
    "america":   ("^GSPC", "S&P 500"),
    "nasdaq":    ("^NDX",  "Nasdaq 100"),
    "nyse":      ("^NYA",  "NYSE Composite"),
    "euronext":  ("^FCHI", "CAC 40"),
    "hong_kong": ("^HSI",  "Hang Seng"),
}

WYCKOFF = {
    "markup":       ("ĐÁNH LÊN",    "#1A5C2B", "#C6EFCE"),
    "distribution": ("PHÂN PHỐI",   "#7A3E00", "#FFE0C2"),
    "markdown":     ("ĐÁNH XUỐNG",  "#9C0006", "#FFC7CE"),
    "accumulation": ("TÍCH LUỸ",    "#1B3A5C", "#DDEEFF"),
    "unclear":      ("CHƯA RÕ",     "#64748B", "#F1F5F9"),
}


# ─────────────────────────────────────────────────────────────────────────────
# Dữ liệu và cấu trúc sóng
# ─────────────────────────────────────────────────────────────────────────────
def fetch_ohlc(market: str = "america", years: int = 3) -> pd.DataFrame | None:
    import yfinance as yf
    sym = YF_INDEX.get(market.lower(), ("^GSPC", ""))[0]
    try:
        h = yf.Ticker(sym).history(period=f"{years}y")
        return h if h is not None and len(h) > 150 else None
    except Exception:
        return None


def zigzag(close: pd.Series, pct: float = 5.0) -> list[tuple]:
    """Các điểm xoay của giá: chỉ giữ lại đoạn chạy vượt `pct` phần trăm.

    Lọc nhiễu để còn lại khung xương của xu hướng — cả Wyckoff lẫn Elliott
    đều đọc trên khung xương này chứ không đọc từng phiên.
    """
    if close is None or len(close) < 30:
        return []
    piv = [(close.index[0], float(close.iloc[0]), 0)]   # 0 = chưa rõ chiều
    last_i, last_p = 0, float(close.iloc[0])
    direction = 0
    for i in range(1, len(close)):
        p = float(close.iloc[i])
        chg = (p / last_p - 1) * 100
        if direction >= 0 and chg <= -pct:
            piv.append((close.index[i], p, -1)); last_i, last_p, direction = i, p, -1
        elif direction <= 0 and chg >= pct:
            piv.append((close.index[i], p, +1)); last_i, last_p, direction = i, p, +1
        elif (direction > 0 and p > last_p) or (direction < 0 and p < last_p):
            piv[-1] = (close.index[i], p, direction); last_i, last_p = i, p
    return piv


def _range_stats(h: pd.DataFrame, lookback: int = 120) -> dict:
    """Vùng giá gần đây rộng hẹp ra sao, giá đang nằm ở đâu trong vùng."""
    w = h.tail(lookback)
    hi, lo = float(w["High"].max()), float(w["Low"].min())
    px = float(w["Close"].iloc[-1])
    width = (hi / lo - 1) * 100 if lo > 0 else None
    pos = (px - lo) / (hi - lo) * 100 if hi > lo else None
    return {"hi": hi, "lo": lo, "width": width, "pos": pos, "px": px}


def _volume_signals(h: pd.DataFrame, lookback: int = 120) -> dict:
    """Nỗ lực so với kết quả — phần cốt lõi của Wyckoff."""
    w = h.tail(lookback).copy()
    if "Volume" not in w or w["Volume"].sum() <= 0:
        return {}
    w["ret"] = w["Close"].pct_change()
    up, dn = w[w["ret"] > 0], w[w["ret"] < 0]
    out = {}
    if len(up) and len(dn):
        # >1: phiên tăng có khối lượng lớn hơn phiên giảm, tức bên mua chủ động
        out["up_dn_vol"] = round(float(up["Volume"].mean() / dn["Volume"].mean()), 2)
    first, second = w.head(lookback // 2), w.tail(lookback // 2)
    if first["Volume"].mean() > 0:
        # <1: khối lượng cạn dần trong vùng — dấu hiệu hàng đã được hấp thụ
        out["vol_trend"] = round(float(second["Volume"].mean() / first["Volume"].mean()), 2)
    return out


def _spring_upthrust(h: pd.DataFrame, lookback: int = 120, probe: int = 25) -> dict:
    """Cú rũ bỏ dưới đáy vùng, hoặc cú vượt đỉnh hụt.

    Đây là hai tín hiệu đặc trưng nhất của Wyckoff: phá vỡ rồi quay lại ngay
    nghĩa là phá vỡ đó giả.
    """
    base = h.tail(lookback).head(lookback - probe)
    tail = h.tail(probe)
    if base.empty or tail.empty:
        return {}
    b_hi, b_lo = float(base["High"].max()), float(base["Low"].min())
    out = {}
    pierced_lo = tail["Low"].min() < b_lo
    pierced_hi = tail["High"].max() > b_hi
    last = float(tail["Close"].iloc[-1])
    if pierced_lo and last > b_lo:
        out["spring"] = round((b_lo / float(tail["Low"].min()) - 1) * 100, 1)
    if pierced_hi and last < b_hi:
        out["upthrust"] = round((float(tail["High"].max()) / b_hi - 1) * 100, 1)
    return out


# ─────────────────────────────────────────────────────────────────────────────
# Wyckoff
# ─────────────────────────────────────────────────────────────────────────────
def wyckoff(h: pd.DataFrame) -> dict:
    close = h["Close"]
    ma50 = float(close.rolling(50).mean().iloc[-1])
    ma200 = float(close.rolling(200).mean().iloc[-1]) if len(close) >= 200 else None
    px = float(close.iloc[-1])
    rs = _range_stats(h)
    vs = _volume_signals(h)
    su = _spring_upthrust(h)
    piv = zigzag(close, pct=5.0)

    # Cấu trúc đỉnh đáy của bốn điểm xoay gần nhất
    highs = [p for p in piv if p[2] > 0][-3:]
    lows = [p for p in piv if p[2] < 0][-3:]
    hh = len(highs) >= 2 and highs[-1][1] > highs[-2][1]
    hl = len(lows) >= 2 and lows[-1][1] > lows[-2][1]
    lh = len(highs) >= 2 and highs[-1][1] < highs[-2][1]
    ll = len(lows) >= 2 and lows[-1][1] < lows[-2][1]

    above200 = ma200 is not None and px > ma200
    narrow = rs["width"] is not None and rs["width"] < 18     # vùng đi ngang
    reasons = []

    if narrow:
        pos = rs["pos"] if rs["pos"] is not None else 50
        # Phân phối phải có BẰNG CHỨNG, không chỉ là "đang ở nửa trên vùng".
        # Gọi nhầm phân phối giữa sóng tăng là sai đắt giá: nó khuyên tránh mua
        # đúng lúc nên mua.
        weak = (su.get("upthrust", 0) or 0) > 1.0 or vs.get("up_dn_vol", 1) < 0.95
        if above200:
            if pos > 70 and weak:
                phase, conf = "distribution", 0.55
                reasons.append(f"đi ngang {rs['width']:.0f}% sát đỉnh vùng sau một đợt tăng")
                if (su.get("upthrust", 0) or 0) > 1.0:
                    conf += 0.15; reasons.append(f"cú vượt đỉnh hụt {su['upthrust']:+.1f}%")
                if vs.get("up_dn_vol", 1) < 0.95:
                    conf += 0.15; reasons.append("phiên giảm có khối lượng lớn hơn phiên tăng")
            else:
                # Đi ngang trên MA200 mà chưa có dấu hiệu yếu: nghỉ giữa xu hướng tăng
                phase, conf = "markup", 0.5
                reasons.append(f"đi ngang {rs['width']:.0f}% nhưng vẫn trên MA200 — "
                               f"tái tích luỹ giữa xu hướng tăng")
                if vs.get("vol_trend", 1) < 0.85:
                    conf += 0.15; reasons.append("khối lượng cạn dần, hàng đang được hấp thụ")
        else:
            # Tích luỹ chỉ có nghĩa khi đến SAU một pha giảm
            phase, conf = "accumulation", 0.5
            reasons.append(f"đi ngang {rs['width']:.0f}% dưới MA200 sau một đợt giảm")
            if su.get("spring"):
                conf += 0.2; reasons.append(f"có cú rũ bỏ dưới đáy {su['spring']:+.1f}% rồi hồi")
            if vs.get("vol_trend", 1) < 0.85:
                conf += 0.15; reasons.append("khối lượng cạn dần trong vùng")
    elif above200 and (hh or hl):
        phase, conf = "markup", 0.6
        reasons.append("đỉnh sau cao hơn đỉnh trước" if hh else "đáy sau cao hơn đáy trước")
        reasons.append("giá trên MA200")
        if vs.get("up_dn_vol", 1) > 1.05:
            conf += 0.15; reasons.append("phiên tăng có khối lượng lớn hơn phiên giảm")
        if su.get("upthrust"):
            conf -= 0.15; reasons.append("nhưng vừa có cú vượt đỉnh hụt")
    elif (not above200) and (lh or ll):
        phase, conf = "markdown", 0.6
        reasons.append("đỉnh sau thấp hơn đỉnh trước" if lh else "đáy sau thấp hơn đáy trước")
        reasons.append("giá dưới MA200")
        if su.get("spring"):
            conf -= 0.15; reasons.append("nhưng vừa có cú rũ bỏ rồi hồi")
    else:
        phase, conf = "unclear", 0.3
        reasons.append("cấu trúc đỉnh đáy chưa rõ chiều")

    label, fg, bg = WYCKOFF[phase]
    return {
        "phase": phase, "label": label, "fg": fg, "bg": bg,
        "confidence": round(min(max(conf, 0.2), 0.9), 2),
        "reasons": reasons,
        "range": rs, "volume": vs, "events": su,
        "ma50": ma50, "ma200": ma200, "price": px,
        "advice": _wyckoff_advice(phase),
    }


def _wyckoff_advice(phase: str) -> str:
    return {
        "markup":       "Pha đi lên. Nhịp chỉnh trong pha này thường được mua lại — "
                        "bối cảnh thuận cho việc mua khi giá giảm.",
        "distribution": "Có dấu hiệu hàng đang được phân phối ra. Mã giảm sâu lúc này "
                        "rủi ro cao hơn bình thường vì xu hướng chung có thể sắp đổi.",
        "markdown":     "Pha đi xuống. Mua khi giá giảm trong pha này thường là bắt dao rơi; "
                        "nên chờ dấu hiệu rũ bỏ và tích luỹ trước.",
        "accumulation": "Có dấu hiệu hàng đang được gom. Đây là pha mà mua khi giá giảm "
                        "cho kết quả tốt nhất trong lịch sử, nhưng cũng là pha dài nhất.",
        "unclear":      "Chưa đủ dấu hiệu để xếp pha. Giảm quy mô và chờ cấu trúc rõ hơn.",
    }.get(phase, "")


# ─────────────────────────────────────────────────────────────────────────────
# Elliott — giả thuyết, không phải kết luận
# ─────────────────────────────────────────────────────────────────────────────
def elliott(h: pd.DataFrame) -> dict:
    """Kiểm ba quy tắc cứng của sóng đẩy trên năm điểm xoay gần nhất."""
    piv = zigzag(h["Close"], pct=6.0)
    out = {"ok": False, "note": "", "alt": "", "points": len(piv)}
    if len(piv) < 6:
        out["note"] = "Chưa đủ điểm xoay để thử đếm sóng."
        return out

    pts = piv[-6:]
    p = [float(x[1]) for x in pts]
    up = p[-1] > p[0]

    # Ba quy tắc không được vi phạm của một sóng đẩy
    w1 = abs(p[1] - p[0]); w3 = abs(p[3] - p[2]); w5 = abs(p[5] - p[4])
    r2 = abs(p[2] - p[1]) / w1 if w1 else 9
    broken = []
    if r2 > 1.0:
        broken.append("sóng 2 đã hồi quá 100% sóng 1")
    if w3 < w1 and w3 < w5:
        broken.append("sóng 3 là sóng ngắn nhất")
    overlap = (p[4] < p[1]) if up else (p[4] > p[1])
    if overlap:
        broken.append("sóng 4 chồng lấn vùng giá sóng 1")

    out["rule_breaks"] = broken
    out["retrace_w2"] = round(r2 * 100, 0) if r2 < 9 else None
    out["direction"] = "tăng" if up else "giảm"
    if broken:
        out["ok"] = False
        out["note"] = ("Năm đoạn gần nhất KHÔNG thoả quy tắc sóng đẩy ("
                       + "; ".join(broken) + "), nhiều khả năng đây là một mô hình "
                       "điều chỉnh chứ không phải sóng đẩy.")
        out["alt"] = "Cách đếm thay thế: chuỗi a-b-c trong một đợt điều chỉnh lớn hơn."
    else:
        out["ok"] = True
        out["note"] = (f"Năm đoạn gần nhất thoả cả ba quy tắc của sóng đẩy {out['direction']}, "
                       f"tức có thể đang ở cuối sóng 5.")
        out["alt"] = ("Cách đếm thay thế hợp lệ không kém: đây mới là sóng 3 của một cấp "
                      "lớn hơn, hoặc toàn bộ chỉ là sóng b của một điều chỉnh.")
    out["caveat"] = ("Đếm sóng tự động chỉ kiểm được quy tắc cứng, không thay được mắt "
                     "người. Cùng một đồ thị luôn có nhiều cách đếm hợp lệ.")
    return out


def assess(market: str = "america") -> dict | None:
    h = fetch_ohlc(market)
    if h is None:
        return None
    return {"wyckoff": wyckoff(h), "elliott": elliott(h),
            "index": YF_INDEX.get(market.lower(), ("", ""))[1],
            "bars": len(h)}
