# -*- coding: utf-8 -*-
"""Đoán giai đoạn thị trường theo Wyckoff.

Trả lời đúng một câu: thị trường đang ở NỀN TÍCH LUỸ, NỀN PHÂN PHỐI, hay
không phải nền mà đang trong xu hướng.

Wyckoff cơ giới hoá được tương đối vì nó dựa vào quan hệ giữa giá và khối
lượng — vùng đi ngang, cú rũ bỏ dưới đáy, khối lượng lớn mà giá không đi
được, tất cả đều đo được bằng quy tắc. Nhưng đây vẫn là một cách đọc có
căn cứ, không phải chân lý, và nó mô tả hiện tại chứ không dự báo tương lai.
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

    # Kết luận một câu: đây có phải là nền không, và nền loại gì
    in_base = narrow or phase in ("accumulation", "distribution")
    if phase == "accumulation":
        verdict, v_fg, v_bg = "NỀN TÍCH LUỸ", "#1A5C2B", "#C6EFCE"
        v_note = ("Hàng đang được gom lại sau một đợt giảm. Đây là nền mà mua khi "
                  "giá giảm cho kết quả tốt nhất — nhưng cũng là nền kéo dài lâu nhất.")
    elif phase == "distribution":
        verdict, v_fg, v_bg = "NỀN PHÂN PHỐI", "#9C0006", "#FFC7CE"
        v_note = ("Hàng đang được bán ra cho đám đông sau một đợt tăng. Mua khi giá "
                  "giảm trong nền này là đi ngược dòng tiền lớn.")
    elif narrow and phase == "markup":
        verdict, v_fg, v_bg = "NỀN TÁI TÍCH LUỸ", "#1B3A5C", "#DDEEFF"
        v_note = ("Đi ngang để nghỉ giữa xu hướng tăng, chưa có dấu hiệu bị phân phối. "
                  "Nhịp chỉnh trong nền này thường được mua lại.")
    elif phase == "markup":
        verdict, v_fg, v_bg = "KHÔNG PHẢI NỀN — ĐANG ĐI LÊN", "#1A5C2B", "#C6EFCE"
        v_note = "Đang trong xu hướng tăng, chưa hình thành nền đi ngang."
    elif phase == "markdown":
        verdict, v_fg, v_bg = "KHÔNG PHẢI NỀN — ĐANG ĐI XUỐNG", "#9C0006", "#FFC7CE"
        v_note = ("Đang rơi, chưa tạo nền. Mua lúc này là bắt dao rơi — chờ giá đi "
                  "ngang và có cú rũ bỏ trước.")
    else:
        verdict, v_fg, v_bg = "CHƯA XÁC ĐỊNH ĐƯỢC NỀN", "#64748B", "#F1F5F9"
        v_note = "Cấu trúc chưa đủ rõ để kết luận."

    label, fg, bg = WYCKOFF[phase]
    return {
        "verdict": verdict, "v_fg": v_fg, "v_bg": v_bg, "v_note": v_note,
        "in_base": in_base,
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


def assess(market: str = "america") -> dict | None:
    h = fetch_ohlc(market)
    if h is None:
        return None
    return {"wyckoff": wyckoff(h),
            "index": YF_INDEX.get(market.lower(), ("", ""))[1],
            "bars": len(h)}
