# -*- coding: utf-8 -*-
"""Trạng thái thị trường: uptrend, sideways hay downtrend.

Vì sao cần, với một bộ lọc mua khi giá giảm: kiểm thử lùi cho thấy mua mã
đã sụt sâu cho kết quả rất tốt — nhưng toàn bộ mẫu rơi vào giai đoạn
2023–2025, tức là mua vào đúng đợt hồi phục sau downtrend 2022. Trong chữ V
thì bắt đáy lúc nào cũng thắng. Biết thị trường đang ở pha nào là điều kiện
để đọc tín hiệu "giảm sâu" cho đúng: trong uptrend nó thường là nhịp điều
chỉnh, trong downtrend nó thường mới chỉ là chặng đầu.

Trạng thái tính từ hai nguồn độc lập:
  · Chỉ số: giá so với MA50, MA200 và mức sụt từ đỉnh một năm.
  · Độ rộng: bao nhiêu phần trăm mã trong chính lần quét đang trên MA250.
"""
from __future__ import annotations

import pandas as pd

INDEX = {
    "america":   ("america",  "SP:SPX",        "S&P 500"),
    "nasdaq":    ("america",  "NASDAQ:NDX",    "Nasdaq 100"),
    "nyse":      ("america",  "NYSE:NYA",      "NYSE Composite"),
    "euronext":  ("global",   "TVC:CAC40",     "CAC 40"),
    "hong_kong": ("hongkong", "HSI:HSI",       "Hang Seng"),
    "vietnam":   ("vietnam",  "HOSE:VNINDEX",  "VN-Index"),
}

# nhãn, màu chữ, màu nền
STATES = {
    "uptrend_strong":   ("📈 UPTREND MẠNH",   "#1A5C2B", "#C6EFCE"),
    "uptrend":          ("📈 UPTREND",        "#1B5E20", "#DCEDC8"),
    "sideways":         ("➡️ SIDEWAYS",       "#7D6608", "#FFF2CC"),
    "downtrend":        ("📉 DOWNTREND",      "#9C0006", "#FFC7CE"),
    "downtrend_strong": ("📉 DOWNTREND MẠNH", "#FFFFFF", "#9C0006"),
    "unknown":          ("⚪ THIẾU DỮ LIỆU",   "#475569", "#F1F5F9"),
}


def fetch_index(market: str = "america") -> dict:
    """Số liệu chỉ số của sàn đang quét."""
    import requests

    tv_market, symbol, name = INDEX.get(market.lower(), INDEX["america"])
    out = {"symbol": symbol, "name": name}
    try:
        r = requests.post(
            f"https://scanner.tradingview.com/{tv_market}/scan",
            json={"symbols": {"tickers": [symbol], "query": {"types": []}},
                  "columns": ["close", "SMA50", "SMA200", "price_52_week_high",
                              "Perf.3M", "Perf.6M", "Perf.Y"]},
            timeout=12)
        d = r.json()["data"][0]["d"]
        for k, v in zip(["close", "ma50", "ma200", "high52",
                         "p3m", "p6m", "p1y"], d):
            out[k] = float(v) if v is not None else None
    except Exception as e:
        out["error"] = f"{type(e).__name__}: {e}"
    return out


def breadth(df: pd.DataFrame) -> float | None:
    """Phần trăm mã trong lần quét đang nằm trên MA250.

    Độ rộng là nguồn độc lập với chỉ số: chỉ số có thể được kéo bởi vài mã
    vốn hoá lớn trong khi phần còn lại đã gãy từ lâu.
    """
    if df is None or df.empty or "Giá/MA250%" not in df.columns:
        return None
    v = pd.to_numeric(df["Giá/MA250%"], errors="coerce").dropna()
    return round((v > 0).mean() * 100, 1) if len(v) else None


def classify(idx: dict, breadth_pct: float | None = None) -> dict:
    """Gộp chỉ số và độ rộng thành một trạng thái."""
    close, ma50, ma200 = idx.get("close"), idx.get("ma50"), idx.get("ma200")
    high52 = idx.get("high52")
    res = {"breadth": breadth_pct, "reasons": []}

    if not close or not ma200:
        res["state"] = "unknown"
        res["reasons"].append("thiếu dữ liệu chỉ số")
        res["label"], res["fg"], res["bg"] = STATES["unknown"]
        res["advice"] = _advice("unknown")
        return res

    dd = (close / high52 - 1) * 100 if high52 else None
    res["dd"] = round(dd, 1) if dd is not None else None
    res["vs_ma200"] = round((close / ma200 - 1) * 100, 1)
    res["vs_ma50"] = round((close / ma50 - 1) * 100, 1) if ma50 else None

    # Mỗi điều kiện một phiếu; cộng lại ra trạng thái
    score = 0
    if close > ma200:
        score += 1; res["reasons"].append("chỉ số trên MA200")
    else:
        score -= 1; res["reasons"].append("chỉ số dưới MA200")
    if ma50 and ma50 > ma200:
        score += 1; res["reasons"].append("MA50 trên MA200")
    elif ma50:
        score -= 1; res["reasons"].append("MA50 dưới MA200")
    if dd is not None:
        if dd > -5:
            score += 1; res["reasons"].append(f"sát đỉnh ({dd:+.1f}%)")
        elif dd < -20:
            score -= 2; res["reasons"].append(f"sụt {dd:.1f}% từ đỉnh")
        elif dd < -10:
            score -= 1; res["reasons"].append(f"sụt {dd:.1f}% từ đỉnh")
    if breadth_pct is not None:
        if breadth_pct >= 60:
            score += 1; res["reasons"].append(f"{breadth_pct:.0f}% số mã trên MA250")
        elif breadth_pct < 35:
            score -= 1; res["reasons"].append(f"chỉ {breadth_pct:.0f}% số mã trên MA250")

    res["score"] = score
    # DOWNTREND MẠNH đòi hỏi cả hai đường trung bình gãy VÀ chỉ số sụt trên
    # 20% — chỉ riêng việc nằm dưới MA200 thì mới là downtrend thường.
    res["state"] = ("uptrend_strong" if score >= 3 else "uptrend" if score >= 1
                    else "sideways" if score == 0
                    else "downtrend" if score >= -3 else "downtrend_strong")
    res["label"], res["fg"], res["bg"] = STATES[res["state"]]
    res["advice"] = _advice(res["state"])
    return res


def _advice(state: str) -> str:
    """Tín hiệu này nên đọc ra sao khi đang săn mã giảm sâu."""
    return {
        "uptrend_strong":   "Mã giảm sâu giữa một uptrend mạnh thường là vấn đề riêng "
                            "của doanh nghiệp đó, không phải của thị trường — càng phải "
                            "tìm hiểu kỹ lý do.",
        "uptrend":          "Nhịp điều chỉnh trong uptrend. Đây là bối cảnh thuận lợi nhất "
                            "cho việc mua khi giá giảm.",
        "sideways":         "Thị trường đi ngang, chưa rõ xu hướng. Vào từng phần thay vì "
                            "một lần.",
        "downtrend":        "Giá giảm lúc này nhiều khả năng mới là chặng đầu. Giữ tiền mặt "
                            "và chia nhỏ lệnh mua.",
        "downtrend_strong": "Downtrend đã lan rộng. Mã rẻ còn rẻ nữa — toàn bộ kết quả kiểm "
                            "thử lùi của bộ lọc đến từ một giai đoạn hồi phục, không áp "
                            "dụng được cho pha này.",
        "unknown":          "Không lấy được số liệu chỉ số của sàn này — có thể do mạng. "
                            "Không có bối cảnh thị trường thì tín hiệu giảm sâu của từng mã "
                            "không đọc được là nhịp chỉnh hay chặng đầu của một đợt giảm.",
    }.get(state, "")


def assess(market: str, df: pd.DataFrame | None = None) -> dict:
    return classify(fetch_index(market), breadth(df))
