# -*- coding: utf-8 -*-
"""Danh sách theo dõi — lưu vào watchlist.json.

Khác Portfolio: Portfolio là những mã ĐÃ mua, có giá mua và số lượng, nên
nó tính lãi lỗ. Watchlist là những mã CHƯA mua mà đang chờ — nên thứ cần
biết không phải lãi lỗ, mà là từ lúc để mắt tới giờ mã đã rẻ đi bao nhiêu
và chất lượng còn nguyên không.

Vì vậy lúc thêm, mỗi mã được chụp lại giá cùng ba điểm số. Có cái mốc đó
mới trả lời được câu đáng hỏi với lối mua khi giá giảm: giá đã về vùng
muốn mua chưa, và trong lúc chờ thì doanh nghiệp có xấu đi không. Hai thứ
đó đi cùng nhau mới ra quyết định — giá rẻ đi mà điểm cũng tụt theo thì
thường là thị trường biết điều gì mà bộ lọc chưa thấy.

Tách khỏi GUI để kiểm được mà không cần dựng Qt.
"""
from __future__ import annotations

import json
import os
import sys

FIELDS = ("ticker", "added", "note", "px", "signal", "fin", "real", "base")


def path() -> str:
    """Cạnh file chạy, giống portfolio.json."""
    base = os.path.dirname(os.path.abspath(sys.argv[0]))
    return os.path.join(base, "watchlist.json")


def load() -> list[dict]:
    """Đọc danh sách. File hỏng hay thiếu thì trả về danh sách rỗng.

    Không để lỗi đọc file làm sập app — mất watchlist đã đủ tệ rồi.
    """
    try:
        with open(path(), encoding="utf-8") as f:
            data = json.load(f)
    except Exception:
        return []
    if not isinstance(data, list):
        return []
    # Bỏ những mục không ra hình dạng gì, giữ phần còn lại
    return [it for it in data if isinstance(it, dict) and it.get("ticker")]


def save(items: list[dict]) -> bool:
    try:
        with open(path(), "w", encoding="utf-8") as f:
            json.dump(items, f, ensure_ascii=False, indent=2)
        return True
    except Exception:
        return False


def tickers(items: list[dict] | None = None) -> set[str]:
    return {str(it.get("ticker", "")).upper()
            for it in (items if items is not None else load())}


def has(ticker: str, items: list[dict] | None = None) -> bool:
    return str(ticker).strip().upper() in tickers(items)


def _snapshot(row: dict | None) -> dict:
    """Chụp lại giá và ba điểm số tại thời điểm thêm vào."""
    if not row:
        return {"px": None, "signal": "", "fin": None, "real": None, "base": ""}

    def _f(key):
        v = row.get(key)
        try:
            v = float(v)
        except (TypeError, ValueError):
            return None
        return None if v != v else v      # NaN của pandas

    return {
        "px":     _f("Price ($)"),
        "signal": str(row.get("Signal") or ""),
        "fin":    _f("FIN_Score"),
        "real":   _f("REAL_Score"),
        "base":   str(row.get("Nền giá") or ""),
    }


def add(ticker: str, note: str = "", row: dict | None = None,
        today: str | None = None) -> bool:
    """Thêm một mã. Trả về False nếu mã đã có trong danh sách."""
    tk = str(ticker).strip().upper()
    if not tk:
        return False
    items = load()
    if has(tk, items):
        return False
    if today is None:
        import datetime
        today = datetime.date.today().isoformat()
    items.append({"ticker": tk, "added": today, "note": str(note or ""),
                  **_snapshot(row)})
    return save(items)


def remove(ticker: str) -> bool:
    tk = str(ticker).strip().upper()
    items = load()
    keep = [it for it in items if str(it.get("ticker", "")).upper() != tk]
    if len(keep) == len(items):
        return False
    return save(keep)


def set_note(ticker: str, note: str) -> bool:
    tk = str(ticker).strip().upper()
    items = load()
    for it in items:
        if str(it.get("ticker", "")).upper() == tk:
            it["note"] = str(note or "")
            return save(items)
    return False


def toggle(ticker: str, row: dict | None = None) -> bool:
    """Thêm nếu chưa có, bỏ nếu đã có. Trả về True nếu kết quả là ĐANG theo dõi."""
    if has(ticker):
        remove(ticker)
        return False
    add(ticker, row=row)
    return True
