"""
Đồ thị giá nến kèm vùng định giá
================================
Nến ngày (6 tháng, 1 năm) hoặc nến tuần (3 năm, 5 năm), MA50/MA200 tính trên
giá đóng cửa ngày, khối lượng và đỉnh 52 tuần. Với mã Mỹ đã có nhãn định giá,
bốn vùng Rẻ / Hợp lý / Đắt / Rất đắt được vẽ thành dải giá ngang — lấy từ
valuation.price_bands, tức đúng mức giá mà nhãn trong bảng sẽ đổi.

Dải giá dùng dòng tiền và số cổ phiếu lúc quét nên chỉ đúng cho hiện tại:
ba năm trước dòng tiền khác, vùng khác. Vì vậy dải được tô đậm ở mép phải
(hiện tại) và chỉ để mờ phía sau.
"""
from __future__ import annotations

import html as _html
from dataclasses import dataclass
from datetime import date, timedelta
from typing import Callable

import pandas as pd
import plotly.graph_objects as go
from plotly.subplots import make_subplots

import valuation
from fund_chart import C_BLUE, C_GOLD, C_GREEN, C_RED, _page, _palette, message_html
from fund_holders import US_EXCHANGES

__all__ = ["RANGES", "DEFAULT_RANGE", "PriceHistory", "Snapshot", "yahoo_candidates",
           "fetch", "prepare", "weekly", "window", "snapshot", "build_figure",
           "build_html", "message_html"]

# khóa → (nhãn, số ngày lịch, khung nến: D ngày / W tuần)
RANGES = {"6m": ("6 tháng", 183, "D"), "1y": ("1 năm", 365, "D"),
          "3y": ("3 năm", 1096, "W"), "5y": ("5 năm", 1827, "W")}
DEFAULT_RANGE = "3y"
MA_FAST, MA_SLOW = 50, 200
HIGH_DAYS = 365
# 5 năm + ~220 phiên để MA200 đã có giá trị ngay ở nến đầu tiên của khung 5 năm
FETCH_DAYS = 1827 + 320
# Giá hôm nay lệch giá lúc quét ngoài khoảng này thì vùng định giá không còn
# khớp (chia tách cổ phiếu, quét quá cũ…) — bỏ dải thay vì vẽ sai
SPLIT_GUARD = (0.5, 2.0)
# Ranh giới vùng nằm trong khoảng này so với giá đóng cửa thì nới trục giá để thấy
NEAR_EDGE = (0.55, 1.6)
RECENT = 0.82          # dải tô đậm từ 82% chiều ngang trở về phải
X_PAD = 0.07           # khoảng trống bên phải nến cuối cho nhãn giá

US_MARKETS = frozenset({"america", "nasdaq", "nyse"})
EURONEXT_SUFFIXES = (".PA", ".AS", ".BR", ".LS", ".IR", ".OL", ".MI")
ZONE_COLORS = {"Rẻ": C_GREEN, "Hợp lý": C_BLUE, "Đắt": C_GOLD, "Rất đắt": C_RED}
_ZONE_RGB = {"Rẻ": (52, 196, 114), "Hợp lý": (74, 158, 255),
             "Đắt": (232, 169, 61), "Rất đắt": (232, 72, 61)}
_OHLC = ("Open", "High", "Low", "Close")


@dataclass(frozen=True)
class PriceHistory:
    ticker: str          # mã gốc trong bảng
    symbol: str          # ký hiệu Yahoo đã tải được
    currency: str
    bars: pd.DataFrame   # nến ngày sau prepare()


@dataclass(frozen=True)
class Snapshot:
    close: float
    day: pd.Timestamp
    change: float | None      # % so với phiên trước
    ma50: float | None
    ma200: float | None
    high52: float
    high52_day: pd.Timestamp


# ── Tải dữ liệu ─────────────────────────────────────────────────────────────

def yahoo_candidates(ticker: str, exchange: str = "", market: str = "") -> list[str]:
    """Các ký hiệu Yahoo có thể ứng với mã TradingView, thử lần lượt."""
    sym = str(ticker).strip().upper()
    ex = str(exchange or "").strip().upper()
    mk = str(market or "").strip().lower()
    if not sym:
        return []
    if ex in US_EXCHANGES or (not ex and mk in US_MARKETS):
        return [sym.replace(".", "-")]
    if ex == "HKEX" or mk == "hong_kong":
        return [(sym.zfill(4) if sym.isdigit() else sym) + ".HK"]
    if ex in ("HOSE", "HNX", "UPCOM") or mk == "vietnam":
        return [sym + ".VN", sym + ".HN"]
    if ex.startswith("EURONEXT") or mk == "euronext":
        return [sym + s for s in EURONEXT_SUFFIXES]
    return [sym]


def _yf_history(symbol: str, start: date) -> tuple[pd.DataFrame, str]:
    import yfinance as yf
    tk = yf.Ticker(symbol)
    raw = tk.history(start=start.isoformat(), auto_adjust=True)
    meta = tk.history_metadata or {}
    return raw, str(meta.get("currency") or "")


class Cancelled(Exception):
    """Bên gọi xin dừng giữa chừng (ứng dụng đang thoát)."""


def fetch(ticker: str, exchange: str = "", market: str = "", expect: float | None = None,
          today: date | None = None,
          history: Callable[[str, date], tuple[pd.DataFrame, str]] = _yf_history,
          stop: Callable[[], bool] | None = None,
          ) -> PriceHistory:
    """Tải ~5 năm nến ngày. `expect` là giá lúc quét: khi nhiều ký hiệu cùng có
    dữ liệu (Euronext niêm yết chéo), ưu tiên ký hiệu có giá gần với nó.
    `stop` được hỏi trước mỗi ký hiệu; trả True thì ném Cancelled.
    ValueError khi không ký hiệu nào có giá."""
    cands = yahoo_candidates(ticker, exchange, market)
    start = (today or date.today()) - timedelta(days=FETCH_DAYS)
    fallback: PriceHistory | None = None
    errors: list[str] = []     # ký hiệu nào cũng lỗi (mất mạng, bị chặn…) thì nói rõ
    for sym in cands:
        if stop is not None and stop():
            raise Cancelled(ticker)
        try:
            raw, currency = history(sym, start)
        except Exception as e:     # yfinance ném lỗi đủ kiểu với mã không tồn tại
            errors.append(f"{type(e).__name__}: {e}")
            continue
        bars = prepare(raw)
        if bars.empty:
            continue
        found = PriceHistory(str(ticker), sym, currency, bars)
        if not expect or expect <= 0 or _within(bars["Close"].iloc[-1] / expect, SPLIT_GUARD):
            return found
        fallback = fallback or found
    if fallback is not None:
        return fallback
    tried = ", ".join(cands) or "—"
    if cands and len(errors) == len(cands):
        raise ValueError(f"Không tải được giá {ticker} từ Yahoo Finance (đã thử {tried}) — "
                         f"{errors[-1][:200]}")
    raise ValueError(f"Yahoo Finance không có giá cho {ticker} (đã thử {tried})")


def _within(x: float, bounds: tuple[float, float]) -> bool:
    return bounds[0] <= x <= bounds[1]


def prepare(raw: pd.DataFrame | None) -> pd.DataFrame:
    """Nến ngày sạch: chỉ mục ngày không múi giờ, không trùng, bỏ phiên thiếu
    giá đóng cửa, Open/High/Low hỏng thì lấy theo Close, thêm MA50/MA200."""
    cols = [*_OHLC, "Volume", "MA50", "MA200"]
    if raw is None or raw.empty or "Close" not in raw.columns:
        return pd.DataFrame(columns=cols, index=pd.DatetimeIndex([]))
    idx = pd.DatetimeIndex(raw.index)
    if idx.tz is not None:
        idx = idx.tz_localize(None)
    close = pd.to_numeric(raw["Close"], errors="coerce")
    out = pd.DataFrame({"Close": close.to_numpy()}, index=idx.normalize())
    for c in ("Open", "High", "Low"):
        v = pd.to_numeric(raw[c], errors="coerce") if c in raw.columns else close
        out[c] = v.where(v > 0).to_numpy()          # Yahoo đôi khi trả Open = 0
    vol = raw["Volume"] if "Volume" in raw.columns else pd.Series(0, index=raw.index)
    out["Volume"] = pd.to_numeric(vol, errors="coerce").fillna(0).to_numpy()
    out = out[out["Close"] > 0]
    out = out[~out.index.duplicated(keep="last")].sort_index()
    for c in ("Open", "High", "Low"):
        out[c] = out[c].fillna(out["Close"])
    out["High"] = out[["Open", "High", "Close"]].max(axis=1)
    out["Low"] = out[["Open", "Low", "Close"]].min(axis=1)
    out["MA50"] = out["Close"].rolling(MA_FAST, min_periods=MA_FAST).mean()
    out["MA200"] = out["Close"].rolling(MA_SLOW, min_periods=MA_SLOW).mean()
    return out[cols]


def weekly(bars: pd.DataFrame) -> pd.DataFrame:
    """Gộp nến ngày thành nến tuần (thứ Hai – thứ Sáu). Chỉ mục là phiên cuối
    tuần; cột `first` giữ phiên đầu để hiện trong ô di chuột. MA lấy giá trị
    cuối tuần của MA ngày, nên cùng đường với khung ngày."""
    if bars.empty:
        return bars.assign(first=pd.Series(dtype="datetime64[ns]"))
    days = pd.Series(bars.index, index=bars.index)
    g = bars.groupby(bars.index.to_period("W-FRI"))
    out = pd.DataFrame({"Open": g["Open"].first(), "High": g["High"].max(),
                        "Low": g["Low"].min(), "Close": g["Close"].last(),
                        "Volume": g["Volume"].sum(), "MA50": g["MA50"].last(),
                        "MA200": g["MA200"].last()})
    dg = days.groupby(bars.index.to_period("W-FRI"))
    out["first"] = dg.first().to_numpy()
    out.index = pd.DatetimeIndex(dg.last().to_numpy())
    return out


def window(bars: pd.DataFrame, rng: str) -> pd.DataFrame:
    """Các nến trong khung, tính lùi từ phiên cuối (không từ hôm nay, để dữ
    liệu dừng giao dịch vẫn hiện đủ khung)."""
    _label, days, freq = RANGES.get(rng, RANGES[DEFAULT_RANGE])
    if bars.empty:
        return weekly(bars) if freq == "W" else bars
    start = bars.index[-1] - pd.Timedelta(days=days)
    view = bars[bars.index > start]
    return weekly(view) if freq == "W" else view


def snapshot(bars: pd.DataFrame) -> Snapshot:
    last = bars.iloc[-1]
    day = bars.index[-1]
    prev = bars["Close"].iloc[-2] if len(bars) > 1 else None
    year = bars[bars.index > day - pd.Timedelta(days=HIGH_DAYS)]
    hi_day = year["High"].idxmax()
    return Snapshot(
        close=float(last["Close"]), day=day,
        change=(float(last["Close"]) / prev - 1) * 100 if prev else None,
        ma50=None if pd.isna(last["MA50"]) else float(last["MA50"]),
        ma200=None if pd.isna(last["MA200"]) else float(last["MA200"]),
        high52=float(year.loc[hi_day, "High"]), high52_day=hi_day)


def usable_bands(hist: PriceHistory, bands: valuation.PriceBands | None
                 ) -> tuple[valuation.PriceBands | None, str]:
    """Bỏ vùng định giá khi không còn khớp với giá đang vẽ, kèm lý do."""
    if bands is None:
        return None, ""
    if hist.currency and hist.currency.upper() != "USD":
        return None, f"vùng định giá tính bằng USD, còn giá đang vẽ bằng {hist.currency}"
    close = float(hist.bars["Close"].iloc[-1])
    if not _within(close / bands.scan_price, SPLIT_GUARD):
        return None, (f"giá hiện tại {_money(close)} lệch xa giá lúc quét "
                      f"{_money(bands.scan_price)} (có thể đã chia tách cổ phiếu) — "
                      f"quét lại để có vùng định giá đúng")
    return bands, ""


# ── Định dạng ───────────────────────────────────────────────────────────────

def _decimals(v: float, currency: str = "") -> int:
    if currency.upper() == "VND" or abs(v) >= 1000:
        return 0
    return 2


def _num(v: float, nd: int) -> str:
    """1.234,56 — kiểu Việt Nam."""
    return f"{v:,.{nd}f}".replace(",", " ").replace(".", ",").replace(" ", ".")


def _money(v: float, currency: str = "", nd: int | None = None) -> str:
    return _num(v, _decimals(v, currency) if nd is None else nd)


def _pct(v: float, nd: int = 1) -> str:
    return ("+" if v >= 0 else "−") + _num(abs(v), nd) + "%"


def _span(text: str, color: str, size: int = 11, weight: int = 400) -> str:
    return (f'<span style="color:{color};font-size:{size}px;font-weight:{weight}">'
            f'{text}</span>')


_SEP = '<span style="font-size:11px">  ·  </span>'


def _rgba(zone: str, alpha: float) -> str:
    r, g, b = _ZONE_RGB[zone]
    return f"rgba({r},{g},{b},{alpha})"


# ── Đồ thị ──────────────────────────────────────────────────────────────────

def _zone_spans(bands: valuation.PriceBands, lo: float, hi: float, nd: int = 2
                ) -> list[tuple[str, float, float, str]]:
    """(vùng, đáy, đỉnh trong trục, chữ khoảng giá) cho mỗi vùng thấy được."""
    e = [_money(x, nd=nd) for x in bands.edges]
    full = [(valuation.VERDICTS[0], float("-inf"), bands.cheap, f"≤ {e[0]}"),
            (valuation.VERDICTS[1], bands.cheap, bands.rich, f"{e[0]} – {e[1]}"),
            (valuation.VERDICTS[2], bands.rich, bands.very_rich, f"{e[1]} – {e[2]}"),
            (valuation.VERDICTS[3], bands.very_rich, float("inf"), f"≥ {e[2]}")]
    return [(z, max(a, lo), min(b, hi), txt) for z, a, b, txt in full
            if min(b, hi) > max(a, lo)]


def _y_range(view: pd.DataFrame, close: float,
             bands: valuation.PriceBands | None) -> tuple[float, float]:
    vals = pd.concat([view["Low"], view["High"], view["MA50"], view["MA200"]]).dropna()
    lo, hi = float(vals.min()), float(vals.max())
    if bands:
        for e in bands.edges:
            if _within(e / close, NEAR_EDGE):
                lo, hi = min(lo, e), max(hi, e)
    pad = (hi - lo) * 0.06 or hi * 0.05
    return max(lo - pad, 0.0), hi + pad


def _hover(view: pd.DataFrame, freq: str, nd: int) -> list[str]:
    prev = view["Close"].shift(1)
    out = []
    for i, (day, r) in enumerate(view.iterrows()):
        head = (f"Tuần {r['first']:%d/%m} – {day:%d/%m/%Y}" if freq == "W"
                else f"{day:%d/%m/%Y}")
        ch = "" if pd.isna(prev.iloc[i]) else f"  ({_pct((r['Close'] / prev.iloc[i] - 1) * 100)})"
        out.append(f"{head}<br>Mở {_num(r['Open'], nd)} · Cao {_num(r['High'], nd)}"
                   f" · Thấp {_num(r['Low'], nd)}<br>Đóng <b>{_num(r['Close'], nd)}</b>{ch}")
    return out


def _gaps(view: pd.DataFrame) -> list[str]:
    """Ngày làm việc không có phiên (lễ) — để trục ngày khỏi để trống."""
    if view.empty:
        return []
    every = pd.bdate_range(view.index[0], view.index[-1])
    return [d.strftime("%Y-%m-%d") for d in every.difference(view.index)]


def _headline(hist: PriceHistory, snap: Snapshot, bands, c: dict) -> str:
    cur = hist.currency
    parts = [_span(_money(snap.close, cur), c["text"], 15, 600)
             + _span(f"  {_html.escape(cur)}  ·  đóng cửa {snap.day:%d/%m/%Y}", c["dim2"])]
    if snap.change is not None:
        parts.append(_span(_pct(snap.change), C_GREEN if snap.change >= 0 else C_RED, 11, 600)
                     + _span(" phiên gần nhất", c["dim2"]))
    if bands:
        zone = bands.zone(snap.close)
        chip = _span("vùng ", c["dim2"]) + _span(zone, ZONE_COLORS[zone], 12, 700)
        if not bands.reliable:
            chip += _span("  ⚠ kém tin cậy", C_GOLD, 11, 600)
        parts.append(chip)
    return _SEP.join(parts)


def _targets(snap: Snapshot, bands, cur: str, c: dict) -> str:
    """Dòng thứ hai: còn cách vùng kế bên bao xa, vị trí so với MA và đỉnh."""
    parts = []
    nd = _decimals(snap.close, cur)
    if bands:
        zone = bands.zone(snap.close)
        i = valuation.VERDICTS.index(zone)
        e = bands.edges
        if i > 0:
            below = valuation.VERDICTS[i - 1]
            sign = "≤" if i == 1 else "<"
            parts.append(_span(f"↓ về {below} {sign} {_money(e[i - 1], cur, nd)} ", ZONE_COLORS[below], 11, 600)
                         + _span(f"({_pct((e[i - 1] / snap.close - 1) * 100)})", c["dim2"]))
        if i < 3:
            above = valuation.VERDICTS[i + 1]
            sign = ">" if i == 0 else "≥"
            parts.append(_span(f"↑ sang {above} {sign} {_money(e[i], cur, nd)} ", ZONE_COLORS[above], 11, 600)
                         + _span(f"({_pct((e[i] / snap.close - 1) * 100)})", c["dim2"]))
    for name, ma in (("MA50", snap.ma50), ("MA200", snap.ma200)):
        if ma:
            d = (snap.close / ma - 1) * 100
            parts.append(_span(f"{'trên' if d >= 0 else 'dưới'} {name} ", c["dim2"])
                         + _span(_pct(d), C_GREEN if d >= 0 else C_RED, 11, 600))
    gap = (snap.close / snap.high52 - 1) * 100
    parts.append(_span("đang ở đỉnh 52 tuần", C_GREEN, 11, 600) if gap > -0.5
                 else _span("cách đỉnh 52 tuần ", c["dim2"]) + _span(_pct(gap), C_RED, 11, 600))
    return _SEP.join(parts)


def _add_bands(fig: go.Figure, bands, lo: float, hi: float, close: float,
               nd: int, dark: bool) -> None:
    spans = _zone_spans(bands, lo, hi, nd)
    faint, strong = (0.05, 0.16) if dark else (0.06, 0.15)
    price_f = (close - lo) / (hi - lo)
    for zone, a, b, txt in spans:
        for x0, alpha in ((0.0, faint), (RECENT, strong)):
            fig.add_hrect(y0=a, y1=b, x0=x0, x1=1, fillcolor=_rgba(zone, alpha),
                          line_width=0, layer="below", row=1, col=1)
        if (b - a) / (hi - lo) < 0.035:
            continue
        # Nhãn vùng nằm ở lề phải, giữa phần thấy được của vùng; vùng chứa giá
        # thì đặt vào nửa rộng hơn để khỏi đè lên nhãn giá
        mid = (a + b) / 2
        if a < close < b:
            mid = (a + close) / 2 if close - a > b - close else (close + b) / 2
        if abs((mid - lo) / (hi - lo) - price_f) < 0.05:
            mid = a + (b - a) * (0.25 if mid > close else 0.75)
        fig.add_annotation(
            text=f"<b>{zone}</b><br>{txt}", x=1, xref="paper", y=mid, yref="y",
            xanchor="left", xshift=8, yanchor="middle", showarrow=False, align="left",
            font=dict(size=10, color=ZONE_COLORS[zone], family="Segoe UI"))
    for zone, edge in zip(valuation.VERDICTS[1:], bands.edges):
        if lo < edge < hi:
            fig.add_shape(type="line", xref="x domain", x0=0, x1=1, yref="y", y0=edge, y1=edge,
                          line=dict(color=ZONE_COLORS[zone], width=1, dash="dot"),
                          opacity=0.55, layer="below", row=1, col=1)


def build_figure(hist: PriceHistory, rng: str = DEFAULT_RANGE,
                 bands: valuation.PriceBands | None = None, band_note: str = "",
                 dark: bool = False) -> go.Figure:
    c = _palette(dark)
    rng = rng if rng in RANGES else DEFAULT_RANGE
    label, _days, freq = RANGES[rng]
    bands, guard_note = usable_bands(hist, bands)
    band_note = band_note or guard_note
    snap = snapshot(hist.bars)
    view = window(hist.bars, rng)
    cur = hist.currency
    nd = _decimals(snap.close, cur)
    lo, hi = _y_range(view, snap.close, bands)

    fig = make_subplots(rows=2, cols=1, shared_xaxes=True, row_heights=[0.78, 0.22],
                        vertical_spacing=0.025)
    fig.add_trace(go.Candlestick(
        x=view.index, open=view["Open"], high=view["High"], low=view["Low"], close=view["Close"],
        increasing=dict(line=dict(color=C_GREEN, width=1), fillcolor=C_GREEN),
        decreasing=dict(line=dict(color=C_RED, width=1), fillcolor=C_RED),
        whiskerwidth=0, hovertext=_hover(view, freq, nd), hoverinfo="text", name="",
    ), row=1, col=1)
    for col, color, name in (("MA50", C_GOLD, "MA50"), ("MA200", C_BLUE, "MA200")):
        fig.add_trace(go.Scatter(
            x=view.index, y=view[col], mode="lines", line=dict(color=color, width=1.6),
            connectgaps=False, name=name,
            hovertemplate=f"{name} %{{y:,.{nd}f}}<extra></extra>"), row=1, col=1)
    up = view["Close"] >= view["Open"]
    fig.add_trace(go.Bar(
        x=view.index, y=view["Volume"] / 1e6, marker_color=[C_GREEN if u else C_RED for u in up],
        marker_opacity=0.5, marker_line_width=0, name="KL",
        hovertemplate="KL %{y:,.2f} triệu<extra></extra>"), row=2, col=1)

    if bands:
        _add_bands(fig, bands, lo, hi, snap.close, nd, dark)

    # Đỉnh 52 tuần: đoạn nét đứt phủ đúng 52 tuần gần nhất
    x0 = max(snap.day - pd.Timedelta(days=HIGH_DAYS), view.index[0])
    fig.add_shape(type="line", x0=x0, x1=snap.day, y0=snap.high52, y1=snap.high52,
                  line=dict(color=c["dim2"], width=1, dash="dash"), row=1, col=1)
    fig.add_annotation(x=x0, y=snap.high52, xanchor="left", yanchor="bottom", showarrow=False,
                       text=f"đỉnh 52 tuần {_money(snap.high52, cur, nd)}",
                       font=dict(size=9, color=c["dim2"], family="Segoe UI"), row=1, col=1)

    # Giá hiện tại: đường chấm và nhãn ở khoảng trống bên phải nến cuối
    zone = bands.zone(snap.close) if bands else ""
    zone_col = ZONE_COLORS[zone] if bands else c["text"]
    # chữ trắng trên nền vàng khó đọc; không có vùng thì nền là màu chữ của giao diện
    tag_text = "#1A202C" if zone == "Đắt" else "#FFFFFF" if bands else c["bg"]
    fig.add_hline(y=snap.close, line=dict(color=zone_col, width=1, dash="dot"),
                  opacity=0.8, row=1, col=1)
    fig.add_annotation(x=1, xref="x domain", y=snap.close, yref="y", xanchor="right",
                       yanchor="middle", showarrow=False, text=f"<b>{_money(snap.close, cur, nd)}</b>",
                       bgcolor=zone_col, borderpad=3,
                       font=dict(size=10, color=tag_text, family="Segoe UI"))

    span_days = (view.index[-1] - view.index[0]).days
    x_end = view.index[-1] + pd.Timedelta(days=max(round(span_days * X_PAD), 4 if freq == "D" else 21))
    breaks = [dict(bounds=["sat", "mon"]), dict(values=_gaps(view))] if freq == "D" else []
    fig.update_xaxes(range=[view.index[0] - pd.Timedelta(days=2 if freq == "D" else 5), x_end],
                     rangebreaks=breaks, rangeslider_visible=False, gridcolor=c["grid"],
                     showline=True, linecolor=c["grid"], zeroline=False,
                     tickfont=dict(size=9, color=c["dim"]), hoverformat="%d/%m/%Y",
                     tickformat="%m/%Y" if freq == "W" else "%d/%m")
    fig.update_yaxes(gridcolor=c["grid"], zerolinecolor=c["zero"],
                     tickfont=dict(size=9, color=c["dim"]))
    fig.update_yaxes(range=[lo, hi], tickformat=f",.{0 if hi >= 1000 or nd == 0 else 2}f",
                     row=1, col=1)
    vol_max = float(view["Volume"].max() / 1e6) if len(view) else 0.0
    fig.update_yaxes(rangemode="tozero", tickformat=",.0f" if vol_max >= 20 else ",.1f",
                     nticks=3, row=2, col=1)

    symbol = (f'<span style="color:{c["dim2"]};font-size:11px">  ({_html.escape(hist.symbol)})</span>'
              if hist.symbol.upper() != hist.ticker.upper() else "")
    fig.update_layout(
        paper_bgcolor=c["bg"], plot_bgcolor=c["panel"],
        font=dict(family="Segoe UI, system-ui, sans-serif", color=c["dim"], size=10),
        title=dict(text=(f'<span style="color:{c["text"]};font-size:14px;font-weight:600;'
                         f'letter-spacing:1px">{_html.escape(hist.ticker)}</span>{symbol}'
                         f'<span style="color:{c["dim2"]};font-size:12px;font-weight:300">'
                         f'  ·  Giá · nến {"tuần" if freq == "W" else "ngày"} · {label}</span>'),
                   x=0.0, xanchor="left", y=1, yref="container", yanchor="top",
                   pad=dict(l=14, t=10)),
        showlegend=False, margin=dict(l=58, r=124 if bands else 22, t=100, b=66),
        hovermode="x unified", bargap=0.15, separators=",.",
        hoverlabel=dict(bgcolor=c["hover_bg"], bordercolor=c["hover_border"],
                        font=dict(family="Segoe UI", size=11, color=c["text"])),
    )
    fig.add_annotation(text=_headline(hist, snap, bands, c), xref="paper", yref="paper",
                       x=0, y=1.0, yshift=34, xanchor="left", yanchor="bottom",
                       showarrow=False, align="left")
    fig.add_annotation(text=_targets(snap, bands, cur, c), xref="paper", yref="paper",
                       x=0, y=1.0, yshift=10, xanchor="left", yanchor="bottom",
                       showarrow=False, align="left")
    legend = (_span("━ MA50", C_GOLD, 9, 600) + _span("   ━ MA200", C_BLUE, 9, 600)
              + _span("   ┄ đỉnh 52 tuần", c["dim2"], 9))
    # chú giải nằm cùng dòng tiêu đề, góc phải — hai dòng số liệu cần cả bề ngang
    fig.add_annotation(text=legend, xref="paper", yref="paper", x=1, y=1.0, yshift=74,
                       xanchor="right", yanchor="bottom", showarrow=False)

    foot = ["Nguồn: Yahoo Finance, giá đã điều chỉnh cổ tức và chia tách. "
            "MA tính trên giá đóng cửa ngày. Khối lượng: triệu cổ phiếu."]
    if bands:
        flow = "dòng tiền tự do" if bands.basis == "FCF" else "lợi nhuận ròng"
        foot.append(f"Vùng định giá tính từ {flow} lúc quét — chỉ đúng hiện tại, "
                    f"không áp cho giá quá khứ. Không phải tín hiệu mua bán.")
    elif band_note:
        foot.append(f"Không vẽ vùng định giá: {band_note}.")
    fig.add_annotation(text="<br>".join(foot), xref="paper", yref="paper", x=0, y=0,
                       yshift=-58, xanchor="left", yanchor="bottom", showarrow=False,
                       align="left", font=dict(size=9, color=c["dim2"], family="Segoe UI"))
    return fig


def build_html(hist: PriceHistory, rng: str = DEFAULT_RANGE,
               bands: valuation.PriceBands | None = None, band_note: str = "",
               dark: bool = False) -> str:
    fig = build_figure(hist, rng, bands, band_note, dark)
    fig.update_layout(autosize=True)
    div = fig.to_html(include_plotlyjs="cdn", full_html=False,
                      config={"displayModeBar": False, "displaylogo": False,
                              "scrollZoom": False})
    return _page(div, dark)
