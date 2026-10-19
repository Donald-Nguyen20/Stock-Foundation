"""
Theo dõi thay đổi giữa các lần quét và kiểm chứng tín hiệu
==========================================================
Một tín hiệu đứng yên không nói được nhiều. Mã vừa rơi từ TỐT & THỰC xuống
KHÁ đáng xem hơn hẳn mã nằm ở KHÁ cả năm — nên mỗi lần quét được đặt cạnh
lần quét trước của cùng thị trường.

Phần kiểm chứng trả lời câu hỏi khó chịu hơn: các nhãn này trong quá khứ
có thực sự đi trước lợi nhuận không? Mỗi ảnh chụp đủ cũ được đối chiếu với
giá thật sau 1/3/6/12 tháng, so với S&P 500 cùng kỳ.

Chỉ dùng ảnh chụp của chính máy này. Không có lịch sử nghĩa là chưa kiểm
được — không có số giả lập nào thay thế.
"""
from __future__ import annotations

import html as _html
import re
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path

import pandas as pd

GOOD = "🟢 TỐT & THỰC"
FLAG = "🔴 CỜ ĐỎ"
SIGNAL_RANK = {GOOD: 6, "🔵 KHÁ": 5, "🟠 THẬT NHƯNG YẾU": 4, "🟡 THEO DÕI": 3,
               "⚫ LOẠI": 2, FLAG: 1}
SIGNAL_ORDER = (GOOD, "🔵 KHÁ", "🟠 THẬT NHƯNG YẾU", "🟡 THEO DÕI", "⚫ LOẠI", FLAG,
                "⚪ THIẾU DỮ LIỆU")

CHANGE_COLS = ("Thay đổi", "Thay đổi_rank", "Thay đổi_chi tiết", "Tín hiệu trước")

US_TAGS = frozenset({"AMERICA", "NASDAQ", "NYSE"})
BENCH = "^GSPC"
CHUNK = 50                    # số mã mỗi lần gọi yf.download
HORIZONS = (1, 3, 6, 12)      # tháng
MIN_N = 5                     # dưới mức này nhóm chỉ để tham khảo
YF_SHARE = 0.3                # tỷ lệ mã có dữ liệu nhiều năm để coi là quét kèm yfinance

_SNAP_RE = re.compile(r"^(\d{8})_([A-Za-z0-9]+)_top(\d+)\.csv\.gz$")


# ── Ảnh chụp ────────────────────────────────────────────────────────────────

@dataclass(frozen=True)
class Snap:
    path: Path
    day: date
    tag: str
    top: int


def market_tag(market: str) -> str:
    """Cùng quy tắc đặt tên với save_snapshot."""
    return re.sub(r"[^A-Za-z0-9]", "", str(market))[:16] or "unknown"


def parse_name(path: Path) -> Snap | None:
    m = _SNAP_RE.match(Path(path).name)
    if not m:
        return None
    try:
        day = datetime.strptime(m.group(1), "%Y%m%d").date()
    except ValueError:
        return None
    return Snap(Path(path), day, m.group(2).upper(), int(m.group(3)))


def list_snaps(folder: Path) -> list[Snap]:
    try:
        found = (parse_name(p) for p in Path(folder).glob("*.csv.gz"))
        return sorted((s for s in found if s), key=lambda s: (s.day, s.tag, s.top))
    except OSError:
        return []


def _bool_back(v):
    """CSV trả True/False thành chuỗi khi cột có ô trống."""
    if isinstance(v, str):
        if v == "True":
            return True
        if v == "False":
            return False
    return v


def read_snap(path: Path) -> pd.DataFrame:
    # Chỉ ô rỗng mới là thiếu — mặc định pandas đọc mã "NA" (Nano Labs) thành NaN
    df = pd.read_csv(path, compression="gzip", encoding="utf-8-sig", low_memory=False,
                     keep_default_na=False, na_values=[""])
    for c in df.columns:
        if c.startswith(("FIN_", "REAL_")) and df[c].dtype == object:
            df[c] = df[c].map(_bool_back)
    return df


def load_previous(folder: Path, market: str, today: date | None = None,
                  fallback: tuple | None = None) -> tuple[pd.DataFrame | None, date | None]:
    """Lần quét gần nhất của cùng thị trường, trước hôm nay.

    `fallback` = (df, saved_at, market) của bộ nhớ đệm GUI — dùng khi nó mới
    hơn ảnh chụp mới nhất (ví dụ ảnh chụp chưa từng được ghi)."""
    today = today or date.today()
    tag = market_tag(market).upper()
    cands = [s for s in list_snaps(folder) if s.tag == tag and s.day < today]
    best = max(cands, key=lambda s: (s.day, s.top)) if cands else None

    fb_df, fb_day = None, None
    if fallback:
        df, saved_at, fb_market = fallback
        if (df is not None and isinstance(saved_at, datetime)
                and market_tag(fb_market).upper() == tag and saved_at.date() < today):
            fb_df, fb_day = df, saved_at.date()

    if fb_df is not None and (best is None or fb_day > best.day):
        return fb_df, fb_day
    if best is not None:
        try:
            return read_snap(best.path), best.day
        except Exception:
            return (fb_df, fb_day) if fb_df is not None else (None, None)
    return None, None


# ── So sánh hai lần quét ────────────────────────────────────────────────────

def short_signal(sig) -> str:
    """'🟢 TỐT & THỰC' → 'TỐT & THỰC'."""
    s = str(sig or "").strip()
    return s.split(" ", 1)[1] if " " in s and not s[0].isalnum() else s


def _num(v) -> float | None:
    try:
        f = float(v)
    except (TypeError, ValueError):
        return None
    return None if pd.isna(f) else f


def _text(v) -> str:
    return "" if v is None or (isinstance(v, float) and pd.isna(v)) else str(v)


def _vn_money(v: float) -> str:
    """1.234,56 — ngăn cách kiểu Việt Nam."""
    return f"{v:,.2f}".replace(",", " ").replace(".", ",").replace(" ", ".")


def _fmt(v: float) -> str:
    """+1,5 / −2,0 — có dấu, phẩy thập phân."""
    return ("+" if v >= 0 else "−") + f"{abs(v):.1f}".replace(".", ",")


def _criteria_flips(cur: dict, prev: dict, labels: dict) -> tuple[list[str], list[str]]:
    """(tiêu chí mới đạt, tiêu chí mới trượt) — chỉ tính khi cả hai lần đều có kết quả."""
    gained, lost = [], []
    for col, label in labels.items():
        a, b = prev.get(col), cur.get(col)
        if isinstance(a, bool) and isinstance(b, bool) and a != b:
            (gained if b else lost).append(label)
    return gained, lost


def _score(row: dict, grp: str) -> str:
    s, a = _num(row.get(f"{grp}_Score")), _num(row.get(f"{grp}_App"))
    return f"{int(s)}/{int(a)}" if s is not None and a is not None else "—"


def _describe(cur: dict, prev: dict | None, prev_day: date, labels: dict) -> tuple[str, float, str]:
    """(nhãn ngắn, giá trị sắp xếp, chi tiết cho tooltip) của một mã."""
    if prev is None:
        return "MỚI", 0.5, f"Không có trong lần quét {prev_day:%d/%m/%Y}."
    cs, ps = _text(cur.get("Signal")), _text(prev.get("Signal"))
    cr, pr = SIGNAL_RANK.get(cs), SIGNAL_RANK.get(ps)
    gained, lost = _criteria_flips(cur, prev, labels)
    cv, pv = _text(cur.get("Định giá")), _text(prev.get("Định giá"))

    lines = [f"So với lần quét {prev_day:%d/%m/%Y}:"]
    lines.append(f"Tín hiệu: {short_signal(ps) or '—'} → {short_signal(cs) or '—'}" if cs != ps
                 else f"Tín hiệu giữ nguyên: {short_signal(cs) or '—'}")
    lines.append(f"Tài chính {_score(prev, 'FIN')} → {_score(cur, 'FIN')}   ·   "
                 f"Thực {_score(prev, 'REAL')} → {_score(cur, 'REAL')}")
    if gained:
        lines.append("Đạt thêm: " + ", ".join(gained))
    if lost:
        lines.append("Mất: " + ", ".join(lost))
    if cv and pv and cv != pv:
        lines.append(f"Định giá: {pv} → {cv}")
    p0, p1 = _num(prev.get("Price ($)")), _num(cur.get("Price ($)"))
    if p0 and p1:
        ch = (p1 / p0 - 1) * 100
        lines.append(f"Giá: {_vn_money(p0)} → {_vn_money(p1)} ({_fmt(ch)}%)")
    detail = "\n".join(lines)

    if cs != ps:
        if cr is not None and pr is not None:
            arrow = "▲" if cr > pr else "▼"
            return f"{arrow} từ {short_signal(ps)}", float(cr - pr), detail
        return f"≠ từ {short_signal(ps) or '—'}", 0.0, detail
    net = len(gained) - len(lost)
    if gained or lost:
        return f"• {len(gained) + len(lost)} tiêu chí", net / 100, detail
    if cv and pv and cv != pv:
        return f"• giá {cv}", 0.0, detail
    return "", 0.0, detail


def annotate_changes(cur: pd.DataFrame, prev: pd.DataFrame | None, prev_day: date | None,
                     labels: dict | None = None) -> pd.DataFrame:
    """Bản sao của cur có thêm CHANGE_COLS. Không có lần trước thì để trống."""
    out = cur.copy()
    if prev is None or prev_day is None or "Ticker" not in prev.columns:
        for c in CHANGE_COLS:
            out[c] = "" if c != "Thay đổi_rank" else 0.0
        return out
    labels = {c: l for c, l in (labels or {}).items() if c in out.columns and c in prev.columns}
    prev_rows = {str(r["Ticker"]): r for r in prev.to_dict("records")}
    res = []
    for r in out.to_dict("records"):
        p = prev_rows.get(str(r.get("Ticker")))
        text, rank, detail = _describe(r, p, prev_day, labels)
        res.append((text, rank, detail, _text(p.get("Signal")) if p else ""))
    for i, c in enumerate(CHANGE_COLS):
        out[c] = [x[i] for x in res]
    return out


def _yf_share(df: pd.DataFrame) -> float | None:
    if "Số năm" not in df.columns or not len(df):
        return None
    return float(pd.to_numeric(df["Số năm"], errors="coerce").notna().mean())


def change_summary(cur: pd.DataFrame, prev: pd.DataFrame | None,
                   prev_day: date | None) -> dict | None:
    """Tóm tắt để hiện trên trang tổng quan. `cur` phải đã qua annotate_changes."""
    if prev is None or prev_day is None or "Tín hiệu trước" not in cur.columns:
        return None
    up, down, new_good, new_flags = [], [], [], []
    for r in cur.to_dict("records"):
        t, cs, ps = str(r.get("Ticker")), _text(r.get("Signal")), _text(r.get("Tín hiệu trước"))
        if not ps or cs == ps:
            continue
        cr, pr = SIGNAL_RANK.get(cs), SIGNAL_RANK.get(ps)
        item = (t, ps, cs)
        if cr is not None and pr is not None:
            (up if cr > pr else down).append(item)
        if cs == GOOD:
            new_good.append(item)
        if cs == FLAG:
            new_flags.append(item)

    cur_t = set(cur["Ticker"].astype(str))
    prev_t = set(prev["Ticker"].astype(str))
    prev_sig = dict(zip(prev["Ticker"].astype(str), prev.get("Signal", pd.Series(dtype=object))))
    cur_sig = dict(zip(cur["Ticker"].astype(str), cur["Signal"]))
    same_size = abs(len(cur) - len(prev)) <= max(5, 0.05 * len(prev))
    # Đổi phạm vi quét thì mã rơi khỏi danh sách không có nghĩa là xấu đi
    lost_good = sorted(t for t, s in prev_sig.items()
                       if s == GOOD and cur_sig.get(t) != GOOD and (same_size or t in cur_sig))

    warnings = []
    if not same_size:
        warnings.append(f"Lần trước quét {len(prev)} mã, lần này {len(cur)} — mã vào/ra danh "
                        f"sách phần lớn do đổi phạm vi quét nên không liệt kê.")
    a, b = _yf_share(cur), _yf_share(prev)
    if a is not None and b is not None and (a >= YF_SHARE) != (b >= YF_SHARE):
        on = "lần này" if a >= YF_SHARE else "lần trước"
        warnings.append(f"Chỉ {on} quét kèm dữ liệu yfinance nhiều năm — một phần thay đổi tín "
                        f"hiệu đến từ việc có thêm/mất dữ liệu, không phải từ doanh nghiệp.")

    return {
        "prev_date": prev_day.isoformat(),
        "up": up, "down": down, "new_good": new_good, "new_flags": new_flags,
        "lost_good": lost_good,
        "entered": sorted(cur_t - prev_t) if same_size else None,
        "exited": sorted(prev_t - cur_t) if same_size else None,
        "warnings": warnings,
    }


# ── Kiểm chứng tín hiệu bằng giá thật ──────────────────────────────────────

def validation_snaps(folder: Path, today: date | None = None) -> list[Snap]:
    """Ảnh chụp Mỹ, mỗi (thị trường, tháng) lấy một — ngày sớm nhất trong tháng,
    bản quét rộng nhất trong ngày đó. Quét nhiều lần một tháng không nhân bản quan sát."""
    snaps = [s for s in list_snaps(folder) if s.tag in US_TAGS]
    chosen: dict[tuple, Snap] = {}
    for s in sorted(snaps, key=lambda s: (s.day, -s.top)):
        chosen.setdefault((s.tag, s.day.year, s.day.month), s)
    return sorted(chosen.values(), key=lambda s: s.day)


def observations(snaps: list[Snap]) -> pd.DataFrame:
    """Mỗi dòng: một mã trong một ảnh chụp. Một mã chỉ tính một lần mỗi tháng
    dù có mặt trong cả ảnh AMERICA lẫn NASDAQ."""
    frames = []
    for s in snaps:
        try:
            df = read_snap(s.path)
        except Exception:
            continue
        if "Ticker" not in df.columns or "Signal" not in df.columns:
            continue
        keep = pd.DataFrame({
            "day": pd.Timestamp(s.day), "Ticker": df["Ticker"].astype(str),
            "Signal": df["Signal"].astype(str),
            "Định giá": (df["Định giá"].fillna("").astype(str)
                         if "Định giá" in df.columns else ""),
        })
        frames.append(keep)
    if not frames:
        return pd.DataFrame(columns=["day", "Ticker", "Signal", "Định giá"])
    obs = pd.concat(frames, ignore_index=True)
    obs["ym"] = obs["day"].dt.to_period("M")
    return obs.drop_duplicates(["Ticker", "ym"]).drop(columns="ym").reset_index(drop=True)


def yahoo_symbol(ticker: str) -> str:
    return str(ticker).replace(".", "-")


def download_prices(tickers: list[str], start: date, stop=None) -> pd.DataFrame:
    """Giá đóng cửa đã điều chỉnh, cột là mã gốc (BRK.A chứ không phải BRK-A).

    Tải theo lô để `stop()` (ví dụ QThread.isInterruptionRequested) được hỏi
    giữa các lô — yf.download không ngắt được giữa chừng."""
    import yfinance as yf
    syms = {yahoo_symbol(t): t for t in tickers}
    syms[BENCH] = BENCH
    order = list(syms)
    parts = []
    for i in range(0, len(order), CHUNK):
        if stop is not None and stop():
            raise InterruptedError("đã dừng giữa chừng")
        chunk = order[i:i + CHUNK]
        raw = yf.download(chunk, start=start.isoformat(), auto_adjust=True,
                          progress=False, threads=True)
        if raw is None or raw.empty:
            continue
        if isinstance(raw.columns, pd.MultiIndex):
            close = raw["Close"]
        else:                                   # yfinance cũ trả cột phẳng khi lô chỉ có 1 mã
            close = raw[["Close"]].set_axis(chunk[:1], axis=1)
        parts.append(close.to_frame() if isinstance(close, pd.Series) else close)
    if not parts:
        return pd.DataFrame()
    return pd.concat(parts, axis=1).rename(columns=syms).sort_index()


def _price_at(series: pd.Series, when: pd.Timestamp) -> tuple[pd.Timestamp, float] | None:
    """Giá đóng cửa đầu tiên từ ngày `when` trở đi."""
    s = series.dropna()
    i = s.index.searchsorted(when)
    if i >= len(s):
        return None
    return s.index[i], float(s.iloc[i])


def returns(obs: pd.DataFrame, prices: pd.DataFrame,
            horizons: tuple = HORIZONS) -> pd.DataFrame:
    """Lợi nhuận của từng quan sát và phần vượt S&P 500 ở mỗi mốc đã tới hạn."""
    if BENCH not in prices.columns or prices.empty:
        return pd.DataFrame()
    last = prices.index.max()
    bench = prices[BENCH]
    rows = []
    for o in obs.to_dict("records"):
        t = o["Ticker"]
        if t not in prices.columns:
            continue
        entry = _price_at(prices[t], o["day"])
        b_entry = _price_at(bench, o["day"])
        if not entry or not b_entry or entry[1] <= 0 or b_entry[1] <= 0:
            continue
        for h in horizons:
            due = o["day"] + pd.DateOffset(months=h)
            if due > last:
                continue
            ex, b_ex = _price_at(prices[t], due), _price_at(bench, due)
            if not ex or not b_ex:
                continue
            r = ex[1] / entry[1] - 1
            br = b_ex[1] / b_entry[1] - 1
            rows.append({**o, "h": h, "ret": r * 100, "excess": (r - br) * 100})
    return pd.DataFrame(rows)


def group_stats(rets: pd.DataFrame) -> pd.DataFrame:
    """Mỗi nhóm × mốc: số quan sát, vượt S&P trung bình/trung vị, tỷ lệ thắng."""
    if rets.empty:
        return pd.DataFrame()
    parts = [rets.assign(kind="Toàn bộ", group="Toàn bộ danh sách"),
             rets.assign(kind="Tín hiệu", group=rets["Signal"])]
    if (rets["Định giá"] != "").any():
        v = rets[rets["Định giá"] != ""]
        parts.append(v.assign(kind="Định giá", group=v["Định giá"]))
    g = pd.concat(parts).groupby(["kind", "group", "h"])
    st = g.agg(n=("excess", "size"), mean_excess=("excess", "mean"),
               median_excess=("excess", "median"), mean_ret=("ret", "mean"),
               hit=("excess", lambda x: (x > 0).mean() * 100)).reset_index()
    return st


@dataclass(frozen=True)
class Report:
    stats: pd.DataFrame
    n_snaps: int = 0
    n_obs: int = 0
    first_day: date | None = None
    price_day: date | None = None
    missing: tuple = ()
    message: str = ""


def run_validation(folder: Path, today: date | None = None, downloader=download_prices) -> Report:
    """Toàn bộ quy trình. `downloader(tickers, start)` thay được khi kiểm thử."""
    today = today or date.today()
    snaps = validation_snaps(folder, today)
    if not snaps:
        return Report(pd.DataFrame(), message=(
            "Chưa có ảnh chụp nào của thị trường Mỹ. Mỗi lần quét tự lưu một ảnh "
            "vào thư mục snapshots; một tháng sau lần quét đầu tiên là kiểm được mốc 1 tháng."))
    first = snaps[0].day
    ready = [s for s in snaps if (pd.Timestamp(s.day) + pd.DateOffset(months=HORIZONS[0])).date() <= today]
    if not ready:
        due = (pd.Timestamp(first) + pd.DateOffset(months=HORIZONS[0])).date()
        return Report(pd.DataFrame(), n_snaps=len(snaps), first_day=first, message=(
            f"Ảnh chụp sớm nhất là ngày {first:%d/%m/%Y}. Mốc 1 tháng kiểm được từ "
            f"ngày {due:%d/%m/%Y} — cứ quét đều mỗi tháng là số liệu tự tích lũy."))
    obs = observations(ready)
    if obs.empty:
        return Report(pd.DataFrame(), n_snaps=len(ready), first_day=first,
                      message="Không đọc được ảnh chụp nào (tệp hỏng hoặc thiếu cột Signal).")
    tickers = sorted(obs["Ticker"].unique())
    try:
        prices = downloader(tickers, ready[0].day)
    except Exception as e:
        return Report(pd.DataFrame(), n_snaps=len(ready), n_obs=len(obs), first_day=first,
                      message=f"Không tải được giá từ Yahoo Finance: {e}")
    missing = tuple(t for t in tickers if t not in prices.columns or prices[t].dropna().empty)
    rets = returns(obs, prices)
    price_day = prices.index.max().date() if len(prices) else None
    if rets.empty:
        return Report(pd.DataFrame(), n_snaps=len(ready), n_obs=len(obs), first_day=first,
                      price_day=price_day, missing=missing,
                      message="Có ảnh chụp nhưng chưa mốc nào đủ dữ liệu giá để tính.")
    return Report(group_stats(rets), n_snaps=len(ready), n_obs=len(obs), first_day=first,
                  price_day=price_day, missing=missing)


# ── Hiển thị ────────────────────────────────────────────────────────────────

def report_html(rep: Report, c: dict) -> str:
    """HTML cho QTextBrowser. `c` là bảng màu của giao diện: text, dim, faint,
    green, red, amber, border, head."""
    esc = _html.escape
    css = (f"body{{font-family:'Segoe UI';font-size:12px;color:{c['text']}}}"
           f"h2{{font-size:15px;margin:0 0 4px 0}}"
           f"h3{{font-size:11px;letter-spacing:1px;color:{c['dim']};margin:18px 0 6px 0}}"
           f"p{{margin:4px 0;line-height:150%}}"
           f"td,th{{padding:4px 8px;border-bottom:1px solid {c['border']}}}"
           f"th{{color:{c['dim']};font-weight:600;text-align:right;background:{c['head']}}}")
    head = "<h2>Hiệu quả tín hiệu — thị trường Mỹ</h2>"
    if rep.stats is None or rep.stats.empty:
        body = f"<p>{esc(rep.message)}</p>"
        return f"<html><head><style>{css}</style></head><body>{head}{body}</body></html>"

    meta = (f"{rep.n_snaps} ảnh chụp từ {rep.first_day:%d/%m/%Y} · {rep.n_obs:,} quan sát "
            f"(mã × tháng) · giá tới {rep.price_day:%d/%m/%Y}").replace(",", ".")
    hs = sorted(rep.stats["h"].unique())

    def cell(txt: str, tone: float, n: int) -> str:
        """Ô số: xanh khi tone > 0, đỏ khi < 0, mờ khi quá ít quan sát."""
        col = (c["faint"] if n < MIN_N else c["green"] if tone > 0
               else c["red"] if tone < 0 else c["text"])
        return f'<td align="right" style="color:{col}">{txt}</td>'

    def table(kind: str, order: tuple) -> str:
        st = rep.stats[rep.stats["kind"].isin([kind, "Toàn bộ"])]
        groups = [g for g in order if g in set(st["group"])]
        groups = ["Toàn bộ danh sách"] + [g for g in groups if g != "Toàn bộ danh sách"]
        hdr = "".join(f'<th colspan="3" style="text-align:center">{h} tháng</th>' for h in hs)
        sub = "".join("<th>vượt S&amp;P</th><th>thắng</th><th>n</th>" for _ in hs)
        rows = []
        for g in groups:
            tds = []
            for h in hs:
                r = st[(st["group"] == g) & (st["h"] == h)]
                if r.empty:
                    tds.append('<td align="right">—</td>' * 3)
                    continue
                r = r.iloc[0]
                n = int(r["n"])
                ex, hit = float(r["mean_excess"]), float(r["hit"])
                tds.append(cell(_fmt(ex), ex, n)
                           + cell(f"{hit:.0f}%", hit - 50, n)
                           + f'<td align="right" style="color:{c["dim"]}">{n}</td>')
            style = "font-weight:600" if g == "Toàn bộ danh sách" else ""
            rows.append(f'<tr><td style="{style}">{esc(g)}</td>{"".join(tds)}</tr>')
        return (f'<table cellspacing="0" width="100%"><tr><th></th>{hdr}</tr>'
                f'<tr><th style="text-align:left">Nhóm</th>{sub}</tr>{"".join(rows)}</table>')

    parts = [head, f'<p style="color:{c["dim"]}">{esc(meta)}</p>',
             "<h3>THEO TÍN HIỆU CHẤT LƯỢNG</h3>", table("Tín hiệu", SIGNAL_ORDER)]
    if (rep.stats["kind"] == "Định giá").any():
        parts += ["<h3>THEO ĐỊNH GIÁ</h3>", table("Định giá", ("Rẻ", "Hợp lý", "Đắt", "Rất đắt"))]
    notes = [
        "«vượt S&amp;P» là lợi nhuận trung bình trừ lợi nhuận S&amp;P 500 cùng kỳ, tính bằng "
        "điểm phần trăm; «thắng» là tỷ lệ quan sát vượt S&amp;P. Nhóm tốt phải vượt dòng "
        "<b>Toàn bộ danh sách</b> — vì cả danh sách vốn hóa lớn đã có thể đánh bại hoặc thua chỉ số.",
        f"Nhóm dưới {MIN_N} quan sát in mờ: quá ít để kết luận.",
        "Các quan sát chồng lấn nhau (ảnh tháng 1 và tháng 2 cùng chứa giai đoạn tháng 2–12), "
        "nên số quan sát lớn hơn lượng bằng chứng độc lập thật. Vài năm dữ liệu mới đủ tin.",
    ]
    if rep.missing:
        notes.append(f"{len(rep.missing)} mã không có giá trên Yahoo (đã hủy niêm yết, sáp nhập, "
                     f"đổi mã) bị loại — nếu đó là mã sụp đổ thì kết quả đẹp hơn thực tế: "
                     + esc(", ".join(rep.missing[:12])) + ("…" if len(rep.missing) > 12 else ""))
    parts.append(f'<h3>CÁCH ĐỌC</h3>' + "".join(
        f'<p style="color:{c["dim"]}">• {n}</p>' for n in notes))
    return f"<html><head><style>{css}</style></head><body>{''.join(parts)}</body></html>"
