# -*- coding: utf-8 -*-
"""Biểu đồ so sánh toàn bộ mã đã quét.

Bốn biểu đồ sẵn có đều là chuỗi thời gian của một mã — chúng giúp hiểu sâu
một doanh nghiệp. Nhưng chọn cổ phiếu là việc so giữa các mã, mà cho tới
giờ app không có hình nào làm việc đó.

Các bản đồ ở đây dùng đúng dữ liệu đã có trong bảng, không gọi mạng thêm.
Mỗi lần chỉ vẽ một bản đồ cho rộng chỗ; khung biểu đồ có ô chọn để đổi.

Ba nguyên tắc để bản đồ không nói sai:
  • Không mã nào biến mất lặng lẽ — mã nằm ngoài khung được ghim ở mép bằng
    hình tam giác chỉ hướng, mã thiếu số liệu được đếm ở dòng chú thích.
  • Mã trùng điểm được xếp tỏa quanh vị trí thật thay vì chồng lên nhau.
  • Nhóm Tài chính đo bằng ROE và E/P — ROIC và dòng tiền tự do của ngân
    hàng không mang nghĩa như doanh nghiệp thường.
"""
from __future__ import annotations

import math

import pandas as pd

SIG_COLORS = {
    "🟢 TỐT & THỰC":     "#34C472",
    "🔵 KHÁ":            "#4A9EFF",
    "🔴 CỜ ĐỎ":          "#E8483D",
    "🟠 THẬT NHƯNG YẾU": "#E87C3D",
    "🟡 THEO DÕI":       "#E8A93D",
    "⚫ LOẠI":            "#7A8899",
    "⚪ THIẾU DỮ LIỆU":   "#4A5568",
}

# Khớp với _FIN_SECTORS bên bộ chấm điểm
FIN_SECTORS = ("Financial Services", "Finance")

VIEWS = {
    "quality": "Tài chính × Thực",
    "value":   "Sinh lời × Lợi suất",
    "growth":  "Sinh lời × Tăng trưởng",
    "timing":  "Giá so với đỉnh",
}
DEFAULT_VIEW = "quality"

# Đường dẫn giả mà trang bấm vào để báo mã được chọn — khung Qt chặn lại
PICK_HOST = "pick.local"

_ROIC_LINE = 15     # ngưỡng ROIC của tiêu chí Tài chính
_ROE_LINE  = 12     # ngưỡng thay thế cho nhóm Tài chính
_PEERS_LABEL_MAX = 15


def _theme(dark: bool) -> dict:
    if dark:
        return dict(bg="#0A1628", panel="#0B0E18", grid="#1A2133",
                    text="#DCE4EE", dim="#7A8899", zero="#243044")
    return dict(bg="#FFFFFF", panel="#F8FAFC", grid="#E2E8F0",
                text="#1A202C", dim="#64748B", zero="#CBD5E0")


def _num(df, col):
    if col not in df.columns:
        return pd.Series(float("nan"), index=df.index)
    return pd.to_numeric(df[col], errors="coerce")


def _text(df, col, default=""):
    if col not in df.columns:
        return pd.Series(default, index=df.index)
    return df[col].fillna(default).astype(str)


# ═════════════════════════════════════════════════════════════════════════
# Chuẩn bị số liệu
# ═════════════════════════════════════════════════════════════════════════
def _prep(df: pd.DataFrame) -> pd.DataFrame:
    """Gom các trường bản đồ cần về một bảng gọn, đã đổi sang đơn vị %."""
    tick   = _text(df, "Ticker").str.upper()
    sector = _text(df, "Sector", "—")
    is_fin = sector.isin(FIN_SECTORS)
    mcap   = _num(df, "MCap ($B)")
    pe     = _num(df, "P/E")

    # Lợi suất tiền mặt: FCF / vốn hoá. Ngân hàng thì dòng tiền tự do lẫn cả
    # tiền gửi khách hàng, nên dùng lợi suất lợi nhuận E/P thay thế.
    fcf_yield = _num(df, "FCF_abs") / (mcap * 1e9) * 100
    e_yield   = (100 / pe).where(pe > 0)

    return pd.DataFrame({
        "ticker":  tick,
        "name":    _text(df, "Tên Công Ty").str.slice(0, 38),
        "sector":  sector,
        "signal":  _text(df, "Signal"),
        "is_fin":  is_fin,
        "mcap":    mcap,
        "fin":     _num(df, "FIN_Pct") * 100,
        "real":    _num(df, "REAL_Pct") * 100,
        "ret":     _num(df, "ROE%").where(is_fin, _num(df, "ROIC%")),
        "yield":   e_yield.where(is_fin, fcf_yield),
        "growth":  _num(df, "Rev Growth%"),
        "peak":    _num(df, "% từ đỉnh"),
        "ma250":   _num(df, "Giá/MA250%"),
    }, index=df.index)


def _marker_sizes(mcap: pd.Series) -> pd.Series:
    """Cỡ chấm theo log vốn hoá — mã nghìn tỷ không nuốt mất mã vài chục tỷ."""
    lg = mcap.where(mcap > 0).apply(lambda v: math.log10(v) if v == v else float("nan"))
    lo, hi = lg.min(), lg.max()
    if not (hi > lo):
        return pd.Series(10.0, index=mcap.index)
    return (7 + 13 * (lg - lo) / (hi - lo)).fillna(8.0)


def _spread_ties(x: pd.Series, y: pd.Series, weight: pd.Series):
    """Xếp tỏa các mã trùng điểm theo hình hoa hướng dương.

    Điểm chỉ nhận vài giá trị rời rạc nên hàng chục mã có thể trùng một chỗ.
    Mã lớn nhất ở tâm, các mã khác tỏa dần ra; bán kính tối đa giữ dưới nửa
    bước điểm để không lấn sang ô bên cạnh. Kết quả cố định giữa các lần vẽ.
    """
    px, py = x.copy(), y.copy()
    key = x.round(0).astype(str) + "|" + y.round(0).astype(str)
    valid = x.notna() & y.notna()
    for _, idx in key[valid].groupby(key[valid]).groups.items():
        n = len(idx)
        if n < 2:
            continue
        order = weight.loc[idx].fillna(0).sort_values(ascending=False).index
        step = min(1.7, 6.0 / math.sqrt(n))
        for k, i in enumerate(order):
            r, a = step * math.sqrt(k), k * 2.39996      # góc vàng
            px.at[i] = x.at[i] + r * math.cos(a)
            py.at[i] = y.at[i] + r * math.sin(a)
    return px, py


def _axis_range(s: pd.Series, must=(), lo_q=0.02, hi_q=0.98):
    """Khung trục bỏ 2% ngoại lai mỗi đầu, nhưng luôn chứa các vạch ngưỡng."""
    s = s.dropna()
    if s.empty:
        return -1.0, 1.0, 0.1
    lo, hi = s.quantile(lo_q), s.quantile(hi_q)
    lo, hi = min([lo, *must]), max([hi, *must])
    pad = max((hi - lo) * 0.08, 1.0)
    return lo - pad, hi + pad, pad


def _pin(s: pd.Series, rng):
    """Ghim giá trị ngoài khung vào mép. Trả (giá trị vẽ, hướng −1/0/+1)."""
    lo, hi, pad = rng
    lo_edge, hi_edge = lo + pad * 0.4, hi - pad * 0.4
    side = pd.Series(0, index=s.index)
    side[s < lo_edge] = -1
    side[s > hi_edge] = 1
    return s.clip(lo_edge, hi_edge), side


_EDGE_SYMBOL = {
    (0, 1): "triangle-up",    (0, -1): "triangle-down",
    (1, 0): "triangle-right", (-1, 0): "triangle-left",
    (1, 1): "triangle-ne",    (-1, 1): "triangle-nw",
    (1, -1): "triangle-se",   (-1, -1): "triangle-sw",
}


def _symbols(sx: pd.Series, sy: pd.Series, is_fin: pd.Series) -> pd.Series:
    base = is_fin.map({True: "diamond", False: "circle"})
    edge = [_EDGE_SYMBOL.get((a, b)) for a, b in zip(sx, sy)]
    return pd.Series([e or b for e, b in zip(edge, base)], index=sx.index)


# ═════════════════════════════════════════════════════════════════════════
# Cấu hình từng bản đồ
# ═════════════════════════════════════════════════════════════════════════
def _view_spec(view: str, d: pd.DataFrame) -> dict:
    """Trục, nhãn và vạch tham chiếu cho một bản đồ."""
    has_fin = bool(d["is_fin"].any())
    ret_lines = [(_ROIC_LINE, "dash", f"ROIC {_ROIC_LINE}%")]
    if has_fin:
        ret_lines.append((_ROE_LINE, "dot", f"◆ ROE {_ROE_LINE}%"))
    ret_title = "ROIC (%)  ·  ◆ Tài chính: ROE"

    if view == "value":
        med = d.loc[~d["is_fin"], "yield"].median()
        vlines = [(0, "solid", "")]
        if med == med:
            vlines.append((med, "dot", f"trung vị {med:.1f}%"))
        return dict(
            x="yield", y="ret", pin=True,
            title="Sinh lời trên vốn × Lợi suất tiền mặt  —  chất lượng so với giá phải trả",
            xtitle="Lợi suất FCF = FCF / vốn hoá (%)  ·  ◆ Tài chính: E/P   →  càng phải càng rẻ",
            ytitle=ret_title, xfmt=".1f", yfmt=".1f", xunit="%", yunit="%",
            xname="Lợi suất", yname="Sinh lời",
            hlines=ret_lines, vlines=vlines,
            corners={"tr": "sinh lời cao · giá rẻ", "tl": "sinh lời cao · giá đắt",
                     "br": "rẻ nhưng sinh lời thấp"},
            peers_median=True,
        )
    if view == "growth":
        return dict(
            x="growth", y="ret", pin=True, x_hi_q=0.95,     # đuôi phải rất dài
            title="Sinh lời trên vốn × Tăng trưởng doanh thu  —  lớn lên có tạo ra giá trị không",
            xtitle="Tăng trưởng doanh thu 12 tháng (%)", ytitle=ret_title,
            xfmt="+.1f", yfmt=".1f", xunit="%", yunit="%",
            xname="Tăng trưởng", yname="Sinh lời",
            hlines=ret_lines, vlines=[(0, "solid", "")],
            corners={"tr": "cỗ máy kép: sinh lời cao, đang lớn",
                     "br": "lớn nhanh nhưng tốn vốn",
                     "tl": "sinh lời cao, đang chững"},
            peers_median=True,
        )
    if view == "timing":
        return dict(
            x="peak", y="ma250", pin=True,
            title="Giá hiện tại so với đỉnh 52 tuần và MA250  —  nhịp chỉnh hay xu hướng giảm",
            xtitle="Cách đỉnh 52 tuần (%)   ←  giảm sâu hơn",
            ytitle="Giá so với MA250 (%)", xfmt="+.1f", yfmt="+.1f", xunit="%", yunit="%",
            xname="Cách đỉnh", yname="So với MA250",
            hlines=[(0, "solid", "MA250")], vlines=[(-10, "dot", "−10%"), (-20, "dot", "−20%")],
            corners={"tl": "giảm khỏi đỉnh, vẫn trên MA250 — nhịp chỉnh",
                     "bl": "dưới MA250 — xu hướng giảm",
                     "tr": "sát đỉnh"},
            peers_median=True, x_max=0.5,
        )
    return dict(
        x="fin", y="real", pin=False, spread=True,
        title="Tài chính × Thực  —  số đẹp có được chứng minh không",
        xtitle="Điểm Tài chính (%)", ytitle="Điểm Thực (%)",
        xfmt=".0f", yfmt=".0f", xunit="%", yunit="%",
        xname="Tài chính", yname="Thực",
        hlines=[(80, "dash", "")], vlines=[(80, "dash", "")],
        corners={"tr": "tốt &amp; thật", "br": "số đẹp, chưa chứng minh"},
        fixed=(-8, 108), peers_median=False,
    )


# ═════════════════════════════════════════════════════════════════════════
# Vẽ
# ═════════════════════════════════════════════════════════════════════════
def build_html(df: pd.DataFrame, dark: bool = False,
               view: str = DEFAULT_VIEW, focus: str | None = None) -> str:
    """Một bản đồ so sánh toàn bộ mã.

    focus: mã đang chọn — được khoanh tròn, các mã cùng ngành giữ nguyên màu
    và có nhãn, các mã khác ngành mờ đi.
    """
    import plotly.graph_objects as go

    T = _theme(dark)
    if df is None or df.empty:
        return _empty(T, "Chưa có dữ liệu — bấm SCAN trước.")
    if view not in VIEWS:
        view = DEFAULT_VIEW

    d    = _prep(df)
    spec = _view_spec(view, d)
    xs, ys = d[spec["x"]], d[spec["y"]]
    ok   = xs.notna() & ys.notna()
    missing = int((~ok).sum())
    if not ok.any():
        return _empty(T, "Không mã nào đủ số liệu cho bản đồ này.")

    # ── Toạ độ vẽ ────────────────────────────────────────────────────────
    sx = sy = pd.Series(0, index=d.index)
    if spec.get("fixed"):
        lo, hi = spec["fixed"]
        xr = yr = (lo, hi)
        px, py = _spread_ties(xs, ys, d["mcap"]) if spec.get("spread") else (xs, ys)
    else:
        must_x = [v for v, _, _ in spec["vlines"]]
        must_y = [v for v, _, _ in spec["hlines"]]
        rx = _axis_range(xs[ok], must_x, hi_q=spec.get("x_hi_q", 0.98))
        ry = _axis_range(ys[ok], must_y)
        if "x_max" in spec:                       # trục "cách đỉnh" không vượt 0
            rx = (rx[0], spec["x_max"] + rx[2], rx[2])
        px, sx = _pin(xs, rx)
        py, sy = _pin(ys, ry)
        xr, yr = (rx[0], rx[1]), (ry[0], ry[1])
    pinned = int(((sx != 0) | (sy != 0))[ok].sum())

    size = _marker_sizes(d["mcap"])
    sym  = _symbols(sx, sy, d["is_fin"])

    # ── Mã đang chọn và nhóm cùng ngành ──────────────────────────────────
    focus = (focus or "").strip().upper()
    f_idx = d.index[d["ticker"] == focus]
    f_row = d.loc[f_idx[0]] if len(f_idx) else None
    peers = (d["sector"] == f_row["sector"]) & ok if f_row is not None else ok
    opac  = pd.Series(0.85, index=d.index)
    if f_row is not None:
        opac[~peers] = 0.16
        opac[peers]  = 0.95

    fig = go.Figure()
    hover = (f"<b>%{{customdata[0]}}</b>  %{{customdata[1]}}<br>"
             f"%{{customdata[2]}} · %{{customdata[3]}}<br>"
             f"{spec['xname']} %{{customdata[4]:{spec['xfmt']}}}{spec['xunit']}"
             f" · {spec['yname']} %{{customdata[5]:{spec['yfmt']}}}{spec['yunit']}<br>"
             f"Vốn hoá %{{customdata[6]:,.1f}} tỷ<extra></extra>")
    cdata = d[["ticker", "name", "sector", "signal"]].assign(
        xv=xs, yv=ys, mc=d["mcap"].fillna(0))

    signals = list(SIG_COLORS) + sorted(set(d["signal"]) - set(SIG_COLORS))
    for s in signals:
        m = ok & (d["signal"] == s)
        if not m.any():
            continue
        color = SIG_COLORS.get(s, T["dim"])
        # Chú giải lấy hình của chấm đầu tiên — có thể là tam giác hay chấm
        # mờ — nên dùng một mục riêng, cùng nhóm để bấm vẫn ẩn/hiện được
        fig.add_trace(go.Scatter(
            x=[None], y=[None], mode="markers", name=s or "—", legendgroup=s,
            marker=dict(size=9, color=color, symbol="circle"),
        ))
        fig.add_trace(go.Scatter(
            x=px[m], y=py[m], mode="markers", name=s or "—",
            legendgroup=s, showlegend=False,
            marker=dict(size=size[m], color=color, symbol=sym[m], opacity=opac[m],
                        line=dict(width=0.6, color=T["bg"])),
            customdata=cdata[m].values, hovertemplate=hover,
        ))

    if f_row is not None:
        _add_focus(fig, go, T, d, px, py, ok, peers, f_row, spec)

    # ── Vạch tham chiếu và nhãn góc ──────────────────────────────────────
    for v, dash, label in spec["hlines"]:
        fig.add_hline(y=v, line=dict(color=T["zero"], width=1, dash=dash))
        if label:
            fig.add_annotation(x=xr[0], y=v, text=label, showarrow=False, yshift=8,
                               xanchor="left", font=dict(size=8, color=T["dim"]))
    for v, dash, label in spec["vlines"]:
        fig.add_vline(x=v, line=dict(color=T["zero"], width=1, dash=dash))
        if label:
            fig.add_annotation(x=v, y=yr[0], text=label, showarrow=False, xshift=4,
                               yanchor="bottom", xanchor="left",
                               font=dict(size=8, color=T["dim"]))
    _add_corners(fig, spec["corners"], xr, yr, T)

    sub = "".join(f"<br><span style='font-size:9px;color:{T['dim']}'>{line}</span>"
                  for line in _note(spec, d, ok, missing, pinned, focus, f_row))
    fig.update_xaxes(title_text=spec["xtitle"], range=list(xr), gridcolor=T["grid"],
                     zeroline=False, title_font=dict(size=10))
    fig.update_yaxes(title_text=spec["ytitle"], range=list(yr), gridcolor=T["grid"],
                     zeroline=False, title_font=dict(size=10))
    fig.update_layout(
        autosize=True, margin=dict(l=54, r=16, t=104, b=48),
        paper_bgcolor=T["bg"], plot_bgcolor=T["panel"],
        font=dict(family="Segoe UI, sans-serif", size=10, color=T["text"]),
        title=dict(text=spec["title"] + sub,
                   x=0.01, xanchor="left", y=0.985, yanchor="top",
                   font=dict(size=12, color=T["text"])),
        legend=dict(orientation="h", yanchor="bottom", y=1.005, x=0,
                    font=dict(size=9), bgcolor="rgba(0,0,0,0)"),
        hoverlabel=dict(font_size=11), clickmode="event",
    )
    return _wrap(fig, T)


def _add_focus(fig, go, T, d, px, py, ok, peers, f_row, spec):
    """Khoanh mã đang chọn, ghi nhãn các mã cùng ngành và trung vị ngành."""
    peer_only = peers & (d["ticker"] != f_row["ticker"])
    # Ngành đông mã thì chỉ ghi nhãn các mã lớn nhất, kẻo chữ đè nhau
    label_idx = d.loc[peer_only, "mcap"].fillna(0).nlargest(_PEERS_LABEL_MAX).index
    if len(label_idx):
        fig.add_trace(go.Scatter(
            x=px[label_idx], y=py[label_idx], mode="text",
            text=d.loc[label_idx, "ticker"], textposition="top center",
            textfont=dict(size=9, color=T["text"]),
            hoverinfo="skip", showlegend=False,
        ))
    if spec.get("peers_median") and peers.sum() >= 3:
        mx = d.loc[peers, spec["x"]].median()
        my = d.loc[peers, spec["y"]].median()
        fig.add_trace(go.Scatter(
            x=[mx], y=[my], mode="markers", name="trung vị ngành",
            marker=dict(symbol="cross-thin", size=16,
                        line=dict(width=2, color=T["text"])),
            hovertemplate=(f"Trung vị {int(peers.sum())} mã {f_row['sector']}<br>"
                           f"{spec['xname']} {mx:{spec['xfmt']}}{spec['xunit']}"
                           f" · {spec['yname']} {my:{spec['yfmt']}}{spec['yunit']}"
                           "<extra></extra>"),
        ))
    i = d.index[d["ticker"] == f_row["ticker"]][0]
    if not ok.at[i]:
        return
    fig.add_trace(go.Scatter(
        x=[px.at[i]], y=[py.at[i]], mode="markers+text", name=f_row["ticker"],
        text=[f"<b>{f_row['ticker']}</b>"], textposition="middle right",
        textfont=dict(size=11, color=T["text"]),
        marker=dict(symbol="circle-open", size=_marker_sizes(d["mcap"]).at[i] + 12,
                    color=T["text"], line=dict(width=2.5)),
        hoverinfo="skip",
    ))


def _add_corners(fig, corners, xr, yr, T):
    pos = {"tr": (xr[1], yr[1], "right", "top"),    "tl": (xr[0], yr[1], "left", "top"),
           "br": (xr[1], yr[0], "right", "bottom"), "bl": (xr[0], yr[0], "left", "bottom")}
    colors = {"tr": "#34C472", "tl": "#7A8899", "br": "#E8483D", "bl": "#E8483D"}
    for k, text in corners.items():
        x, y, xa, ya = pos[k]
        fig.add_annotation(x=x, y=y, text=text, showarrow=False, xanchor=xa, yanchor=ya,
                           xshift=-6 if xa == "right" else 6,
                           yshift=-6 if ya == "top" else 18,
                           bgcolor=T["panel"], font=dict(size=9, color=colors[k]))


def _note(spec, d, ok, missing, pinned, focus, f_row) -> tuple[str, str]:
    """Hai dòng chú thích dưới tiêu đề.

    Dòng một nói rõ mọi mã không hiện đúng chỗ; dòng hai là cách đọc hình.
    """
    data = [f"Vẽ {int(ok.sum())}/{len(d)} mã"]
    if pinned:
        data.append(f"▲ {pinned} mã ngoài khung, ghim ở mép")
    if missing:
        data.append(f"{missing} mã thiếu số liệu không vẽ được")
    if focus and f_row is None:
        data.append(f"{focus} không có trong danh sách quét")
    elif f_row is not None and not ok[d["ticker"] == focus].any():
        data.append(f"{focus} thiếu số liệu cho bản đồ này")

    read = ["cỡ chấm = vốn hoá"]
    if d["is_fin"].any() and spec["y"] == "ret":
        read.append("◆ = nhóm Tài chính")
    if spec.get("spread"):
        read.append("mã trùng điểm xếp tỏa quanh vị trí thật")
    if f_row is not None:
        read.append(f"đậm = cùng ngành {f_row['sector']}")
    read.append("bấm chấm để chọn mã")
    return "  ·  ".join(data), "  ·  ".join(read)


def _wrap(fig, T) -> str:
    click = ("var gd=document.getElementById('{plot_id}');"
             "gd.on('plotly_click',function(e){var p=e.points&&e.points[0];"
             "if(!p||!p.customdata)return;"
             f"window.location.href='https://{PICK_HOST}/'+encodeURIComponent(p.customdata[0]);}});")
    div = fig.to_html(include_plotlyjs="cdn", full_html=False, post_script=click,
                      config={"displayModeBar": False, "responsive": True})
    return (f'<!DOCTYPE html><html><head><meta charset="utf-8"><style>'
            f'html,body{{margin:0;padding:0;width:100%;height:100%;'
            f'overflow:hidden;background:{T["bg"]};}}'
            f'.plotly-graph-div{{height:100vh!important;width:100%!important;}}'
            f'</style></head><body>{div}</body></html>')


def _empty(T, msg) -> str:
    return (f'<html><body style="background:{T["bg"]};margin:0;height:100vh;'
            f'display:flex;align-items:center;justify-content:center;">'
            f'<p style="color:{T["dim"]};font-family:Segoe UI,sans-serif;'
            f'font-size:11px;letter-spacing:1px;">{msg}</p></body></html>')
