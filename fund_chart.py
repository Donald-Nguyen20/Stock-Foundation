"""
Biểu đồ số quỹ sở hữu theo quý
==============================
Trên: số quỹ nộp 13F có giữ mã — cột xanh khi tăng so với quý trước,
đỏ khi giảm. Dưới: tổng số cổ phiếu các quỹ đó nắm.

Dữ liệu lấy từ fund_holders.history(); ở đây chỉ dựng HTML Plotly cùng
phong cách với fundamental_chart.
"""
from __future__ import annotations

import html as _html

import pandas as pd
import plotly.graph_objects as go
from plotly.subplots import make_subplots

import fund_holders

C_BLUE = "#4A9EFF"; C_GOLD = "#E8A93D"; C_GREEN = "#34C472"; C_RED = "#E8483D"


def _palette(dark: bool) -> dict:
    if dark:
        return dict(bg="#06080F", panel="#0B0E18", grid="#151E2E", zero="#1F2D40",
                    text="#B8C4D0", dim="#3D4D60", dim2="#5B6B80",
                    hover_bg="#0F1622", hover_border="#1F2D40")
    return dict(bg="#FFFFFF", panel="#F8FAFC", grid="#E2E8F0", zero="#CBD5E0",
                text="#1A202C", dim="#64748B", dim2="#64748B",
                hover_bg="#FFFFFF", hover_border="#CBD5E0")


def _n(v: float) -> str:
    """1.629 — ngăn cách hàng nghìn kiểu Việt Nam."""
    return f"{v:,.0f}".replace(",", ".")


def _shares(v: float) -> str:
    if abs(v) >= 1e9:
        return f"{v / 1e9:,.2f} tỷ".replace(".", ",")
    return f"{v / 1e6:,.1f} triệu".replace(",", " ").replace(".", ",").replace(" ", ".")


def _pct(v: float) -> str:
    return ("+" if v >= 0 else "−") + f"{abs(v):.1f}".replace(".", ",") + "%"


def _shift_key(key: str, lag: int) -> str:
    """'2026Q2' lùi 4 quý → '2025Q2'."""
    n = int(key[:4]) * 4 + int(key[-1]) - 1 - lag
    return f"{n // 4}Q{n % 4 + 1}"


def _change(series: pd.Series, key: str, lag: int) -> tuple[float, float] | None:
    """(chênh lệch, %) giữa quý key và đúng lag quý trước — tính theo lịch,
    không theo vị trí, để quý thiếu dữ liệu không làm lệch mốc so sánh."""
    a, b = series.get(key), series.get(_shift_key(key, lag))
    if a is None or b is None or pd.isna(a) or pd.isna(b) or not b:
        return None
    return a - b, (a - b) / b * 100


def _headline(hist: pd.DataFrame, c: dict) -> str:
    last = hist.dropna(subset=["holders"]).iloc[-1]
    parts = [f'<span style="color:{c["text"]};font-size:15px;font-weight:600">'
             f'{_n(last["holders"])} quỹ</span>'
             f'<span style="color:{c["dim2"]};font-size:11px">  sở hữu tại '
             f'{last["label"]}</span>']
    for lag, name in ((1, "1 quý"), (4, "1 năm")):
        ch = _change(hist.set_index("quarter")["holders"], last["quarter"], lag)
        if ch:
            col = C_GREEN if ch[0] >= 0 else C_RED
            sign = "+" if ch[0] >= 0 else "−"
            parts.append(f'<span style="color:{col};font-size:11px;font-weight:600">'
                         f'{sign}{_n(abs(ch[0]))} ({_pct(ch[1])})</span>'
                         f'<span style="color:{c["dim2"]};font-size:11px"> {name}</span>')
    return '<span style="font-size:11px">   ·   </span>'.join(parts)


def build_figure(hist: pd.DataFrame, ticker: str, dark: bool = False,
                 note: str = "") -> go.Figure:
    c = _palette(dark)
    labels = hist["label"].tolist()
    holders = hist["holders"]
    shares = hist["shares"]
    prev = holders.shift(1)
    diff = holders - prev

    colors = [c["dim"] if pd.isna(h) else C_BLUE if pd.isna(d) or d == 0
              else C_GREEN if d > 0 else C_RED for h, d in zip(holders, diff)]
    hover = []
    for lb, h, d, p in zip(labels, holders, diff, prev):
        if pd.isna(h):
            hover.append(f"{lb}: không có dữ liệu")
        elif pd.isna(d):
            hover.append(f"{lb}: {_n(h)} quỹ")
        else:
            hover.append(f"{lb}: {_n(h)} quỹ  ({'+' if d >= 0 else '−'}{_n(abs(d))}"
                         f" · {_pct(d / p * 100) if p else '—'} so với quý trước)")

    # Số quỹ vẽ bằng đường để trục khỏi phải bắt đầu từ 0 (cột cắt chân
    # sẽ phóng đại chênh lệch); phần tăng giảm tách riêng thành cột từ 0.
    fig = make_subplots(rows=3, cols=1, shared_xaxes=True,
                        row_heights=[0.5, 0.22, 0.28], vertical_spacing=0.075,
                        subplot_titles=("Số quỹ báo cáo có giữ cổ phiếu",
                                        "Tăng / giảm số quỹ so với quý trước",
                                        "Tổng số cổ phiếu các quỹ nắm (triệu)"))
    fig.add_trace(go.Scatter(
        x=labels, y=holders, mode="lines+markers+text", connectgaps=False,
        line=dict(color=C_BLUE, width=2),
        marker=dict(size=9, color=colors, line=dict(width=1.5, color=c["panel"])),
        text=[_n(v) if pd.notna(v) else "" for v in holders],
        textposition="top center", cliponaxis=False,
        textfont=dict(size=10, color=c["text"]),
        customdata=hover, hovertemplate="%{customdata}<extra></extra>",
    ), row=1, col=1)
    fig.add_trace(go.Bar(
        x=labels, y=diff, marker_color=[C_GREEN if (d or 0) >= 0 else C_RED for d in diff],
        marker_opacity=0.88, cliponaxis=False,
        text=[("+" if d >= 0 else "−") + _n(abs(d)) if pd.notna(d) else "" for d in diff],
        textposition="outside", textfont=dict(size=9, color=c["text"]),
        hoverinfo="skip",
    ), row=2, col=1)
    fig.add_trace(go.Scatter(
        x=labels, y=shares / 1e6, mode="lines+markers", connectgaps=False,
        line=dict(color=C_GOLD, width=2), marker=dict(size=6, color=C_GOLD),
        customdata=[_shares(v) if pd.notna(v) else "—" for v in shares],
        hovertemplate="%{customdata} cổ phiếu<extra></extra>",
    ), row=3, col=1)

    if holders.notna().any():
        top, lo = holders.max(), holders.min()
        pad = max((top - lo) * 0.18, top * 0.01, 1)
        fig.update_yaxes(range=[lo - pad, top + pad * 1.6], row=1, col=1)
    if diff.notna().any():
        big = max(diff.abs().max(), 1)
        fig.update_yaxes(range=[min(diff.min(), 0) - big * 0.35,
                                max(diff.max(), 0) + big * 0.35], row=2, col=1)

    fig.update_layout(
        paper_bgcolor=c["bg"], plot_bgcolor=c["panel"],
        font=dict(family="Segoe UI, system-ui, sans-serif", color=c["dim"], size=10),
        title=dict(text=(f'<span style="color:{c["text"]};font-size:14px;font-weight:600;'
                         f'letter-spacing:1px">{_html.escape(ticker)}</span>'
                         f'<span style="color:{c["dim2"]};font-size:12px;font-weight:300">'
                         f'  ·  Quỹ sở hữu theo quý</span>'),
                   x=0.0, xanchor="left", y=1, yref="container", yanchor="top",
                   pad=dict(l=14, t=10)),
        showlegend=False, margin=dict(l=56, r=22, t=86, b=56),
        hovermode="x unified", bargap=0.32, separators=",.",
        hoverlabel=dict(bgcolor=c["hover_bg"], bordercolor=c["hover_border"],
                        font=dict(family="Segoe UI", size=12, color=c["text"])),
    )
    fig.update_xaxes(gridcolor=c["grid"], showline=True, linecolor=c["grid"],
                     zeroline=False, tickfont=dict(size=9, color=c["dim"]))
    fig.update_yaxes(gridcolor=c["grid"], zerolinecolor=c["zero"],
                     tickfont=dict(size=9, color=c["dim"]))
    fig.update_yaxes(tickformat=",.0f")
    fig.update_yaxes(rangemode="tozero", row=3, col=1)
    for ann in fig.layout.annotations:
        ann.update(font=dict(size=10, color=c["dim2"], family="Segoe UI"),
                   x=0, xanchor="left")

    if holders.notna().any():
        fig.add_annotation(text=_headline(hist, c), xref="paper", yref="paper",
                           x=0, y=1.0, yshift=28, xanchor="left", yanchor="bottom",
                           showarrow=False, align="left")
    latest = hist["quarter"].iloc[-1] if len(hist) else ""
    foot = ("Nguồn: SEC Form 13F — quỹ quản lý từ 100 triệu USD phải báo cáo, "
            "chậm nhất 45 ngày sau quý; không gồm quyền chọn")
    if latest:
        foot += f"  ·  quý gần nhất kết thúc {fund_holders.quarter_end(latest):%d/%m/%Y}"
    if note:
        foot += f"  ·  {note}"
    fig.add_annotation(text=foot, xref="paper", yref="paper", x=0, y=0, yshift=-46,
                       xanchor="left", yanchor="bottom", showarrow=False,
                       font=dict(size=9, color=c["dim2"], family="Segoe UI"))
    return fig


def _page(body: str, dark: bool) -> str:
    bg = _palette(dark)["bg"]
    return (f'<!DOCTYPE html><html><head><meta charset="utf-8"><style>'
            f'html,body{{margin:0;padding:0;width:100%;height:100%;'
            f'overflow:hidden;background:{bg};}}'
            f'.plotly-graph-div{{height:100vh!important;width:100%!important;}}'
            f'</style></head><body>{body}</body></html>')


def build_html(hist: pd.DataFrame, ticker: str, dark: bool = False, note: str = "") -> str:
    fig = build_figure(hist, ticker, dark, note)
    fig.update_layout(autosize=True)
    div = fig.to_html(include_plotlyjs="cdn", full_html=False,
                      config={"displayModeBar": False, "displaylogo": False})
    return _page(div, dark)


def message_html(title: str, body: str, dark: bool = False, tone: str = "dim") -> str:
    """Trang chữ giữa khung: đang tải, chưa có dữ liệu, không áp dụng…"""
    c = _palette(dark)
    col = {"dim": c["dim2"], "warn": C_GOLD, "error": C_RED}.get(tone, c["dim2"])
    return _page(
        f'<div style="height:100vh;display:flex;align-items:center;justify-content:center;'
        f'font-family:Segoe UI,sans-serif;text-align:center;padding:0 40px;box-sizing:border-box">'
        f'<div style="max-width:520px"><div style="color:{col};font-size:13px;font-weight:600;'
        f'letter-spacing:.5px;margin-bottom:10px">{_html.escape(title)}</div>'
        f'<div style="color:{c["dim2"]};font-size:11px;line-height:1.6">{body}</div>'
        f'</div></div>', dark)
