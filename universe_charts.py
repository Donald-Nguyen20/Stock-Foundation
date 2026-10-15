# -*- coding: utf-8 -*-
"""Biểu đồ so sánh toàn bộ mã đã quét.

Bốn biểu đồ sẵn có đều là chuỗi thời gian của một mã — chúng giúp hiểu sâu
một doanh nghiệp. Nhưng chọn cổ phiếu là việc so giữa các mã, mà cho tới
giờ app không có hình nào làm việc đó.

Hai bản đồ ở đây dùng đúng dữ liệu đã có trong bảng, không gọi mạng thêm.
"""
from __future__ import annotations

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


def _theme(dark: bool) -> dict:
    if dark:
        return dict(bg="#0A1628", panel="#0B0E18", grid="#1A2133",
                    text="#DCE4EE", dim="#7A8899", zero="#243044")
    return dict(bg="#FFFFFF", panel="#F8FAFC", grid="#E2E8F0",
                text="#1A202C", dim="#64748B", zero="#CBD5E0")


def _num(df, col):
    return pd.to_numeric(df[col], errors="coerce") if col in df.columns else pd.Series(dtype=float)


def build_html(df: pd.DataFrame, dark: bool = False, height: int = 580) -> str:
    """Hai bản đồ xếp dọc: Tài chính×Thực và ROIC×Định giá."""
    from plotly.subplots import make_subplots
    import plotly.graph_objects as go

    T = _theme(dark)
    if df is None or df.empty:
        return _empty(T, "Chưa có dữ liệu — bấm SCAN trước.")

    fig = make_subplots(
        rows=2, cols=1, vertical_spacing=0.13,
        subplot_titles=("Bản đồ Tài chính × Thực  —  mỗi chấm là một mã",
                        "ROIC × Định giá  —  chất lượng so với giá phải trả"),
    )

    fin  = _num(df, "FIN_Pct") * 100
    real = _num(df, "REAL_Pct") * 100
    sig  = df["Signal"] if "Signal" in df.columns else pd.Series([""] * len(df))
    tick = df["Ticker"] if "Ticker" in df.columns else pd.Series([""] * len(df))
    name = df["Tên Công Ty"] if "Tên Công Ty" in df.columns else tick

    # ── 1. Bản đồ hai điểm ────────────────────────────────────────────────
    # Vẽ từng nhóm tín hiệu riêng để chú giải bấm ẩn/hiện được
    for s, color in SIG_COLORS.items():
        m = (sig == s) & fin.notna() & real.notna()
        if not m.any():
            continue
        fig.add_trace(go.Scatter(
            x=fin[m], y=real[m], mode="markers", name=s,
            marker=dict(size=9, color=color, opacity=0.85,
                        line=dict(width=0.5, color=T["bg"])),
            text=[f"{t}<br>{str(n)[:38]}" for t, n in zip(tick[m], name[m])],
            hovertemplate="<b>%{text}</b><br>Tài chính %{x:.0f}%"
                          "<br>Thực %{y:.0f}%<extra></extra>",
            legendgroup=s,
        ), row=1, col=1)

    # Vạch chia bốn góc — đúng các ngưỡng tín hiệu đang dùng
    for v in (80,):
        fig.add_vline(x=v, line=dict(color=T["zero"], width=1, dash="dash"), row=1, col=1)
        fig.add_hline(y=v, line=dict(color=T["zero"], width=1, dash="dash"), row=1, col=1)
    fig.add_annotation(x=97, y=97, text="tốt &amp; thật", showarrow=False,
                       font=dict(size=9, color="#34C472"), row=1, col=1)
    fig.add_annotation(x=97, y=6, text="số đẹp,<br>chưa chứng minh", showarrow=False,
                       font=dict(size=9, color="#E8483D"), align="right", row=1, col=1)

    # ── 2. ROIC × Định giá ────────────────────────────────────────────────
    roic = _num(df, "ROIC%")
    pe   = _num(df, "P/E")
    ok   = roic.notna() & pe.notna() & (pe > 0) & (pe < 120) & (roic > -20) & (roic < 200)
    for s, color in SIG_COLORS.items():
        m = ok & (sig == s)
        if not m.any():
            continue
        fig.add_trace(go.Scatter(
            x=pe[m], y=roic[m], mode="markers", name=s, showlegend=False,
            marker=dict(size=9, color=color, opacity=0.85,
                        line=dict(width=0.5, color=T["bg"])),
            text=[f"{t}<br>{str(n)[:38]}" for t, n in zip(tick[m], name[m])],
            hovertemplate="<b>%{text}</b><br>P/E %{x:.1f}×"
                          "<br>ROIC %{y:.1f}%<extra></extra>",
            legendgroup=s,
        ), row=2, col=1)

    if ok.any():
        fig.add_hline(y=15, line=dict(color=T["zero"], width=1, dash="dash"), row=2, col=1)
        fig.add_annotation(x=pe[ok].min(), y=15, text="ngưỡng ROIC 15%",
                           showarrow=False, yshift=9, xanchor="left",
                           font=dict(size=8, color=T["dim"]), row=2, col=1)

    fig.update_xaxes(title_text="Điểm Tài chính (%)", range=[-3, 103],
                     gridcolor=T["grid"], zeroline=False, row=1, col=1)
    fig.update_yaxes(title_text="Điểm Thực (%)", range=[-3, 103],
                     gridcolor=T["grid"], zeroline=False, row=1, col=1)
    fig.update_xaxes(title_text="P/E (lần)", gridcolor=T["grid"], zeroline=False, row=2, col=1)
    fig.update_yaxes(title_text="ROIC (%)", gridcolor=T["grid"], zeroline=False, row=2, col=1)

    fig.update_layout(
        autosize=True, margin=dict(l=48, r=14, t=46, b=42),
        paper_bgcolor=T["bg"], plot_bgcolor=T["panel"],
        font=dict(family="Segoe UI, sans-serif", size=10, color=T["text"]),
        legend=dict(orientation="h", yanchor="bottom", y=1.055, x=0,
                    font=dict(size=9), bgcolor="rgba(0,0,0,0)"),
        hoverlabel=dict(font_size=11),
    )
    for a in fig.layout.annotations[:2]:
        a.font.size = 11
        a.font.color = T["text"]
        a.x = 0
        a.xanchor = "left"

    div = fig.to_html(include_plotlyjs="cdn", full_html=False,
                      config={"displayModeBar": False})
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
