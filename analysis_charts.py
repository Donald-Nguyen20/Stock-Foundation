# -*- coding: utf-8 -*-
"""Đồ thị bằng chứng cho tab Phân tích.

Khác với sáu ô biểu đồ chuỗi thời gian: những hình ở đây trả lời câu
"mã này đứng ở đâu so với phần còn lại", tức là đúng thứ mà phần phân tích
định tính đang khẳng định. Toàn bộ lấy từ dữ liệu đã có trong bảng, không
gọi mạng thêm.
"""
from __future__ import annotations

import pandas as pd

# Các chỉ số đem ra xếp hạng, kèm chiều tốt
RANK_METRICS = [
    ("ROIC%",         "ROIC",            True),
    ("Gross Margin%", "Biên gộp",        True),
    ("Op Margin%",    "Biên hoạt động",  True),
    ("NI Growth%",    "Tăng lợi nhuận",  True),
    ("Rev Growth%",   "Tăng doanh thu",  True),
    ("FCF/NI",        "FCF / Lợi nhuận", True),
    ("FCF_Margin%",   "Biên FCF",        True),
    ("D/E",           "Nợ trên vốn chủ", False),
]


def _theme(dark: bool) -> dict:
    if dark:
        return dict(bg="#0A1628", panel="#0B0E18", grid="#1A2133",
                    text="#DCE4EE", dim="#7A8899", soft="#243044")
    return dict(bg="#FFFFFF", panel="#F8FAFC", grid="#E2E8F0",
                text="#1A202C", dim="#64748B", soft="#CBD5E0")


def _num(df, col):
    return (pd.to_numeric(df[col], errors="coerce")
            if col in df.columns else pd.Series(dtype=float))


def build_html(row: dict, df: pd.DataFrame, dark: bool = False) -> str:
    """Hai hình: vị trí trong vũ trụ, và xếp hạng theo từng chỉ số."""
    from plotly.subplots import make_subplots
    import plotly.graph_objects as go

    T = _theme(dark)
    if not row or df is None or df.empty:
        return _empty(T, "Chọn một mã để xem bằng chứng")

    tk = str(row.get("Ticker", "") or "")
    fig = make_subplots(
        rows=2, cols=1, vertical_spacing=0.14,
        row_heights=[0.42, 0.58],
        subplot_titles=(f"{tk} nằm ở đâu trên bản đồ Tài chính × Thực",
                        "Xếp hạng theo từng chỉ số, so với toàn bộ mã đã quét"),
    )

    # ── 1. Vị trí trên bản đồ hai điểm ────────────────────────────────────
    fin  = _num(df, "FIN_Pct") * 100
    real = _num(df, "REAL_Pct") * 100
    ok   = fin.notna() & real.notna()
    fig.add_trace(go.Scatter(
        x=fin[ok], y=real[ok], mode="markers", showlegend=False,
        marker=dict(size=7, color=T["soft"], opacity=0.55),
        text=df.loc[ok, "Ticker"] if "Ticker" in df.columns else None,
        hovertemplate="%{text}<extra></extra>",
    ), row=1, col=1)

    try:
        mf = float(row.get("FIN_Pct") or 0) * 100
        mr = float(row.get("REAL_Pct") or 0) * 100
        fig.add_trace(go.Scatter(
            x=[mf], y=[mr], mode="markers+text", showlegend=False,
            marker=dict(size=17, color="#E8483D", symbol="circle",
                        line=dict(width=2, color=T["bg"])),
            text=[tk], textposition="top center",
            textfont=dict(size=11, color=T["text"]),
            hovertemplate=f"<b>{tk}</b><br>Tài chính %{{x:.0f}}%"
                          f"<br>Thực %{{y:.0f}}%<extra></extra>",
        ), row=1, col=1)
    except (TypeError, ValueError):
        pass

    for v in (80,):
        fig.add_vline(x=v, line=dict(color=T["soft"], width=1, dash="dash"), row=1, col=1)
        fig.add_hline(y=v, line=dict(color=T["soft"], width=1, dash="dash"), row=1, col=1)

    # ── 2. Phân vị theo từng chỉ số ───────────────────────────────────────
    labels, pcts, texts, colors = [], [], [], []
    for col, label, higher_better in RANK_METRICS:
        series = _num(df, col).dropna()
        try:
            v = float(row.get(col))
        except (TypeError, ValueError):
            continue
        if v != v or len(series) < 5:
            continue
        p = (series < v).mean() * 100
        if not higher_better:
            p = 100 - p
        labels.append(label)
        pcts.append(round(p, 0))
        unit = "×" if col == "FCF/NI" else ("" if col == "D/E" else "%")
        texts.append(f"  {v:,.2f}{unit}  ·  top {100 - p:.0f}%")
        colors.append("#34C472" if p >= 70 else "#4A9EFF" if p >= 40
                      else "#E8A93D" if p >= 20 else "#E8483D")

    if labels:
        fig.add_trace(go.Bar(
            x=pcts, y=labels, orientation="h", showlegend=False,
            marker_color=colors, marker_opacity=0.9,
            text=texts, textposition="outside",
            textfont=dict(size=10, color=T["dim"]),
            hovertemplate="%{y}: phân vị %{x:.0f}<extra></extra>",
        ), row=2, col=1)
        fig.add_vline(x=50, line=dict(color=T["soft"], width=1, dash="dot"), row=2, col=1)

    fig.update_xaxes(title_text="Điểm Tài chính (%)", range=[-3, 103],
                     gridcolor=T["grid"], zeroline=False, row=1, col=1)
    fig.update_yaxes(title_text="Điểm Thực (%)", range=[-3, 103],
                     gridcolor=T["grid"], zeroline=False, row=1, col=1)
    fig.update_xaxes(title_text="Phân vị trong nhóm đã quét", range=[0, 128],
                     gridcolor=T["grid"], zeroline=False, row=2, col=1)
    fig.update_yaxes(autorange="reversed", gridcolor=T["grid"], row=2, col=1)

    fig.update_layout(
        autosize=True, margin=dict(l=96, r=16, t=40, b=36),
        paper_bgcolor=T["bg"], plot_bgcolor=T["panel"],
        font=dict(family="Segoe UI, sans-serif", size=10, color=T["text"]),
        bargap=0.34,
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
