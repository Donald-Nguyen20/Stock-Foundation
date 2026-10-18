# -*- coding: utf-8 -*-
"""Dựng báo cáo phân tích từ Markdown thành HTML cho QTextBrowser.

Bố cục theo lối báo cáo phân tích của công ty chứng khoán: đầu trang có câu
tiêu đề luận điểm, hộp luận điểm chính, dải chỉ số của bộ lọc, các mục đánh số
có kẻ dòng, rồi chân trang ghi nguồn và lưu ý. Qt chỉ hiểu một phần HTML4/CSS
nên mọi bố cục đều dựng bằng bảng.

Đọc được cả bản lưu cũ (năm mục, chưa có TÓM TẮT, tin dạng
"**ngày — nguồn:** nội dung") lẫn định dạng mới mà prompt hiện tại yêu cầu.
"""
from __future__ import annotations

import html
import re
from datetime import datetime

import ai_analysis

SERIF = "Cambria,'Times New Roman',serif"
SANS = "'Segoe UI',sans-serif"

TITLES = {
    "tomtat":   "Tóm tắt",
    "nganh":    "Ngành kinh doanh",
    "loithe":   "Lợi thế cạnh tranh",
    "tiemnang": "Tiềm năng phát triển",
    "tintuc":   "Tin tức quan trọng",
    "doichieu": "Đối chiếu với số liệu",
}

OUTLINE = [
    ("Tóm tắt",               "Luận điểm chính trong ba đến bốn ý"),
    ("Ngành kinh doanh",      "Mô hình kiếm tiền, cấu trúc ngành và đối thủ"),
    ("Lợi thế cạnh tranh",    "Hào kinh tế và mức độ giải thích được ROIC, biên lợi nhuận"),
    ("Tiềm năng phát triển",  "Động lực và rủi ro cụ thể trong 2–3 năm tới"),
    ("Tin tức quan trọng",    "Tối đa năm tin gần đây, mỗi tin có ngày và nguồn"),
    ("Đối chiếu với số liệu", "Điểm khớp, điểm lệch và các tiêu chí cần lưu ý"),
]

# (nhãn, cột số liệu, hậu tố, số lẻ, cột kết quả tiêu chí của bộ lọc, ngưỡng)
METRICS = [
    ("ROIC",          "ROIC%",         "%",     1, "FIN_ROIC",   "> 15%"),
    ("Biên gộp",      "Gross Margin%", "%",     1, "FIN_GM",     "> 40%"),
    ("Tăng LN",       "NI Growth%",    "%",     0, "FIN_NIG",    "> 20%"),
    ("Tăng DT",       "Rev Growth%",   "%",     0, "FIN_REVG",   "> 15%"),
    ("FCF / LN",      "FCF/NI",        "×",     2, "REAL_CASH",  "≥ 0,8×"),
    ("ROE − ROIC",    "ROE_ROIC_Gap",  " điểm", 1, "REAL_LEV",   "< 15 điểm"),
    ("Nợ ròng / FCF", "Nợ ròng/FCF",   " năm",  1, "REAL_BAL",   "< 3 năm"),
    ("Dốc biên gộp",  "Dốc biên",      " điểm", 1, "REAL_EROD",  "≥ −1 điểm/năm"),
]

_HEAD = re.compile(r"^\s*#{1,4}\s*(.+?)\s*#*\s*$")
_EMOJI = re.compile("[\U0001F000-\U0001FAFF️‍]")
_DATE = r"\d{1,2}(?:[–-]\d{1,2})?[/.-]\d{1,2}[/.-]\d{2,4}"
_BADGE = re.compile(r"(?<!\w)(ĐẠT|TRƯỢT|THIẾU DỮ LIỆU)(?!\w)")
_LABEL_MAX = 48


# ─────────────────────────────────────────────────────────────────────────────
# Tiện ích
# ─────────────────────────────────────────────────────────────────────────────
def _mix(fg: str, bg: str, a: float) -> str:
    """Pha màu fg lên nền bg với độ đậm a — Qt rich text không hiểu rgba."""
    f = [int(fg[i:i + 2], 16) for i in (1, 3, 5)]
    b = [int(bg[i:i + 2], 16) for i in (1, 3, 5)]
    return "#%02X%02X%02X" % tuple(round(x * a + y * (1 - a)) for x, y in zip(f, b))


def _esc(t) -> str:
    return html.escape(str(t or ""), quote=False)


def inline(t: str, P: dict) -> str:
    """Markdown trong dòng → HTML: đậm, nghiêng, liên kết. Bỏ emoji."""
    t = re.sub(r"  +", " ", _EMOJI.sub("", t))
    t = _esc(t)
    t = re.sub(r"\*\*(.+?)\*\*", r"<b>\1</b>", t)
    t = re.sub(r"(?<!\*)\*(?!\s)(.+?)(?<!\s)\*(?!\*)", r"<i>\1</i>", t)
    t = re.sub(r"\[([^\]]+)\]\((https?://[^)\s]+)\)",
               rf"<a href='\2' style='color:{P['BLUE']};text-decoration:none;'>\1</a>", t)
    return t


def _badge(word: str, P: dict) -> str:
    col = {"ĐẠT": P["GREEN"], "TRƯỢT": P["RED"]}.get(word, P["AMBER"])
    return (f"<span style='color:{col};background-color:{_mix(col, P['SURFACE'], 0.13)};"
            f"font-size:10px;font-weight:700;letter-spacing:0.5px;'>&nbsp;{word}&nbsp;</span>")


def _badges(t: str, P: dict) -> str:
    return _BADGE.sub(lambda m: _badge(m.group(1), P), t)


def _when(info: dict | None) -> datetime:
    try:
        return datetime.fromisoformat((info or {}).get("created", ""))
    except (TypeError, ValueError):
        return datetime.now()


def _rule(color: str, width: int = 1, margin: str = "0") -> str:
    return (f'<table width="100%" cellspacing="0" cellpadding="0" style="margin:{margin};">'
            f'<tr><td style="border-top:{width}px solid {color};font-size:1px;">&nbsp;</td>'
            f'</tr></table>')


def _caption(text: str, color: str, margin: str = "0 0 6px 0") -> str:
    return (f'<p style="margin:{margin};color:{color};font-size:10px;font-weight:700;'
            f'letter-spacing:1.5px;">{text}</p>')


def page(P: dict, inner: str) -> str:
    return (f'<body style="color:{P["TEXT1"]};font-family:{SANS};font-size:13px;">'
            f'{inner}</body>')


def notice(P: dict, text: str, color: str, title: str = "") -> str:
    head = (f'<p style="margin:0 0 4px 0;color:{color};font-weight:700;font-size:12px;">'
            f'{title}</p>') if title else ""
    return (f'<table width="100%" cellspacing="0" cellpadding="0" style="margin:0 0 14px 0;">'
            f'<tr><td style="background-color:{_mix(color, P["SURFACE"], 0.08)};'
            f'border-left:2px solid {color};padding:10px 14px;font-size:12px;'
            f'color:{P["TEXT1"]};line-height:140%;">{head}{text}</td></tr></table>')


# ─────────────────────────────────────────────────────────────────────────────
# Tách Markdown
# ─────────────────────────────────────────────────────────────────────────────
def split(md: str) -> tuple[list[str], list[tuple[str, list[str]]]]:
    """(các dòng trước tiêu đề đầu tiên, [(khóa mục, các dòng)])."""
    titles = {t: k for k, t in ai_analysis.SECTIONS}
    pre: list[str] = []
    secs: list[tuple[str, list[str]]] = []
    for line in (md or "").splitlines():
        m = _HEAD.match(line)
        name = re.sub(r"^\d+[.)]\s*", "", m.group(1).strip("* ")).upper() if m else ""
        if name in titles:
            secs.append((titles[name], []))
        elif secs:
            secs[-1][1].append(line)
        else:
            pre.append(line)
    return pre, secs


def _blocks(lines: list[str]) -> list[tuple]:
    """Chia thân mục thành khối: ("ul", chữ), ("ol", số, chữ), ("label", chữ), ("p", chữ)."""
    out = []
    for ln in lines:
        t = ln.strip()
        if not t or re.fullmatch(r"[-*_]{3,}", t):
            continue
        if m := re.match(r"^[-*•]\s+(.*)$", t):
            out.append(("ul", m.group(1)))
        elif m := re.match(r"^(\d{1,2})[.)]\s+(.*)$", t):
            out.append(("ol", m.group(1), m.group(2)))
        elif (m := re.fullmatch(r"\*\*([^*]+?)\*\*\s*:?", t)) and len(m.group(1)) <= _LABEL_MAX:
            out.append(("label", m.group(1).strip().rstrip(":").strip()))
        else:
            out.append(("p", t))
    return out


# ─────────────────────────────────────────────────────────────────────────────
# Thân mục
# ─────────────────────────────────────────────────────────────────────────────
def _li(P: dict, mark: str, body: str, width: int = 16) -> str:
    return (f'<tr><td width="{width}" valign="top" style="padding:3px 0;color:{P["TEXT2"]};'
            f'font-size:12px;">{mark}</td>'
            f'<td valign="top" style="padding:3px 0;line-height:140%;">{body}</td></tr>')


def _label(P: dict, text: str, key: str | None) -> str:
    low = text.lower()
    col = None
    if key == "doichieu":
        col = (P["GREEN"] if "khớp" in low else P["RED"] if "lệch" in low
               else P["AMBER"] if ("tiêu chí" in low or "kiểm tra" in low) else None)
    mark = f'<span style="color:{col};">■</span>&nbsp;&nbsp;' if col else ""
    return (f'<p style="margin:14px 0 5px 0;font-size:12px;font-weight:700;'
            f'color:{P["TEXT1"]};">{mark}{inline(text, P)}</p>')


def _criteria_row(P: dict, parts: list[str]) -> str:
    name, status, why = (parts + ["", ""])[:3]
    st = status.strip().upper()
    st_html = _badge(st, P) if _BADGE.fullmatch(st) else inline(status, P)
    return (f'<tr><td width="30%" valign="top" style="padding:7px 12px 7px 0;font-weight:600;'
            f'border-bottom:1px solid {P["BORDER"]};"><b>{inline(name, P)}</b></td>'
            f'<td width="96" valign="top" style="padding:7px 12px 7px 0;'
            f'border-bottom:1px solid {P["BORDER"]};">{st_html}</td>'
            f'<td valign="top" style="padding:7px 0;line-height:145%;'
            f'border-bottom:1px solid {P["BORDER"]};">{_badges(inline(why, P), P)}</td></tr>')


def body(P: dict, lines: list[str], key: str | None) -> str:
    """Thân một mục. Tin tức dựng thành bảng, tiêu chí đối chiếu có nhãn trạng thái."""
    if key == "tintuc":
        return _news(P, lines)
    out, group = [], None   # group: loại danh sách đang mở ("ul", "ol", "crit")

    def close():
        nonlocal group
        if group:
            out.append("</table>")
            group = None

    def open_(kind, margin="2px 0 8px 0"):
        nonlocal group
        if group != kind:
            close()
            out.append(f'<table width="100%" cellspacing="0" cellpadding="0" '
                       f'style="margin:{margin};">')
            group = kind

    fix = (lambda s: _badges(s, P)) if key == "doichieu" else (lambda s: s)
    for b in _blocks(lines):
        kind = b[0]
        if kind == "ul" and key == "doichieu" and b[1].count(" | ") >= 2:
            open_("crit", "4px 0 10px 0")
            out.append(_criteria_row(P, [x.strip() for x in b[1].split(" | ", 2)]))
        elif kind == "ul":
            open_("ul")
            out.append(_li(P, "•", fix(inline(b[1], P))))
        elif kind == "ol":
            open_("ol")
            out.append(_li(P, f"{b[1]}.", fix(inline(b[2], P)), 22))
        elif kind == "label":
            close()
            out.append(_label(P, b[1], key))
        else:
            close()
            out.append(f'<p style="margin:0 0 9px 0;line-height:140%;">'
                       f'{fix(inline(b[1], P))}</p>')
    close()
    return "".join(out)


def parse_news(txt: str) -> tuple[str, str, str]:
    """(ngày, nguồn, nội dung) từ một dòng tin — đọc cả định dạng mới lẫn cũ."""
    parts = [p.strip() for p in txt.split(" | ", 2)]
    if len(parts) == 3 and re.fullmatch(rf"\**\s*{_DATE}\s*\**", parts[0]):
        return parts[0].strip("* "), parts[1], parts[2]
    m = re.match(rf"^\*\*\s*({_DATE})\s*[—–-]\s*(.+?):?\*\*:?\s*(.*)$", txt)
    if m:
        return m.group(1), m.group(2), m.group(3)
    m = re.match(rf"^\**\s*({_DATE})\s*\**\s*[—–:|-]?\s*(.*)$", txt)
    if m:
        return m.group(1), "", m.group(2)
    return "", "", txt


def _news(P: dict, lines: list[str]) -> str:
    rows, rest = [], []
    for b in _blocks(lines):
        if b[0] in ("ul", "ol"):
            rows.append(parse_news(b[-1]))
        else:
            rest.append(f'<p style="margin:0 0 9px 0;line-height:140%;">{inline(b[-1], P)}</p>')
    if not rows:
        return "".join(rest)

    th = (f'padding:0 12px 6px 0;color:{P["TEXT3"]};font-size:9px;font-weight:700;'
          f'letter-spacing:1px;border-bottom:1px solid {P["BORDER2"]};')
    td = f'padding:8px 12px 8px 0;border-bottom:1px solid {P["BORDER"]};'
    out = [f'<table width="100%" cellspacing="0" cellpadding="0" style="margin:4px 0 10px 0;">'
           f'<tr><td width="86" style="{th}">NGÀY</td><td width="150" style="{th}">NGUỒN</td>'
           f'<td style="{th}">NỘI DUNG</td></tr>']
    impact = (f'<br><span style="color:{P["TEXT2"]};font-weight:600;">Tác động:</span> ')
    for day, src, txt in rows:
        srcs = "<br>".join(inline(s, P) for s in re.split(r"\s+/\s+", src) if s.strip())
        text = re.sub(r"\s*Tác động:\s*", impact, inline(txt, P), count=1)
        out.append(
            f'<tr><td valign="top" style="{td}color:{P["TEXT2"]};font-size:11px;'
            f'font-weight:600;">{_esc(day)}</td>'
            f'<td valign="top" style="{td}font-size:11px;line-height:140%;">{srcs}</td>'
            f'<td valign="top" style="{td}line-height:140%;">{text}</td></tr>')
    out.append("</table>")
    return "".join(out) + "".join(rest)


# ─────────────────────────────────────────────────────────────────────────────
# Các phần của trang
# ─────────────────────────────────────────────────────────────────────────────
def _summary(lines: list[str]) -> tuple[str, list[str]]:
    """Mục TÓM TẮT → (câu tiêu đề, các ý chính)."""
    head, points = "", []
    for b in _blocks(lines):
        if b[0] in ("ul", "ol"):
            points.append(b[-1])
        elif not head and not points:
            head = b[-1].strip("* ").rstrip(":")
    return head, points


def _masthead(P: dict, row: dict | None, headline: str, info: dict | None) -> str:
    row = row or {}
    tk, co = _esc(row.get("Ticker", "")), _esc(row.get("Tên Công Ty", ""))
    sector = _esc(row.get("Sector", "") or "")
    title = inline(headline, P) if headline else (co or tk)
    sub = [f"<b>{tk}</b>"] + ([co] if headline and co else []) + ([sector] if sector else [])
    return (f'<table width="100%" cellspacing="0" cellpadding="0"><tr>'
            f'<td>{_caption("PHÂN TÍCH DOANH NGHIỆP", P["BLUE"], "0")}</td>'
            f'<td align="right" style="color:{P["TEXT3"]};font-size:10px;">'
            f'{_when(info):%d/%m/%Y}</td></tr></table>'
            f'<p style="margin:8px 0 5px 0;font-family:{SERIF};font-size:20px;font-weight:700;'
            f'color:{P["TEXT1"]};line-height:125%;">{title}</p>'
            f'<p style="margin:0 0 12px 0;color:{P["TEXT2"]};font-size:12px;">'
            f'{"&nbsp;&nbsp;·&nbsp;&nbsp;".join(sub)}</p>'
            + _rule(P["TEXT1"], 2))


def _key_points(P: dict, points: list[str]) -> str:
    items = "".join(_li(P, f'<span style="color:{P["BLUE"]};">■</span>', inline(p, P), 18)
                    for p in points)
    return (f'<table width="100%" cellspacing="0" cellpadding="0" style="margin:16px 0 0 0;">'
            f'<tr><td style="background-color:{_mix(P["BLUE"], P["SURFACE"], 0.05)};'
            f'border-top:2px solid {P["BLUE"]};padding:12px 18px 10px 18px;">'
            f'{_caption("LUẬN ĐIỂM CHÍNH", P["BLUE"])}'
            f'<table width="100%" cellspacing="0" cellpadding="0">{items}</table>'
            f'</td></tr></table>')


def _num(v: float, dec: int) -> str:
    """Số kiểu Việt Nam: 1.234,5 và dấu trừ chuẩn."""
    s = f"{abs(v):,.{dec}f}".replace(",", " ").replace(".", ",").replace(" ", ".")
    return ("−" if v < 0 else "") + s


def _metric(P: dict, row: dict, spec: tuple) -> tuple[str, str]:
    """(giá trị đã định dạng, màu) — màu theo đúng kết quả tiêu chí của bộ lọc."""
    _, col, _, dec, flag, _ = spec
    try:
        v = float(row.get(col))
        if v != v:
            raise ValueError
    except (TypeError, ValueError):
        return "—", P["TEXT3"]
    got = row.get(flag)
    if got is None or got not in (True, False):
        return _num(v, dec), P["TEXT1"]
    return _num(v, dec), P["GREEN"] if got else P["RED"]


def _metrics(P: dict, row: dict | None) -> str:
    if not row:
        return ""
    edge = f'border-top:1px solid {P["TEXT2"]};border-bottom:1px solid {P["BORDER"]};'
    cells = []
    for i, spec in enumerate(METRICS):
        val, col = _metric(P, row, spec)
        sep = f'border-left:1px solid {P["BORDER"]};' if i else ""
        # Đơn vị chữ in nhỏ cạnh số để ô hẹp không phải xuống dòng
        unit = spec[2] if val != "—" else ""
        if unit.startswith(" "):
            unit = (f'<span style="font-size:10px;font-weight:400;'
                    f'color:{P["TEXT2"]};">{unit}</span>')
        cells.append(
            f'<td width="12.5%" valign="top" style="{edge}{sep}padding:8px 10px 9px 10px;">'
            f'<span style="color:{P["TEXT2"]};font-size:9px;font-weight:700;'
            f'letter-spacing:0.5px;">{_esc(spec[0].upper())}</span><br>'
            f'<span style="color:{col};font-size:16px;font-weight:600;'
            f'white-space:nowrap;">{val}{unit}</span><br>'
            f'<span style="color:{P["TEXT3"]};font-size:9px;">'
            f'ngưỡng {_esc(spec[5])}</span></td>')
    return (_caption("CHỈ SỐ TỪ BỘ LỌC", P["TEXT2"], "20px 0 6px 0")
            + f'<table width="100%" cellspacing="0" cellpadding="0"><tr>{"".join(cells)}'
              f'</tr></table>')


def _heading(P: dict, n: int, title: str) -> str:
    return (f'<table width="100%" cellspacing="0" cellpadding="0" style="margin:24px 0 10px 0;">'
            f'<tr><td style="border-bottom:1px solid {P["BORDER2"]};padding:0 0 6px 0;">'
            f'<span style="font-family:{SERIF};font-size:17px;font-weight:700;'
            f'color:{P["BLUE"]};">{n}</span>'
            f'<span style="font-family:{SERIF};font-size:17px;font-weight:700;'
            f'color:{P["TEXT1"]};">&nbsp;&nbsp;&nbsp;{title}</span></td></tr></table>')


def _footer(P: dict, info: dict | None) -> str:
    n = (info or {}).get("searches") or 0
    news = (f"tin tức tổng hợp từ {n} lượt tìm kiếm web, dẫn nguồn theo từng tin"
            if n else "tin tức dẫn nguồn theo từng tin")
    small = f'margin:0 0 4px 0;color:{P["TEXT3"]};font-size:10px;line-height:140%;'
    return (_rule(P["BORDER2"], 1, "28px 0 8px 0")
            + f'<p style="{small}"><b>Nguồn:</b> chỉ số tài chính lấy từ bộ lọc định lượng '
              f'của ứng dụng tại lần quét gần nhất; {news}. '
              f'Lập lúc {_when(info):%H:%M ngày %d/%m/%Y}.</p>'
              f'<p style="{small}"><b>Lưu ý:</b> tài liệu chỉ nhằm mục đích tham khảo, '
              f'không phải khuyến nghị mua, bán hay nắm giữ chứng khoán.</p>')


# ─────────────────────────────────────────────────────────────────────────────
# Trang hoàn chỉnh
# ─────────────────────────────────────────────────────────────────────────────
def build(md: str, P: dict, row: dict | None = None, info: dict | None = None,
          note: str = "", streaming: bool = False) -> str:
    """Báo cáo đầy đủ. `streaming` bỏ chân trang khi nội dung còn đang về."""
    pre, secs = split(md)
    headline, points = _summary(next((ln for k, ln in secs if k == "tomtat"), []))
    parts = [_masthead(P, row, headline, info)]
    if note:
        parts.append(f'<div style="margin-top:12px;">{notice(P, note, P["AMBER"])}</div>')
    if points:
        parts.append(_key_points(P, points))
    parts.append(_metrics(P, row))
    if any(x.strip() for x in pre):
        parts.append(f'<div style="margin-top:16px;">{body(P, pre, None)}</div>')
    n = 0
    for key, lines in secs:
        if key == "tomtat":
            continue
        n += 1
        parts.append(_heading(P, n, TITLES.get(key, key)))
        parts.append(body(P, lines, key))
    if not streaming and secs:
        parts.append(_footer(P, info))
    return page(P, "".join(parts))


def _intro(P: dict, title: str, text: str) -> str:
    return (_caption("PHÂN TÍCH DOANH NGHIỆP", P["BLUE"], "0")
            + f'<p style="margin:8px 0 6px 0;font-family:{SERIF};font-size:22px;'
              f'font-weight:700;">{title}</p>'
              f'<p style="margin:0 0 14px 0;color:{P["TEXT2"]};font-size:12px;'
              f'line-height:155%;">{text}</p>'
            + _rule(P["TEXT1"], 2))


def empty(P: dict, row: dict | None) -> str:
    """Chưa có báo cáo: nêu báo cáo sẽ gồm những gì."""
    tk = _esc((row or {}).get("Ticker", ""))
    rows = "".join(
        f'<tr><td width="34" valign="top" style="padding:8px 0;font-family:{SERIF};'
        f'font-size:15px;font-weight:700;color:{P["BLUE"]};'
        f'border-bottom:1px solid {P["BORDER"]};">{i}</td>'
        f'<td valign="top" style="padding:8px 0;border-bottom:1px solid {P["BORDER"]};">'
        f'<span style="font-weight:600;">{t}</span><br>'
        f'<span style="color:{P["TEXT2"]};font-size:11px;">{d}</span></td></tr>'
        for i, (t, d) in enumerate(OUTLINE, 1))
    return page(P,
        _intro(P, f"Chưa có báo cáo cho {tk}" if tk else "Chọn một mã để lập báo cáo",
               "Báo cáo giải thích các chỉ số mà bộ lọc đã tính cho mã này, đánh giá lợi thế "
               "cạnh tranh và tổng hợp tin tức gần đây có dẫn nguồn. Bấm <b>LẬP BÁO CÁO</b> "
               "để bắt đầu; thời gian lập khoảng một đến hai phút.")
        + _caption("CẤU TRÚC BÁO CÁO", P["TEXT2"], "16px 0 2px 0")
        + f'<table width="100%" cellspacing="0" cellpadding="0">{rows}</table>'
          f'<p style="margin:14px 0 0 0;color:{P["TEXT3"]};font-size:10px;">'
          f'Mỗi mã lưu một bản trong ngày — mở lại trong ngày không phải lập lại.</p>')


def running(P: dict, row: dict | None) -> str:
    tk = _esc((row or {}).get("Ticker", ""))
    return page(P, _intro(
        P, f"Đang lập báo cáo {tk}",
        "Đang tìm tin tức và đối chiếu số liệu. Thường mất một đến hai phút; "
        "nội dung sẽ hiện dần ngay khi có."))


def error(P: dict, msg: str) -> str:
    lines = "".join(f'<p style="margin:4px 0;">{inline(x, P)}</p>'
                    for x in msg.splitlines() if x.strip())
    return page(P, notice(P, lines, P["RED"], "Không lập được báo cáo"))
