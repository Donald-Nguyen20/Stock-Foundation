# -*- coding: utf-8 -*-
"""
Sinh file Word: "Hướng dẫn sử dụng — Hiểu từng thông số · Cách tính · Cách chọn mã"

Chạy:  ./venv/Scripts/python.exe make_guide_doc.py
Nguồn tiêu chí: FIN_CRITERIA / REAL_CRITERIA / score_stocks
                trong "python screen_top100.py".
"""
import os
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.section import WD_SECTION
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

# ── Bảng màu (đồng bộ với app) ──────────────────────────────────────────────
NAVY   = RGBColor(0x1B, 0x3A, 0x5C)
BLUE   = RGBColor(0x1E, 0x40, 0xAF)
GREEN  = RGBColor(0x1A, 0x5C, 0x2B)
AMBER  = RGBColor(0x9C, 0x65, 0x00)
RED    = RGBColor(0x9C, 0x00, 0x06)
GREY   = RGBColor(0x64, 0x74, 0x8B)
BLACK  = RGBColor(0x1A, 0x20, 0x2C)

HDR_BG   = "1B3A5C"
ALT_BG   = "F0F4FA"
BOX_BG   = "EBF4FF"
WARN_BG  = "FFF6E5"

FONT = "Calibri"


# ── Tiện ích ────────────────────────────────────────────────────────────────
def shade(cell, hexcolor):
    el = OxmlElement("w:shd")
    el.set(qn("w:val"), "clear")
    el.set(qn("w:color"), "auto")
    el.set(qn("w:fill"), hexcolor)
    cell._tc.get_or_add_tcPr().append(el)


def cell_text(cell, text, bold=False, size=9, color=BLACK, align=None):
    cell.text = ""
    p = cell.paragraphs[0]
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after = Pt(2)
    if align is not None:
        p.alignment = align
    for i, line in enumerate(str(text).split("\n")):
        if i:
            p.add_run().add_break()
        for seg, seg_bold in _parse_bold(line):
            r = p.add_run(seg)
            r.bold = bold or seg_bold
            r.font.size = Pt(size)
            r.font.color.rgb = color
            r.font.name = FONT
    return cell


def add_table(doc, headers, rows, widths=None, size=9, hdr_size=9, colors=None):
    """colors: dict {(row_idx, col_idx): RGBColor} tô chữ cho ô cụ thể."""
    t = doc.add_table(rows=1, cols=len(headers))
    t.style = "Table Grid"
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    t.autofit = False

    for i, h in enumerate(headers):
        c = t.rows[0].cells[i]
        shade(c, HDR_BG)
        cell_text(c, h, bold=True, size=hdr_size,
                  color=RGBColor(0xFF, 0xFF, 0xFF),
                  align=WD_ALIGN_PARAGRAPH.CENTER)

    for ri, row in enumerate(rows):
        cells = t.add_row().cells
        for ci, val in enumerate(row):
            col = (colors or {}).get((ri, ci), BLACK)
            bold = ci == 0
            al = WD_ALIGN_PARAGRAPH.CENTER if (widths and ci > 0 and len(str(val)) <= 14) else None
            cell_text(cells[ci], val, bold=bold, size=size, color=col, align=al)
            if ri % 2 == 0:
                shade(cells[ci], ALT_BG)

    if widths:
        for r in t.rows:
            for ci, w in enumerate(widths):
                r.cells[ci].width = Cm(w)
    doc.add_paragraph().paragraph_format.space_after = Pt(4)
    return t


def h1(doc, text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(18)
    p.paragraph_format.space_after = Pt(8)
    r = p.add_run(text)
    r.bold = True
    r.font.size = Pt(15)
    r.font.color.rgb = NAVY
    r.font.name = FONT
    pr = p._p.get_or_add_pPr()
    bd = OxmlElement("w:pBdr")
    bt = OxmlElement("w:bottom")
    bt.set(qn("w:val"), "single")
    bt.set(qn("w:sz"), "12")
    bt.set(qn("w:color"), "1B3A5C")
    bt.set(qn("w:space"), "4")
    bd.append(bt)
    pr.append(bd)
    return p


def h2(doc, text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(12)
    p.paragraph_format.space_after = Pt(4)
    r = p.add_run(text)
    r.bold = True
    r.font.size = Pt(12)
    r.font.color.rgb = BLUE
    r.font.name = FONT
    return p


def h3(doc, text, color=None):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(9)
    p.paragraph_format.space_after = Pt(2)
    r = p.add_run(text)
    r.bold = True
    r.font.size = Pt(10.5)
    r.font.color.rgb = color or NAVY
    r.font.name = FONT
    return p


def para(doc, text, size=10, color=BLACK, italic=False, space_after=6, indent=0):
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(space_after)
    p.paragraph_format.line_spacing = 1.15
    if indent:
        p.paragraph_format.left_indent = Cm(indent)
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    for seg, bold in _parse_bold(text):
        r = p.add_run(seg)
        r.bold = bold
        r.italic = italic
        r.font.size = Pt(size)
        r.font.color.rgb = color
        r.font.name = FONT
    return p


def _parse_bold(text):
    """Tách chuỗi theo **đậm** thành các đoạn (text, bold)."""
    out, buf, bold = [], "", False
    i = 0
    while i < len(text):
        if text[i:i + 2] == "**":
            if buf:
                out.append((buf, bold))
                buf = ""
            bold = not bold
            i += 2
        else:
            buf += text[i]
            i += 1
    if buf:
        out.append((buf, bold))
    return out or [(text, False)]


def bullet(doc, text, size=10, color=BLACK, level=0):
    p = doc.add_paragraph(style="List Bullet")
    p.paragraph_format.space_after = Pt(3)
    p.paragraph_format.left_indent = Cm(0.75 + level * 0.6)
    p.paragraph_format.line_spacing = 1.12
    for seg, bold in _parse_bold(text):
        r = p.add_run(seg)
        r.bold = bold
        r.font.size = Pt(size)
        r.font.color.rgb = color
        r.font.name = FONT
    return p


def callout(doc, title, text, bg=BOX_BG, fg=NAVY):
    t = doc.add_table(rows=1, cols=1)
    t.style = "Table Grid"
    c = t.rows[0].cells[0]
    shade(c, bg)
    c.text = ""
    p = c.paragraphs[0]
    p.paragraph_format.space_before = Pt(4)
    p.paragraph_format.space_after = Pt(2)
    r = p.add_run(title)
    r.bold = True
    r.font.size = Pt(10)
    r.font.color.rgb = fg
    r.font.name = FONT
    p2 = c.add_paragraph()
    p2.paragraph_format.space_after = Pt(4)
    p2.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    for seg, bold in _parse_bold(text):
        rr = p2.add_run(seg)
        rr.bold = bold
        rr.font.size = Pt(9.5)
        rr.font.color.rgb = BLACK
        rr.font.name = FONT
    doc.add_paragraph().paragraph_format.space_after = Pt(4)
    return t


def code(doc, text):
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Cm(0.6)
    p.paragraph_format.space_after = Pt(6)
    r = p.add_run(text)
    r.font.size = Pt(9)
    r.font.name = "Consolas"
    r.font.color.rgb = RGBColor(0x2D, 0x37, 0x48)
    return p


# ═══════════════════════════════════════════════════════════════════════════
# XÂY DỰNG TÀI LIỆU — HƯỚNG DẪN SỬ DỤNG
# ═══════════════════════════════════════════════════════════════════════════
doc = Document()

st = doc.styles["Normal"]
st.font.name = FONT
st.font.size = Pt(10)
st.element.rPr.rFonts.set(qn("w:eastAsia"), FONT)

sec = doc.sections[0]
sec.top_margin = Cm(2.0); sec.bottom_margin = Cm(2.0)
sec.left_margin = Cm(2.2); sec.right_margin = Cm(2.0)

ORANGE = RGBColor(0xC2, 0x41, 0x0C)


def metric(doc, name, formula, source, meaning, threshold=None):
    """Một mục trong từ điển thông số."""
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(11)
    p.paragraph_format.space_after = Pt(2)
    r = p.add_run(name)
    r.bold = True; r.font.size = Pt(11); r.font.color.rgb = NAVY; r.font.name = FONT
    r2 = p.add_run(f"     ({source})")
    r2.font.size = Pt(8.5); r2.font.color.rgb = GREY; r2.font.name = FONT

    if formula:
        code(doc, formula)
    para(doc, meaning, size=9.5, space_after=3)
    if threshold:
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(7)
        p.paragraph_format.left_indent = Cm(0.6)
        r = p.add_run("▸  Ngưỡng đạt: ")
        r.bold = True; r.font.size = Pt(9.5); r.font.color.rgb = GREEN; r.font.name = FONT
        for seg, b in _parse_bold(threshold):
            rr = p.add_run(seg)
            rr.bold = b; rr.font.size = Pt(9.5); rr.font.color.rgb = GREEN; rr.font.name = FONT


# ── Trang bìa ───────────────────────────────────────────────────────────────
for _ in range(5):
    doc.add_paragraph()

p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = p.add_run("STOCK FOUNDATION SCREENER")
r.bold = True; r.font.size = Pt(13); r.font.color.rgb = GREY; r.font.name = FONT

p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_after = Pt(4)
r = p.add_run("HƯỚNG DẪN SỬ DỤNG")
r.bold = True; r.font.size = Pt(28); r.font.color.rgb = NAVY; r.font.name = FONT

p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = p.add_run("Hiểu từng thông số  ·  Cách tính điểm  ·  Cách chọn mã")
r.font.size = Pt(14); r.font.color.rgb = BLUE; r.font.name = FONT

for _ in range(6):
    doc.add_paragraph()

t = doc.add_table(rows=1, cols=1); t.style = "Table Grid"
c = t.rows[0].cells[0]; shade(c, ALT_BG)
cell_text(c,
          "Tài liệu này giải thích mọi con số bạn nhìn thấy trong app: nó là gì, tính bằng công thức nào, "
          "và đọc nó ra sao. Sau đó hướng dẫn cách app chấm điểm và cách bạn dùng kết quả đó để chọn mã.\n\n"
          "Đọc mục 1 và 2 là dùng được ngay. Mục 3 là từ điển tra cứu khi gặp một cột lạ. "
          "Mục 6 và 7 là phần quan trọng nhất: cách đọc kết quả cho đúng.",
          size=10, color=BLACK, align=WD_ALIGN_PARAGRAPH.JUSTIFY)

doc.add_page_break()

# ═══════════════════════════════════════════════════════════════════════════
h1(doc, "1.  APP NÀY LÀM GÌ")

para(doc,
     "App quét toàn bộ cổ phiếu vốn hóa lớn của một sàn, rồi trả lời **hai câu hỏi tách rời nhau** "
     "cho từng mã:")

add_table(doc,
          ["", "Câu hỏi", "Ví dụ tiêu chí"],
          [
              ["Nhóm A\nTÀI CHÍNH", "Số liệu tài chính của doanh nghiệp này có tốt không?",
               "Sinh lời trên vốn > 15%, biên lợi nhuận cao, lợi nhuận đang tăng, nợ thấp"],
              ["Nhóm B\nKIỂM CHỨNG", "Những con số đó có thật không?",
               "Lợi nhuận có kèm tiền mặt không, ROE cao là do kinh doanh hay do vay nợ"],
          ],
          widths=[3.0, 6.2, 7.4], size=9.5)

para(doc,
     "Lý do phải tách làm hai: con số đẹp trên báo cáo tài chính có thể đến từ nhiều nguồn, "
     "và không phải nguồn nào cũng phản ánh một doanh nghiệp tốt. Lợi nhuận có thể được ghi nhận "
     "trước khi thu được tiền. Biên lợi nhuận có thể được nâng lên bằng cách hoãn đầu tư. "
     "EPS có thể tăng chỉ vì công ty mua lại cổ phiếu làm giảm mẫu số. ROE có thể cao vì vay nhiều "
     "chứ không phải vì kinh doanh giỏi.")

callout(doc, "Hai điểm số không bao giờ được cộng gộp",
        "Nếu gộp thành một thang 12 điểm, mã đạt **6 tài chính + 1 kiểm chứng** (rất đáng ngờ) và mã đạt "
        "**3 + 4** (trung bình đều) sẽ cùng ra 7/12 và trông y hệt nhau. Tách riêng thì trường hợp đầu "
        "hiện lên thành tín hiệu 🔴 CỜ ĐỎ. Đó là lý do app luôn hiển thị hai điểm cạnh nhau.")

h2(doc, "1.1  Nguồn dữ liệu")
add_table(doc,
          ["Nguồn", "Cung cấp", "Đặc điểm"],
          [
              ["TradingView", "Toàn bộ 12 trên 13 tiêu chí, cùng mọi số liệu hiển thị",
               "Một lệnh gọi cho cả danh sách. Nhanh, có cho mọi thị trường."],
              ["yfinance\n(tùy chọn)", "Duy nhất tiêu chí Bền vững 5 năm (ROE và biên gộp trung bình 5 năm)",
               "Gọi từng mã, khoảng 0,3 giây mỗi mã. Chỉ có dữ liệu cho thị trường Mỹ."],
          ],
          widths=[3.0, 7.0, 6.6], size=9.5)
para(doc,
     "Nếu **không bật** ô «yfinance Moat», app vẫn chạy đầy đủ với 12 tiêu chí và nhanh hơn rất nhiều. "
     "Tiêu chí Bền vững 5 năm khi đó để trống và không bị tính vào mẫu số — xem mục 5.", size=9.5)

# ═══════════════════════════════════════════════════════════════════════════
h1(doc, "2.  BẮT ĐẦU NHANH")

for i, (title, body) in enumerate([
    ("Chọn thị trường và số mã",
     "Ô đầu tiên chọn sàn (america, nasdaq, nyse, vietnam, hong_kong, euronext). "
     "Ô số bên cạnh là lấy bao nhiêu mã vốn hóa lớn nhất. 100–200 là hợp lý để bắt đầu."),
    ("Cân nhắc bật «yfinance Moat»",
     "Bật thì có thêm tiêu chí Bền vững 5 năm nhưng quét chậm hơn nhiều (100 mã mất khoảng 5–8 phút). "
     "Tắt thì quét xong trong vài giây. Với thị trường ngoài Mỹ, bật cũng không có tác dụng."),
    ("Bấm SCAN",
     "App tải dữ liệu, chấm điểm cả hai nhóm rồi tự chuyển sang Dashboard."),
    ("Đọc Dashboard trước",
     "Đây là nơi tóm tắt: có bao nhiêu mã TỐT & THỰC, bao nhiêu mã CỜ ĐỎ, và vì sao. "
     "Xem mục 7.1 để biết đọc từng bảng."),
    ("Vào hai tab để xem chi tiết",
     "Tab **Tài chính** cho biết số liệu có đẹp không. Tab **Kiểm chứng** cho biết số liệu đó có thật không. "
     "Click một dòng để xem breakdown từng tiêu chí, rê chuột lên bất kỳ cột nào để đọc giải thích."),
    ("Xuất Excel nếu cần lưu lại",
     "File gồm ba sheet: Screener (toàn bộ), Cờ đỏ (mã số đẹp mà chưa chứng minh được), "
     "và Tổng hợp (thống kê)."),
]):
    h3(doc, f"Bước {i+1} — {title}")
    para(doc, body, size=9.5, space_after=4)

callout(doc, "Mẹo: rê chuột là ra giải thích",
        "Mọi cột trong hai bảng đều có chú giải. Rê chuột lên **tiêu đề cột** hoặc lên **ô dữ liệu** sẽ hiện "
        "một hộp nói chỉ số đó là gì, tính thế nào và đọc ra sao. Với cột tiêu chí ✓/✗ thì chú giải còn nói "
        "rõ tiêu chí đó đang bác bỏ khả năng gì. Bạn không cần thuộc lòng tài liệu này.")

doc.add_page_break()

# ═══════════════════════════════════════════════════════════════════════════
h1(doc, "3.  TỪ ĐIỂN THÔNG SỐ")

para(doc,
     "Phần này giải thích từng con số trong app. Với mỗi thông số: công thức, nguồn, ý nghĩa và cách đọc. "
     "Thông số nào là tiêu chí chấm điểm thì có thêm dòng **Ngưỡng đạt**; thông số không có dòng đó "
     "là chỉ để tham khảo, không ảnh hưởng tới điểm.")

# ── 3.1 ─────────────────────────────────────────────────────────────────────
h2(doc, "3.1  Quy mô doanh nghiệp")

metric(doc, "MCap — Vốn hóa thị trường", "MCap = Giá cổ phiếu × Số cổ phiếu đang lưu hành",
       "TradingView",
       "Tổng giá trị thị trường của toàn bộ doanh nghiệp. App dùng chỉ số này để chọn ra top N mã lớn nhất "
       "trước khi chấm điểm, nên mọi kết quả chỉ có giá trị trong nhóm vốn hóa lớn.")

metric(doc, "Doanh thu", "Doanh thu = tổng tiền bán hàng của 4 quý gần nhất (TTM)",
       "TradingView · total_revenue_ttm",
       "Tiền thu về từ bán hàng, **trước khi trừ bất kỳ chi phí nào**. Đây là dòng trên cùng của báo cáo "
       "kết quả kinh doanh và cũng là con số khó nguỵ tạo nhất — muốn tăng doanh thu thì phải bán được "
       "nhiều hàng hơn thật. Chỉ hiển thị để bạn thấy quy mô, không chấm điểm, vì số tuyệt đối không "
       "so sánh được giữa doanh nghiệp lớn và nhỏ.")

metric(doc, "Lợi nhuận — Lợi nhuận ròng", "Lợi nhuận ròng = Doanh thu − mọi chi phí − lãi vay − thuế",
       "TradingView · net_income",
       "Phần còn lại cuối cùng, thuộc về cổ đông. Đây là **con số kế toán**: nó tuân theo các quy tắc "
       "ghi nhận, và một doanh nghiệp có thể ghi nhận doanh thu ngay khi giao hàng dù chưa nhận được tiền. "
       "Vì vậy lợi nhuận tăng không đồng nghĩa tiền mặt tăng — đó chính là lý do có cột kế tiếp.")

metric(doc, "FCF — Dòng tiền tự do",
       "FCF = Tiền thu từ hoạt động kinh doanh − Tiền chi mua sắm tài sản",
       "TradingView · free_cash_flow_ttm",
       "Tiền mặt thật còn lại sau khi doanh nghiệp đã trả mọi chi phí vận hành và đã chi đầu tư để duy trì "
       "và mở rộng. Đây là số tiền thực sự có thể dùng để trả nợ, chia cổ tức hoặc mua lại cổ phiếu.\n"
       "Trong app, cột này được đặt **ngay cạnh cột Lợi nhuận** ở tab Kiểm chứng — cố ý, để bạn nhìn "
       "một cái là thấy hai con số có khớp nhau không.")

callout(doc, "Ba cột này đọc cùng nhau mới có ý nghĩa",
        "Ví dụ từ một lần quét thật:\n"
        "**NVDA** — doanh thu 253 tỷ, lợi nhuận 120 tỷ, FCF 119 tỷ. Lợi nhuận và tiền mặt gần như bằng nhau: "
        "lợi nhuận này là tiền thật.\n"
        "**AMZN** — doanh thu 776 tỷ, lợi nhuận 78 tỷ, FCF **âm 11,6 tỷ**. Báo cáo có lãi lớn nhưng tiền mặt "
        "đang chảy ra vì đầu tư rất nặng. Không phải gian lận, nhưng là một thực tế mà nếu chỉ nhìn cột "
        "lợi nhuận thì bạn sẽ không thấy.")

# ── 3.2 ─────────────────────────────────────────────────────────────────────
h2(doc, "3.2  Nhóm sinh lời — doanh nghiệp kiếm tiền tốt đến đâu")

metric(doc, "ROIC — Sinh lời trên vốn đầu tư",
       "ROIC = Lợi nhuận sau thuế từ hoạt động kinh doanh / (Vốn chủ sở hữu + Nợ vay)",
       "TradingView · return_on_invested_capital",
       "Đây là **thước đo chất lượng doanh nghiệp tốt nhất trong toàn bộ bảng**. Nó hỏi: bỏ ra một đồng vốn "
       "thì thu về bao nhiêu lợi nhuận — và quan trọng là nó tính trên **toàn bộ** vốn, cả phần tự có lẫn "
       "phần đi vay. Nhờ vậy nó không bị thổi phồng bởi việc vay nợ, khác với ROE.\n"
       "Doanh nghiệp huy động vốn với chi phí khoảng 8–10% mỗi năm. Nếu ROIC thấp hơn mức đó, "
       "doanh nghiệp càng tăng trưởng càng phá hủy giá trị.",
       "trên **15%**. Riêng ngân hàng và bảo hiểm dùng ROE > 12% thay thế.")

metric(doc, "Op Margin — Biên lợi nhuận hoạt động",
       "Op Margin = Lợi nhuận từ hoạt động kinh doanh / Doanh thu",
       "TradingView · operating_margin",
       "Trong 100 đồng doanh thu, còn lại bao nhiêu đồng sau khi trừ giá vốn hàng bán và chi phí bán hàng, "
       "quản lý — nhưng **trước** lãi vay và thuế. Vì bỏ qua lãi vay nên nó phản ánh hiệu quả vận hành "
       "thuần túy, không bị ảnh hưởng bởi việc doanh nghiệp vay nhiều hay ít.",
       "trên **15%**")

metric(doc, "GM — Biên lợi nhuận gộp",
       "GM = (Doanh thu − Giá vốn hàng bán) / Doanh thu",
       "TradingView · gross_margin",
       "Bán một sản phẩm giá 100 mà chi phí làm ra nó là 40 thì biên gộp là 60%. Chỉ số này phản ánh "
       "**sức mạnh định giá**: doanh nghiệp có bán được cao hơn hẳn chi phí sản xuất không, và có còn dư "
       "tiền cho nghiên cứu, quảng cáo không.\n"
       "**Lưu ý quan trọng:** biên gộp lệch rất mạnh theo ngành. Phần mềm thường 60–80% vì chi phí nhân bản "
       "gần bằng không; bán lẻ hiếm khi vượt 30% vì phải nhập hàng về bán. Một mã bán lẻ trượt tiêu chí này "
       "không có nghĩa nó kém — chỉ là ngưỡng chung chưa phân biệt được ngành.",
       "trên **40%**")

metric(doc, "ROE — Sinh lời trên vốn chủ sở hữu",
       "ROE = Lợi nhuận ròng / Vốn chủ sở hữu",
       "TradingView · return_on_equity",
       "Mỗi đồng vốn của cổ đông sinh ra bao nhiêu lợi nhuận. Nghe rất hợp lý, nhưng **app cố ý không dùng "
       "ROE làm tiêu chí chấm điểm** vì nó dễ bị thổi phồng theo hai cách: vay nhiều (mẫu số không đổi mà "
       "tử số tăng), hoặc mua lại cổ phiếu quỹ khiến vốn chủ sở hữu teo lại.\n"
       "Thay vào đó app dùng ROIC, rồi so hai con số với nhau ở cột ROE−ROIC để biết phần chênh đến từ đâu.")

metric(doc, "Net Margin — Biên lợi nhuận ròng",
       "Net Margin = Lợi nhuận ròng / Doanh thu", "TradingView · net_margin",
       "Trong 100 đồng doanh thu còn lại bao nhiêu đồng lợi nhuận cuối cùng. Chỉ hiển thị tham khảo.")

# ── 3.3 ─────────────────────────────────────────────────────────────────────
h2(doc, "3.3  Nhóm tăng trưởng — doanh nghiệp đang lớn lên hay đứng yên")

para(doc,
     "Cả ba chỉ số dưới đây dùng số **TTM** — cộng dồn bốn quý gần nhất rồi so với bốn quý cùng kỳ năm trước. "
     "Cách này mượt hơn so với lấy một quý đơn lẻ, vì nhiều ngành có mùa vụ rõ rệt.", size=9.5)

metric(doc, "NI Growth — Tăng trưởng lợi nhuận ròng",
       "NI Growth = (Lợi nhuận 4 quý gần nhất / Lợi nhuận 4 quý cùng kỳ năm trước − 1) × 100",
       "TradingView · net_income_yoy_growth_ttm",
       "Lợi nhuận năm nay lớn hơn năm ngoái bao nhiêu phần trăm.\n"
       "App đo tăng trưởng **lợi nhuận ròng** chứ không đo tăng trưởng EPS — đây là lựa chọn có chủ đích. "
       "EPS còn phụ thuộc vào số cổ phiếu lưu hành, nên EPS có thể tăng dù lợi nhuận đứng yên. "
       "Phần chênh giữa hai con số này chính là nguyên liệu cho một tiêu chí kiểm chứng ở mục 3.5.",
       "trên **20%**")

metric(doc, "Rev Growth — Tăng trưởng doanh thu",
       "Rev Growth = (Doanh thu 4 quý gần nhất / Doanh thu 4 quý cùng kỳ năm trước − 1) × 100",
       "TradingView · total_revenue_yoy_growth_ttm",
       "Doanh thu năm nay lớn hơn năm ngoái bao nhiêu phần trăm.\n"
       "Tiêu chí này bổ trợ cho tiêu chí trên. Lợi nhuận có thể tăng nhờ cắt giảm chi phí — nhưng cắt chi phí "
       "chỉ làm được vài lần rồi hết dư địa. Doanh thu tăng mới là dấu hiệu doanh nghiệp thực sự bán được "
       "nhiều hơn. Khi lợi nhuận tăng mạnh mà doanh thu đứng yên, tăng trưởng đó khó lặp lại.",
       "trên **15%**")

metric(doc, "NI Qtr% — Tăng trưởng lợi nhuận quý gần nhất",
       "NI Qtr% = (Lợi nhuận quý vừa công bố / Lợi nhuận cùng quý năm trước - 1) x 100",
       "TradingView · net_income_yoy_growth_fq",
       "Chỉ riêng quý vừa công bố, không cộng dồn. Trong app cột này nằm **ngay cạnh NI Grw%** "
       "(mức TTM) — cố ý, để bạn so hai con số ngay lập tức.\n"
       "Vì sao cần: số TTM cộng dồn bốn quý nên **có độ trễ**. Một doanh nghiệp mới bắt đầu xấu đi "
       "vẫn giữ số TTM đẹp thêm hai ba quý nữa, vì các quý tốt cũ còn đang đè lên. Nhìn riêng cột "
       "TTM thì bạn hoàn toàn không thấy điều đó.")

metric(doc, "Rev Qtr% — Tăng trưởng doanh thu quý gần nhất",
       "Rev Qtr% = (Doanh thu quý vừa công bố / Doanh thu cùng quý năm trước - 1) x 100",
       "TradingView · total_revenue_yoy_growth_fq",
       "Tương tự chỉ số trên, đặt cạnh cột Rev Grw%. Chỉ hiển thị tham khảo, không chấm điểm.")

metric(doc, "EPS Growth — Tăng trưởng lãi trên mỗi cổ phiếu",
       "EPS Growth = mức tăng của (Lợi nhuận ròng / Số cổ phiếu lưu hành)",
       "TradingView · earnings_per_share_diluted_yoy_growth_ttm",
       "Không phải tiêu chí chấm điểm. App lấy con số này **chỉ để so với NI Growth** — xem cột EPS−NI "
       "ở mục 3.5.")

# ── 3.4 ─────────────────────────────────────────────────────────────────────
h2(doc, "3.4  Nhóm an toàn — doanh nghiệp có chịu được cú sốc không")

metric(doc, "D/E — Nợ trên vốn chủ sở hữu",
       "D/E = Tổng nợ vay / Vốn chủ sở hữu", "TradingView · debt_to_equity",
       "D/E bằng 1,0 nghĩa là cứ một đồng vốn tự có thì có một đồng đi vay. Vay nợ không xấu — nó khuếch đại "
       "lợi nhuận khi làm ăn tốt. Vấn đề là nó khuếch đại cả theo chiều ngược lại, và **lãi vay vẫn phải trả "
       "kể cả năm doanh thu sụt giảm**.\n"
       "**Không áp dụng cho ngân hàng và bảo hiểm.** Với các doanh nghiệp này, tiền gửi của khách hàng được "
       "hạch toán là nợ, nên D/E 8–15 lần là cấu trúc bình thường chứ không phải rủi ro. App tự nhận diện "
       "và bỏ qua tiêu chí này cho nhóm Financial Services.",
       "dưới **1,0**")

metric(doc, "Net Cash — Tiền ròng",
       "Net Cash = Tiền và tương đương tiền + Đầu tư ngắn hạn − Tổng nợ vay",
       "TradingView · cash_n_short_term_invest_fq và total_debt_fq",
       "Nếu doanh nghiệp trả hết nợ ngay hôm nay thì còn lại bao nhiêu tiền. Số dương nghĩa là có nhiều tiền "
       "hơn nợ — đây là vị thế rất vững, doanh nghiệp không phụ thuộc vào việc vay đảo nợ khi thị trường "
       "tín dụng khó khăn.\n"
       "Số âm không tự động là xấu; nhiều doanh nghiệp tốt vẫn duy trì nợ ròng. Vì vậy tiêu chí chấm điểm "
       "cho phép **hoặc** Net Cash dương **hoặc** thanh khoản tốt.")

metric(doc, "Nợ/FCF — Số năm dòng tiền cần để trả hết nợ ròng",
       "Nợ/FCF = Nợ ròng / Dòng tiền tự do một năm     (đơn vị: năm)",
       "Tính trong app",
       "Nợ ròng 52 tỷ với dòng tiền tự do 67 tỷ mỗi năm thì chỉ số là 0,8 — trả hết trong chưa tới "
       "một năm. Cùng khoản nợ đó ở một doanh nghiệp chỉ tạo 10 tỷ mỗi năm thì mất hơn 5 năm.\n"
       "Đây là cách đo gánh nặng nợ đúng bản chất hơn Current Ratio, vốn là thước đo sinh ra cho "
       "doanh nghiệp sản xuất có hàng tồn kho và khoản phải thu.\n"
       "Ô để trống khi doanh nghiệp đã có tiền ròng dương (không còn nợ để trả) hoặc khi dòng tiền âm "
       "(không trả được bằng dòng tiền).",
       "dưới **3 năm** — là vế thứ ba của tiêu chí Bảng cân đối")

metric(doc, "Curr.R — Hệ số thanh toán hiện hành",
       "Current Ratio = Tài sản ngắn hạn / Nợ ngắn hạn", "TradingView · current_ratio",
       "Đo khả năng trả các khoản đến hạn trong vòng một năm bằng tài sản có thể chuyển thành tiền trong "
       "cùng khoảng thời gian đó. Dưới 1,0 nghĩa là nợ ngắn hạn nhiều hơn tài sản ngắn hạn — cần theo dõi.")

# ── 3.5 ─────────────────────────────────────────────────────────────────────
h2(doc, "3.5  Nhóm kiểm chứng — những con số kia có thật không")

para(doc,
     "Đây là phần đặc trưng nhất của app. Mỗi chỉ số dưới đây không được chọn vì «nó là một chỉ số hay», "
     "mà vì nó **bác bỏ một cách cụ thể** mà số liệu ở các mục trên có thể bị làm đẹp.")

add_table(doc,
          ["Số liệu đẹp", "Có thể ảo vì", "Chỉ số phát hiện"],
          [
              ["Lợi nhuận cao", "Ghi nhận doanh thu chưa thu được tiền", "FCF/NI"],
              ["Biên lợi nhuận cao", "Hoãn đầu tư, vốn hóa chi phí cho đẹp sổ", "FCF Mgn"],
              ["EPS tăng mạnh", "Mua cổ phiếu quỹ làm giảm mẫu số", "EPS−NI"],
              ["ROE cao", "Vay nhiều, hoặc vốn chủ teo vì mua cổ phiếu quỹ", "ROE−ROIC"],
              ["Bảng cân đối sạch", "Nợ nặng so với dòng tiền tạo ra", "Net Cash · Curr.R · Nợ/FCF"],
              ["Số TTM đang tăng", "Ba quý cũ đè lên, quý gần nhất đã xấu đi", "Đà LN · Còn đúng"],
              ["Một năm xuất sắc", "Một kỳ may mắn, không phải năng lực", "Bền vững 5N"],
          ],
          widths=[3.8, 7.2, 5.6], size=9)

metric(doc, "FCF/NI — Dòng tiền tự do chia lợi nhuận ròng",
       "FCF/NI = Dòng tiền tự do / Lợi nhuận ròng",
       "Tính trong app từ hai trường TradingView",
       "**Phép thử quan trọng nhất của toàn bộ app.** Lợi nhuận là con số kế toán; dòng tiền tự do là tiền "
       "thật. Trong một doanh nghiệp lành mạnh, hai con số này bám sát nhau qua thời gian.\n"
       "Khi lợi nhuận cao mà dòng tiền thấp kéo dài, phần chênh nằm ở các khoản phải thu chưa đòi được, "
       "hàng tồn kho chất đống, hoặc các bút toán ghi nhận trước. Kết quả 0,99 nghĩa là gần như mỗi đồng "
       "lợi nhuận đều có một đồng tiền mặt đi kèm. Kết quả âm nghĩa là doanh nghiệp báo lãi nhưng tiền "
       "đang chảy ra.\n"
       "Ô để trống khi doanh nghiệp đang lỗ, vì khi đó tỷ lệ này không có ý nghĩa.",
       "từ **0,80 lần** trở lên")

metric(doc, "FCF Mgn — Biên dòng tiền tự do",
       "FCF Margin = Dòng tiền tự do / Doanh thu",
       "TradingView · free_cash_flow_margin_ttm",
       "Mỗi 100 đồng doanh thu để lại bao nhiêu đồng tiền mặt thật. Chỉ số này bổ sung cho FCF/NI ở chỗ: "
       "một doanh nghiệp lợi nhuận thấp có thể có FCF/NI rất cao mà tiền tuyệt đối vẫn ít. Đo trên doanh thu "
       "thì tránh được điều đó.",
       "trên **8%**")

metric(doc, "EPS−NI — Chênh lệch tăng trưởng EPS so với lợi nhuận",
       "EPS−NI = EPS Growth% − NI Growth%     (đơn vị: điểm phần trăm)",
       "Tính trong app",
       "Giả sử lợi nhuận ròng tăng 10% nhưng EPS tăng 25%. Phần chênh 15 điểm **không đến từ kinh doanh** — "
       "nó đến từ việc số cổ phiếu lưu hành giảm, tức doanh nghiệp đã mua lại cổ phiếu quỹ.\n"
       "Mua cổ phiếu quỹ không xấu, cổ đông vẫn hưởng lợi. Nhưng đó là một cách phân bổ vốn, "
       "**không phải tăng trưởng nội tại**, và nó không lặp lại được vô hạn. Nếu bạn nhìn EPS mà tưởng "
       "doanh nghiệp đang tăng trưởng 25% thì bạn đã hiểu sai bản chất.\n"
       "Số âm nghĩa là ngược lại — lợi nhuận tăng nhanh hơn EPS, tức doanh nghiệp đang phát hành thêm "
       "cổ phiếu làm pha loãng. Trường hợp này vẫn được coi là đạt tiêu chí.",
       "không vượt quá **10 điểm %**")

metric(doc, "ROE−ROIC — Chênh lệch sinh lời vốn chủ so với toàn bộ vốn",
       "ROE−ROIC = ROE% − ROIC%     (đơn vị: điểm phần trăm)",
       "Tính trong app",
       "ROE tính lợi nhuận trên riêng vốn của cổ đông; ROIC tính trên toàn bộ vốn gồm cả nợ vay. "
       "Vì vậy **khoảng cách giữa hai con số chính là phần mà đòn bẩy đóng góp**.\n"
       "Một doanh nghiệp có ROE 25% và ROIC 22% đang thực sự hiệu quả — chênh 3 điểm là không đáng kể. "
       "Một doanh nghiệp có ROE 45% nhưng ROIC 8% thì con số 45% kia chủ yếu là kết quả của việc vay nợ "
       "hoặc mua lại cổ phiếu tới mức vốn chủ sở hữu teo lại làm mẫu số nhỏ đi.\n"
       "Không áp dụng cho ngân hàng và bảo hiểm, vì các doanh nghiệp này vận hành bằng đòn bẩy theo bản chất.",
       "không vượt quá **15 điểm %**")

metric(doc, "Đà LN — Đà tăng trưởng lợi nhuận",
       "Đà LN = NI Qtr% - NI Growth%     (đơn vị: điểm phần trăm)",
       "Tính trong app",
       "So tốc độ tăng trưởng của quý gần nhất với tốc độ trung bình cả năm. Dương nghĩa là đang "
       "tăng tốc; âm nhiều nghĩa là đang giảm tốc, và con số TTM đẹp chủ yếu là dư âm của các quý cũ.\n"
       "Ví dụ có thật từ một lần quét: IBM có lợi nhuận TTM tăng **84%**, nhưng lợi nhuận quý gần nhất "
       "**giảm 1,3%**. Nhìn cột TTM thì đây là doanh nghiệp tăng trưởng rất mạnh; nhìn cột quý thì "
       "đà tăng đã tắt từ mấy tháng trước.")

metric(doc, "Còn đúng — Số liệu TTM có còn đúng ở quý gần nhất không",
       "Đạt khi:  NI Qtr%  >=  10%\n"
       "   hoặc:  NI Qtr%  >=  NI Growth% - 15 điểm %",
       "Tính trong app",
       "Tiêu chí này bác bỏ khả năng con số TTM đẹp chỉ là dư âm của quá khứ. Nó đạt trong hai "
       "trường hợp: quý gần nhất **vẫn đang tăng trưởng tốt** (trên 10%), hoặc **không tụt quá xa** "
       "so với chính mức TTM của nó.\n"
       "Phải có vế thứ nhất, vì phép so bằng điểm phần trăm sẽ vô nghĩa với doanh nghiệp tăng trưởng "
       "rất cao: một mã có TTM 295% và quý 225% thì chênh nhau tới 70 điểm, nhưng 225% nào có phải "
       "là xấu đi. Ngược lại, một mã TTM 33% mà quý chỉ còn 8,8% thì đó là giảm tốc thật.",
       "quý gần nhất **trên 10%**, hoặc **không tụt quá 15 điểm %** so với mức TTM")

metric(doc, "Bền vững 5N — Số liệu tốt có lặp lại nhiều năm không",
       "Xếp hạng từ ROE trung bình 5 năm và Biên gộp trung bình 5 năm:\n"
       "  WIDE   ★★★  khi ROE ≥ 20% và GM ≥ 50%\n"
       "  NARROW ★★   khi ROE ≥ 15% và GM ≥ 35%\n"
       "  UNCERTAIN ★ khi ROE ≥ 10%\n"
       "  WEAK        các trường hợp còn lại",
       "yfinance — cần bật ô «yfinance Moat», chỉ có dữ liệu cho thị trường Mỹ",
       "Mọi chỉ số phía trên đều đo **kỳ gần nhất**. Một doanh nghiệp có thể có một năm xuất sắc nhờ may mắn, "
       "nhờ một hợp đồng lớn, hoặc nhờ giá nguyên liệu thuận lợi. Tiêu chí này hỏi: kết quả tốt đó có lặp lại "
       "qua năm năm không.\n"
       "**Khi không có dữ liệu, ô để trống chứ app không thay bằng số liệu một kỳ.** Đây là chủ ý: tiêu chí "
       "này tồn tại để đo tính bền vững, mà lấy một kỳ rồi gọi là bền vững thì chính nó đã là một con số ảo.",
       "xếp hạng **WIDE** hoặc **NARROW**")

# ── 3.6 ─────────────────────────────────────────────────────────────────────
h2(doc, "3.6  Nhóm tham khảo — hiển thị nhưng không chấm điểm")

add_table(doc,
          ["Chỉ số", "Công thức", "Đọc thế nào"],
          [
              ["P/E", "Giá / Lãi trên mỗi cổ phiếu",
               "Trả bao nhiêu đồng cho một đồng lợi nhuận mỗi năm. Cao thường phản ánh kỳ vọng tăng trưởng."],
              ["PEG", "P/E / Tốc độ tăng trưởng lợi nhuận",
               "Cân đối định giá với tăng trưởng. Dưới 1 thường được coi là hợp lý."],
              ["EV/EBITDA", "Giá trị doanh nghiệp / Lợi nhuận trước lãi, thuế, khấu hao",
               "Định giá không phụ thuộc cấu trúc vốn, dùng để so giữa các mã có mức nợ khác nhau."],
              ["1Y%", "Thay đổi giá 12 tháng gần nhất",
               "Hiệu suất giá. Không nói gì về chất lượng doanh nghiệp."],
          ],
          widths=[2.4, 5.4, 8.8], size=9)

callout(doc, "Vì sao định giá không được chấm điểm",
        "App trả lời câu hỏi «doanh nghiệp này có tốt và số liệu có thật không», **không** trả lời "
        "«giá hiện tại có rẻ không». Hai câu hỏi đó khác nhau và cần hai loại phân tích khác nhau. "
        "Hệ quả bạn cần nhớ: một mã đạt 🟢 TỐT & THỰC vẫn có thể đang giao dịch ở mức giá rất đắt. "
        "Hãy tự nhìn cột P/E và PEG trước khi quyết định.",
        bg=WARN_BG, fg=ORANGE)

doc.add_page_break()

# ═══════════════════════════════════════════════════════════════════════════
h1(doc, "4.  MƯỜI BA TIÊU CHÍ CHẤM ĐIỂM")

h2(doc, "4.1  Nhóm A — Tài chính tốt")
add_table(doc,
          ["Tiêu chí", "Ngưỡng đạt", "Hỏi điều gì"],
          [
              ["ROIC", "> 15%", "Đồng vốn bỏ ra có sinh lời cao hơn chi phí huy động không?"],
              ["Op Margin", "> 15%", "Vận hành có hiệu quả, chi phí có được kiểm soát không?"],
              ["Gross Mgn", "> 40%", "Có bán được cao hơn hẳn chi phí sản xuất không?"],
              ["NI Growth", "> 20%", "Lợi nhuận có đang tăng không?"],
              ["Rev Growth", "> 15%", "Doanh thu có đang tăng không?"],
              ["D/E", "< 1,0", "Có phụ thuộc quá nhiều vào vay nợ không?"],
          ],
          widths=[3.0, 2.6, 11.0], size=9.5)
para(doc,
     "Với sector **Financial Services**: ROIC thay bằng ROE > 12%, và D/E bị bỏ qua hoàn toàn.", size=9)

h2(doc, "4.2  Nhóm B — Chứng minh là thực")
add_table(doc,
          ["Tiêu chí", "Ngưỡng đạt", "Bác bỏ khả năng nào"],
          [
              ["FCF/NI", "≥ 0,80×", "Lợi nhuận ghi sổ nhưng chưa thu được tiền"],
              ["FCF Mgn", "> 8%", "Biên đẹp nhờ hoãn đầu tư hoặc vốn hóa chi phí"],
              ["Không buyback", "EPS−NI ≤ 10 điểm %", "EPS tăng nhờ mua cổ phiếu quỹ, không phải kinh doanh"],
              ["ROE thật", "ROE−ROIC ≤ 15 điểm %", "ROE cao nhờ đòn bẩy hoặc vốn chủ teo lại"],
              ["Bảng cân đối", "Tiền ròng > 0, hoặc Curr.R > 1,5,\nhoặc Nợ/FCF < 3 năm",
               "Bảng cân đối trông sạch nhưng thực tế không gánh nổi nợ"],
              ["Còn đúng", "Quý > 10% hoặc ≥ TTM − 15 điểm %", "Số TTM đẹp chỉ vì các quý cũ còn đè lên"],
              ["Bền vững 5N", "WIDE hoặc NARROW", "Một kỳ may mắn được trình bày như năng lực dài hạn"],
          ],
          widths=[3.0, 4.4, 9.2], size=9.5)
para(doc,
     "Với sector **Financial Services**: hai tiêu chí ROE thật và Bảng cân đối bị bỏ qua, vì ngân hàng "
     "vận hành bằng đòn bẩy theo bản chất và các chỉ số thanh khoản thông thường không áp dụng được.", size=9)

# ═══════════════════════════════════════════════════════════════════════════
h1(doc, "5.  CÁCH TÍNH ĐIỂM")

h2(doc, "5.1  Ba trạng thái của một tiêu chí")
add_table(doc,
          ["Ký hiệu", "Nghĩa", "Cộng điểm?", "Tính vào mẫu số?"],
          [
              ["✓", "Đạt tiêu chí", "Có", "Có"],
              ["✗", "Không đạt", "Không", "Có"],
              ["—", "Không áp dụng cho ngành này, hoặc thiếu dữ liệu", "Không", "**Không**"],
          ],
          widths=[2.0, 8.0, 3.0, 3.6], size=9.5)

para(doc,
     "Trạng thái thứ ba là điểm quan trọng cần hiểu. Điểm số **không** phải «số tiêu chí đạt trên 6» — "
     "nó là **số tiêu chí đạt trên số tiêu chí thực sự áp dụng được**. Nhờ vậy một ngân hàng, vốn có hai "
     "tiêu chí nhóm B không áp dụng được, không bị coi như đã trượt hai tiêu chí đó.")

code(doc, "Điểm hiển thị  =  số ✓  /  (số ✓ + số ✗)\n"
          "Tỷ lệ đạt      =  số ✓  ÷  (số ✓ + số ✗)")

para(doc,
     "Vì tính theo tỷ lệ nên khi sắp xếp, một mã **4/4** đứng trên một mã **5/6** — đạt tuyệt đối trên "
     "bốn tiêu chí áp dụng được thì tốt hơn là trượt một trong sáu.", size=9.5)

callout(doc, "Khi nào app từ chối xếp hạng",
        "Nếu một nhóm có **dưới 3 tiêu chí** áp dụng được, app trả về ⚪ THIẾU DỮ LIỆU thay vì tính điểm. "
        "Nó không đoán một kết luận từ hai con số lẻ. Bạn sẽ thấy điều này ở các mã mới niêm yết "
        "hoặc quỹ ETF lọt vào danh sách theo vốn hóa.")

h2(doc, "5.2  Ví dụ tính tay — AMZN")
para(doc,
     "Lấy từ một lần quét thật để bạn theo dõi được từng bước. Đây là ví dụ minh hoạ cách app tính, "
     "không phải khuyến nghị đầu tư; số liệu cũng thay đổi theo thời gian.", size=9)

add_table(doc,
          ["Nhóm A", "Kết quả", "", "Nhóm B", "Kết quả"],
          [
              ["ROIC 21,8% > 15%", "✓", "", "FCF/NI = −0,15", "✗"],
              ["Op Margin < 15%", "✗", "", "FCF Mgn = −1,5%", "✗"],
              ["Gross Mgn > 40%", "✓", "", "EPS−NI = −1,8 điểm %", "✓"],
              ["NI Growth > 20%", "✓", "", "ROE−ROIC = 8,8 điểm %", "✓"],
              ["Rev Growth > 15%", "✓", "", "Net Cash âm, Curr.R thấp", "✗"],
              ["D/E < 1,0", "✓", "", "Bền vững 5N — không có dữ liệu", "—"],
          ],
          widths=[5.0, 2.0, 0.6, 5.6, 2.0], size=9,
          colors={(0, 1): GREEN, (1, 1): RED, (2, 1): GREEN, (3, 1): GREEN, (4, 1): GREEN, (5, 1): GREEN,
                  (0, 4): RED, (1, 4): RED, (2, 4): GREEN, (3, 4): GREEN, (4, 4): RED, (5, 4): GREY})

code(doc,
     "Điểm Tài chính  =  5 ✓ / 6 áp dụng  =  5/6  =  83%\n"
     "Điểm Thực       =  2 ✓ / 5 áp dụng  =  2/5  =  40%      (Bền vững 5N không tính vào mẫu số)\n\n"
     "Tài chính ≥ 80%  và  Thực < 50%   →   🔴 CỜ ĐỎ")

para(doc,
     "Đây chính là trường hợp mà cấu trúc hai điểm sinh ra để phát hiện. Nhìn riêng nhóm A thì đây là "
     "một doanh nghiệp rất mạnh — 5 trên 6 tiêu chí đạt. Nhưng dòng tiền tự do đang **âm 11,6 tỷ** trong khi "
     "báo cáo lãi 77,7 tỷ, vì doanh nghiệp đang đầu tư cực nặng. Nếu app gộp hai điểm lại thành 7/11, "
     "mã này sẽ trông ngang với một mã trung bình đều và thực tế đó bị che mất.")

# ═══════════════════════════════════════════════════════════════════════════
h1(doc, "6.  BẢY TÍN HIỆU VÀ CÁCH ĐỌC")

add_table(doc,
          ["Tín hiệu", "Điều kiện", "Nghĩa là gì", "Nên làm gì"],
          [
              ["🟢 TỐT & THỰC", "A ≥ 80%\nB ≥ 80%",
               "Số liệu mạnh và vượt được các phép kiểm chứng",
               "Nhóm đáng đọc kỹ trước tiên. Kiểm tra thêm định giá."],
              ["🔵 KHÁ", "A ≥ 60%\nB ≥ 60%",
               "Cả hai mặt đều khá, còn vài điểm chưa đạt",
               "Mở breakdown xem tiêu chí nào trượt và trượt vì sao."],
              ["🔴 CỜ ĐỎ", "A ≥ 80%\nB < 60%",
               "Số liệu rất đẹp nhưng không chứng minh được",
               "Đọc cột «Vì sao chưa chứng minh được». Không loại ngay — "
               "có thể có lý do chính đáng như đang đầu tư lớn — nhưng phải hiểu lý do đó."],
              ["🟠 THẬT NHƯNG YẾU", "B ≥ 80%\nA < 50%",
               "Sổ sách trung thực, dòng tiền thật, nhưng biên và tăng trưởng chưa hấp dẫn",
               "Doanh nghiệp lành mạnh nhưng chưa hấp dẫn. Có thể theo dõi chờ cải thiện."],
              ["🟡 THEO DÕI", "A ≥ 50%", "Chưa đủ điều kiện", "Đưa vào watchlist, quét lại sau."],
              ["⚫ LOẠI", "còn lại", "Không đạt cả hai mặt", "Bỏ qua."],
              ["⚪ THIẾU DỮ LIỆU", "< 3 tiêu chí\náp dụng được",
               "Không đủ cơ sở để kết luận", "Tra cứu thủ công nếu quan tâm."],
          ],
          widths=[3.2, 2.4, 5.0, 6.0], size=8.5,
          colors={(0, 0): GREEN, (1, 0): BLUE, (2, 0): RED, (3, 0): ORANGE,
                  (4, 0): AMBER, (5, 0): GREY, (6, 0): GREY})

para(doc,
     "**A** là tỷ lệ đạt nhóm Tài chính, **B** là tỷ lệ đạt nhóm Kiểm chứng. "
     "Lưu ý 🔴 CỜ ĐỎ **không phải** là «xấu hơn ⚫ LOẠI». Nó là một cảnh báo có nội dung: "
     "mã này có số liệu tốt, nhưng có một khoảng cách giữa những gì báo cáo nói và những gì tiền mặt xác nhận, "
     "và bạn cần hiểu khoảng cách đó trước khi làm gì tiếp.", size=9.5)

doc.add_page_break()

# ═══════════════════════════════════════════════════════════════════════════
h1(doc, "7.  CÁCH CHỌN MÃ")

h2(doc, "7.1  Đọc Dashboard")
para(doc, "Dashboard có bốn phần, đọc theo thứ tự này:")

add_table(doc,
          ["Phần", "Cho biết điều gì", "Dùng thế nào"],
          [
              ["Dải KPI trên cùng",
               "Tổng số mã, số mã TỐT & THỰC, số CỜ ĐỎ, số KHÁ, và điểm trung bình hai nhóm",
               "Nhìn nhanh xem lần quét này có gì. Điểm trung bình Thực thấp hơn hẳn Tài chính "
               "nghĩa là cả thị trường đang có khoảng cách giữa lợi nhuận và tiền mặt."],
              ["🟢 TỐT & THỰC",
               "Danh sách đầy đủ mọi mã đạt cả hai mặt, sắp xếp theo điểm Thực rồi tới điểm Tài chính",
               "Đây là danh sách rút gọn để bạn đọc kỹ. Không bị cắt bớt — số lượng ghi ngay trên tiêu đề."],
              ["🔴 CỜ ĐỎ",
               "Mã có số liệu đẹp nhưng không chứng minh được, kèm cột nêu **chính xác** trượt ở điểm nào",
               "Đọc cột lý do. Đây thường là phần cho bạn nhiều thông tin nhất trong cả Dashboard."],
              ["📊 Tỷ lệ đạt từng tiêu chí",
               "Trong lần quét này, mỗi tiêu chí có bao nhiêu phần trăm số mã đạt",
               "Tiêu chí nào trên 90% là quá lỏng, dưới 5% là quá chặt — cả hai đều không phân loại được gì. "
               "Dùng để biết nên chỉnh ngưỡng nào."],
              ["🏢 Theo ngành",
               "Mỗi ngành có bao nhiêu mã tốt, bao nhiêu cờ đỏ, điểm trung bình hai nhóm",
               "Cho thấy ngành nào số liệu đẹp mà khó chứng minh nhất."],
          ],
          widths=[3.6, 6.4, 6.6], size=9)

h2(doc, "7.2  Quy trình chọn mã gợi ý")
for i, (title, body) in enumerate([
    ("Bắt đầu từ danh sách TỐT & THỰC trên Dashboard",
     "Đây đã là kết quả lọc hai lớp. Với 100–200 mã quét, danh sách này thường còn 20–40 mã."),
    ("Loại bớt bằng định giá",
     "App không chấm điểm định giá, nên bước này bạn phải tự làm. Mở tab Tài chính, xem cột P/E. "
     "Một mã đạt mọi tiêu chí nhưng P/E 90 lần là một quyết định khác hẳn so với cùng mã đó ở P/E 20."),
    ("Mở breakdown từng mã còn lại",
     "Click vào dòng. Thẻ chi tiết bên dưới hiện cả 6 tiêu chí của tab đang xem kèm giá trị thực tế. "
     "Chuyển sang tab kia để xem 6 tiêu chí còn lại. Rê chuột lên từng ô để đọc giải thích."),
    ("Đọc kỹ ba cột Doanh thu · Lợi nhuận · FCF",
     "Ở tab Kiểm chứng, ba cột này nằm cạnh nhau. Lợi nhuận và FCF xấp xỉ nhau là dấu hiệu tốt. "
     "Chênh lệch lớn thì phải tìm hiểu nguyên nhân trước khi đi tiếp."),
    ("Xem qua danh sách CỜ ĐỎ",
     "Đừng bỏ qua phần này. Đôi khi một mã cờ đỏ hóa ra là cơ hội — doanh nghiệp đang đầu tư nặng "
     "cho tăng trưởng tương lai. Nhưng bạn phải chủ động biết điều đó, thay vì mua nhầm vì tưởng "
     "số liệu đều đẹp."),
    ("Đọc báo cáo tài chính thật trước khi quyết định",
     "App thu hẹp từ hàng trăm mã xuống vài chục mã đáng đọc. Nó không thay thế việc đọc báo cáo, "
     "tìm hiểu mô hình kinh doanh và ban lãnh đạo."),
]):
    h3(doc, f"Bước {i+1} — {title}")
    para(doc, body, size=9.5, space_after=4)

h2(doc, "7.3  Ba tình huống điển hình")

h3(doc, "Số liệu tốt và tiền mặt xác nhận — NVDA", GREEN)
para(doc,
     "Doanh thu 253 tỷ, lợi nhuận 120 tỷ, FCF 119 tỷ. Tỷ lệ FCF/NI là **0,99** — gần như mỗi đồng lợi nhuận "
     "đều có một đồng tiền mặt tương ứng. ROIC 106% với chênh ROE−ROIC chỉ 8,1 điểm nghĩa là mức sinh lời "
     "khổng lồ này đến từ kinh doanh chứ không từ đòn bẩy. Điểm 6/6 và 5/5.", size=9.5)

h3(doc, "Lợi nhuận lớn nhưng tiền chảy ra — AMZN", RED)
para(doc,
     "Doanh thu 776 tỷ, lợi nhuận 78 tỷ, nhưng FCF **âm 11,6 tỷ**. Tài chính 5/6 rất đẹp; kiểm chứng chỉ 2/5. "
     "Nguyên nhân là đầu tư rất nặng vào hạ tầng. Đây không phải gian lận và cũng chưa chắc là mã xấu — "
     "nhưng nếu bạn chỉ nhìn cột lợi nhuận thì bạn hoàn toàn không thấy được điều này. Vai trò của app là "
     "đưa nó lên bề mặt để bạn tự quyết định.", size=9.5)

h3(doc, "ROE cao nhưng phần lớn là đòn bẩy — AAPL", AMBER)
para(doc,
     "ROE lên tới 148,8% — con số này nếu đứng một mình sẽ trông phi thường. Nhưng ROIC là 78,8%, tức "
     "chênh lệch **69,9 điểm phần trăm**. Khoảng cách đó đến từ việc doanh nghiệp mua lại cổ phiếu quỹ "
     "trong nhiều năm khiến vốn chủ sở hữu — mẫu số của ROE — teo lại. ROIC 78,8% vẫn là xuất sắc, "
     "nhưng nó mới là con số phản ánh đúng hiệu quả kinh doanh. Tiêu chí ROE thật vì vậy báo ✗.", size=9.5)

h2(doc, "7.4  Bốn sai lầm thường gặp")
bullet(doc, "**Cộng hai điểm lại với nhau.** 6/6 và 1/5 không phải là «7 trên 11». Đó là một cảnh báo. "
            "Hai điểm được thiết kế để đọc riêng.")
bullet(doc, "**Coi 🔴 CỜ ĐỎ là mã xấu.** Nó nghĩa là «có khoảng cách chưa giải thích được», không phải "
            "«tránh xa». Đọc cột lý do rồi mới kết luận.")
bullet(doc, "**Quên rằng app không xét giá.** Điểm cao không có nghĩa giá đang hợp lý. "
            "Định giá là việc bạn phải tự làm.")
bullet(doc, "**Nhầm dấu — với dấu ✗.** Dấu gạch ngang nghĩa là tiêu chí không áp dụng được hoặc thiếu "
            "dữ liệu, và nó **không** bị tính vào mẫu số. Một mã 4/4 không hề kém một mã 6/6.")

# ═══════════════════════════════════════════════════════════════════════════
h1(doc, "8.  GIỚI HẠN CẦN BIẾT")

para(doc,
     "App được xây để thu hẹp danh sách, không phải để ra quyết định thay bạn. Những giới hạn dưới đây "
     "là thật và bạn nên biết trước khi tin vào điểm số.")

h3(doc, "Chỉ quét cổ phiếu vốn hóa lớn", AMBER)
para(doc,
     "App lấy top N mã theo vốn hóa rồi mới chấm điểm. Mọi kết luận chỉ có giá trị bên trong nhóm đó — "
     "nó không tìm được cơ hội ở nhóm vốn hóa vừa và nhỏ.", size=9.5)

h3(doc, "Ngưỡng biên lợi nhuận dùng chung cho mọi ngành", AMBER)
para(doc,
     "Biên gộp 40% là mức bình thường với phần mềm nhưng gần như không thể đạt với bán lẻ hay công nghiệp "
     "nặng. Hai tiêu chí Gross Mgn và Op Margin vì thế đo **cổ phiếu thuộc ngành nào** nhiều hơn là đo chất "
     "lượng doanh nghiệp. Hệ quả: hệ thống thiên lệch về phía công nghệ và y tế. Khi xem một mã bán lẻ, "
     "hãy tự điều chỉnh kỳ vọng ở hai tiêu chí này.", size=9.5)

h3(doc, "Ngưỡng là đúng hoặc sai, không có mức độ", AMBER)
para(doc,
     "Một mã ROIC 14,9% và một mã 15,1% chênh nhau trọn một điểm dù thực tế gần như bằng nhau. Ngược lại, "
     "16% và 60% được tính điểm giống hệt nhau. Vì vậy đừng chỉ nhìn ✓/✗ — hãy nhìn cả giá trị thực tế "
     "hiện ngay trong ô ở thẻ chi tiết.", size=9.5)

h3(doc, "Phân loại ngành đôi khi lệch", AMBER)
para(doc,
     "App dựa vào sector của TradingView để quyết định tiêu chí nào được bỏ qua. Có trường hợp lệch: "
     "Mastercard bị xếp vào Financial Services nên được miễn hai tiêu chí như thể nó là ngân hàng, "
     "dù đây là một mạng lưới thanh toán. Nếu thấy một mã có mẫu số nhỏ bất thường, hãy kiểm tra cột Ngành.",
     size=9.5)

h3(doc, "Không có định giá và không có quy tắc bán", RED)
para(doc,
     "Không tiêu chí nào liên quan tới giá phải trả. Và app chỉ trả lời «xem mã nào», hoàn toàn không có "
     "quy tắc thoát lệnh hay cắt lỗ. Hai mảnh này bạn phải tự bổ sung.", size=9.5)

h3(doc, "Các ngưỡng chưa được kiểm định trên dữ liệu lịch sử", RED)
para(doc,
     "Toàn bộ ngưỡng — 15%, 20%, 40%, 0,80 lần, 10 và 15 điểm phần trăm — đều lấy từ chuẩn mực phổ biến "
     "hoặc chọn theo lập luận. App chưa lưu dữ liệu theo thời điểm nên chưa thể kiểm thử lùi, và hiện "
     "**chưa có bằng chứng nào cho thấy mã 6/6 thực sự cho kết quả đầu tư tốt hơn mã 4/6**. "
     "Hãy đọc điểm số như một thứ tự ưu tiên để nghiên cứu, không phải như một dự báo.", size=9.5)

callout(doc, "Bảng tỷ lệ đạt giúp bạn tự hiệu chỉnh",
        "Dashboard và sheet Tổng hợp đều có bảng tỷ lệ đạt từng tiêu chí, tự gắn nhãn «quá lỏng» khi trên "
        "90% số mã đạt và «quá chặt» khi dưới 5% đạt. Đây chưa phải kiểm thử lùi, nhưng nó cho biết ngay "
        "tiêu chí nào đang không phân loại được gì — thông tin cần có trước khi chỉnh bất kỳ ngưỡng nào.")

# ═══════════════════════════════════════════════════════════════════════════
h1(doc, "9.  CÂU HỎI THƯỜNG GẶP")

for q, a in [
    ("Có bắt buộc bật «yfinance Moat» không?",
     "Không. Tắt thì app chạy đầy đủ 12 trên 13 tiêu chí và nhanh hơn rất nhiều — vài giây thay vì 5–8 phút "
     "cho 100 mã. Chỉ mất tiêu chí Bền vững 5 năm, và tiêu chí đó khi thiếu sẽ để trống chứ không bị tính "
     "là trượt. Với thị trường ngoài Mỹ thì bật cũng không có tác dụng vì không có dữ liệu."),
    ("Vì sao hai mã cùng điểm 5 mà một mã ghi 5/6 còn mã kia ghi 5/5?",
     "Vì mẫu số là số tiêu chí **thực sự áp dụng được**, không phải hằng số. Mã ghi 5/5 có một tiêu chí "
     "không áp dụng cho ngành của nó hoặc thiếu dữ liệu, nên tiêu chí đó bị loại khỏi mẫu số. "
     "Khi sắp xếp, 5/5 đứng trên 5/6 vì tỷ lệ cao hơn."),
    ("Mã tôi quan tâm hiện ⚪ THIẾU DỮ LIỆU, phải làm sao?",
     "Nghĩa là dưới 3 tiêu chí của một nhóm có dữ liệu, thường gặp với mã mới niêm yết hoặc quỹ ETF. "
     "App cố ý không đoán trong trường hợp này. Bạn cần tra cứu thủ công."),
    ("Vì sao một mã tăng giá rất mạnh lại bị xếp ⚫ LOẠI?",
     "Vì app không chấm điểm giá. Nó chỉ nhìn chất lượng tài chính và tính xác thực của số liệu. "
     "Cột 1Y% có hiển thị hiệu suất giá nhưng không tham gia tính điểm."),
    ("Số liệu cập nhật đến thời điểm nào?",
     "TradingView trả về số liệu TTM và quý gần nhất đã công bố. Nghĩa là sau khi doanh nghiệp ra báo cáo "
     "mới, có thể mất vài ngày để dữ liệu được cập nhật. Quét lại để lấy số mới nhất."),
    ("Có dùng được cho thị trường Việt Nam không?",
     "Được. Chọn «vietnam» ở ô thị trường. Toàn bộ 12 tiêu chí lấy từ TradingView đều hoạt động; "
     "chỉ riêng tiêu chí Bền vững 5 năm không có dữ liệu. Giá và vốn hóa sẽ hiển thị theo đồng Việt Nam."),
    ("Excel xuất ra gồm những gì?",
     "Ba sheet. **Screener**: toàn bộ mã kèm số liệu và cả 13 tiêu chí. **Cờ đỏ**: riêng nhóm số đẹp mà "
     "chưa chứng minh được, kèm lý do cụ thể. **Tổng hợp**: phân bố tín hiệu, tỷ lệ đạt từng tiêu chí, "
     "và thống kê theo ngành. Nếu bật yfinance còn có thêm sheet Financials với biểu đồ từng mã."),
] :
    h3(doc, q)
    para(doc, a, size=9.5, space_after=5)

# ═══════════════════════════════════════════════════════════════════════════
doc.add_page_break()
h1(doc, "PHỤ LỤC — TRA CỨU NHANH")

h2(doc, "A.  Công thức")
add_table(doc,
          ["Chỉ số", "Công thức"],
          [
              ["ROIC", "Lợi nhuận sau thuế từ kinh doanh / (Vốn chủ + Nợ vay)"],
              ["ROE", "Lợi nhuận ròng / Vốn chủ sở hữu"],
              ["Op Margin", "Lợi nhuận hoạt động / Doanh thu"],
              ["Gross Margin", "(Doanh thu − Giá vốn hàng bán) / Doanh thu"],
              ["Net Margin", "Lợi nhuận ròng / Doanh thu"],
              ["NI Growth", "Lợi nhuận 4 quý gần nhất so 4 quý cùng kỳ năm trước"],
              ["Rev Growth", "Doanh thu 4 quý gần nhất so 4 quý cùng kỳ năm trước"],
              ["D/E", "Tổng nợ vay / Vốn chủ sở hữu"],
              ["FCF", "Tiền từ hoạt động kinh doanh − Chi mua sắm tài sản"],
              ["FCF/NI", "Dòng tiền tự do / Lợi nhuận ròng"],
              ["FCF Margin", "Dòng tiền tự do / Doanh thu"],
              ["EPS−NI", "EPS Growth% − NI Growth%   (điểm phần trăm)"],
              ["ROE−ROIC", "ROE% − ROIC%   (điểm phần trăm)"],
              ["Net Cash", "Tiền + Đầu tư ngắn hạn − Tổng nợ vay"],
              ["Nợ/FCF", "Nợ ròng / Dòng tiền tự do một năm   (số năm)"],
              ["Current Ratio", "Tài sản ngắn hạn / Nợ ngắn hạn"],
              ["NI Qtr% / Rev Qtr%", "Quý vừa công bố so với cùng quý năm trước"],
              ["Đà LN", "NI Qtr% − NI Growth%   (điểm phần trăm)"],
              ["Bền vững 5N", "Xếp hạng từ ROE 5 năm và Biên gộp 5 năm"],
              ["Điểm hiển thị", "Số tiêu chí ✓ / Số tiêu chí áp dụng được"],
          ],
          widths=[4.0, 12.6], size=9)

h2(doc, "B.  Ngưỡng đạt")
add_table(doc,
          ["Nhóm A — Tài chính", "Ngưỡng", "Nhóm B — Kiểm chứng", "Ngưỡng"],
          [
              ["ROIC", "> 15%", "FCF/NI", "≥ 0,80×"],
              ["Op Margin", "> 15%", "FCF Mgn", "> 8%"],
              ["Gross Mgn", "> 40%", "Không buyback", "≤ 10 điểm %"],
              ["NI Growth", "> 20%", "ROE thật", "≤ 15 điểm %"],
              ["Rev Growth", "> 15%", "Bảng cân đối", "Tiền ròng>0, CR>1,5, hoặc Nợ/FCF<3"],
              ["D/E", "< 1,0", "Còn đúng", "Quý > 10% hoặc ≥ TTM − 15 đ%"],
              ["", "", "Bền vững 5N", "WIDE hoặc NARROW"],
          ],
          widths=[3.8, 2.6, 4.2, 6.0], size=9)

h2(doc, "C.  Ngưỡng tín hiệu")
add_table(doc,
          ["Tín hiệu", "Tài chính (A)", "Kiểm chứng (B)"],
          [
              ["🟢 TỐT & THỰC", "≥ 80%", "≥ 80%"],
              ["🔵 KHÁ", "≥ 60%", "≥ 60%"],
              ["🔴 CỜ ĐỎ", "≥ 80%", "< 60%"],
              ["🟠 THẬT NHƯNG YẾU", "< 50%", "≥ 80%"],
              ["🟡 THEO DÕI", "≥ 50%", "—"],
              ["⚫ LOẠI", "còn lại", "—"],
              ["⚪ THIẾU DỮ LIỆU", "dưới 3 tiêu chí áp dụng được ở một trong hai nhóm", ""],
          ],
          widths=[4.4, 5.2, 7.0], size=9,
          colors={(0, 0): GREEN, (1, 0): BLUE, (2, 0): RED, (3, 0): ORANGE,
                  (4, 0): AMBER, (5, 0): GREY, (6, 0): GREY})

h2(doc, "D.  Tiêu chí bị bỏ qua theo ngành")
add_table(doc,
          ["Ngành", "Tiêu chí bị bỏ qua hoặc thay thế", "Lý do"],
          [
              ["Financial Services\n(ngân hàng, bảo hiểm,\nmôi giới)",
               "ROIC → thay bằng ROE > 12%\nD/E → bỏ qua\nROE thật → bỏ qua\nBảng cân đối → bỏ qua",
               "Tiền gửi khách hàng hạch toán là nợ nên D/E 8–15 lần là bình thường. "
               "Các doanh nghiệp này vận hành bằng đòn bẩy theo bản chất, "
               "nên các chỉ số nợ và thanh khoản thông thường không áp dụng được."],
              ["Mọi ngành khi thiếu dữ liệu",
               "Tiêu chí thiếu dữ liệu để trống",
               "Không tính vào mẫu số, nên không bị coi là trượt."],
          ],
          widths=[4.4, 5.4, 6.8], size=9)

p = doc.add_paragraph()
p.paragraph_format.space_before = Pt(16)
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = p.add_run("— Hết —")
r.italic = True; r.font.size = Pt(9); r.font.color.rgb = GREY; r.font.name = FONT

out = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                   "Huong_dan_su_dung.docx")
doc.save(out)
print("Da tao:", out)
