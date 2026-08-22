# -*- coding: utf-8 -*-
"""
Sinh file Word: "Tiêu chí lựa chọn cổ phiếu — Diễn giải & Đánh giá tính khả thi"

Chạy:  ./venv/Scripts/python.exe make_criteria_doc.py
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
# XÂY DỰNG TÀI LIỆU
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

# ── Trang bìa ───────────────────────────────────────────────────────────────
for _ in range(5):
    doc.add_paragraph()

p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = p.add_run("STOCK FOUNDATION SCREENER")
r.bold = True; r.font.size = Pt(13); r.font.color.rgb = GREY; r.font.name = FONT

p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_after = Pt(4)
r = p.add_run("TIÊU CHÍ LỰA CHỌN CỔ PHIẾU")
r.bold = True; r.font.size = Pt(26); r.font.color.rgb = NAVY; r.font.name = FONT

p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = p.add_run("Số liệu tài chính tốt — và bằng chứng rằng số liệu đó là thật")
r.font.size = Pt(14); r.font.color.rgb = BLUE; r.font.name = FONT

doc.add_paragraph()
p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = p.add_run("Phân tích trực tiếp từ mã nguồn\npython screen_top100.py  ·  screener_gui.py")
r.italic = True; r.font.size = Pt(10); r.font.color.rgb = GREY; r.font.name = FONT

for _ in range(5):
    doc.add_paragraph()

t = doc.add_table(rows=1, cols=1); t.style = "Table Grid"
c = t.rows[0].cells[0]; shade(c, ALT_BG)
cell_text(c,
          "Bộ lọc chỉ hỏi hai câu và cố ý giữ chúng tách rời. Nhóm A hỏi số liệu tài chính có tốt không. "
          "Nhóm B hỏi số liệu đó có thật không — mỗi tiêu chí nhóm B bác bỏ một cách cụ thể mà số liệu "
          "nhóm A có thể bị làm đẹp. Hai điểm không được cộng gộp, vì một mã 6/6 tài chính nhưng 1/6 "
          "kiểm chứng là trường hợp nguy hiểm nhất, và gộp lại thành 7/12 sẽ khiến nó trông y hệt "
          "một mã trung bình đều.",
          size=10, color=BLACK, align=WD_ALIGN_PARAGRAPH.JUSTIFY)

doc.add_page_break()

# ═══════════════════════════════════════════════════════════════════════════
h1(doc, "1.  TRIẾT LÝ CỦA BỘ LỌC")

para(doc,
     "Hầu hết bộ lọc cổ phiếu chỉ làm một việc: tìm những con số đẹp. Vấn đề là con số đẹp trên báo cáo "
     "tài chính có thể đến từ nhiều nguồn, và không phải nguồn nào cũng phản ánh một doanh nghiệp tốt. "
     "Lợi nhuận có thể được ghi nhận trước khi thu tiền. Biên lợi nhuận có thể được nâng lên bằng cách "
     "vốn hóa chi phí hoặc hoãn đầu tư. EPS có thể tăng chỉ vì công ty mua lại cổ phiếu quỹ làm giảm "
     "mẫu số. ROE có thể cao vì đòn bẩy lớn chứ không phải vì kinh doanh hiệu quả.")

para(doc,
     "Bộ lọc này được cấu trúc quanh đúng vấn đề đó. Nó chia làm hai nhóm và **không bao giờ cộng gộp "
     "điểm của chúng**:")

add_table(doc,
          ["Nhóm", "Câu hỏi", "Số tiêu chí", "Nguồn dữ liệu"],
          [
              ["A — Tài chính tốt", "Số liệu tài chính có mạnh không?", "6",
               "TradingView (có cho mọi thị trường)"],
              ["B — Chứng minh là thực", "Số liệu đó có thật không?", "7",
               "TradingView (6) + yfinance (1)"],
          ],
          widths=[4.2, 6.4, 2.4, 5.0], size=9.5)

callout(doc, "Vì sao không gộp thành một điểm",
        "Nếu cộng thành một thang 12 điểm, mã đạt **6 tài chính + 1 kiểm chứng** (rất nguy hiểm) và mã đạt "
        "**3 + 4** (trung bình đều) sẽ cùng ra 7/12 và trông giống hệt nhau. Tách riêng thì trường hợp đầu "
        "hiện lên thành tín hiệu **🔴 CỜ ĐỎ** — đúng thứ mà bộ lọc được xây để phát hiện.")

h2(doc, "1.1  Vũ trụ cổ phiếu")
para(doc,
     "Chương trình lấy **top N cổ phiếu theo vốn hóa** của sàn được chọn rồi mới chấm điểm trên tập đó. "
     "Mọi kết luận vì thế chỉ có giá trị bên trong nhóm vốn hóa lớn. Các sàn hỗ trợ: america, nasdaq, "
     "nyse, euronext, hong_kong, vietnam.", size=9.5)

# ═══════════════════════════════════════════════════════════════════════════
h1(doc, "2.  NHÓM A — TÀI CHÍNH TỐT")

para(doc,
     "Toàn bộ sáu tiêu chí lấy từ TradingView trong một lệnh gọi duy nhất, nên có dữ liệu cho mọi thị trường "
     "và không phụ thuộc nguồn phụ nào.")

add_table(doc,
          ["Tiêu chí", "Trường dữ liệu", "Ngưỡng", "Đo cái gì"],
          [
              ["ROIC", "ROIC%", "> 15%", "Sinh lời trên vốn đầu tư, cao hơn chi phí vốn khoảng 8–10%"],
              ["Op Margin", "Op Margin%", "> 15%", "Biên lợi nhuận hoạt động — kỷ luật chi phí, đòn bẩy vận hành"],
              ["Gross Mgn", "Gross Margin%", "> 40%", "Biên lợi nhuận gộp — lợi thế cấu trúc về giá bán"],
              ["NI Growth", "NI Growth%", "> 20%", "Lợi nhuận ròng TTM tăng so cùng kỳ"],
              ["Rev Growth", "Rev Growth%", "> 15%", "Doanh thu TTM tăng so cùng kỳ"],
              ["D/E", "D/E", "< 1,0", "Nợ trên vốn chủ sở hữu"],
          ],
          widths=[2.6, 3.4, 2.2, 9.8], size=9)

para(doc,
     "**Hai điểm đáng lưu ý về thiết kế.** Thứ nhất, tăng trưởng dùng số **TTM** (bốn quý gần nhất) chứ "
     "không dùng số một quý đơn lẻ — mượt hơn và ít nhiễu theo mùa vụ. Thứ hai, hệ thống đo tăng trưởng "
     "**lợi nhuận ròng** chứ không đo tăng trưởng EPS; đây là lựa chọn có chủ đích, vì EPS chịu ảnh hưởng "
     "của số cổ phiếu lưu hành, và phần chênh lệch giữa hai con số này chính là nguyên liệu cho một tiêu chí "
     "kiểm chứng ở nhóm B.", size=9.5)

h2(doc, "2.1  Ba cột số tuyệt đối để đọc bằng mắt")
para(doc,
     "Ngoài các tiêu chí chấm điểm, bảng còn hiển thị **Doanh thu**, **Lợi nhuận ròng** và **FCF** ở dạng "
     "số tuyệt đối (đơn vị tỷ). Chúng **không tham gia chấm điểm** — số tuyệt đối không so sánh được giữa "
     "các mã có quy mô khác nhau, nên tiêu chí phải dùng tăng trưởng và biên. Nhưng đặt ba cột này cạnh "
     "nhau thì phép kiểm chứng quan trọng nhất đọc được ngay bằng mắt: nếu lợi nhuận 120 tỷ mà FCF chỉ 20 tỷ, "
     "khoảng cách đó hiện ra lập tức mà không cần nhìn tỷ số. Cột Lợi nhuận và FCF vì thế được đặt liền nhau "
     "ở tab Kiểm chứng, ngay trước cột FCF/NI.", size=9.5)

h2(doc, "2.1  Xử lý riêng cho nhóm Tài chính")
para(doc,
     "Với sector Financial Services, ROIC được thay bằng **ROE > 12%**, và tiêu chí D/E bị **bỏ qua hoàn toàn** "
     "vì tiền gửi khách hàng được hạch toán là nợ nên tỷ lệ 8–15 lần là cấu trúc bình thường của ngân hàng. "
     "Tiêu chí bị bỏ qua trả về «không áp dụng» và **không bị tính vào mẫu số** — xem mục 4.1.", size=9.5)

# ═══════════════════════════════════════════════════════════════════════════
h1(doc, "3.  NHÓM B — CHỨNG MINH LÀ THỰC")

para(doc,
     "Đây là phần khác biệt của hệ thống. Mỗi tiêu chí không được chọn vì «nó là một chỉ số tốt», mà vì "
     "nó **bác bỏ một cách cụ thể** mà số liệu nhóm A có thể bị làm đẹp.")

add_table(doc,
          ["Số liệu đẹp ở nhóm A", "Có thể ảo vì", "Tiêu chí kiểm chứng tương ứng"],
          [
              ["Lợi nhuận cao", "Ghi nhận doanh thu chưa thu được tiền",
               "FCF/NI ≥ 0,8 — dòng tiền tự do phải bằng ít nhất 80% lợi nhuận ròng"],
              ["Biên lợi nhuận cao", "Vốn hóa chi phí, hoãn đầu tư cho đẹp sổ",
               "FCF Mgn > 8% — tiền thật trên mỗi đồng doanh thu"],
              ["EPS tăng mạnh", "Mua cổ phiếu quỹ làm giảm mẫu số",
               "EPS − NI ≤ 10 điểm % — EPS không được tăng nhanh hơn lợi nhuận quá nhiều"],
              ["ROE cao", "Đòn bẩy lớn, hoặc vốn chủ teo vì mua cổ phiếu quỹ",
               "ROE − ROIC ≤ 15 điểm % — phần chênh chính là đóng góp của đòn bẩy"],
              ["Bảng cân đối sạch", "Thanh khoản thực tế yếu, nợ ẩn",
               "Net Cash > 0 hoặc Current Ratio > 1,5"],
              ["Số TTM đang tăng", "Ba quý cũ còn đè lên, quý gần nhất đã xấu đi",
               "Còn đúng — quý gần nhất > 10% hoặc không tụt quá 15 điểm % so với TTM"],
              ["Một năm xuất sắc", "Một kỳ may mắn trình bày như năng lực dài hạn",
               "ROE và biên gộp trung bình 5 năm đều ở mức cao"],
          ],
          widths=[3.6, 5.4, 8.6], size=9)

h2(doc, "3.1  Bốn tiêu chí đáng chú ý nhất")

h3(doc, "FCF/NI — phép thử accrual")
para(doc,
     "Tỷ lệ dòng tiền tự do trên lợi nhuận ròng là phép thử trực tiếp nhất cho câu hỏi «lợi nhuận này có "
     "thật không». Lợi nhuận là con số kế toán, dòng tiền tự do là tiền thật còn lại sau khi đã chi đầu tư. "
     "Khi hai con số này lệch nhau kéo dài, phần chênh nằm ở các khoản phải thu, hàng tồn kho hoặc các bút "
     "toán dồn tích. Đây cũng là nền tảng của hiện tượng accrual anomaly đã được ghi nhận rộng rãi trong "
     "nghiên cứu tài chính. Khi công ty đang lỗ, tỷ lệ này trả về **rỗng** thay vì một con số vô nghĩa.",
     size=9.5)

h3(doc, "Chênh lệch EPS trừ lợi nhuận ròng — phép thử buyback")
para(doc,
     "Nếu lợi nhuận ròng tăng 10% mà EPS tăng 25%, phần chênh 15 điểm phần trăm không đến từ kinh doanh — "
     "nó đến từ việc số cổ phiếu lưu hành giảm. Cổ đông hiện hữu vẫn hưởng lợi, nhưng đây không phải "
     "tăng trưởng nội tại, và nó không lặp lại được mãi. Phép thử này đơn giản nhưng rất ít bộ lọc có, "
     "và nó chỉ khả thi vì hệ thống đã chọn đo tăng trưởng lợi nhuận ròng ở nhóm A thay vì đo EPS.",
     size=9.5)

h3(doc, "Còn đúng — phép thử độ trễ của số TTM")
para(doc,
     "Mọi chỉ số tăng trưởng của nhóm A đều dùng số TTM, tức cộng dồn bốn quý. Cách này mượt và ít nhiễu "
     "theo mùa vụ, nhưng đổi lại nó **có độ trễ**: một doanh nghiệp bắt đầu xấu đi vẫn giữ số TTM đẹp thêm "
     "hai ba quý nữa vì các quý tốt cũ còn đang đè lên. Tiêu chí này so tăng trưởng lợi nhuận của riêng quý "
     "gần nhất với mức TTM để phát hiện khoảng cách đó.\n"
     "Ví dụ có thật: IBM có lợi nhuận TTM tăng 84% nhưng lợi nhuận quý gần nhất giảm 1,3%. Nhìn cột TTM thì "
     "đây là doanh nghiệp tăng trưởng rất mạnh; nhìn cột quý thì đà tăng đã tắt.\n"
     "Điều kiện đạt gồm hai vế: quý gần nhất trên 10%, **hoặc** không tụt quá 15 điểm % so với TTM. "
     "Vế đầu là bắt buộc, vì phép so bằng điểm phần trăm vô nghĩa với doanh nghiệp tăng trưởng ba chữ số — "
     "một mã có TTM 295% và quý 225% chênh nhau 70 điểm nhưng không hề xấu đi.",
     size=9.5)

h3(doc, "Chênh lệch ROE trừ ROIC — phép thử đòn bẩy")
para(doc,
     "ROE đo lợi nhuận trên vốn chủ sở hữu; ROIC đo lợi nhuận trên toàn bộ vốn đầu tư gồm cả nợ. Khoảng cách "
     "giữa hai con số chính là phần đóng góp của cấu trúc vốn. Một công ty có ROE 25% và ROIC 22% đang thật "
     "sự hiệu quả. Một công ty có ROE 45% nhưng ROIC 8% đang vay nhiều, hoặc đã mua lại cổ phiếu tới mức "
     "vốn chủ sở hữu teo lại và làm mẫu số nhỏ đi một cách giả tạo.",
     size=9.5)

h2(doc, "3.2  Nguồn dữ liệu và một cảnh báo cố ý")
para(doc,
     "Năm trong sáu tiêu chí nhóm B lấy trực tiếp từ TradingView, nên chúng chạy được cho mọi thị trường "
     "và không cần vòng lặp yfinance vốn mất 5–8 phút cho 100 mã. Riêng tiêu chí **Bền vững 5 năm** cần "
     "trung bình 5 năm của ROE và biên gộp, thứ TradingView không cung cấp, nên phải lấy từ yfinance "
     "và chỉ có dữ liệu cho thị trường Mỹ.", size=9.5)

callout(doc, "Khi thiếu dữ liệu 5 năm, hệ thống để trống chứ không thay bằng số liệu một kỳ",
        "Đây là lựa chọn có chủ đích. Phiên bản trước, khi không có yfinance, chương trình tính Moat Score "
        "từ ROE và biên gộp của **kỳ gần nhất**. Nhưng tiêu chí này tồn tại để đo tính **bền vững** — "
        "lấy một kỳ rồi gọi đó là bền vững thì chính nó đã là một con số ảo, đúng thứ mà cả nhóm B sinh ra "
        "để chống lại. Nay tiêu chí trả về «thiếu dữ liệu» và không tính vào mẫu số.")

# ═══════════════════════════════════════════════════════════════════════════
h1(doc, "4.  CÁCH CHẤM ĐIỂM VÀ TÍN HIỆU")

h2(doc, "4.1  Chấm theo tỷ lệ, không theo mẫu số cố định")
para(doc,
     "Điểm là **số tiêu chí đạt chia cho số tiêu chí thực sự áp dụng được**, chứ không chia cho hằng số 6. "
     "Cách này xử lý gọn hai vấn đề cùng lúc: mã thiếu dữ liệu không bị phạt như thể đã trượt, và nhóm "
     "Tài chính không bị thiệt vì có những tiêu chí vốn không áp dụng được cho ngân hàng. Một mã đạt 4 trên "
     "4 tiêu chí áp dụng được sẽ xếp trên một mã đạt 5 trên 6.")

add_table(doc,
          ["Ký hiệu", "Nghĩa", "Có tính vào mẫu số không?"],
          [
              ["✓", "Đạt tiêu chí", "Có"],
              ["✗", "Không đạt", "Có"],
              ["—", "Không áp dụng cho ngành này, hoặc thiếu dữ liệu", "Không"],
          ],
          widths=[2.4, 8.6, 5.4], size=9.5)

para(doc,
     "Nếu số tiêu chí áp dụng được của một nhóm **dưới 3**, hệ thống trả về ⚪ THIẾU DỮ LIỆU thay vì xếp hạng. "
     "Nó không đoán từ vài con số lẻ.", size=9.5)

h2(doc, "4.2  Bảy tín hiệu")
add_table(doc,
          ["Tín hiệu", "Điều kiện", "Ý nghĩa"],
          [
              ["🟢 TỐT & THỰC", "A ≥ 80% và B ≥ 80%",
               "Số liệu mạnh và vượt được các phép kiểm chứng — nhóm nên nghiên cứu trước tiên"],
              ["🔵 KHÁ", "A ≥ 60% và B ≥ 60%", "Cả hai mặt đều khá, còn vài điểm cần xem"],
              ["🔴 CỜ ĐỎ", "A ≥ 80% nhưng B < 50%",
               "Số liệu rất đẹp nhưng không chứng minh được. Đây là kết quả đặc trưng của hệ thống"],
              ["🟠 THẬT NHƯNG YẾU", "B ≥ 80% nhưng A < 50%",
               "Sổ sách trung thực, dòng tiền thật, nhưng biên và tăng trưởng chưa hấp dẫn"],
              ["🟡 THEO DÕI", "A ≥ 50%", "Chưa đủ điều kiện, đưa vào danh sách theo dõi"],
              ["⚫ LOẠI", "còn lại", "Không đạt cả hai mặt"],
              ["⚪ THIẾU DỮ LIỆU", "< 3 tiêu chí áp dụng được", "Không đủ cơ sở để kết luận"],
          ],
          widths=[3.8, 3.6, 9.0], size=9,
          colors={(0, 0): GREEN, (1, 0): BLUE, (2, 0): RED, (3, 0): ORANGE,
                  (4, 0): AMBER, (5, 0): GREY, (6, 0): GREY})

para(doc,
     "Tín hiệu **🔴 CỜ ĐỎ** là lý do tồn tại của cấu trúc hai điểm. Trong một lần quét thử 30 cổ phiếu vốn hóa "
     "lớn nhất thị trường Mỹ, AMZN rơi vào nhóm này: đạt 5 trên 6 tiêu chí tài chính, nhưng tỷ lệ FCF trên "
     "lợi nhuận ròng là −0,15 — tức dòng tiền tự do âm do đầu tư rất nặng. Một bộ lọc chỉ nhìn nhóm A sẽ "
     "xếp mã này vào hàng đầu.", size=9.5)

# ═══════════════════════════════════════════════════════════════════════════
h1(doc, "5.  NHỮNG THỨ ĐÃ ĐƯỢC GỠ BỎ")

para(doc,
     "Phiên bản trước có 9 tiêu chí CAN SLIM cải biên cộng 6 tiêu chí Quality Compounder. Bảng dưới ghi lại "
     "những gì bị loại và vì sao — hầu hết không phải vì chỉ số kém, mà vì chúng không phục vụ hai câu hỏi "
     "mà hệ thống muốn trả lời, hoặc vì chúng đếm lặp một thứ đã được đếm rồi.")

add_table(doc,
          ["Đã gỡ", "Lý do"],
          [
              ["MKT — hướng thị trường",
               "Gán cùng một giá trị cho mọi dòng trong bảng, nên không phân biệt được mã nào với mã nào. "
               "Trong thị trường tăng, mọi mã cùng được cộng 1 điểm và số tín hiệu mạnh tăng vọt không phải "
               "vì bản thân các mã đó tốt lên. Nay hiển thị riêng trên Dashboard làm thông tin bối cảnh."],
              ["N — giá gần đỉnh 52 tuần",
               "Chỉ có dữ liệu khi bật yfinance và chỉ với thị trường Mỹ; mọi trường hợp còn lại luôn trả về "
               "«thiếu dữ liệu» nên không bao giờ cộng điểm."],
              ["L — sức mạnh tương đối",
               "Là hành vi giá, không phải chất lượng tài chính. Ngoài ra ngưỡng «vượt chỉ số» quá lỏng nên "
               "khoảng một nửa số mã đều đạt."],
              ["M — EPS tăng tốc",
               "Ý tưởng tốt nhưng cần ít nhất 5 quý liên tiếp có số liệu EPS, mà bảng báo cáo quý của yfinance "
               "thường chỉ có 4–5 kỳ. Trong nhiều trường hợp nó âm thầm rơi về công thức thay thế vốn dùng "
               "chính trường dữ liệu của một tiêu chí khác."],
              ["C và A — tăng trưởng EPS theo quý và theo năm",
               "Thay bằng tăng trưởng lợi nhuận ròng TTM. Ba tiêu chí C, A, M trước đây cùng đo một thứ — "
               "tăng trưởng lợi nhuận — và chiếm trọn một phần ba tổng điểm cho một đặc tính duy nhất."],
              ["Conviction",
               "Nhân điểm với hệ số chất lượng moat, mà moat lại suy ra từ ROE và biên gộp vốn đã được chấm "
               "điểm rồi. Cổ phiếu biên cao được thưởng ba lần cho cùng một đặc tính."],
              ["Moat Proxy",
               "Chỉ là bảng tra cứu tĩnh từ sector sang một chuỗi mô tả, không tham gia bất kỳ phép tính nào."],
              ["EQ Badge",
               "Đo đúng tỷ lệ FCF trên lợi nhuận ròng, nhưng lấy từ yfinance và chỉ để hiển thị. Nay chính "
               "tỷ lệ đó lấy từ TradingView và trở thành tiêu chí kiểm chứng CASH."],
              ["FCF/sh > 0 (bộ Quality Compounder cũ)",
               "Ngưỡng «lớn hơn 0» mà gần như mọi doanh nghiệp vốn hóa lớn đều vượt qua, nên hầu như luôn "
               "cho điểm và không phân loại được gì. Thay bằng FCF/NI và FCF Margin."],
          ],
          widths=[4.6, 12.0], size=9)

callout(doc, "Kết quả phụ đáng giá: bỏ được phần lớn phụ thuộc vào yfinance",
        "Vì nhóm B nay lấy dữ liệu từ TradingView, vòng lặp yfinance chỉ còn cần cho duy nhất tiêu chí "
        "Bền vững 5 năm. Quét 100 mã không bật yfinance mất vài giây thay vì 5–8 phút, và **chạy được cho "
        "thị trường Việt Nam và Hồng Kông** — trước đây toàn bộ nhóm chỉ báo nâng cao đều trống với các "
        "thị trường này.")

# ═══════════════════════════════════════════════════════════════════════════
h1(doc, "6.  ĐÁNH GIÁ TÍNH KHẢ THI")

h2(doc, "6.1  Bảng chấm từng tiêu chí")
add_table(doc,
          ["Tiêu chí", "Nhóm", "Cơ sở lý thuyết", "Độ tin cậy dữ liệu", "Sức phân biệt"],
          [
              ["ROIC", "A", "Cao — chuẩn mực so với chi phí vốn", "Cao", "Tốt"],
              ["Op Margin", "A", "Cao — ít bị thao túng hơn lợi nhuận ròng", "Cao", "Tốt"],
              ["Gross Mgn", "A", "Trung bình — lệch nặng theo ngành", "Cao", "Trung bình"],
              ["NI Growth", "A", "Cao", "Cao — TTM mượt hơn số quý", "Tốt"],
              ["Rev Growth", "A", "Cao — khó nguỵ tạo hơn lợi nhuận", "Cao", "Tốt"],
              ["D/E", "A", "Cao", "Trung bình — méo khi mua cổ phiếu quỹ nhiều", "Trung bình"],
              ["FCF/NI", "B", "Rất cao — nền tảng của accrual anomaly", "Cao", "Rất tốt"],
              ["FCF Mgn", "B", "Cao", "Cao", "Tốt"],
              ["Không buyback", "B", "Cao — phép thử trực tiếp và hiếm gặp", "Cao", "Tốt"],
              ["ROE thật", "B", "Cao — tách bạch đòn bẩy khỏi hiệu quả", "Cao", "Tốt"],
              ["Bảng cân đối", "B", "Trung bình — hai điều kiện hoặc", "Cao", "Trung bình"],
              ["Còn đúng", "B", "Cao — bù đúng điểm mù của số TTM", "Cao", "Tốt"],
              ["Bền vững 5N", "B", "Cao", "Thấp — yfinance, chỉ thị trường Mỹ", "Tốt khi có dữ liệu"],
          ],
          widths=[3.0, 1.4, 5.4, 4.2, 2.6], size=8.5)

h2(doc, "6.2  Những hạn chế còn lại")

h3(doc, "Ngưỡng biên lợi nhuận vẫn dùng chung cho mọi ngành", AMBER)
para(doc,
     "Biên gộp 40% là mức trung bình với phần mềm nhưng gần như không thể đạt với bán lẻ hay công nghiệp "
     "nặng. Hai tiêu chí Gross Mgn và Op Margin vì thế đo **cổ phiếu thuộc ngành nào** nhiều hơn là đo "
     "chất lượng doanh nghiệp, khiến hệ thống thiên lệch về phía công nghệ và y tế. Cách sửa là đặt ngưỡng "
     "theo trung vị ngành thay vì một con số chung.", size=9.5)

h3(doc, "Ngưỡng nhị phân làm mất thông tin", AMBER)
para(doc,
     "Một công ty có ROIC 14,9% và một công ty 15,1% chênh nhau trọn một điểm, trong khi khoảng cách thực "
     "gần như bằng không. Ngược lại, 16% và 60% được đối xử giống hệt nhau. Một thang điểm liên tục theo "
     "khoảng cách tới ngưỡng sẽ giữ lại lượng thông tin đang bị vứt bỏ.", size=9.5)

h3(doc, "Phân loại ngành của TradingView đôi khi lệch", AMBER)
para(doc,
     "Mastercard được xếp vào Financial Services, nên hai tiêu chí ROE thật và Bảng cân đối tự động chuyển "
     "sang «không áp dụng» — dù đây là một mạng lưới thanh toán chứ không phải ngân hàng. Kết quả là mẫu số "
     "nhóm B của mã này nhỏ hơn và một chênh lệch ROE–ROIC rất lớn bị bỏ qua. Nếu cần chặt chẽ hơn, "
     "nên lọc theo ngành con thay vì theo sector.", size=9.5)

h3(doc, "Không có định giá và không có quy tắc bán", RED)
para(doc,
     "Không tiêu chí nào liên quan tới giá phải trả. P/E, PEG và EV/EBITDA vẫn được tính và hiển thị nhưng "
     "không chấm điểm, nên một cổ phiếu xuất sắc đang giao dịch ở P/E 90 vẫn có thể nhận 🟢 TỐT & THỰC. "
     "Đây là lựa chọn có chủ đích — hệ thống trả lời «doanh nghiệp này có tốt và có thật không», "
     "không phải «giá này có rẻ không» — nhưng người dùng cần ý thức rõ khoảng trống đó. Tương tự, "
     "chương trình chỉ trả lời «xem mã nào», hoàn toàn không có quy tắc thoát lệnh hay cắt lỗ.", size=9.5)

h3(doc, "Các ngưỡng chưa được kiểm định trên dữ liệu lịch sử", RED)
para(doc,
     "Toàn bộ ngưỡng — 15%, 20%, 40%, 0,8, 10 điểm %, 15 điểm % — đều mượn từ chuẩn mực phổ biến hoặc chọn "
     "theo lập luận, **chưa từng được kiểm định trên dữ liệu lịch sử của chính hệ thống này**. Chương trình "
     "không lưu dữ liệu theo thời điểm nên chưa thể kiểm thử lùi, và hiện chưa có bằng chứng nào cho thấy "
     "mã 6/6 thực sự có kết quả tốt hơn mã 4/6.", size=9.5)

callout(doc, "Đã có sẵn công cụ để bắt đầu hiệu chỉnh ngưỡng",
        "Dashboard và sheet Tổng hợp trong file Excel đều có bảng **tỷ lệ đạt của từng tiêu chí**, tự động "
        "gắn nhãn «quá lỏng» khi trên 90% số mã đạt và «quá chặt» khi dưới 5% đạt. Đây chưa phải kiểm thử lùi, "
        "nhưng nó cho biết ngay tiêu chí nào đang không phân loại được gì — điều kiện cần trước khi chỉnh "
        "bất kỳ con số nào.")

# ═══════════════════════════════════════════════════════════════════════════
h1(doc, "7.  KHUYẾN NGHỊ TIẾP THEO")

add_table(doc,
          ["#", "Việc cần làm", "Lý do", "Công sức"],
          [
              ["1", "Đặt ngưỡng biên lợi nhuận theo trung vị ngành",
               "Ngưỡng chung đang đo ngành nghề chứ không đo chất lượng doanh nghiệp", "Trung bình"],
              ["2", "Lưu kết quả mỗi lần quét kèm ngày tháng",
               "Không có dữ liệu theo thời điểm thì không thể chứng minh bất kỳ ngưỡng nào là đúng", "Thấp"],
              ["3", "Cho điểm liên tục theo khoảng cách tới ngưỡng thay vì đúng/sai",
               "Giữ lại thông tin đang bị vứt bỏ ở ranh giới", "Trung bình"],
              ["4", "Lọc nhóm Tài chính theo ngành con thay vì sector",
               "Tránh việc mạng lưới thanh toán bị miễn trừ như ngân hàng", "Thấp"],
              ["5", "Thêm cổng định giá tùy chọn, ví dụ gắn cờ khi PEG > 3",
               "Hiện không có ràng buộc nào về giá phải trả", "Thấp"],
              ["6", "Thêm phép thử pha loãng: cảnh báo khi EPS tăng chậm hơn lợi nhuận nhiều",
               "Hiện chỉ bắt buyback thổi phồng, chưa bắt phát hành thêm làm loãng", "Rất thấp"],
          ],
          widths=[0.9, 6.0, 7.2, 2.3], size=8.5)

# ═══════════════════════════════════════════════════════════════════════════
h1(doc, "8.  KẾT LUẬN")

para(doc,
     "Bộ tiêu chí sau khi làm lại có một điều mà bản trước không có: **một câu hỏi rõ ràng và một cấu trúc "
     "phục vụ đúng câu hỏi đó**. Mười ba tiêu chí ít hơn mười lăm tiêu chí cũ, nhưng chúng đo mười ba thứ "
     "khác nhau thay vì đo bốn thứ lặp lại ba lần. Việc tách hai điểm và giữ chúng không cộng gộp là quyết "
     "định quan trọng nhất, vì nó biến một trường hợp trước đây vô hình — số liệu đẹp mà mỏng thực chất — "
     "thành một tín hiệu nổi bật.")

para(doc,
     "Việc chuyển nhóm kiểm chứng sang dữ liệu TradingView mang lại một lợi ích không nằm trong dự tính ban "
     "đầu: hệ thống nay chạy nhanh hơn nhiều và hoạt động được ở những thị trường mà trước đây phần lớn "
     "chỉ báo đều trống. Phụ thuộc vào nguồn dữ liệu không chính thức giảm từ bốn chỉ báo xuống còn một.")

para(doc,
     "Những gì còn thiếu đã được nêu rõ ở mục 6 và không nên bị coi nhẹ: ngưỡng chưa kiểm định, biên lợi "
     "nhuận lệch theo ngành, không có định giá, không có quy tắc bán. Đây là **công cụ thu hẹp danh sách** "
     "để đọc kỹ, không phải hệ thống ra quyết định mua bán. Nhưng trong phạm vi đó, nó làm được một việc "
     "mà phần lớn bộ lọc không làm: nó không chỉ tìm số liệu đẹp, nó còn hỏi xem số liệu ấy có đáng tin không.")

callout(doc, "Một câu tóm tắt",
        "Nhóm A tìm doanh nghiệp trông tốt. Nhóm B kiểm tra xem cái vẻ đó có chịu được soi xét không. "
        "Giá trị của hệ thống nằm ở chỗ nó báo động khi hai câu trả lời không khớp nhau.")

# ── Phụ lục ────────────────────────────────────────────────────────────────
doc.add_page_break()
h1(doc, "PHỤ LỤC — TRA CỨU TRƯỜNG DỮ LIỆU")
add_table(doc,
          ["Tên trong app", "Trường gốc", "Nguồn"],
          [
              ["ROIC%", "return_on_invested_capital", "TradingView"],
              ["Op Margin%", "operating_margin", "TradingView"],
              ["Gross Margin%", "gross_margin", "TradingView"],
              ["NI Growth%", "net_income_yoy_growth_ttm", "TradingView"],
              ["Rev Growth%", "total_revenue_yoy_growth_ttm", "TradingView"],
              ["EPS Growth%", "earnings_per_share_diluted_yoy_growth_ttm", "TradingView"],
              ["D/E", "debt_to_equity", "TradingView"],
              ["ROE%", "return_on_equity", "TradingView"],
              ["FCF_abs · FCF ($B)", "free_cash_flow_ttm", "TradingView"],
              ["FCF_Margin%", "free_cash_flow_margin_ttm", "TradingView"],
              ["NI_abs · Lợi nhuận ($B)", "net_income", "TradingView"],
              ["Rev_abs · Doanh thu ($B)", "total_revenue_ttm", "TradingView"],
              ["Rev_q_abs · DThu quý ($B)", "total_revenue_fq", "TradingView"],
              ["NI_q_abs · LN quý ($B)", "net_income_fq", "TradingView"],
              ["NI Qtr%", "net_income_yoy_growth_fq", "TradingView"],
              ["Rev Qtr%", "total_revenue_yoy_growth_fq", "TradingView"],
              ["NI_Trend_Gap · Đà LN", "NI Qtr% − NI Growth%", "Tính trong code"],
              ["Cash_abs", "cash_n_short_term_invest_fq", "TradingView"],
              ["Debt_abs", "total_debt_fq", "TradingView"],
              ["Current Ratio", "current_ratio", "TradingView"],
              ["FCF/NI", "FCF_abs / NI_abs", "Tính trong code"],
              ["EPS_NI_Gap", "EPS Growth% − NI Growth%", "Tính trong code"],
              ["ROE_ROIC_Gap", "ROE% − ROIC%", "Tính trong code"],
              ["Net Cash ($B)", "(Cash_abs − Debt_abs) / 1e9", "Tính trong code"],
              ["Moat Score", "ROE 5 năm + biên gộp 5 năm", "yfinance"],
              ["FIN_Score / FIN_App", "số đạt / số áp dụng được, nhóm A", "Tính trong code"],
              ["REAL_Score / REAL_App", "số đạt / số áp dụng được, nhóm B", "Tính trong code"],
              ["Signal", "hàm của FIN_Pct và REAL_Pct", "Tính trong code"],
          ],
          widths=[4.4, 8.2, 4.0], size=9)

p = doc.add_paragraph()
p.paragraph_format.space_before = Pt(14)
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = p.add_run("— Hết —")
r.italic = True; r.font.size = Pt(9); r.font.color.rgb = GREY; r.font.name = FONT

out = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                   "Tieu_chi_lua_chon_co_phieu.docx")
doc.save(out)
print("Da tao:", out)
