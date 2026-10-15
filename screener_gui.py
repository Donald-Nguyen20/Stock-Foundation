"""
Fundamental Screener — PySide6 GUI
Wraps python screen_top100.py
"""
import sys, os, time, importlib.util
import pandas as pd

from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLineEdit, QPushButton, QLabel, QFrame, QSplitter,
    QTableWidget, QTableWidgetItem, QHeaderView, QAbstractItemView,
    QComboBox, QSpinBox, QCheckBox, QFileDialog,
    QDialog, QTextBrowser, QTabWidget, QProgressBar,
)
from PySide6.QtCore import Qt, QThread, Signal, QTimer, QSortFilterProxyModel
from PySide6.QtGui import QFont, QColor, QKeySequence, QShortcut, QPixmap, QPainter
from PySide6.QtWidgets import QHeaderView
#pyinstaller StockScreener.spec --clean#
# ─────────────────────────────────────────────────────────────────────────────
# Load python screen_top100.py via importlib (filename has a space)
# ─────────────────────────────────────────────────────────────────────────────
_dir  = os.path.dirname(os.path.abspath(__file__))
_spec = importlib.util.spec_from_file_location(
    "screen_top100", os.path.join(_dir, "python screen_top100.py"))
_m = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(_m)

fetch_data             = _m.fetch_data
clean_df               = _m.clean_df
score_stocks           = _m.score_stocks
fetch_moat_yfinance    = _m.fetch_moat_yfinance
fetch_market_direction = _m.fetch_market_direction
_MKT_INDEX          = _m._MKT_INDEX
write_excel            = _m.write_excel
FIN_CRITERIA        = _m.FIN_CRITERIA
FIN_KEYS            = _m.FIN_KEYS
N_FIN               = _m.N_FIN
REAL_CRITERIA       = _m.REAL_CRITERIA
REAL_KEYS           = _m.REAL_KEYS
N_REAL              = _m.N_REAL
MOAT_SCORE_STYLE    = _m.MOAT_SCORE_STYLE
FETCH_COLS          = _m.FETCH_COLS
RENAME              = _m.RENAME

# ─────────────────────────────────────────────────────────────────────────────
# Design tokens
# ─────────────────────────────────────────────────────────────────────────────
BG       = "#F0F4F9"
SURFACE  = "#FFFFFF"
PANEL    = "#E8EDF5"
INPUT_BG = "#FFFFFF"
BORDER   = "#D1D9E6"
BORDER2  = "#A8B8CC"
BLUE     = "#2563EB"
BLUE_HV  = "#1D4ED8"
TEXT1    = "#1A202C"
TEXT2    = "#4A5568"
TEXT3    = "#94A3B8"
GREEN    = "#16A34A"
RED      = "#DC2626"
AMBER    = "#D97706"
SPINNER  = ["⠋","⠙","⠹","⠸","⠼","⠴","⠦","⠧","⠇","⠏"]

DARK_THEME = dict(
    BG="#06080F", SURFACE="#0B0E18", PANEL="#0E1220",
    INPUT_BG="#131928", BORDER="#1A2133", BORDER2="#243044",
    BLUE="#3D8EF0", BLUE_HV="#2463B4",
    TEXT1="#DCE4EE", TEXT2="#7A8899", TEXT3="#3D4D60",
    GREEN="#34C472", RED="#E8483D", AMBER="#E8A93D",
)
LIGHT_THEME = dict(
    BG="#F0F4F9", SURFACE="#FFFFFF", PANEL="#E8EDF5",
    INPUT_BG="#FFFFFF", BORDER="#D1D9E6", BORDER2="#A8B8CC",
    BLUE="#2563EB", BLUE_HV="#1D4ED8",
    TEXT1="#1A202C", TEXT2="#4A5568", TEXT3="#94A3B8",
    GREEN="#276749", RED="#C53030", AMBER="#975A16",
)

# Signal badge colors  (text, background)
SIGNAL_COLORS = {
    "🟢 TỐT & THỰC":     ("#1A5C2B", "#C6EFCE"),
    "🔵 KHÁ":            ("#1B3A5C", "#DDEEFF"),
    "🔴 CỜ ĐỎ":          ("#9C0006", "#FFC7CE"),
    "🟠 THẬT NHƯNG YẾU": ("#7A3E00", "#FFE0C2"),
    "🟡 THEO DÕI":       ("#7D6608", "#FFF2CC"),
    "⚫ LOẠI":            ("#3D4D60", "#E2E8F0"),
    "⚪ THIẾU DỮ LIỆU":   ("#64748B", "#F1F5F9"),
}

# ─────────────────────────────────────────────────────────────────────────────
# Tab 1 — TÀI CHÍNH  (nhóm A: số liệu tốt)
# ─────────────────────────────────────────────────────────────────────────────
TABLE_COLS = [
    ("No",         "_no_",          36,  Qt.AlignCenter | Qt.AlignVCenter),
    ("Ticker",     "Ticker",        65,  Qt.AlignLeft   | Qt.AlignVCenter),
    ("Company",    "Tên Công Ty",  175,  Qt.AlignLeft   | Qt.AlignVCenter),
    ("Sector",     "Sector",       105,  Qt.AlignLeft   | Qt.AlignVCenter),
    ("Price($)",   "Price ($)",     70,  Qt.AlignRight  | Qt.AlignVCenter),
    ("MCap($B)",   "MCap ($B)",     72,  Qt.AlignRight  | Qt.AlignVCenter),
    ("DThu($B)",   "Doanh thu ($B)", 74, Qt.AlignRight  | Qt.AlignVCenter),
    ("LNhuận($B)", "Lợi nhuận ($B)", 78, Qt.AlignRight  | Qt.AlignVCenter),
    ("DThu Q($B)", "DThu quý ($B)",  76, Qt.AlignRight  | Qt.AlignVCenter),
    ("LN Q($B)",   "LN quý ($B)",    70, Qt.AlignRight  | Qt.AlignVCenter),
    ("ROIC%",      "ROIC%",         62,  Qt.AlignRight  | Qt.AlignVCenter),
    ("OpMgn%",     "Op Margin%",    64,  Qt.AlignRight  | Qt.AlignVCenter),
    ("GM%",        "Gross Margin%", 60,  Qt.AlignRight  | Qt.AlignVCenter),
    ("NI Grw%",    "NI Growth%",    64,  Qt.AlignRight  | Qt.AlignVCenter),
    ("NI Q%",      "NI Qtr%",       60,  Qt.AlignRight  | Qt.AlignVCenter),
    ("Rev Grw%",   "Rev Growth%",   66,  Qt.AlignRight  | Qt.AlignVCenter),
    ("Rev Q%",     "Rev Qtr%",      62,  Qt.AlignRight  | Qt.AlignVCenter),
    ("D/E",        "D/E",           52,  Qt.AlignRight  | Qt.AlignVCenter),
    ("ROE%",       "ROE%",          58,  Qt.AlignRight  | Qt.AlignVCenter),
    ("P/E",        "P/E",           52,  Qt.AlignRight  | Qt.AlignVCenter),
    ("1Y%",        "1Y%",           55,  Qt.AlignRight  | Qt.AlignVCenter),
    # 6 tiêu chí nhóm A
    ("ROIC",       "FIN_ROIC",      40,  Qt.AlignCenter | Qt.AlignVCenter),
    ("OpM",        "FIN_OPM",       40,  Qt.AlignCenter | Qt.AlignVCenter),
    ("GM",         "FIN_GM",        38,  Qt.AlignCenter | Qt.AlignVCenter),
    ("NI↑",        "FIN_NIG",       40,  Qt.AlignCenter | Qt.AlignVCenter),
    ("Rev↑",       "FIN_REVG",      42,  Qt.AlignCenter | Qt.AlignVCenter),
    ("D/E",        "FIN_DE",        40,  Qt.AlignCenter | Qt.AlignVCenter),
    ("Tài chính",  "FIN_Disp",      64,  Qt.AlignCenter | Qt.AlignVCenter),
    ("Thực",       "REAL_Disp",     58,  Qt.AlignCenter | Qt.AlignVCenter),
    ("Tín hiệu",   "Signal",       150,  Qt.AlignCenter | Qt.AlignVCenter),
    ("Cảnh báo",   "Cảnh báo",     140,  Qt.AlignCenter | Qt.AlignVCenter),
]

# ─────────────────────────────────────────────────────────────────────────────
# Tab 2 — KIỂM CHỨNG  (nhóm B: số liệu đó là thực hay ảo)
# Mỗi tiêu chí bác bỏ một cách cụ thể mà nhóm A có thể bị làm đẹp.
# ─────────────────────────────────────────────────────────────────────────────
REAL_COLS = [
    ("No",         "_no_",          36,  Qt.AlignCenter | Qt.AlignVCenter),
    ("Ticker",     "Ticker",        65,  Qt.AlignLeft   | Qt.AlignVCenter),
    ("Company",    "Tên Công Ty",  175,  Qt.AlignLeft   | Qt.AlignVCenter),
    ("Sector",     "Sector",       105,  Qt.AlignLeft   | Qt.AlignVCenter),
    ("LNhuận($B)", "Lợi nhuận ($B)", 78, Qt.AlignRight  | Qt.AlignVCenter),
    ("FCF($B)",    "FCF ($B)",      70,  Qt.AlignRight  | Qt.AlignVCenter),
    ("FCF/NI",     "FCF/NI",        62,  Qt.AlignRight  | Qt.AlignVCenter),
    ("FCF Mgn%",   "FCF_Margin%",   70,  Qt.AlignRight  | Qt.AlignVCenter),
    ("EPS−NI",     "EPS_NI_Gap",    64,  Qt.AlignRight  | Qt.AlignVCenter),
    ("ROE−ROIC",   "ROE_ROIC_Gap",  74,  Qt.AlignRight  | Qt.AlignVCenter),
    ("NetCash$B",  "Net Cash ($B)", 76,  Qt.AlignRight  | Qt.AlignVCenter),
    ("Curr.R",     "Current Ratio", 60,  Qt.AlignRight  | Qt.AlignVCenter),
    ("Nợ/FCF",     "Nợ ròng/FCF",   62,  Qt.AlignRight  | Qt.AlignVCenter),
    ("Đà LN",      "NI_Trend_Gap",  64,  Qt.AlignRight  | Qt.AlignVCenter),
    ("Bền vững",   "Moat Score",   110,  Qt.AlignCenter | Qt.AlignVCenter),
    # 6 tiêu chí nhóm B
    ("Tiền",       "REAL_CASH",     42,  Qt.AlignCenter | Qt.AlignVCenter),
    ("Biên",       "REAL_FCFM",     42,  Qt.AlignCenter | Qt.AlignVCenter),
    ("K.buyback",  "REAL_BUYB",     62,  Qt.AlignCenter | Qt.AlignVCenter),
    ("ROE thật",   "REAL_LEV",      56,  Qt.AlignCenter | Qt.AlignVCenter),
    ("Cân đối",    "REAL_BAL",      52,  Qt.AlignCenter | Qt.AlignVCenter),
    ("Còn đúng",   "REAL_TREND",    58,  Qt.AlignCenter | Qt.AlignVCenter),
    ("5 năm",      "REAL_SUS",      46,  Qt.AlignCenter | Qt.AlignVCenter),
    ("Tài chính",  "FIN_Disp",      64,  Qt.AlignCenter | Qt.AlignVCenter),
    ("Thực",       "REAL_Disp",     58,  Qt.AlignCenter | Qt.AlignVCenter),
    ("Tín hiệu",   "Signal",       150,  Qt.AlignCenter | Qt.AlignVCenter),
    ("Cảnh báo",   "Cảnh báo",     140,  Qt.AlignCenter | Qt.AlignVCenter),
]


# ─────────────────────────────────────────────────────────────────────────────
# Chú giải cho từng cột — hiện khi rê chuột lên tiêu đề cột hoặc lên ô dữ liệu.
# Mỗi mục nói ba điều: chỉ số là gì, tính thế nào, và đọc nó ra sao.
# ─────────────────────────────────────────────────────────────────────────────
COL_TIPS = {
    "Ticker":  "Mã cổ phiếu trên sàn.",
    "Tên Công Ty": "Tên đầy đủ của doanh nghiệp.",
    "Sector":  "Ngành theo phân loại của TradingView. Ngành quyết định một số tiêu chí có được "
               "áp dụng hay không — nhóm Financial Services được xử lý riêng.",
    "Price ($)": "Giá đóng cửa gần nhất.",
    "MCap ($B)": "Vốn hóa thị trường = giá × số cổ phiếu lưu hành. Đơn vị tỷ.",

    "Doanh thu ($B)":
        "<b>Doanh thu</b> bốn quý gần nhất (TTM), đơn vị tỷ.<br>"
        "Tổng tiền bán hàng trước khi trừ bất kỳ chi phí nào. Đây là con số khó nguỵ tạo nhất "
        "trên báo cáo tài chính.<br>"
        "<i>Chỉ để tham khảo quy mô — không chấm điểm.</i>",
    "Lợi nhuận ($B)":
        "<b>Lợi nhuận ròng</b> bốn quý gần nhất (TTM), đơn vị tỷ.<br>"
        "Phần còn lại sau khi trừ hết chi phí, lãi vay và thuế. Đây là con số kế toán — "
        "nó có thể tăng mà tiền mặt không tăng.<br>"
        "<i>So cột này với cột FCF để thấy ngay lợi nhuận có kèm tiền thật hay không.</i>",
    "FCF ($B)":
        "<b>Dòng tiền tự do</b> bốn quý gần nhất (TTM), đơn vị tỷ.<br>"
        "Tiền thật còn lại sau khi đã chi cho vận hành và đầu tư tài sản. Đây là tiền doanh nghiệp "
        "thực sự có thể dùng để trả nợ, chia cổ tức hoặc mua lại cổ phiếu.<br>"
        "<i>FCF thấp hơn lợi nhuận kéo dài là dấu hiệu lợi nhuận nằm trên giấy.</i>",

    "DThu quý ($B)":
        "<b>Doanh thu quý gần nhất</b>, đơn vị tỷ.<br>"
        "Chỉ riêng quý vừa công bố, không cộng dồn. So với cột Doanh thu (TTM) để thấy quý này "
        "đóng góp bao nhiêu.<br>"
        "<i>Chỉ để tham khảo — không chấm điểm.</i>",
    "LN quý ($B)":
        "<b>Lợi nhuận ròng quý gần nhất</b>, đơn vị tỷ.<br>"
        "Chỉ riêng quý vừa công bố.<br>"
        "<i>Chỉ để tham khảo — không chấm điểm.</i>",
    "NI Qtr%":
        "<b>Tăng trưởng lợi nhuận quý gần nhất</b> so với cùng quý năm trước.<br>"
        "Đặt cạnh cột NI Grw% (mức TTM) là thấy ngay doanh nghiệp đang tăng tốc hay giảm tốc. "
        "Số TTM cộng dồn bốn quý nên có độ trễ — một doanh nghiệp mới xấu đi vẫn giữ số TTM đẹp "
        "thêm hai ba quý nữa vì các quý cũ còn đè lên.<br>"
        "<i>Là dữ liệu cho tiêu chí «Còn đúng» ở tab Kiểm chứng.</i>",
    "Rev Qtr%":
        "<b>Tăng trưởng doanh thu quý gần nhất</b> so với cùng quý năm trước.<br>"
        "Đặt cạnh cột Rev Grw% (mức TTM) để so đà tăng trưởng hiện tại với mức trung bình cả năm.<br>"
        "<i>Chỉ để tham khảo — không chấm điểm.</i>",
    "NI_Trend_Gap":
        "<b>Đà tăng trưởng lợi nhuận</b> = NI Qtr% − NI Growth%   (điểm phần trăm).<br>"
        "Dương nghĩa là quý gần nhất tăng nhanh hơn mức trung bình cả năm — đang tăng tốc. "
        "Âm nhiều nghĩa là đang giảm tốc, và số TTM đẹp chủ yếu nhờ các quý cũ.<br>"
        "<i>Là dữ liệu cho tiêu chí «Còn đúng».</i>",
    "ROIC%":
        "<b>Return on Invested Capital</b> — lợi nhuận trên toàn bộ vốn đầu tư, gồm cả vốn chủ và nợ.<br>"
        "Đây là thước đo chất lượng doanh nghiệp tốt nhất trong bảng: nó cho biết mỗi đồng vốn bỏ vào "
        "sinh ra bao nhiêu lợi nhuận, bất kể vốn đó vay hay tự có.<br>"
        "<i>Trên 15% là vượt chi phí vốn (thường 8–10%). Dưới đó, doanh nghiệp đang phá hủy giá trị.</i>",
    "Op Margin%":
        "<b>Biên lợi nhuận hoạt động</b> = lợi nhuận từ hoạt động kinh doanh / doanh thu.<br>"
        "Đo hiệu quả vận hành sau khi trừ giá vốn và chi phí bán hàng, quản lý — nhưng trước lãi vay "
        "và thuế, nên ít bị méo bởi cấu trúc vốn.<br>"
        "<i>Ngưỡng đạt: trên 15%.</i>",
    "Gross Margin%":
        "<b>Biên lợi nhuận gộp</b> = (doanh thu − giá vốn hàng bán) / doanh thu.<br>"
        "Phản ánh khả năng bán cao hơn chi phí sản xuất — tức sức mạnh định giá.<br>"
        "<i>Ngưỡng đạt: trên 40%. Lưu ý biên gộp lệch rất mạnh theo ngành: phần mềm thường 60–80%, "
        "bán lẻ hiếm khi vượt 30%.</i>",
    "NI Growth%":
        "<b>Tăng trưởng lợi nhuận ròng</b> — bốn quý gần nhất so với bốn quý cùng kỳ năm trước.<br>"
        "Dùng số TTM thay vì số một quý để giảm nhiễu theo mùa vụ.<br>"
        "<i>Ngưỡng đạt: trên 20%. Hệ thống đo tăng trưởng lợi nhuận chứ không đo EPS, vì EPS còn "
        "chịu ảnh hưởng của số cổ phiếu lưu hành.</i>",
    "Rev Growth%":
        "<b>Tăng trưởng doanh thu</b> — bốn quý gần nhất so cùng kỳ năm trước.<br>"
        "Lợi nhuận có thể tăng nhờ cắt chi phí, nhưng doanh thu chỉ tăng khi bán được nhiều hơn thật sự.<br>"
        "<i>Ngưỡng đạt: trên 15%.</i>",
    "D/E":
        "<b>Debt to Equity</b> = tổng nợ / vốn chủ sở hữu.<br>"
        "Đo mức độ phụ thuộc vào vay nợ. Lãi vay phải trả kể cả khi doanh thu giảm, nên đòn bẩy "
        "khuếch đại cả lãi lẫn lỗ.<br>"
        "<i>Ngưỡng đạt: dưới 1,0. Không áp dụng cho ngân hàng và bảo hiểm — tiền gửi khách hàng "
        "được hạch toán là nợ nên tỷ lệ 8–15 lần là bình thường.</i>",
    "ROE%":
        "<b>Return on Equity</b> = lợi nhuận ròng / vốn chủ sở hữu.<br>"
        "Lợi nhuận trên phần vốn của cổ đông.<br>"
        "<i>Không phải tiêu chí chấm điểm, vì ROE dễ bị thổi phồng bằng đòn bẩy hoặc bằng việc mua "
        "cổ phiếu quỹ làm vốn chủ teo lại. Hệ thống dùng ROIC và so hai con số với nhau ở cột ROE−ROIC.</i>",
    "P/E":
        "<b>Price to Earnings</b> = giá / lợi nhuận trên mỗi cổ phiếu.<br>"
        "<i>Chỉ hiển thị tham khảo. Hệ thống không chấm điểm định giá — nó trả lời «doanh nghiệp này "
        "có tốt và có thật không», không phải «giá này có rẻ không».</i>",
    "1Y%": "Hiệu suất giá 12 tháng gần nhất. Chỉ hiển thị, không chấm điểm.",
    "Net Margin%": "<b>Biên lợi nhuận ròng</b> = lợi nhuận ròng / doanh thu. Chỉ hiển thị.",
    "EV/EBITDA": "Giá trị doanh nghiệp trên lợi nhuận trước lãi, thuế và khấu hao. Chỉ hiển thị.",
    "PEG": "P/E chia cho tốc độ tăng trưởng lợi nhuận. Chỉ hiển thị.",

    "FCF/NI":
        "<b>Dòng tiền tự do chia lợi nhuận ròng</b> — phép thử quan trọng nhất của nhóm Kiểm chứng.<br>"
        "Lợi nhuận là con số kế toán; dòng tiền tự do là tiền thật. Khi hai con số lệch nhau kéo dài, "
        "phần chênh nằm ở khoản phải thu, hàng tồn kho hoặc các bút toán dồn tích.<br>"
        "<i>Ngưỡng đạt: từ 0,80 lần trở lên. Để trống khi doanh nghiệp đang lỗ, vì khi đó tỷ lệ "
        "không có ý nghĩa.</i>",
    "FCF_Margin%":
        "<b>Biên dòng tiền tự do</b> = dòng tiền tự do / doanh thu.<br>"
        "Mỗi đồng doanh thu để lại bao nhiêu tiền thật. Bác bỏ trường hợp biên lợi nhuận đẹp nhờ "
        "vốn hóa chi phí hoặc hoãn đầu tư.<br>"
        "<i>Ngưỡng đạt: trên 8%.</i>",
    "EPS_NI_Gap":
        "<b>Chênh lệch tăng trưởng EPS trừ tăng trưởng lợi nhuận ròng</b>, tính bằng điểm phần trăm.<br>"
        "Nếu lợi nhuận tăng 10% mà EPS tăng 25%, phần chênh 15 điểm không đến từ kinh doanh — "
        "nó đến từ việc mua lại cổ phiếu làm giảm số cổ phiếu lưu hành.<br>"
        "<i>Ngưỡng đạt: không vượt quá 10 điểm %. Số âm nghĩa là đang pha loãng (phát hành thêm), "
        "vẫn được coi là đạt tiêu chí này.</i>",
    "ROE_ROIC_Gap":
        "<b>Chênh lệch ROE trừ ROIC</b>, tính bằng điểm phần trăm.<br>"
        "ROE tính trên vốn chủ, ROIC tính trên toàn bộ vốn gồm cả nợ — nên khoảng cách giữa chúng "
        "chính là phần đóng góp của đòn bẩy.<br>"
        "<i>Ngưỡng đạt: không vượt quá 15 điểm %. ROE 25% với ROIC 22% là hiệu quả thật; "
        "ROE 45% với ROIC 8% là đòn bẩy.</i>",
    "Net Cash ($B)":
        "<b>Tiền ròng</b> = tiền và tương đương + đầu tư ngắn hạn − tổng nợ, đơn vị tỷ.<br>"
        "Dương nghĩa là doanh nghiệp có nhiều tiền hơn nợ.<br>"
        "<i>Là vế thứ nhất của tiêu chí Bảng cân đối. Số âm không tự động là xấu — xem thêm cột "
        "Nợ/FCF để biết khoản nợ đó có nặng so với dòng tiền không. Không áp dụng cho ngân hàng.</i>",
    "Current Ratio":
        "<b>Hệ số thanh toán hiện hành</b> = tài sản ngắn hạn / nợ ngắn hạn.<br>"
        "Đo khả năng trả các khoản đến hạn trong vòng một năm.<br>"
        "<i>Trên 1,5 được coi là an toàn — là vế thứ hai của tiêu chí Bảng cân đối. Lưu ý đây vốn là "
        "thước đo cho doanh nghiệp sản xuất có tồn kho và khoản phải thu; với công ty dịch vụ tạo nhiều "
        "tiền mặt thì cột Nợ/FCF phản ánh đúng bản chất hơn.</i>",
    "Nợ ròng/FCF":
        "<b>Số năm dòng tiền tự do cần để trả hết nợ ròng</b> = Nợ ròng / FCF một năm.<br>"
        "Ví dụ nợ ròng 52 tỷ với FCF 67 tỷ mỗi năm thì chỉ số là 0,8 — trả hết trong chưa tới một năm.<br>"
        "<i>Để trống khi doanh nghiệp đã có tiền ròng dương (không còn nợ để trả), hoặc khi FCF âm "
        "(không trả được bằng dòng tiền). Là vế thứ ba của tiêu chí Bảng cân đối: dưới 3 năm là đạt.</i>",
    "Moat Score":
        "<b>Độ bền vững qua 5 năm</b> — xếp hạng từ ROE và biên lợi nhuận gộp trung bình 5 năm.<br>"
        "WIDE ★★★ khi ROE ≥ 20% và biên gộp ≥ 50%; NARROW ★★ khi ROE ≥ 15% và biên gộp ≥ 35%.<br>"
        "<i>Đạt tiêu chí khi là WIDE hoặc NARROW. Cần bật «yfinance Moat» và chỉ có dữ liệu cho "
        "thị trường Mỹ — khi thiếu thì để trống chứ không thay bằng số liệu một kỳ, vì lấy một kỳ "
        "mà gọi là bền vững thì chính nó đã là số ảo.</i>",

    "FIN_Disp":
        "<b>Điểm Tài chính</b> — số tiêu chí nhóm A đạt / số tiêu chí thực sự áp dụng được.<br>"
        "Mẫu số thay đổi theo ngành và theo dữ liệu sẵn có, nên mã thiếu dữ liệu và nhóm Tài chính "
        "không bị phạt oan. Sắp xếp theo tỷ lệ, nên 4/4 đứng trên 5/6.",
    "REAL_Disp":
        "<b>Điểm Kiểm chứng</b> — số tiêu chí nhóm B đạt / số tiêu chí thực sự áp dụng được.<br>"
        "Điểm này <b>không được cộng với điểm Tài chính</b>. Điểm Tài chính cao mà điểm này thấp "
        "là tín hiệu cảnh báo, không phải điểm trung bình.",
    "Cảnh báo":
        "<b>Ba chỉ báo cảnh báo về giá</b> — thuần hiển thị, <b>không chấm điểm và không dùng để lọc</b>. "
        "Chúng không nói mã tốt hay xấu, chỉ nói giá đang ở trạng thái nào.<br><br>"
        "<b>▼50</b> — giá thấp hơn đỉnh 52 tuần từ 50% trở lên. Mất hơn một nửa giá trị thường là "
        "đặt lại kỳ vọng về doanh nghiệp, không phải một nhịp điều chỉnh.<br>"
        "<b>&lt;MA250</b> — giá nằm dưới trung bình 250 phiên, tức dưới mức trung bình một năm giao dịch. "
        "Ranh giới thường dùng để phân biệt nhịp điều chỉnh trong xu hướng tăng với đổi xu hướng.<br>"
        "<b>−10%</b> — giá sụt từ 10% trở lên ngay trong phiên. Thường là phản ứng với tin bất ngờ: "
        "hạ dự báo, kết quả hụt kỳ vọng, biến cố pháp lý.<br><br>"
        "<i>Với người mua khi giá giảm: càng nhiều cờ thì càng nghiêng về «đang rơi» hơn là «đang chiết khấu». "
        "Nhưng app không kết luận thay — hãy đối chiếu với tiêu chí «Còn đúng» ở tab Kiểm chứng để biết "
        "số liệu quý gần nhất có xấu đi thật không.</i>",
    "% từ đỉnh":
        "<b>Khoảng cách tới đỉnh 52 tuần</b> = (Giá hiện tại / Đỉnh 52 tuần − 1) × 100.<br>"
        "Số âm càng lớn thì giá càng xa đỉnh.<br>"
        "<i>Chỉ hiển thị — không chấm điểm.</i>",
    "Giá/MA250%":
        "<b>Khoảng cách tới trung bình 250 phiên</b> = (Giá hiện tại / MA250 − 1) × 100.<br>"
        "Dương nghĩa là giá đang trên mức trung bình một năm; âm là đang dưới.<br>"
        "<i>Chỉ hiển thị — không chấm điểm.</i>",
    "Chg%":
        "<b>Thay đổi giá trong phiên hiện tại</b>, tính bằng phần trăm.<br>"
        "<i>Chỉ hiển thị — không chấm điểm.</i>",
    "Signal":
        "<b>Tín hiệu tổng hợp</b> từ hai tỷ lệ đạt.<br>"
        "🟢 TỐT &amp; THỰC: cả hai ≥ 80%<br>"
        "🔵 KHÁ: cả hai ≥ 60%<br>"
        "🔴 CỜ ĐỎ: tài chính ≥ 80% nhưng kiểm chứng &lt; 60% — số đẹp chưa chứng minh được<br>"
        "🟠 THẬT NHƯNG YẾU: kiểm chứng ≥ 80% nhưng tài chính &lt; 50%<br>"
        "🟡 THEO DÕI · ⚫ LOẠI · ⚪ THIẾU DỮ LIỆU (dưới 3 tiêu chí áp dụng được)",
}


def _criteria_tip(group, key):
    """Chú giải cho cột ✓/✗ của một tiêu chí."""
    cfg = (FIN_CRITERIA if group == "FIN" else REAL_CRITERIA).get(key)
    if not cfg:
        return ""
    tip = f"<b>{cfg['label']}</b><br>{cfg['desc']}"
    if "bac_bo" in cfg:
        tip += f"<br><i>Bác bỏ khả năng: {cfg['bac_bo']}</i>"
    tip += "<br><br>✓ đạt · ✗ không đạt · — không áp dụng hoặc thiếu dữ liệu (không tính vào mẫu số)"
    return tip


def _tip_for(key):
    """Chú giải cho một cột bất kỳ, kể cả cột tiêu chí."""
    if key.startswith("FIN_") and key != "FIN_Disp":
        return _criteria_tip("FIN", key[4:])
    if key.startswith("REAL_") and key != "REAL_Disp":
        return _criteria_tip("REAL", key[5:])
    return COL_TIPS.get(key, "")


def _apply_header_tips(table, cols):
    """Gắn chú giải lên tiêu đề cột."""
    for i, (_, key, _, _) in enumerate(cols):
        item = table.horizontalHeaderItem(i)
        if item is not None:
            tip = _tip_for(key)
            if tip:
                item.setToolTip(tip)

from fundamental_chart import ChartWorker
from PySide6.QtWebEngineWidgets import QWebEngineView


# ─────────────────────────────────────────────────────────────────────────────
# Chart panel — Plotly rendered to PNG, displayed as a native QLabel pixmap
# ─────────────────────────────────────────────────────────────────────────────
class ColoredHeader(QHeaderView):
    """Header tự vẽ để hỗ trợ màu nền khác nhau theo nhóm cột."""
    _DEF_BG    = QColor("#0A1628")
    _DEF_FG    = QColor("#7EB8D4")
    _BORDER_B  = QColor("#3D8EF0")
    _BORDER_R  = QColor("#1A2133")

    def __init__(self, parent=None):
        super().__init__(Qt.Horizontal, parent)
        self.setHighlightSections(False)
        self.setSectionsClickable(True)

    def paintSection(self, painter, rect, idx):
        painter.save()
        painter.setClipRect(rect)

        bg_raw = self.model().headerData(idx, Qt.Horizontal, Qt.BackgroundRole)
        fg_raw = self.model().headerData(idx, Qt.Horizontal, Qt.ForegroundRole)
        bg = bg_raw.color() if hasattr(bg_raw, "color") else (bg_raw if isinstance(bg_raw, QColor) else self._DEF_BG)
        fg = fg_raw.color() if hasattr(fg_raw, "color") else (fg_raw if isinstance(fg_raw, QColor) else self._DEF_FG)

        painter.fillRect(rect, bg)
        painter.setPen(self._BORDER_B)
        painter.drawLine(rect.left(), rect.bottom(), rect.right(), rect.bottom())
        painter.setPen(self._BORDER_R)
        painter.drawLine(rect.right(), rect.top(), rect.right(), rect.bottom() - 2)

        # draw sort arrow
        sort_col = self.sortIndicatorSection()
        sort_asc = self.sortIndicatorOrder() == Qt.AscendingOrder
        arrow_w = 0
        if sort_col == idx:
            arrow_w = 10
            ax = rect.right() - arrow_w - 3
            ay = rect.center().y()
            painter.setPen(fg)
            painter.setBrush(fg)
            if sort_asc:
                pts = [(ax, ay + 3), (ax + arrow_w, ay + 3), (ax + arrow_w // 2, ay - 2)]
            else:
                pts = [(ax, ay - 3), (ax + arrow_w, ay - 3), (ax + arrow_w // 2, ay + 2)]
            from PySide6.QtGui import QPolygon
            from PySide6.QtCore import QPoint
            poly = QPolygon([QPoint(x, y) for x, y in pts])
            painter.drawPolygon(poly)

        text = self.model().headerData(idx, Qt.Horizontal, Qt.DisplayRole) or ""
        painter.setPen(fg)
        f = QFont("Segoe UI", 9, QFont.Bold)
        f.setLetterSpacing(QFont.AbsoluteSpacing, 0.8)
        painter.setFont(f)
        text_rect = rect.adjusted(3, 0, -(3 + arrow_w + 4), -2)
        painter.drawText(text_rect, Qt.AlignCenter | Qt.AlignVCenter, str(text))
        painter.restore()


class ChartPanel(QWidget):
    """Chart side panel — uses QWebEngineView + HTML (no kaleido dependency)."""
    def __init__(self, parent=None):
        super().__init__(parent)
        self._dark     = True
        self._html_q   = ""
        self._html_a   = ""
        self._mode     = "quarterly"
        self._worker   = None
        self._ticker   = ""
        self._spin_idx = 0
        self._spin_msg = ""
        self._spin_tmr = QTimer(self)
        self._spin_tmr.timeout.connect(self._tick)
        self.setStyleSheet(f"background:{BG};")

        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)
        self._bar_w    = self._build_bar()
        self._sep_line = self._sep()
        layout.addWidget(self._bar_w)
        layout.addWidget(self._sep_line)

        self._web = QWebEngineView()
        self._web.setHtml(self._placeholder())
        layout.addWidget(self._web, stretch=1)

    def _placeholder(self):
        bg = "#0A1628" if self._dark else "#F0F4F9"
        fg = "#3D4D60" if self._dark else "#94A3B8"
        return (f'<html><body style="background:{bg};margin:0;display:flex;'
                f'align-items:center;justify-content:center;height:100vh;">'
                f'<p style="color:{fg};font-family:Segoe UI,sans-serif;'
                f'font-size:11px;letter-spacing:1px;">Click a row to load chart</p>'
                f'</body></html>')

    def _build_bar(self):
        w = QWidget(); w.setFixedHeight(36)
        w.setStyleSheet(f"background:{SURFACE};")
        h = QHBoxLayout(w); h.setContentsMargins(14, 0, 14, 0); h.setSpacing(8)
        self._chart_lbl = QLabel("—")
        self._chart_lbl.setStyleSheet(
            f"color:{TEXT1}; font-size:13px; font-weight:700;"
            f" font-family:'Consolas',monospace; letter-spacing:2px;")
        h.addWidget(self._chart_lbl); h.addStretch()
        self._chart_status = QLabel("Click a row to load chart")
        self._chart_status.setStyleSheet(f"color:{TEXT3}; font-size:9px; letter-spacing:1px;")
        h.addWidget(self._chart_status); h.addSpacing(10)
        self._btn_q = QPushButton("Quarterly")
        self._btn_a = QPushButton("Annual")
        for btn in (self._btn_q, self._btn_a):
            btn.setFixedSize(84, 22); btn.setCursor(Qt.PointingHandCursor)
        self._btn_q.clicked.connect(lambda: self._switch_mode("quarterly"))
        self._btn_a.clicked.connect(lambda: self._switch_mode("annual"))
        h.addWidget(self._btn_q); h.addWidget(self._btn_a)
        self._refresh_btns()
        return w

    def _sep(self):
        f = QFrame(); f.setFrameShape(QFrame.HLine); f.setFixedHeight(1)
        f.setStyleSheet(f"background:{BORDER}; border:none;"); return f

    def load_ticker(self, ticker: str):
        ticker = ticker.upper()
        if ticker == self._ticker and (self._html_q or self._html_a):
            return
        self._ticker = ticker
        self._html_q = ""; self._html_a = ""
        self._mode   = "quarterly"
        self._chart_lbl.setText(ticker)
        self._web.setHtml(self._placeholder())
        self._refresh_btns()
        if self._worker and self._worker.isRunning():
            self._worker.terminate(); self._worker.wait()
        self._spin_msg = f"Loading {ticker}…"
        self._spin_tmr.start(80)
        self._worker = ChartWorker(ticker, height=580, dark_mode=self._dark)
        self._worker.done.connect(self._on_done)
        self._worker.failed.connect(self._on_failed)
        self._worker.msg.connect(lambda m: setattr(self, "_spin_msg", m))
        self._worker.start()

    def _on_done(self, html_a, html_q, summary):
        self._spin_tmr.stop()
        self._html_q = html_q
        self._html_a = html_a
        self._mode   = "quarterly"
        self._web.setHtml(html_q)
        parts = summary.split("·")
        self._set_status(f"✓  {parts[-1].strip()}" if len(parts) > 1 else "✓", GREEN)
        self._refresh_btns()

    def _on_failed(self, msg):
        self._spin_tmr.stop()
        bg = "#0A1628" if self._dark else "#F0F4F9"
        self._web.setHtml(
            f'<html><body style="background:{bg};margin:0;display:flex;'
            f'align-items:center;justify-content:center;height:100vh;">'
            f'<p style="color:#DC2626;font-family:Segoe UI,sans-serif;'
            f'font-size:11px;">Error: {msg}</p></body></html>')
        self._set_status(f"✕  {msg}", RED)

    def _switch_mode(self, mode):
        if mode == self._mode: return
        html = self._html_a if mode == "annual" else self._html_q
        if not html: return
        self._mode = mode
        self._web.setHtml(html)
        self._refresh_btns()

    def _set_status(self, text, color=TEXT3):
        self._chart_status.setText(text)
        self._chart_status.setStyleSheet(f"color:{color}; font-size:9px; letter-spacing:1px;")

    def _tick(self):
        self._spin_idx = (self._spin_idx + 1) % len(SPINNER)
        self._set_status(f"{SPINNER[self._spin_idx]}  {self._spin_msg}", AMBER)

    def _refresh_btns(self):
        has_q = bool(self._html_q)
        has_a = bool(self._html_a)
        act = (f"QPushButton {{ background:{BLUE}; color:#FFF; border:none;"
               f" border-radius:3px; font-size:9px; font-weight:700;"
               f" letter-spacing:1px; font-family:'Segoe UI',sans-serif; }}")
        off = (f"QPushButton {{ background:transparent; color:{TEXT2};"
               f" border:1px solid {BORDER}; border-radius:3px; font-size:9px;"
               f" font-weight:500; letter-spacing:1px; font-family:'Segoe UI',sans-serif; }}"
               f"QPushButton:hover {{ color:{TEXT1}; border-color:{BORDER2}; }}")
        dis = (f"QPushButton {{ background:transparent; color:{TEXT3};"
               f" border:1px solid {BORDER}; border-radius:3px; font-size:9px;"
               f" font-family:'Segoe UI',sans-serif; }}")
        self._btn_q.setEnabled(has_q)
        self._btn_a.setEnabled(has_a)
        self._btn_q.setStyleSheet(act if (self._mode == "quarterly" and has_q) else off if has_q else dis)
        self._btn_a.setStyleSheet(act if (self._mode == "annual"    and has_a) else off if has_a else dis)

    def set_theme(self, dark: bool):
        self._dark = dark
        self.setStyleSheet(f"background:{BG};")
        self._bar_w.setStyleSheet(f"background:{SURFACE};")
        self._sep_line.setStyleSheet(f"background:{BORDER}; border:none;")
        self._chart_lbl.setStyleSheet(
            f"color:{TEXT1}; font-size:13px; font-weight:700;"
            f" font-family:'Consolas',monospace; letter-spacing:2px;")
        self._chart_status.setStyleSheet(f"color:{TEXT3}; font-size:9px; letter-spacing:1px;")
        self._refresh_btns()
        if self._ticker:
            prev = self._ticker; self._ticker = ""
            self.load_ticker(prev)


# ─────────────────────────────────────────────────────────────────────────────
# Custom table item with numeric sort
# ─────────────────────────────────────────────────────────────────────────────
class NumItem(QTableWidgetItem):
    def __init__(self, val, fmt=""):
        self._num = float(val) if val is not None and val == val else float("-inf")
        if val is None or (isinstance(val, float) and val != val):
            text = "—"
        elif fmt == "pct":
            text = f"{val:+.1f}%" if val != 0 else "—"
        elif fmt == "price":
            text = f"${val:,.2f}"
        elif fmt == "vnd":
            text = f"₫{val:,.0f}"
        elif fmt == "mcap":
            text = f"{val:,.1f}"
        elif fmt == "mcap_vnd":
            text = f"₫{val:,.0f} tỷ"
        elif fmt == "bil":
            text = f"{val:,.1f}"
        elif fmt == "bil_vnd":
            text = f"{val:,.0f}"
        elif fmt == "int":
            text = str(int(val))
        else:
            text = f"{val:.2f}" if isinstance(val, float) else str(val)
        super().__init__(text)

    def __lt__(self, other):
        if isinstance(other, NumItem):
            return self._num < other._num
        return super().__lt__(other)


def _score_cell(got, app):
    """Ô điểm dạng 'đạt/áp dụng'. Sort theo tỷ lệ chứ không theo số tuyệt đối,
    để mã có 4/4 không bị xếp dưới mã 5/6."""
    got = int(got or 0)
    app = int(app or 0)
    if not app:
        cell = NumItem(-1, "int")
        cell.setText("—")
        cell.setForeground(QColor(TEXT3))
        return cell
    pct  = got / app
    cell = NumItem(round(pct, 3))
    cell.setText(f"{got}/{app}")
    c  = "#1A5C2B" if pct >= 0.8 else "#1B3A5C" if pct >= 0.6 \
         else "#7D6608" if pct >= 0.4 else "#9C0006"
    bg = "#C6EFCE" if pct >= 0.8 else "#DDEEFF" if pct >= 0.6 \
         else "#FFF2CC" if pct >= 0.4 else "#FFC7CE"
    cell.setForeground(QColor(c)); cell.setBackground(QColor(bg))
    return cell


# ─────────────────────────────────────────────────────────────────────────────
# Background workers
# ─────────────────────────────────────────────────────────────────────────────
class ScanWorker(QThread):
    progress      = Signal(str)
    ticker_update = Signal(int, int, str, str)   # idx, total, ticker, score
    done          = Signal(object)               # DataFrame
    failed        = Signal(str)
    market_ctx    = Signal(object, object)       # market_ok, index_1y — chỉ để hiển thị

    def __init__(self, market, top, use_yf):
        super().__init__()
        self.market  = market
        self.top     = top
        self.use_yf  = use_yf
        self._cancel = False

    def cancel(self): self._cancel = True

    def run(self):
        try:
            from tradingview_screener import Query
            self.progress.emit(f"Querying TradingView — {self.market.upper()} top {self.top}…")
            _, raw = (Query().set_markets(self.market.lower())
                      .select(*FETCH_COLS)
                      .order_by("market_cap_basic", ascending=False)
                      .limit(self.top)
                      .get_scanner_data())
            df = clean_df(raw)
            self.progress.emit(f"Got {len(df)} stocks.")

            _US_MARKETS = {"america", "nasdaq", "nyse"}
            moat_cache = None
            if self.use_yf and self.market.lower() in _US_MARKETS:
                tickers = df["Ticker"].tolist()
                sectors = dict(zip(df["Ticker"], df["Sector"].fillna("")))
                moat_cache = {}
                total = len(tickers)
                for i, ticker in enumerate(tickers, 1):
                    if self._cancel:
                        self.failed.emit("Cancelled.")
                        return
                    sector = sectors.get(ticker, "")
                    proxy, score, w52_pct, eq_badge, eps_acc, fcf_margin, net_cash = fetch_moat_yfinance(ticker, sector)
                    moat_cache[ticker] = (proxy, score, w52_pct, eq_badge, eps_acc, fcf_margin, net_cash)
                    self.ticker_update.emit(i, total, ticker, score)
                    time.sleep(0.3)

            # Hướng thị trường chỉ dùng để hiển thị bối cảnh trên Dashboard —
            # KHÔNG cộng vào điểm, vì nó giống nhau cho mọi mã nên không phân biệt được mã nào.
            self.progress.emit(f"Fetching market direction ({_MKT_INDEX.get(self.market.lower(), ('','SP:SPX'))[1]})…")
            market_ok, index_1y = fetch_market_direction(self.market)
            self.market_ctx.emit(market_ok, index_1y)
            self.progress.emit("Chấm điểm Tài chính + Kiểm chứng…")
            df = score_stocks(df, moat_cache)
            self.done.emit(df)
        except Exception as e:
            self.failed.emit(str(e))


class TickerWorker(QThread):
    done   = Signal(object)   # single-row DataFrame
    failed = Signal(str)

    def __init__(self, ticker, market="america"):
        super().__init__()
        self.ticker = ticker.upper()
        self.market = market

    def run(self):
        try:
            from tradingview_screener import Query, Column
            _, raw = (Query().set_markets(self.market.lower())
                      .select(*FETCH_COLS)
                      .where(Column("name") == self.ticker)
                      .get_scanner_data())
            if raw.empty:
                self.failed.emit(f"'{self.ticker}' not found on TradingView.")
                return
            df = clean_df(raw)
            moat_cache = {self.ticker: fetch_moat_yfinance(
                self.ticker, df["Sector"].iloc[0] if "Sector" in df.columns else "")}
            df = score_stocks(df, moat_cache)
            self.done.emit(df)
        except Exception as e:
            self.failed.emit(str(e))


# ─────────────────────────────────────────────────────────────────────────────
# Thẻ chi tiết — dùng chung cho cả hai tab.
# Luôn hiện đồng thời hai điểm (Tài chính · Thực) để không bao giờ đọc một
# điểm mà quên điểm kia; phần ô ✓/✗ thì đổi theo nhóm tiêu chí của tab.
# ─────────────────────────────────────────────────────────────────────────────
class CriteriaCard(QWidget):
    GROUP = "FIN"          # lớp con ghi đè

    def __init__(self, parent=None):
        super().__init__(parent)
        self._criteria = FIN_CRITERIA if self.GROUP == "FIN" else REAL_CRITERIA
        self._keys     = FIN_KEYS     if self.GROUP == "FIN" else REAL_KEYS

        self.setStyleSheet(f"background:{PANEL};")
        layout = QVBoxLayout(self)
        layout.setContentsMargins(18, 10, 18, 10)
        layout.setSpacing(8)

        self._lbl_ticker = QLabel("—")
        self._lbl_ticker.setStyleSheet(
            f"color:{TEXT1}; font-size:20px; font-weight:700;"
            f" font-family:'Consolas',monospace; letter-spacing:3px;")
        self._lbl_info = QLabel("")
        self._lbl_info.setWordWrap(True)
        self._lbl_info.setStyleSheet(
            f"color:{TEXT1}; font-size:13px; font-family:'Segoe UI',sans-serif;")
        self._lbl_moat = QLabel("")
        self._lbl_moat.setStyleSheet(
            f"font-size:12px; font-weight:700; font-family:'Segoe UI',sans-serif;"
            f" padding:3px 12px; border-radius:3px;")
        self._lbl_signal = QLabel("")
        self._lbl_signal.setStyleSheet(
            f"font-size:12px; font-weight:700; font-family:'Segoe UI',sans-serif;"
            f" padding:3px 12px; border-radius:3px;")

        top_row = QHBoxLayout()
        top_row.addWidget(self._lbl_ticker)
        top_row.addWidget(self._lbl_info)
        top_row.addStretch()
        top_row.addWidget(self._lbl_moat)
        top_row.addWidget(self._lbl_signal)
        layout.addLayout(top_row)

        box = QWidget()
        self._box = box
        box.setStyleSheet(f"background:{SURFACE}; border-radius:4px;")
        row = QHBoxLayout(box)
        row.setContentsMargins(12, 8, 12, 8)
        row.setSpacing(6)

        self._cells = {}
        for key in self._keys:
            cfg = self._criteria[key]
            col = QVBoxLayout()
            col.setSpacing(3)
            cap = QLabel(cfg["label"])
            cap.setAlignment(Qt.AlignCenter)
            cap.setStyleSheet(
                f"color:{TEXT2}; font-size:10px; font-weight:600;"
                f" font-family:'Segoe UI',sans-serif;")
            val = QLabel("—")
            val.setAlignment(Qt.AlignCenter)
            val.setFixedSize(74, 28)
            val.setStyleSheet(
                f"color:{TEXT1}; font-size:13px; font-weight:700;"
                f" background:{PANEL}; border-radius:3px; font-family:'Consolas',monospace;")
            col.addWidget(cap)
            col.addWidget(val)
            row.addLayout(col)
            self._cells[key] = val

        row.addStretch()

        self._lbl_fin  = QLabel("—")
        self._lbl_real = QLabel("—")
        for lbl, cap_txt in ((self._lbl_fin, "TÀI CHÍNH"), (self._lbl_real, "THỰC")):
            col = QVBoxLayout()
            col.setSpacing(2)
            cap = QLabel(cap_txt)
            cap.setAlignment(Qt.AlignCenter)
            cap.setStyleSheet(
                f"color:{TEXT2}; font-size:9px; font-weight:700; letter-spacing:1px;"
                f" font-family:'Segoe UI',sans-serif;")
            lbl.setAlignment(Qt.AlignCenter)
            lbl.setStyleSheet(
                f"color:{TEXT1}; font-size:17px; font-weight:700;"
                f" font-family:'Consolas',monospace;")
            col.addWidget(cap)
            col.addWidget(lbl)
            row.addLayout(col)
        layout.addWidget(box)

        self._lbl_empty = QLabel("Chọn một dòng để xem chi tiết")
        self._lbl_empty.setAlignment(Qt.AlignCenter)
        self._lbl_empty.setStyleSheet(
            f"color:{TEXT3}; font-size:12px; font-style:italic;"
            f" font-family:'Segoe UI',sans-serif;")
        layout.addWidget(self._lbl_empty)

    # ── hiển thị ────────────────────────────────────────────────────────────
    @staticmethod
    def _fmt(field, val):
        if val is None or (isinstance(val, float) and pd.isna(val)):
            return "—"
        if isinstance(val, (int, float)):
            if field in ("FCF/NI",):                    return f"{val:.2f}×"
            if field in ("Net Cash ($B)",):             return f"{val:+,.1f}B"
            if field in ("D/E", "Current Ratio"):       return f"{val:.2f}"
            if field.endswith("%") or "Gap" in field:   return f"{val:+.1f}%"
            return f"{val:,.1f}"
        return str(val)

    def _score_style(self, lbl, got, app):
        if not app:
            lbl.setText("—")
            lbl.setStyleSheet(
                f"color:{TEXT2}; font-size:17px; font-weight:700;"
                f" font-family:'Consolas',monospace;")
            return
        lbl.setText(f"{got}/{app}")
        pct = got / app
        c = GREEN if pct >= 0.8 else BLUE if pct >= 0.6 else AMBER if pct >= 0.4 else RED
        lbl.setStyleSheet(
            f"color:{c}; font-size:17px; font-weight:700;"
            f" font-family:'Consolas',monospace;")

    def show_row(self, row: dict):
        self._lbl_empty.hide()
        self._lbl_ticker.setText(row.get("Ticker", "—"))
        parts = [p for p in [row.get("Tên Công Ty", ""), row.get("Sector", "")] if p]
        price = row.get("Price ($)"); mcap = row.get("MCap ($B)")
        price_str = f"  ${price:,.2f}" if isinstance(price, (int, float)) else ""
        mcap_str  = f"  MCap ${mcap:,.1f}B" if isinstance(mcap, (int, float)) else ""
        self._lbl_info.setText("  ·  ".join(parts) + price_str + mcap_str)

        moat = row.get("Moat Score", "")
        if moat in MOAT_SCORE_STYLE:
            fg, bg = MOAT_SCORE_STYLE[moat]
            self._lbl_moat.setText(f"🏰 {moat}")
            self._lbl_moat.setStyleSheet(
                f"color:#{fg}; background:#{bg}; font-size:10px; font-weight:700;"
                f" font-family:'Segoe UI',sans-serif; padding:2px 10px; border-radius:3px;")
        else:
            self._lbl_moat.setText("")

        sig = row.get("Signal", "")
        if sig in SIGNAL_COLORS:
            fg, bg = SIGNAL_COLORS[sig]
            self._lbl_signal.setText(sig)
            self._lbl_signal.setStyleSheet(
                f"color:{fg}; background:{bg}; font-size:10px; font-weight:700;"
                f" font-family:'Segoe UI',sans-serif; padding:2px 10px; border-radius:3px;")
        else:
            self._lbl_signal.setText("")

        self._score_style(self._lbl_fin,  row.get("FIN_Score", 0),  row.get("FIN_App", 0))
        self._score_style(self._lbl_real, row.get("REAL_Score", 0), row.get("REAL_App", 0))

        for key in self._keys:
            cfg    = self._criteria[key]
            passed = row.get(f"{self.GROUP}_{key}")
            lbl    = self._cells[key]
            shown  = self._fmt(cfg["field"], row.get(cfg["field"]))
            tip    = f"<b>{cfg['label']}</b> — {shown}<br>{cfg['desc']}"
            if "bac_bo" in cfg:
                tip += f"<br><i>Bác bỏ khả năng: {cfg['bac_bo']}</i>"
            if passed is True:
                lbl.setText(f"✓ {shown}")
                lbl.setStyleSheet(
                    f"color:#1A5C2B; font-size:12px; font-weight:700;"
                    f" background:#C6EFCE; border-radius:3px; font-family:'Consolas',monospace;")
            elif passed is False:
                lbl.setText(f"✗ {shown}")
                lbl.setStyleSheet(
                    f"color:#9C0006; font-size:12px; font-weight:700;"
                    f" background:#FFC7CE; border-radius:3px; font-family:'Consolas',monospace;")
            else:
                lbl.setText("—")
                tip += "<br><i>Không áp dụng hoặc thiếu dữ liệu — không tính vào mẫu số</i>"
                lbl.setStyleSheet(
                    f"color:{TEXT2}; font-size:12px; font-weight:700;"
                    f" background:{PANEL}; border-radius:3px; font-family:'Consolas',monospace;")
            lbl.setToolTip(tip)

    def clear(self):
        self._lbl_ticker.setText("—")
        self._lbl_info.setText("")
        self._lbl_moat.setText("")
        self._lbl_signal.setText("")
        for lbl in (self._lbl_fin, self._lbl_real):
            lbl.setText("—")
            lbl.setStyleSheet(
                f"color:{TEXT2}; font-size:17px; font-weight:700;"
                f" font-family:'Consolas',monospace;")
        for lbl in self._cells.values():
            lbl.setText("—")
            lbl.setStyleSheet(
                f"color:{TEXT1}; font-size:12px; font-weight:700;"
                f" background:{PANEL}; border-radius:3px; font-family:'Consolas',monospace;")
        self._lbl_empty.show()

    def apply_theme(self):
        self.setStyleSheet(f"background:{PANEL};")
        self._box.setStyleSheet(f"background:{SURFACE}; border-radius:4px;")
        self._lbl_ticker.setStyleSheet(
            f"color:{TEXT1}; font-size:20px; font-weight:700;"
            f" font-family:'Consolas',monospace; letter-spacing:3px;")
        self._lbl_info.setStyleSheet(
            f"color:{TEXT1}; font-size:13px; font-family:'Segoe UI',sans-serif;")
        self._lbl_empty.setStyleSheet(
            f"color:{TEXT3}; font-size:12px; font-style:italic;"
            f" font-family:'Segoe UI',sans-serif;")
        for lbl in self._cells.values():
            if lbl.text() == "—":
                lbl.setStyleSheet(
                    f"color:{TEXT1}; font-size:12px; font-weight:700;"
                    f" background:{PANEL}; border-radius:3px; font-family:'Consolas',monospace;")


class DetailCard(CriteriaCard):
    """Tab 1 — 6 tiêu chí tài chính."""
    GROUP = "FIN"


class QualityDetailCard(CriteriaCard):
    """Tab 2 — 6 tiêu chí kiểm chứng."""
    GROUP = "REAL"


# ─────────────────────────────────────────────────────────────────────────────
# Portfolio Panel — Tab My Portfolio
# ─────────────────────────────────────────────────────────────────────────────
class PortfolioPanel(QWidget):
    """Quản lý danh mục cổ phiếu — lưu vào portfolio.json."""

    _COLS = [
        ("#",         40, Qt.AlignCenter),
        ("Ticker",    72, Qt.AlignCenter),
        ("Company",  160, Qt.AlignLeft),
        ("Buy Date",  90, Qt.AlignCenter),
        ("Buy $",     80, Qt.AlignRight),
        ("Qty",       58, Qt.AlignCenter),
        ("Cost",      92, Qt.AlignRight),
        ("Price",     80, Qt.AlignRight),
        ("Value",     92, Qt.AlignRight),
        ("P&L $",     88, Qt.AlignRight),
        ("P&L %",     68, Qt.AlignRight),
        ("CS",        40, Qt.AlignCenter),
        ("QC",        40, Qt.AlignCenter),
        ("Signal",   120, Qt.AlignCenter),
        ("",          34, Qt.AlignCenter),
    ]

    def __init__(self, parent=None):
        super().__init__(parent)
        self._holdings = []
        self._scan_df  = None
        self.setStyleSheet(f"background:{BG};")
        self._build_ui()
        self._load()
        self._refresh_table()

    @staticmethod
    def _pf_path():
        base = os.path.dirname(os.path.abspath(sys.argv[0]))
        return os.path.join(base, "portfolio.json")

    # ── Persistence ───────────────────────────────────────────────────────────
    def _load(self):
        import json
        try:
            with open(self._pf_path(), encoding="utf-8") as f:
                self._holdings = json.load(f)
        except Exception:
            self._holdings = []

    def _save(self):
        import json
        try:
            with open(self._pf_path(), "w", encoding="utf-8") as f:
                json.dump(self._holdings, f, ensure_ascii=False, indent=2)
        except Exception:
            pass

    # ── Build UI ──────────────────────────────────────────────────────────────
    def _sep(self):
        f = QFrame(); f.setFrameShape(QFrame.HLine); f.setFixedHeight(1)
        f.setStyleSheet(f"background:{BORDER}; border:none;"); return f

    def _build_ui(self):
        lay = QVBoxLayout(self)
        lay.setContentsMargins(0, 0, 0, 0); lay.setSpacing(0)

        # KPI bar
        self._kpi_bar = QWidget()
        self._kpi_bar.setFixedHeight(74)
        self._kpi_bar.setStyleSheet(f"background:{SURFACE};")
        self._kpi_hlay = QHBoxLayout(self._kpi_bar)
        self._kpi_hlay.setContentsMargins(14, 6, 14, 6); self._kpi_hlay.setSpacing(6)
        lay.addWidget(self._kpi_bar)
        lay.addWidget(self._sep())

        # Holdings table
        self._table = QTableWidget(0, len(self._COLS))
        self._table.setHorizontalHeaderLabels([c[0] for c in self._COLS])
        self._table.verticalHeader().setVisible(False)
        self._table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self._table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self._table.setAlternatingRowColors(False)
        self._table.horizontalHeader().setHighlightSections(False)
        self._table.setStyleSheet(f"""
            QTableWidget {{ background:{BG}; gridline-color:{BORDER}; outline:none; font-size:12px; }}
            QTableWidget::item {{ padding:3px 6px; }}
            QTableWidget::item:selected {{ background:{BORDER2}; color:{TEXT1}; }}
            QHeaderView::section {{
                background:{SURFACE}; color:{TEXT2}; font-size:9px;
                font-weight:700; letter-spacing:1px;
                padding:5px 6px; border:none; border-bottom:1px solid {BORDER};
            }}
        """)
        for i, (_, w, _) in enumerate(self._COLS):
            self._table.setColumnWidth(i, w)
        self._table.horizontalHeader().setSectionResizeMode(2, QHeaderView.Stretch)
        lay.addWidget(self._table, stretch=1)
        lay.addWidget(self._sep())

        # Add form
        form = QWidget(); form.setFixedHeight(44)
        form.setStyleSheet(f"background:{SURFACE};")
        fh = QHBoxLayout(form)
        fh.setContentsMargins(14, 0, 14, 0); fh.setSpacing(6)

        _is = (f"background:{INPUT_BG};color:{TEXT1};border:1px solid {BORDER};"
               f"border-radius:3px;padding:0 8px;font-size:12px;")

        def _lbl(t):
            l = QLabel(t)
            l.setStyleSheet(f"color:{TEXT3};font-size:9px;letter-spacing:1px;")
            return l

        self._inp_ticker = QLineEdit()
        self._inp_ticker.setPlaceholderText("AAPL")
        self._inp_ticker.setFixedSize(72, 28); self._inp_ticker.setMaxLength(12)
        self._inp_ticker.setStyleSheet(_is)

        self._inp_date = QLineEdit()
        self._inp_date.setPlaceholderText("YYYY-MM-DD")
        self._inp_date.setFixedSize(100, 28); self._inp_date.setStyleSheet(_is)
        self._inp_date.setText(pd.Timestamp.now().strftime("%Y-%m-%d"))

        self._inp_price = QLineEdit()
        self._inp_price.setPlaceholderText("Giá mua")
        self._inp_price.setFixedSize(90, 28); self._inp_price.setStyleSheet(_is)

        self._inp_qty = QLineEdit()
        self._inp_qty.setPlaceholderText("Số CP")
        self._inp_qty.setFixedSize(70, 28); self._inp_qty.setStyleSheet(_is)

        self._inp_note = QLineEdit()
        self._inp_note.setPlaceholderText("Ghi chú (tuỳ chọn)")
        self._inp_note.setFixedHeight(28); self._inp_note.setStyleSheet(_is)

        btn_add = QPushButton("＋  THÊM")
        btn_add.setFixedSize(84, 28); btn_add.setCursor(Qt.PointingHandCursor)
        btn_add.setStyleSheet(
            f"QPushButton{{background:{BLUE};color:#FFF;border:none;border-radius:3px;"
            f"font-size:9px;font-weight:700;letter-spacing:1px;}}"
            f"QPushButton:hover{{background:{BLUE_HV};}}")
        btn_add.clicked.connect(self._add_holding)
        self._inp_ticker.returnPressed.connect(self._add_holding)

        fh.addWidget(_lbl("TICKER")); fh.addWidget(self._inp_ticker); fh.addSpacing(4)
        fh.addWidget(_lbl("NGÀY MUA")); fh.addWidget(self._inp_date); fh.addSpacing(4)
        fh.addWidget(_lbl("GIÁ")); fh.addWidget(self._inp_price); fh.addSpacing(4)
        fh.addWidget(_lbl("SỐ CP")); fh.addWidget(self._inp_qty); fh.addSpacing(4)
        fh.addWidget(self._inp_note, stretch=1)
        fh.addWidget(btn_add)
        lay.addWidget(form)

    # ── Actions ───────────────────────────────────────────────────────────────
    def _add_holding(self):
        ticker = self._inp_ticker.text().strip().upper()
        date   = self._inp_date.text().strip()
        try:
            price = float(self._inp_price.text().strip().replace(",", ""))
            qty   = float(self._inp_qty.text().strip().replace(",", ""))
        except ValueError:
            return
        if not ticker or price <= 0 or qty <= 0:
            return
        self._holdings.append({
            "ticker":    ticker,
            "buy_date":  date,
            "buy_price": price,
            "qty":       qty,
            "note":      self._inp_note.text().strip(),
        })
        self._save()
        self._inp_ticker.clear(); self._inp_price.clear()
        self._inp_qty.clear();    self._inp_note.clear()
        self._refresh_table()

    def _delete_holding(self, idx: int):
        if 0 <= idx < len(self._holdings):
            self._holdings.pop(idx)
            self._save(); self._refresh_table()

    # ── Data helpers ──────────────────────────────────────────────────────────
    def _get_price(self, ticker):
        if self._scan_df is not None and not self._scan_df.empty:
            m = self._scan_df[self._scan_df["Ticker"] == ticker]
            if not m.empty:
                return m.iloc[0].get("Price ($)")
        return None

    def _get_scores(self, ticker):
        if self._scan_df is not None and not self._scan_df.empty:
            m = self._scan_df[self._scan_df["Ticker"] == ticker]
            if not m.empty:
                r = m.iloc[0]
                return (r.get("FIN_Score"), r.get("REAL_Score"),
                        r.get("Signal",""), r.get("Tên Công Ty",""))
        return None, None, "", ""

    def refresh_scan(self, df):
        self._scan_df = df
        self._refresh_table()

    # ── Table render ──────────────────────────────────────────────────────────
    def _refresh_table(self):
        self._table.setRowCount(0)
        total_cost = total_val = 0.0
        pl_pcts = []

        _C = Qt.AlignVCenter | Qt.AlignCenter
        _R = Qt.AlignVCenter | Qt.AlignRight
        _L = Qt.AlignVCenter | Qt.AlignLeft

        SIG_STYLE = {k: (bg.lstrip("#"), fg.lstrip("#"))
                     for k, (fg, bg) in SIGNAL_COLORS.items()}

        for i, h in enumerate(self._holdings):
            ticker    = h.get("ticker","")
            buy_px    = float(h.get("buy_price", 0))
            qty       = float(h.get("qty", 0))
            buy_date  = h.get("buy_date","")
            cost      = buy_px * qty
            curr_px   = self._get_price(ticker)
            curr_val  = curr_px * qty if curr_px else None
            pl_d      = (curr_val - cost)  if curr_val is not None else None
            pl_p      = ((curr_px - buy_px) / buy_px * 100) if curr_px else None
            cs, qc, sig, company = self._get_scores(ticker)

            total_cost += cost
            if curr_val: total_val += curr_val
            if pl_p is not None: pl_pcts.append((ticker, pl_p))

            row_bg = "F0F4FA" if i % 2 == 0 else "FFFFFF"
            row = self._table.rowCount()
            self._table.insertRow(row)
            self._table.setRowHeight(row, 26)

            def _cell(txt, align=_R, color=TEXT1, bold=False, bg=None):
                item = QTableWidgetItem(str(txt))
                item.setTextAlignment(align)
                item.setForeground(QColor(f"#{color}"))
                item.setBackground(QColor(f"#{bg}" if bg else f"#{row_bg}"))
                if bold:
                    fn = item.font(); fn.setBold(True); item.setFont(fn)
                return item

            self._table.setItem(row, 0,  _cell(i+1,   _C, TEXT3))
            self._table.setItem(row, 1,  _cell(ticker, _C, BLUE[1:], bold=True))
            self._table.setItem(row, 2,  _cell(company[:28] or "—", _L))
            self._table.setItem(row, 3,  _cell(buy_date, _C, TEXT2))
            self._table.setItem(row, 4,  _cell(f"${buy_px:,.2f}", _R))
            self._table.setItem(row, 5,  _cell(f"{qty:,.0f}", _C))
            self._table.setItem(row, 6,  _cell(f"${cost:,.0f}", _R, TEXT2))

            if curr_px:
                self._table.setItem(row, 7, _cell(f"${curr_px:,.2f}", _R))
                self._table.setItem(row, 8, _cell(f"${curr_val:,.0f}", _R))
                pl_c = GREEN[1:] if pl_d >= 0 else RED[1:]
                self._table.setItem(row, 9,  _cell(f"{pl_d:+,.0f}", _R, pl_c, bold=True))
                self._table.setItem(row, 10, _cell(f"{pl_p:+.1f}%",  _R, pl_c, bold=True))
            else:
                for col in range(7, 11):
                    self._table.setItem(row, col, _cell("—", _C, TEXT3))

            cs_c = GREEN[1:] if cs is not None and int(cs) >= 7 else TEXT2
            self._table.setItem(row, 11, _cell(str(int(cs)) if cs is not None else "—", _C, cs_c))
            qc_c = GREEN[1:] if qc is not None and int(qc) >= 4 else TEXT2
            self._table.setItem(row, 12, _cell(str(int(qc)) if qc is not None else "—", _C, qc_c))

            if sig in SIG_STYLE:
                sbg, sfg = SIG_STYLE[sig]
                self._table.setItem(row, 13, _cell(sig, _C, sfg, bg=sbg))
            else:
                self._table.setItem(row, 13, _cell("—", _C, TEXT3))

            # Delete button
            btn_del = QPushButton("✕")
            btn_del.setFixedSize(24, 22); btn_del.setCursor(Qt.PointingHandCursor)
            btn_del.setStyleSheet(
                f"QPushButton{{background:transparent;color:{TEXT3};border:none;font-size:10px;}}"
                f"QPushButton:hover{{color:{RED};}}")
            btn_del.clicked.connect(lambda _, idx=i: self._delete_holding(idx))
            cw = QWidget(); cw.setStyleSheet(f"background:#{row_bg};")
            cl = QHBoxLayout(cw); cl.setContentsMargins(4,0,4,0); cl.addWidget(btn_del)
            self._table.setCellWidget(row, 14, cw)

        self._refresh_kpi(total_cost, total_val, pl_pcts)

    def _refresh_kpi(self, total_cost, total_val, pl_pcts):
        while self._kpi_hlay.count():
            it = self._kpi_hlay.takeAt(0)
            if it.widget(): it.widget().deleteLater()

        pl_d   = total_val - total_cost if total_val > 0 else 0
        pl_p   = (pl_d / total_cost * 100) if total_cost > 0 else 0
        best   = max(pl_pcts, key=lambda x: x[1]) if pl_pcts else None
        worst  = min(pl_pcts, key=lambda x: x[1]) if pl_pcts else None

        def _card(val, lbl, col=None):
            card = QWidget()
            card.setStyleSheet(
                f"background:{SURFACE};border:1px solid {BORDER};border-radius:4px;")
            cv = QVBoxLayout(card); cv.setContentsMargins(10,5,10,4); cv.setSpacing(1)
            vl = QLabel(val)
            vl.setStyleSheet(f"color:{col or TEXT1};font-size:17px;font-weight:700;")
            vl.setAlignment(Qt.AlignCenter)
            ll = QLabel(lbl)
            ll.setStyleSheet(f"color:{TEXT3};font-size:9px;letter-spacing:1px;")
            ll.setAlignment(Qt.AlignCenter)
            cv.addWidget(vl); cv.addWidget(ll)
            return card

        self._kpi_hlay.addWidget(_card(f"${total_cost:,.0f}", "TỔNG ĐẦU TƯ"))
        self._kpi_hlay.addWidget(_card(
            f"${total_val:,.0f}" if total_val > 0 else "—", "GIÁ TRỊ HIỆN TẠI"))
        c = GREEN if pl_d >= 0 else RED
        self._kpi_hlay.addWidget(_card(
            f"{pl_d:+,.0f}" if total_val > 0 else "—", "P&L ($)", c))
        self._kpi_hlay.addWidget(_card(
            f"{pl_p:+.1f}%" if total_val > 0 else "—", "P&L (%)", c))
        self._kpi_hlay.addWidget(_card(str(len(self._holdings)), "# MÃ"))
        if best:
            self._kpi_hlay.addWidget(_card(
                f"{best[0]}  {best[1]:+.1f}%", "BEST", GREEN))
        if worst and (best is None or worst[0] != best[0]):
            self._kpi_hlay.addWidget(_card(
                f"{worst[0]}  {worst[1]:+.1f}%", "WORST", RED))
        self._kpi_hlay.addStretch()

    def apply_theme(self):
        self.setStyleSheet(f"background:{BG};")
        self._kpi_bar.setStyleSheet(f"background:{SURFACE};")
        self._refresh_table()


# ─────────────────────────────────────────────────────────────────────────────
# Help dialog — F1
# ─────────────────────────────────────────────────────────────────────────────
class HelpDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Help  ·  F1")
        self.resize(800, 660)
        self.setStyleSheet(f"""
            QDialog      {{ background:{BG}; }}
            *            {{ background:transparent; color:{TEXT1}; }}
            QTextBrowser {{
                background:{SURFACE}; color:{TEXT1};
                border:none; padding:20px 26px;
                font-family:'Segoe UI',sans-serif;
            }}
            QPushButton  {{
                background:{BLUE}; color:#FFFFFF;
                border:none; border-radius:3px;
                font-size:10px; font-weight:700; letter-spacing:1.5px;
                font-family:'Segoe UI',sans-serif;
            }}
            QPushButton:hover {{ background:{BLUE_HV}; }}
        """)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)

        browser = QTextBrowser()
        browser.setOpenExternalLinks(False)
        browser.setHtml(self._html())
        layout.addWidget(browser, stretch=1)

        foot = QWidget()
        foot.setFixedHeight(44)
        foot.setStyleSheet(f"background:{SURFACE}; border-top:1px solid {BORDER};")
        fh = QHBoxLayout(foot)
        fh.setContentsMargins(18, 0, 18, 0)
        hint = QLabel("F1 / Esc  to close")
        hint.setStyleSheet(f"color:{TEXT3}; font-size:9px; letter-spacing:1px;")
        btn = QPushButton("CLOSE")
        btn.setFixedSize(80, 28)
        btn.setCursor(Qt.PointingHandCursor)
        btn.clicked.connect(self.close)
        fh.addWidget(hint); fh.addStretch(); fh.addWidget(btn)
        layout.addWidget(foot)

    def keyPressEvent(self, event):
        if event.key() in (Qt.Key_Escape, Qt.Key_F1):
            self.close()
        else:
            super().keyPressEvent(event)

    def _html(self):
        H_DIM = "#64748B"; H_BLUE = "#1E40AF"; H_GREEN = "#16A34A"
        H_RED = "#B91C1C"; H_ORANGE = "#C2410C"; H_TEXT = "#1A202C"

        def rows(criteria, keys, accent):
            out = ""
            for i, k in enumerate(keys):
                c   = criteria[k]
                op  = {"gt": "&gt;", "lt": "&lt;"}.get(c["op"], "")
                thr = c["thr"]
                if k == "CASH":     thrs = "≥ 0.80×"
                elif k == "BUYB":   thrs = "≤ 10 điểm %"
                elif k == "LEV":    thrs = "≤ 15 điểm %"
                elif k == "BAL":    thrs = "&gt; 0 hoặc CR &gt; 1.5"
                elif k == "SUS":    thrs = "WIDE / NARROW"
                elif thr is None:   thrs = "—"
                else:               thrs = f"{op} {thr}{'%' if c['op'] == 'gt' and thr and thr > 2 else ''}"
                bg   = "#F0F4FA" if i % 2 == 0 else "#FFFFFF"
                extra = (f'<br><span style="color:{H_RED};font-size:11px;">Bác bỏ khả năng: '
                         f'{c["bac_bo"]}</span>') if "bac_bo" in c else ""
                out += f"""
                <tr style="background-color:{bg};">
                  <td style="color:{accent};font-weight:700;font-size:12px;padding:8px 12px;width:92px;">{c['label']}</td>
                  <td style="color:{H_DIM};font-family:Consolas,monospace;font-size:11px;padding:8px 12px;width:110px;">{c['field']}</td>
                  <td style="color:{H_GREEN};font-weight:700;padding:8px 12px;width:120px;">{thrs}</td>
                  <td style="color:{H_TEXT};font-size:12px;padding:8px 12px;line-height:1.55;">{c['desc']}{extra}</td>
                </tr>"""
            return out

        fin_rows  = rows(FIN_CRITERIA, FIN_KEYS, H_BLUE)
        real_rows = rows(REAL_CRITERIA, REAL_KEYS, H_ORANGE)

        sig_rows = ""
        for i, (icon, name, cond, fg, bg, desc) in enumerate([
            ("🟢", "TỐT &amp; THỰC", "A ≥ 80%  và  B ≥ 80%", "#1A5C2B", "#C6EFCE",
             "Số liệu tài chính mạnh và vượt được các phép kiểm chứng. Đây là nhóm đáng nghiên cứu sâu trước tiên."),
            ("🔵", "KHÁ", "A ≥ 60%  và  B ≥ 60%", "#1B3A5C", "#DDEEFF",
             "Cả hai mặt đều khá. Xem tiêu chí nào chưa đạt trước khi kết luận."),
            ("🔴", "CỜ ĐỎ", "A ≥ 80%  nhưng  B &lt; 60%", "#9C0006", "#FFC7CE",
             "Số liệu rất đẹp nhưng không chứng minh được bằng dòng tiền, cơ cấu vốn hay lịch sử. "
             "Đây chính là nhóm mà hệ thống được thiết kế để phát hiện — đẹp trên báo cáo, mỏng ở thực chất."),
            ("🟠", "THẬT NHƯNG YẾU", "B ≥ 80%  nhưng  A &lt; 50%", "#7A3E00", "#FFE0C2",
             "Sổ sách trung thực, dòng tiền thật, nhưng biên lợi nhuận hoặc tăng trưởng chưa đủ hấp dẫn."),
            ("🟡", "THEO DÕI", "A ≥ 50%", "#7D6608", "#FFF2CC",
             "Chưa đủ điều kiện. Đưa vào danh sách theo dõi và quét lại sau."),
            ("⚫", "LOẠI", "còn lại", "#3D4D60", "#E2E8F0",
             "Không đạt cả hai mặt."),
            ("⚪", "THIẾU DỮ LIỆU", "&lt; 3 tiêu chí áp dụng được", "#64748B", "#F1F5F9",
             "Không đủ dữ liệu để kết luận. Hệ thống không đoán — nó nói thẳng là không biết."),
        ]):
            sig_rows += f"""
            <tr style="background-color:{bg};">
              <td style="color:{fg};font-weight:700;font-size:13px;padding:9px 13px;width:150px;">{icon}&nbsp;{name}</td>
              <td style="color:{fg};font-weight:700;font-size:12px;padding:9px 13px;width:170px;font-family:Consolas,monospace;">{cond}</td>
              <td style="color:{fg};font-size:12px;padding:9px 13px;line-height:1.55;">{desc}</td>
            </tr>"""

        pair_rows = ""
        for i, (claim, risk, proof) in enumerate([
            ("Lợi nhuận cao", "Ghi nhận doanh thu chưa thu được tiền", "FCF/NI — dòng tiền tự do ≥ 80% lợi nhuận"),
            ("Biên lợi nhuận cao", "Vốn hóa chi phí, hoãn đầu tư cho đẹp sổ", "FCF Mgn — biên dòng tiền tự do &gt; 8%"),
            ("EPS tăng mạnh", "Mua cổ phiếu quỹ làm giảm mẫu số", "Không buyback — EPS không vượt lợi nhuận quá 10 điểm %"),
            ("ROE cao", "Đòn bẩy lớn, hoặc vốn chủ teo vì mua cổ phiếu quỹ", "ROE thật — ROE không vượt ROIC quá 15 điểm %"),
            ("Bảng cân đối sạch", "Thanh khoản thực tế yếu, nợ ẩn", "Cân đối — Net Cash &gt; 0 hoặc Current Ratio &gt; 1.5"),
            ("Một năm xuất sắc", "Một kỳ may mắn được trình bày như năng lực", "Bền vững — ROE &amp; biên gộp 5 năm đều cao"),
        ]):
            bg = "#F0F4FA" if i % 2 == 0 else "#FFFFFF"
            pair_rows += f"""
            <tr style="background-color:{bg};">
              <td style="color:{H_BLUE};font-weight:700;font-size:12px;padding:8px 12px;width:150px;">{claim}</td>
              <td style="color:{H_RED};font-size:12px;padding:8px 12px;width:270px;">{risk}</td>
              <td style="color:{H_GREEN};font-weight:700;font-size:12px;padding:8px 12px;">{proof}</td>
            </tr>"""

        return f"""<html><head><style>
  body  {{ background-color:#F8FAFC; color:#1A202C;
           font-family:'Segoe UI',sans-serif; font-size:13px; margin:0; padding:0; }}
  h1    {{ color:#1E40AF; font-size:20px; letter-spacing:3px; font-weight:700; margin:0 0 5px 0; }}
  .sub  {{ color:#64748B; font-size:11px; letter-spacing:2px; margin:0 0 20px 0; }}
  h2    {{ color:#374151; font-size:12px; letter-spacing:2px; font-weight:700;
           margin:26px 0 9px 0; border-bottom:2px solid #CBD5E0; padding-bottom:5px; }}
  table {{ border-collapse:collapse; width:100%; margin-bottom:10px; }}
  th    {{ background-color:#E2E8F0; color:#2D3748; text-align:left; padding:8px 13px;
           font-size:11px; letter-spacing:1px; font-weight:700; border-bottom:2px solid #CBD5E0; }}
  .tip  {{ color:#64748B; font-size:11px; font-style:italic; margin:5px 0 0 0; }}
  li    {{ margin:7px 0; color:#1A202C; font-size:12px; }}
  b     {{ color:#0F172A; font-weight:700; }}
  .box  {{ background:#EBF4FF; border-left:3px solid #1E40AF; padding:11px 15px;
           margin:12px 0; font-size:12px; line-height:1.6; }}
</style></head><body>
<h1>SCREENER TÀI CHÍNH</h1>
<p class="sub">SỐ LIỆU TỐT  ·  VÀ BẰNG CHỨNG RẰNG SỐ LIỆU ĐÓ LÀ THẬT</p>

<div class="box">
  Hệ thống chỉ hỏi hai câu, và cố ý giữ chúng tách rời nhau.<br><br>
  <b style="color:#1E40AF;">Nhóm A — Tài chính có tốt không?</b> Sinh lời trên vốn, biên lợi nhuận,
  tăng trưởng, đòn bẩy.<br>
  <b style="color:#C2410C;">Nhóm B — Số liệu đó có thật không?</b> Mỗi tiêu chí nhóm B bác bỏ một cách
  cụ thể mà số liệu nhóm A có thể bị làm đẹp.<br><br>
  Hai điểm <b>không được cộng gộp</b>. Một mã đạt 6/6 tài chính nhưng 1/6 kiểm chứng là trường hợp
  nguy hiểm nhất, và nếu gộp lại thành 7/12 thì nó sẽ trông y hệt một mã trung bình đều — nên hệ thống
  hiển thị hai điểm cạnh nhau và gắn cờ đỏ cho trường hợp đó.
</div>

<h2>① NHÓM A — TÀI CHÍNH TỐT ({N_FIN} tiêu chí)</h2>
<p style="color:#64748B;font-size:12px;margin:0 0 10px 0;">
  Toàn bộ lấy từ TradingView nên có dữ liệu cho mọi thị trường.
  Nhóm <b>Financial Services</b> được xử lý riêng: ROIC thay bằng ROE &gt; 12%, và D/E bị bỏ qua
  vì tiền gửi khách hàng hạch toán là nợ.
</p>
<table>
  <tr><th style="width:92px;">Tiêu chí</th><th style="width:110px;">Trường dữ liệu</th>
      <th style="width:120px;">Ngưỡng</th><th>Ý nghĩa</th></tr>
  {fin_rows}
</table>

<h2>② NHÓM B — CHỨNG MINH LÀ THỰC ({N_REAL} tiêu chí)</h2>
<p style="color:#64748B;font-size:12px;margin:0 0 10px 0;">
  Năm tiêu chí đầu lấy từ TradingView. Riêng <b>Bền vững 5N</b> cần bật <b>yfinance Moat</b>
  và chỉ có dữ liệu cho thị trường Mỹ — khi thiếu, nó để trống chứ không thay bằng số liệu một kỳ,
  vì lấy một kỳ mà gọi là bền vững thì chính là số ảo.
</p>
<table>
  <tr><th style="width:92px;">Tiêu chí</th><th style="width:110px;">Trường dữ liệu</th>
      <th style="width:120px;">Ngưỡng</th><th>Ý nghĩa &amp; điều nó bác bỏ</th></tr>
  {real_rows}
</table>

<h2>③ TỪNG TIÊU CHÍ NHÓM B BÁC BỎ ĐIỀU GÌ</h2>
<table>
  <tr><th style="width:150px;">Số liệu đẹp</th><th style="width:270px;">Có thể ảo vì</th>
      <th>Tiêu chí kiểm chứng tương ứng</th></tr>
  {pair_rows}
</table>

<h2>④ TÍN HIỆU TỔNG HỢP</h2>
<table>
  <tr><th style="width:150px;">Tín hiệu</th><th style="width:170px;">Điều kiện</th><th>Ý nghĩa</th></tr>
  {sig_rows}
</table>
<p class="tip">A = tỷ lệ đạt nhóm Tài chính · B = tỷ lệ đạt nhóm Kiểm chứng.</p>

<h2>⑤ CÁCH CHẤM ĐIỂM</h2>
<ul>
  <li><b>Chấm theo tỷ lệ, không theo mẫu số cố định.</b> Điểm là <i>số tiêu chí đạt / số tiêu chí thực sự
      áp dụng được</i>. Mã thiếu dữ liệu và nhóm Tài chính vì thế không bị phạt oan.</li>
  <li><b>✓ đạt · ✗ không đạt · — không áp dụng hoặc thiếu dữ liệu.</b> Dấu gạch ngang không bị tính vào
      mẫu số, nên nó khác hẳn dấu ✗.</li>
  <li><b>Dưới 3 tiêu chí áp dụng được thì không xếp hạng</b> — hệ thống trả về ⚪ THIẾU DỮ LIỆU thay vì
      đoán bừa từ vài con số lẻ.</li>
  <li><b>Sắp xếp cột điểm theo tỷ lệ</b>, nên 4/4 đứng trên 5/6 chứ không ngược lại.</li>
</ul>

<h2>⑥ NHỮNG THỨ CỐ Ý KHÔNG ĐƯA VÀO ĐIỂM</h2>
<ul>
  <li><b>Hướng thị trường</b> — giống hệt nhau cho mọi mã trong một lần quét, nên nó không phân biệt được
      mã nào với mã nào; đưa vào điểm chỉ làm trôi ngưỡng tín hiệu theo thị trường. Nay hiển thị riêng
      trên Dashboard như thông tin bối cảnh.</li>
  <li><b>Động lượng giá và sức mạnh tương đối</b> — là hành vi giá, không phải chất lượng tài chính.</li>
  <li><b>Định giá (P/E, PEG)</b> — vẫn hiển thị để tham khảo nhưng không chấm điểm, vì hệ thống trả lời
      câu hỏi «doanh nghiệp này có tốt và có thật không», không phải «giá này có rẻ không».</li>
  <li><b>Hệ số nhân Conviction</b> — trước đây nhân điểm với chất lượng moat, mà moat lại suy ra từ ROE
      và biên gộp vốn đã được chấm rồi. Đó là đếm ba lần cùng một thứ.</li>
</ul>

<h2>⑦ CÁCH DÙNG</h2>
<ul>
  <li><b>SCAN</b> — lấy top N cổ phiếu theo vốn hóa từ TradingView và chấm cả hai nhóm.</li>
  <li><b>Tab Tài chính</b> — {N_FIN} tiêu chí nhóm A cùng số liệu gốc.</li>
  <li><b>Tab Kiểm chứng</b> — {N_REAL} tiêu chí nhóm B cùng các tỷ số dùng để kiểm chứng.
      Cả hai tab đều hiện đồng thời hai điểm và tín hiệu gộp.</li>
  <li><b>Dashboard</b> — bảng <b style="color:#B91C1C;">CỜ ĐỎ</b> nêu rõ từng mã chưa chứng minh được
      ở điểm nào; bảng <b>tỷ lệ đạt</b> cho biết tiêu chí nào đang quá lỏng hoặc quá chặt.</li>
  <li><b>yfinance Moat</b> — bật trước khi quét để có tiêu chí Bền vững 5 năm. Chậm hơn đáng kể
      (khoảng 0,3 giây mỗi mã) và chỉ có dữ liệu cho thị trường Mỹ.</li>
  <li><b>Click vào một dòng</b> — xem chi tiết từng tiêu chí; rê chuột lên ô ✓/✗ để đọc giải thích
      và giá trị thực tế.</li>
</ul>

<div class="box" style="background:#FFF6E5;border-left-color:#C2410C;">
  <b style="color:#C2410C;">Giới hạn cần biết.</b> Các ngưỡng ở đây được chọn theo chuẩn mực phổ biến,
  chưa được kiểm định trên dữ liệu lịch sử của chính hệ thống này — chương trình không lưu dữ liệu theo
  thời điểm nên chưa thể kiểm thử lùi. Ngưỡng biên lợi nhuận dùng chung cho mọi ngành nên vẫn thiên lệch
  về phía công nghệ và y tế. Và hệ thống chỉ trả lời «mua gì», không có quy tắc bán hay cắt lỗ.
  Hãy đọc điểm số như một thứ tự ưu tiên để nghiên cứu, không phải một khuyến nghị đầu tư.
</div>
</body></html>"""



# ─────────────────────────────────────────────────────────────────────────────
# AI Settings — khung giao diện.
# Phần đăng nhập OAuth mới dựng vỏ, chưa nối vào luồng xác thực thật:
# _on_login() hiện chỉ đổi trạng thái hiển thị để kiểm tra bố cục.
# ─────────────────────────────────────────────────────────────────────────────
class AISettingsDialog(QDialog):
    STATE_OUT = "chua_dang_nhap"
    STATE_IN  = "da_dang_nhap"

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("AI Settings")
        self.setFixedSize(460, 340)
        self._state   = self.STATE_OUT
        self._account = ""

        self.setStyleSheet(f"""
            QDialog {{ background:{BG}; }}
            *       {{ background:transparent; color:{TEXT1};
                       font-family:'Segoe UI',sans-serif; }}
        """)

        root = QVBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(0)

        # ── Tiêu đề ────────────────────────────────────────────────────────
        head = QWidget(); head.setFixedHeight(52)
        head.setStyleSheet(f"background:{SURFACE}; border-bottom:1px solid {BORDER};")
        hh = QHBoxLayout(head); hh.setContentsMargins(20, 0, 20, 0)
        t = QLabel("AI  SETTINGS")
        t.setStyleSheet(f"color:{TEXT1}; font-size:12px; font-weight:700; letter-spacing:3px;")
        hh.addWidget(t); hh.addStretch()
        root.addWidget(head)

        # ── Thân ───────────────────────────────────────────────────────────
        body = QWidget(); body.setStyleSheet(f"background:{BG};")
        bv = QVBoxLayout(body)
        bv.setContentsMargins(20, 18, 20, 18)
        bv.setSpacing(10)

        sec = QLabel("TÀI KHOẢN")
        sec.setStyleSheet(f"color:{TEXT3}; font-size:9px; font-weight:700; letter-spacing:2px;")
        bv.addWidget(sec)

        card = QWidget()
        self._card = card
        card.setStyleSheet(f"background:{SURFACE}; border:1px solid {BORDER}; border-radius:5px;")
        cv = QVBoxLayout(card)
        cv.setContentsMargins(16, 14, 16, 14)
        cv.setSpacing(10)

        row = QHBoxLayout(); row.setSpacing(10)
        self._dot = QLabel("●")
        self._dot.setFixedWidth(12)
        self._lbl_state = QLabel()
        self._lbl_state.setWordWrap(True)
        self._lbl_state.setStyleSheet(f"color:{TEXT1}; font-size:12px;")
        row.addWidget(self._dot)
        row.addWidget(self._lbl_state, stretch=1)
        cv.addLayout(row)

        self._btn_login = QPushButton()
        self._btn_login.setFixedHeight(34)
        self._btn_login.setCursor(Qt.PointingHandCursor)
        self._btn_login.clicked.connect(self._on_login)
        cv.addWidget(self._btn_login)

        bv.addWidget(card)

        note = QLabel(
            "Đăng nhập bằng OAuth — app mở trình duyệt để bạn xác nhận, "
            "không bao giờ yêu cầu bạn dán mật khẩu hay API key vào đây.")
        note.setWordWrap(True)
        note.setStyleSheet(f"color:{TEXT3}; font-size:10px; line-height:150%;")
        bv.addWidget(note)

        bv.addSpacing(4)
        sec2 = QLabel("TÍNH NĂNG")
        sec2.setStyleSheet(f"color:{TEXT3}; font-size:9px; font-weight:700; letter-spacing:2px;")
        bv.addWidget(sec2)

        soon = QLabel(
            "Các tính năng AI sẽ được bổ sung sau khi phần đăng nhập hoàn tất.")
        soon.setWordWrap(True)
        soon.setStyleSheet(
            f"color:{TEXT3}; font-size:11px; font-style:italic;"
            f" background:{PANEL}; border-radius:4px; padding:10px 12px;")
        bv.addWidget(soon)

        bv.addStretch()
        root.addWidget(body, stretch=1)

        # ── Chân ───────────────────────────────────────────────────────────
        foot = QWidget(); foot.setFixedHeight(46)
        foot.setStyleSheet(f"background:{SURFACE}; border-top:1px solid {BORDER};")
        fh = QHBoxLayout(foot); fh.setContentsMargins(20, 0, 20, 0)
        hint = QLabel("Esc  để đóng")
        hint.setStyleSheet(f"color:{TEXT3}; font-size:9px; letter-spacing:1px;")
        btn_close = QPushButton("ĐÓNG")
        btn_close.setFixedSize(84, 28)
        btn_close.setCursor(Qt.PointingHandCursor)
        btn_close.setStyleSheet(
            f"QPushButton {{ background:{BORDER2}; color:{TEXT1}; border:none;"
            f" border-radius:3px; font-size:10px; font-weight:700; letter-spacing:1.5px; }}"
            f"QPushButton:hover {{ background:{BLUE}; color:#FFFFFF; }}")
        btn_close.clicked.connect(self.close)
        fh.addWidget(hint); fh.addStretch(); fh.addWidget(btn_close)
        root.addWidget(foot)

        self._render()

    # ── Trạng thái ─────────────────────────────────────────────────────────
    def _render(self):
        if self._state == self.STATE_IN:
            self._dot.setStyleSheet(f"color:{GREEN}; font-size:13px;")
            self._lbl_state.setText(
                f"Đã đăng nhập{('  ·  ' + self._account) if self._account else ''}")
            self._btn_login.setText("ĐĂNG XUẤT")
            self._btn_login.setStyleSheet(
                f"QPushButton {{ background:{BORDER2}; color:{TEXT1}; border:none;"
                f" border-radius:4px; font-size:11px; font-weight:700; letter-spacing:1px; }}"
                f"QPushButton:hover {{ background:{RED}; color:#FFFFFF; }}")
        else:
            self._dot.setStyleSheet(f"color:{TEXT3}; font-size:13px;")
            self._lbl_state.setText("Chưa đăng nhập")
            self._btn_login.setText("ĐĂNG NHẬP CLAUDE  (OAuth)")
            self._btn_login.setStyleSheet(
                f"QPushButton {{ background:{BLUE}; color:#FFFFFF; border:none;"
                f" border-radius:4px; font-size:11px; font-weight:700; letter-spacing:1px; }}"
                f"QPushButton:hover {{ background:{BLUE_HV}; }}")

    def _on_login(self):
        """Chưa nối OAuth — mới chỉ đổi trạng thái để xem bố cục.

        Khi làm thật, chỗ này sẽ mở trình duyệt tới trang uỷ quyền, chạy một
        máy chủ callback cục bộ để nhận authorization code, đổi lấy token và
        lưu vào kho bảo mật của hệ điều hành. Không lưu token vào mã nguồn
        hay vào file cấu hình dạng văn bản thuần.
        """
        if self._state == self.STATE_IN:
            self._state, self._account = self.STATE_OUT, ""
        else:
            self._state, self._account = self.STATE_IN, "(chưa nối OAuth)"
        self._render()

    def keyPressEvent(self, event):
        if event.key() == Qt.Key_Escape:
            self.close()
        else:
            super().keyPressEvent(event)


class ExportWorker(QThread):
    progress = Signal(int, str)   # (0-100, message)
    done     = Signal(str)        # output path
    failed   = Signal(str)        # error message

    def __init__(self, df, path, market, top, use_yf):
        super().__init__()
        self._df     = df
        self._path   = path
        self._market = market
        self._top    = top
        self._use_yf = use_yf

    def run(self):
        try:
            def cb(pct, msg=""):
                self.progress.emit(pct, msg)
            write_excel(self._df, self._path, self._market, self._top,
                        use_yf=self._use_yf, progress_cb=cb)
            self.done.emit(self._path)
        except Exception as e:
            self.failed.emit(str(e))


# Main window
# ─────────────────────────────────────────────────────────────────────────────
class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Fundamental Screener")
        self.resize(1700, 960)
        self._is_dark     = False
        globals().update(LIGHT_THEME)
        self._df          = None
        self._worker      = None
        self._ticker_wkr  = None
        self._spin_idx    = 0
        self._spin_msg    = ""
        self._spin_tmr    = QTimer(self)
        self._spin_tmr.timeout.connect(self._tick)
        self._lookup_tmr  = QTimer(self)
        self._lookup_tmr.setSingleShot(True)
        self._lookup_tmr.timeout.connect(self._do_lookup)

        QShortcut(QKeySequence(Qt.Key_F1), self).activated.connect(self._show_help)

        self.setStyleSheet(f"""
            QMainWindow {{ background:{BG}; }}
            * {{ background:transparent; color:{TEXT1}; }}
            QScrollBar:vertical {{ background:{SURFACE}; width:8px; border-radius:4px; }}
            QScrollBar::handle:vertical {{ background:{BORDER2}; border-radius:4px; }}
            QScrollBar:horizontal {{ background:{SURFACE}; height:8px; border-radius:4px; }}
            QScrollBar::handle:horizontal {{ background:{BORDER2}; border-radius:4px; }}
            QTableWidget {{ gridline-color:{BORDER}; outline:none; }}
            QTableWidget::item:selected {{ background:{BORDER2}; }}
            QToolTip {{
                background:{SURFACE}; color:{TEXT1};
                border:1px solid {BORDER2}; font-size:10px; padding:4px;
            }}
        """)

        root = QWidget()
        self.setCentralWidget(root)
        vbox = QVBoxLayout(root)
        vbox.setContentsMargins(0, 0, 0, 0)
        vbox.setSpacing(0)

        vbox.addWidget(self._build_header())
        vbox.addWidget(self._sep())
        vbox.addWidget(self._build_controls())
        vbox.addWidget(self._sep())

        # Right: chart panel (always visible)
        self._chart_panel = ChartPanel()

        # Left: tab widget (CAN SLIM | Quality Compounder)
        self._tabs_w = self._build_tabs()
        self._tabs_w.currentChanged.connect(self._on_tab_changed)

        # Horizontal splitter: tabs | chart
        h_split = QSplitter(Qt.Horizontal)
        h_split.setStyleSheet(
            f"QSplitter::handle {{ background:{BORDER}; width:2px; }}")
        h_split.addWidget(self._tabs_w)
        h_split.addWidget(self._chart_panel)
        h_split.setSizes([920, 580])
        # Dashboard là tab mặc định (index 0) → ẩn chart panel ngay từ đầu
        self._chart_panel.setVisible(False)
        vbox.addWidget(h_split, stretch=1)

        vbox.addWidget(self._sep())
        vbox.addWidget(self._build_status())

        # Load cache từ lần scan trước (nếu có)
        QTimer.singleShot(200, self._load_cache)

    # ── Section builders ──────────────────────────────────────────────────────

    def _build_header(self):
        self._header_w = QWidget(); self._header_w.setFixedHeight(50)
        self._header_w.setStyleSheet(f"background:{SURFACE};")
        h = QHBoxLayout(self._header_w); h.setContentsMargins(22, 0, 16, 0); h.setSpacing(0)
        lbl = QLabel()
        lbl.setTextFormat(Qt.RichText)
        lbl.setText(
            f'<span style="color:{TEXT1};font-size:14px;font-weight:700;letter-spacing:4px">FUNDAMENTAL</span>'
            f'<span style="color:{BLUE};font-size:14px;font-weight:700;letter-spacing:4px"> SCREENER</span>'
        )
        self._title_lbl = lbl
        h.addWidget(lbl); h.addStretch()
        self._tag_lbl = QLabel("Catalyst · TradingView · yfinance")
        self._tag_lbl.setStyleSheet(f"color:{TEXT3}; font-size:9px; letter-spacing:2px;")
        h.addWidget(self._tag_lbl)
        h.addSpacing(10)

        self._btn_theme = QPushButton("☀")
        self._btn_theme.setFixedSize(28, 28)
        self._btn_theme.setToolTip("Switch Light / Dark")
        self._btn_theme.setCursor(Qt.PointingHandCursor)
        self._btn_theme.clicked.connect(self._toggle_theme)
        h.addWidget(self._btn_theme)
        h.addSpacing(6)

        self._btn_ai = QPushButton("✦")
        self._btn_ai.setFixedSize(28, 28)
        self._btn_ai.setToolTip("AI Settings")
        self._btn_ai.setCursor(Qt.PointingHandCursor)
        self._btn_ai.clicked.connect(self._show_ai_settings)
        h.addWidget(self._btn_ai)
        h.addSpacing(6)

        self._btn_help = QPushButton("?")
        self._btn_help.setFixedSize(28, 28)
        self._btn_help.setToolTip("Help  (F1)")
        self._btn_help.setCursor(Qt.PointingHandCursor)
        self._btn_help.clicked.connect(self._show_help)
        h.addWidget(self._btn_help)
        self._refresh_icon_btns()
        return self._header_w

    def _build_controls(self):
        w = QWidget(); w.setFixedHeight(118)
        self._controls_w = w
        w.setStyleSheet(f"background:{BG};")
        v = QVBoxLayout(w); v.setContentsMargins(22, 6, 22, 6); v.setSpacing(5)

        # Row 1: scan controls
        row1 = QHBoxLayout(); row1.setSpacing(10)

        lbl_market = QLabel("MARKET")
        self._lbl_market = lbl_market
        lbl_market.setStyleSheet(
            f"color:{TEXT1}; font-size:9px; font-weight:600; letter-spacing:1.5px;")
        self._market = QComboBox()
        self._market.addItems(["AMERICA", "NASDAQ", "NYSE", "EURONEXT", "HONG_KONG", "VIETNAM"])
        self._market.setFixedHeight(32)
        self._market.setStyleSheet(self._combo_style())

        lbl_top = QLabel("TOP")
        self._lbl_top = lbl_top
        lbl_top.setStyleSheet(
            f"color:{TEXT1}; font-size:9px; font-weight:600; letter-spacing:1.5px;")
        self._top_spin = QSpinBox()
        self._top_spin.setRange(50, 1000); self._top_spin.setValue(300); self._top_spin.setSingleStep(50)
        self._top_spin.setFixedSize(60, 32)
        self._top_spin.setButtonSymbols(QSpinBox.ButtonSymbols.NoButtons)
        self._top_spin.setStyleSheet(self._spinbox_style())

        self._spin_minus = QPushButton("−")
        self._spin_plus  = QPushButton("+")
        for btn in (self._spin_minus, self._spin_plus):
            btn.setFixedSize(24, 32)
            btn.setCursor(Qt.PointingHandCursor)
            btn.setStyleSheet(self._spin_btn_style())
        self._spin_minus.clicked.connect(lambda: self._top_spin.setValue(self._top_spin.value() - self._top_spin.singleStep()))
        self._spin_plus.clicked.connect( lambda: self._top_spin.setValue(self._top_spin.value() + self._top_spin.singleStep()))

        self._use_yf = QCheckBox("yfinance Moat")
        self._use_yf.setChecked(True)
        self._use_yf.setStyleSheet(
            f"color:{TEXT1}; font-size:10px; font-family:'Segoe UI',sans-serif;")

        self._btn_scan = QPushButton("▶  SCAN")
        self._btn_scan.setFixedSize(110, 32)
        self._btn_scan.setCursor(Qt.PointingHandCursor)
        self._btn_scan.setStyleSheet(self._btn_style(BLUE, BLUE_HV))
        self._btn_scan.clicked.connect(self._on_scan)

        self._btn_export = QPushButton("💾  Export")
        self._btn_export.setFixedSize(95, 32)
        self._btn_export.setCursor(Qt.PointingHandCursor)
        self._btn_export.setEnabled(False)
        self._btn_export.setStyleSheet(self._btn_style("#2E5C2E", "#1A4C1A"))
        self._btn_export.clicked.connect(self._on_export)

        row1.addWidget(lbl_market); row1.addWidget(self._market)
        row1.addWidget(lbl_top)
        row1.addWidget(self._spin_minus)
        row1.addWidget(self._top_spin)
        row1.addWidget(self._spin_plus)
        row1.addWidget(self._use_yf)
        row1.addStretch()
        row1.addWidget(self._btn_scan)
        row1.addWidget(self._btn_export)
        v.addLayout(row1)

        # Row 2: ticker lookup + table filter
        row2 = QHBoxLayout(); row2.setSpacing(10)

        lbl_tk = QLabel("TICKER")
        self._lbl_tk = lbl_tk
        lbl_tk.setStyleSheet(
            f"color:{TEXT1}; font-size:9px; font-weight:600; letter-spacing:1.5px;")
        self._ticker_inp = QLineEdit()
        self._ticker_inp.setPlaceholderText("e.g. AAPL  · NVDA · MSFT")
        self._ticker_inp.setFixedHeight(28); self._ticker_inp.setMaxLength(12)
        self._ticker_inp.setStyleSheet(self._input_style(font="Consolas,monospace", size="13px", ls="3px"))
        self._ticker_inp.returnPressed.connect(self._do_lookup)
        self._ticker_inp.textChanged.connect(self._on_ticker_typed)

        self._btn_lookup = QPushButton("LOOKUP")
        self._btn_lookup.setFixedSize(80, 28)
        self._btn_lookup.setCursor(Qt.PointingHandCursor)
        self._btn_lookup.setStyleSheet(self._btn_style(BLUE, BLUE_HV, size="9px"))
        self._btn_lookup.clicked.connect(self._do_lookup)

        lbl_filter = QLabel("FILTER")
        self._lbl_filter = lbl_filter
        lbl_filter.setStyleSheet(
            f"color:{TEXT1}; font-size:9px; font-weight:600; letter-spacing:1.5px;")
        self._filter_inp = QLineEdit()
        self._filter_inp.setPlaceholderText("Search ticker or company name…")
        self._filter_inp.setFixedHeight(28)
        self._filter_inp.setStyleSheet(self._input_style())
        self._filter_inp.textChanged.connect(self._apply_filter)

        self._lbl_count = QLabel("0 / 0  rows")
        self._lbl_count.setStyleSheet(
            f"color:{TEXT3}; font-size:9px; letter-spacing:1px;")

        row2.addWidget(lbl_tk); row2.addWidget(self._ticker_inp)
        row2.addWidget(self._btn_lookup)
        sep = QFrame(); sep.setFrameShape(QFrame.VLine)
        sep.setStyleSheet(f"background:{BORDER}; max-width:1px; margin:2px 4px;")
        row2.addWidget(sep)
        row2.addWidget(lbl_filter); row2.addWidget(self._filter_inp)
        row2.addStretch(); row2.addWidget(self._lbl_count)
        v.addLayout(row2)

        # Row 3: quick filter chips
        row3 = QHBoxLayout(); row3.setSpacing(10)
        lbl_qf = QLabel("QUICK FILTER")
        lbl_qf.setStyleSheet(
            f"color:{TEXT1}; font-size:9px; font-weight:600; letter-spacing:1.5px;")
        self._lbl_qf = lbl_qf

        self._chk_sb   = QCheckBox("🟢  TỐT & THỰC")
        self._chk_moat = QCheckBox("🔴  CỜ ĐỎ")
        self._chk_1y   = QCheckBox("💵  FCF ≥ lợi nhuận")
        self._chk_52w  = QCheckBox("🏰  Bền vững 5 năm")
        self._chip_checks = (self._chk_sb, self._chk_moat, self._chk_1y, self._chk_52w)

        for chk in self._chip_checks:
            chk.setStyleSheet(self._chip_style())
            chk.stateChanged.connect(self._apply_filter)

        self._btn_clr = QPushButton("✕")
        self._btn_clr.setFixedSize(24, 22)
        self._btn_clr.setToolTip("Clear all filters")
        self._btn_clr.setCursor(Qt.PointingHandCursor)
        self._btn_clr.setStyleSheet(self._btn_style(PANEL, BORDER2, "9px"))
        self._btn_clr.clicked.connect(self._clear_chips)

        sep2 = QFrame(); sep2.setFrameShape(QFrame.VLine)
        sep2.setStyleSheet(f"background:{BORDER}; max-width:1px; margin:2px 4px;")
        row3.addWidget(lbl_qf); row3.addWidget(sep2)
        for chk in self._chip_checks:
            row3.addWidget(chk)
        row3.addStretch(); row3.addWidget(self._btn_clr)
        v.addLayout(row3)
        return w

    def _build_table(self):
        w = QWidget(); w.setStyleSheet(f"background:{BG};")
        v = QVBoxLayout(w); v.setContentsMargins(0, 0, 0, 0); v.setSpacing(0)
        self._table = QTableWidget(0, len(TABLE_COLS))
        self._table.setStyleSheet(f"""
            QTableWidget {{
                background:{BG}; alternate-background-color:{PANEL};
                font-size:11px; font-family:'Segoe UI',sans-serif;
                selection-background-color:{BORDER2};
                border:none;
            }}
        """)
        self._table.setAlternatingRowColors(True)
        self._table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self._table.setSelectionMode(QAbstractItemView.SingleSelection)
        self._table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self._table.setSortingEnabled(True)
        self._table.verticalHeader().setVisible(False)
        self._table.verticalHeader().setDefaultSectionSize(22)
        self._table.setHorizontalScrollMode(QAbstractItemView.ScrollPerPixel)
        self._table.setVerticalScrollMode(QAbstractItemView.ScrollPerPixel)
        _hdr = ColoredHeader(self._table)
        self._table.setHorizontalHeader(_hdr)

        headers = [c[0] for c in TABLE_COLS]
        self._table.setHorizontalHeaderLabels(headers)
        _apply_header_tips(self._table, TABLE_COLS)
        for i, (_, _, w_px, _) in enumerate(TABLE_COLS):
            self._table.setColumnWidth(i, w_px)

        self._color_header_cols()

        self._table.itemSelectionChanged.connect(self._on_row_selected)
        v.addWidget(self._table)
        return w

    def _build_detail(self):
        self._detail = DetailCard()
        return self._detail

    # ── Tab widget ────────────────────────────────────────────────────────────

    def _build_tabs(self):
        tabs = QTabWidget()
        tabs.setDocumentMode(True)
        tabs.setStyleSheet(self._tab_style())

        # Tab 0 — Dashboard
        tabs.addTab(self._build_dashboard_panel(), "  📊 Dashboard  ")

        # Tab 1 — CAN SLIM
        self._cs_w = QWidget(); self._cs_w.setStyleSheet(f"background:{BG};")
        cs_v = QVBoxLayout(self._cs_w)
        cs_v.setContentsMargins(0, 0, 0, 0); cs_v.setSpacing(0)
        self._cs_vsplit = QSplitter(Qt.Vertical)
        self._cs_vsplit.setStyleSheet(f"QSplitter::handle {{ background:{BORDER}; height:2px; }}")
        self._cs_vsplit.addWidget(self._build_table())
        self._cs_vsplit.addWidget(self._build_detail())
        self._cs_vsplit.setSizes([580, 180])
        cs_v.addWidget(self._cs_vsplit)
        tabs.addTab(self._cs_w, "  Catalyst  ")

        # Tab 2 — Quality Compounder
        self._qc_w = QWidget(); self._qc_w.setStyleSheet(f"background:{BG};")
        qc_v = QVBoxLayout(self._qc_w)
        qc_v.setContentsMargins(0, 0, 0, 0); qc_v.setSpacing(0)
        self._qc_vsplit = QSplitter(Qt.Vertical)
        self._qc_vsplit.setStyleSheet(f"QSplitter::handle {{ background:{BORDER}; height:2px; }}")
        self._qc_vsplit.addWidget(self._build_qc_table())
        self._qc_vsplit.addWidget(self._build_qc_detail())
        self._qc_vsplit.setSizes([580, 180])
        qc_v.addWidget(self._qc_vsplit)
        tabs.addTab(self._qc_w, "  Quality Compounder  ")

        # Tab 3 — My Portfolio
        self._portfolio_panel = PortfolioPanel()
        tabs.addTab(self._portfolio_panel, "  💼 Portfolio  ")

        return tabs

    def _build_dashboard_panel(self):
        w = QWidget(); w.setStyleSheet(f"background:{BG};")
        lay = QVBoxLayout(w); lay.setContentsMargins(0,0,0,0); lay.setSpacing(0)
        self._dash_browser = QTextBrowser()
        self._dash_browser.setOpenExternalLinks(False)
        self._dash_browser.setStyleSheet(f"""
            QTextBrowser {{
                background:{BG}; color:#DCE4EE;
                border:none; padding:16px 20px;
                font-family:'Segoe UI',sans-serif;
            }}
        """)
        self._dash_browser.setHtml(self._dash_placeholder())
        lay.addWidget(self._dash_browser)
        return w

    def _dash_placeholder(self):
        return f"""<html><body style="background:{BG};color:#4A6080;
            font-family:'Segoe UI',sans-serif;text-align:center;padding-top:120px;">
            <p style="font-size:48px;margin:0;">📊</p>
            <p style="font-size:18px;font-weight:700;color:#3D5070;margin:12px 0 6px 0;">
              Dashboard chưa có dữ liệu</p>
            <p style="font-size:13px;color:#2A3A50;">
              Nhấn <b style="color:#3D8EF0;">▶ SCAN</b> để tải dữ liệu và xem Dashboard</p>
        </body></html>"""

    def _tab_style(self):
        return f"""
            QTabWidget::pane {{
                border:none;
                background:{BG};
            }}
            QTabBar {{
                background:{SURFACE};
            }}
            QTabBar::tab {{
                background:{SURFACE}; color:{TEXT3};
                border:none;
                border-right:1px solid {BORDER};
                border-bottom:2px solid transparent;
                padding:9px 22px; font-size:9px; font-weight:700;
                letter-spacing:2px; font-family:'Segoe UI',sans-serif;
                min-width:80px;
            }}
            QTabBar::tab:selected {{
                background:{BG}; color:{TEXT1};
                border-bottom:2px solid {BLUE};
            }}
            QTabBar::tab:hover:!selected {{
                color:{TEXT2};
            }}
        """

    # ── QC table ──────────────────────────────────────────────────────────────

    def _build_qc_table(self):
        w = QWidget(); w.setStyleSheet(f"background:{BG};")
        v = QVBoxLayout(w); v.setContentsMargins(0, 0, 0, 0); v.setSpacing(0)
        self._qc_table = QTableWidget(0, len(REAL_COLS))
        self._qc_table.setStyleSheet(f"""
            QTableWidget {{
                background:{BG}; alternate-background-color:{PANEL};
                font-size:11px; font-family:'Segoe UI',sans-serif;
                selection-background-color:{BORDER2};
                border:none;
            }}
        """)
        self._qc_table.setAlternatingRowColors(True)
        self._qc_table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self._qc_table.setSelectionMode(QAbstractItemView.SingleSelection)
        self._qc_table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self._qc_table.setSortingEnabled(True)
        self._qc_table.verticalHeader().setVisible(False)
        self._qc_table.verticalHeader().setDefaultSectionSize(22)
        self._qc_table.setHorizontalScrollMode(QAbstractItemView.ScrollPerPixel)
        self._qc_table.setVerticalScrollMode(QAbstractItemView.ScrollPerPixel)
        _hdr = ColoredHeader(self._qc_table)
        self._qc_table.setHorizontalHeader(_hdr)
        self._qc_table.setHorizontalHeaderLabels([c[0] for c in REAL_COLS])
        _apply_header_tips(self._qc_table, REAL_COLS)
        for i, (_, _, w_px, _) in enumerate(REAL_COLS):
            self._qc_table.setColumnWidth(i, w_px)
        self._color_qc_header_cols()
        self._qc_table.itemSelectionChanged.connect(self._on_qc_row_selected)
        v.addWidget(self._qc_table)
        return w

    def _build_qc_detail(self):
        self._qc_detail = QualityDetailCard()
        return self._qc_detail

    def _color_qc_header_cols(self):
        dark = self._is_dark
        _qc_bg = "#0D1F10" if dark else "#E8F5E9"
        _qc_fg = "#5EC472" if dark else "#276749"
        _mo_bg = "#1A0D2E" if dark else "#F3E8FF"
        _mo_fg = "#B07ED4" if dark else "#7C3AED"
        _fu_bg = "#0D1828" if dark else "#EEF2FB"
        _fu_fg = "#7EB8D4" if dark else "#1E40AF"
        for i, (_, key, _, _) in enumerate(REAL_COLS):
            item = self._qc_table.horizontalHeaderItem(i)
            if not item:
                continue
            if key.startswith("REAL_"):
                item.setBackground(QColor(_qc_bg)); item.setForeground(QColor(_qc_fg))
            elif key == "Moat Score":
                item.setBackground(QColor(_mo_bg)); item.setForeground(QColor(_mo_fg))
            elif key in ("Gross Margin%", "Op Margin%", "ROIC%", "FCF_Margin%", "D/E",
                         "Net Cash ($B)", "Current Ratio", "FCF/NI", "EPS_NI_Gap", "ROE_ROIC_Gap",
                         "NI_Trend_Gap", "LN quý ($B)", "Nợ ròng/FCF"):
                item.setBackground(QColor(_fu_bg)); item.setForeground(QColor(_fu_fg))
        self._qc_table.horizontalHeader().update()

    def _populate_qc_table(self, df: pd.DataFrame):
        self._qc_data = {}
        self._qc_table.setSortingEnabled(False)
        self._qc_table.setRowCount(0)
        self._qc_table.setRowCount(len(df))

        # Columns that are %-type and get green/red coloring
        PCT_GREEN_RED = {"Gross Margin%", "Op Margin%", "ROIC%", "1Y%", "FCF_Margin%",
                         "NI Qtr%", "Rev Qtr%", "NI_Trend_Gap"}
        is_vn = self._market.currentText().lower() == "vietnam"

        for ri, (_, row) in enumerate(df.iterrows()):
            rd = row.to_dict()
            ticker = rd.get("Ticker", "")
            if ticker:
                self._qc_data[ticker] = rd

            for ci, (_, key, _, align) in enumerate(REAL_COLS):
                val  = rd.get(key)
                cell = None

                if key == "_no_":
                    cell = NumItem(ri + 1, "int")
                    cell.setForeground(QColor(TEXT1))
                elif key == "Ticker":
                    cell = QTableWidgetItem(str(val) if val else "")
                    cell.setFont(QFont("Consolas", 9, QFont.Bold))
                    cell.setForeground(QColor("#00BFFF"))
                elif key in ("Tên Công Ty", "Sector"):
                    cell = QTableWidgetItem(str(val) if val else "")
                    cell.setForeground(QColor(TEXT1))
                elif key == "Moat Score":
                    cell = QTableWidgetItem(str(val) if val else "—")
                    if val in MOAT_SCORE_STYLE:
                        fg, bg = MOAT_SCORE_STYLE[val]
                        cell.setForeground(QColor(f"#{fg}"))
                        cell.setBackground(QColor(f"#{bg}"))
                elif key == "Signal":
                    cell = QTableWidgetItem(str(val) if val else "")
                    if val in SIGNAL_COLORS:
                        fg, bg = SIGNAL_COLORS[val]
                        cell.setForeground(QColor(fg))
                        cell.setBackground(QColor(bg))
                elif key == "Cảnh báo":
                    txt  = str(val) if val else ""
                    cell = QTableWidgetItem(txt if txt else "—")
                    nflag = len(txt.split()) if txt else 0
                    if nflag >= 2:
                        cell.setForeground(QColor("#9C0006")); cell.setBackground(QColor("#FFC7CE"))
                    elif nflag == 1:
                        cell.setForeground(QColor("#7A3E00")); cell.setBackground(QColor("#FFE0C2"))
                    else:
                        cell.setForeground(QColor(TEXT3))
                    cell.setFont(QFont("Consolas", 8, QFont.Bold))

                elif key == "FIN_Disp":
                    cell = _score_cell(rd.get("FIN_Score"), rd.get("FIN_App"))
                elif key == "REAL_Disp":
                    cell = _score_cell(rd.get("REAL_Score"), rd.get("REAL_App"))
                elif key.startswith("REAL_"):
                    if val is True:
                        cell = QTableWidgetItem("✓")
                        cell.setForeground(QColor("#1A5C2B"))
                        cell.setBackground(QColor("#C6EFCE"))
                    elif val is False:
                        cell = QTableWidgetItem("✗")
                        cell.setForeground(QColor("#9C0006"))
                        cell.setBackground(QColor("#FFC7CE"))
                    else:
                        cell = QTableWidgetItem("—")
                        cell.setForeground(QColor(TEXT3))
                elif key == "Price ($)":
                    cell = NumItem(val, "vnd" if is_vn else "price")
                elif key == "MCap ($B)":
                    cell = NumItem(val, "mcap_vnd" if is_vn else "mcap")

                elif key in ("Doanh thu ($B)", "Lợi nhuận ($B)", "FCF ($B)",
                             "DThu quý ($B)", "LN quý ($B)"):
                    cell = NumItem(val, "bil_vnd" if is_vn else "bil")
                    if isinstance(val, float) and val == val:
                        # Lợi nhuận hoặc dòng tiền âm là thông tin quan trọng — tô đỏ ngay
                        cell.setForeground(QColor(RED if val < 0 else TEXT1))
                elif key == "Net Cash ($B)":
                    cell = NumItem(val)
                    if isinstance(val, float) and val == val:
                        if val > 5:
                            cell.setForeground(QColor("#1A5C2B"))
                            cell.setBackground(QColor("#C6EFCE"))
                        elif val > 0:
                            cell.setForeground(QColor("#1A5C2B"))
                        elif val > -5:
                            cell.setForeground(QColor("#7D6608"))
                        else:
                            cell.setForeground(QColor(RED))
                    else:
                        cell.setForeground(QColor(TEXT3))

                elif key == "FCF_Margin%":
                    cell = NumItem(val, "pct")
                    if isinstance(val, float):
                        if val >= 25:
                            cell.setForeground(QColor("#1A5C2B"))
                            cell.setBackground(QColor("#D5F5E3"))
                        elif val >= 15:
                            cell.setForeground(QColor("#1A5C2B"))
                        elif val < 0:
                            cell.setForeground(QColor(RED))
                        else:
                            cell.setForeground(QColor(TEXT3))

                elif key in PCT_GREEN_RED:
                    cell = NumItem(val, "pct")
                    if isinstance(val, float):
                        cell.setForeground(QColor(GREEN) if val > 0 else
                                           QColor(RED)   if val < 0 else
                                           QColor(TEXT3))
                else:
                    cell = NumItem(val)

                if cell:
                    cell.setTextAlignment(int(align))
                    _tip = _tip_for(key)
                    if _tip:
                        cell.setToolTip(_tip)
                    self._qc_table.setItem(ri, ci, cell)

        self._qc_table.setSortingEnabled(True)
        self._apply_qc_filter()

    def _recolor_qc_table(self):
        PCT_GREEN_RED = {"Gross Margin%", "Op Margin%", "ROIC%", "1Y%", "FCF_Margin%",
                         "NI Qtr%", "Rev Qtr%", "NI_Trend_Gap"}
        for ri in range(self._qc_table.rowCount()):
            for ci, (_, key, _, _) in enumerate(REAL_COLS):
                item = self._qc_table.item(ri, ci)
                if not item:
                    continue
                if key in ("_no_", "Tên Công Ty", "Sector"):
                    item.setForeground(QColor(TEXT1))
                elif key in PCT_GREEN_RED and hasattr(item, "_num"):
                    v = item._num
                    if v != float("-inf"):
                        item.setForeground(QColor(GREEN) if v > 0 else
                                           QColor(RED)   if v < 0 else
                                           QColor(TEXT3))

    def _on_qc_row_selected(self):
        rows = self._qc_table.selectedItems()
        if not rows:
            return
        ri = self._qc_table.currentRow()
        ticker_item = self._qc_table.item(ri, 1)
        if not ticker_item:
            return
        ticker = ticker_item.text()
        rd = getattr(self, "_qc_data", {}).get(ticker)
        if rd:
            self._qc_detail.show_row(rd)
            self._ticker_inp.setText(ticker)
            self._chart_panel.load_ticker(ticker)

    def _on_tab_changed(self, idx: int):
        # Tab 0 = Dashboard → full width, ẩn chart panel
        is_dash = (idx == 0)
        if hasattr(self, "_chart_panel"):
            self._chart_panel.setVisible(not is_dash)
        self._apply_filter()

    def _apply_qc_filter(self):
        text      = self._filter_inp.text().strip().lower()
        only_sb   = self._chk_sb.isChecked()
        only_moat = self._chk_moat.isChecked()
        only_1y   = self._chk_1y.isChecked()
        only_52w  = self._chk_52w.isChecked()
        any_chip  = only_sb or only_moat or only_1y or only_52w
        qd        = getattr(self, "_qc_data", {})

        visible = 0
        total   = self._qc_table.rowCount()
        for ri in range(total):
            show = True
            if text:
                t_item = self._qc_table.item(ri, 1)
                n_item = self._qc_table.item(ri, 2)
                t_txt  = (t_item.text() if t_item else "").lower()
                n_txt  = (n_item.text() if n_item else "").lower()
                show   = (text in t_txt) or (text in n_txt)
            if show and any_chip:
                ticker = (self._qc_table.item(ri, 1).text()
                          if self._qc_table.item(ri, 1) else "")
                rd = qd.get(ticker, {})
                if only_sb and rd.get("Signal") != "🟢 TỐT & THỰC":
                    show = False
                if show and only_moat and rd.get("Signal") != "🔴 CỜ ĐỎ":
                    show = False
                if show and only_1y:
                    v = rd.get("FCF/NI")
                    if v is None or (isinstance(v, float) and pd.isna(v)) or v < 0.8:
                        show = False
                if show and only_52w and rd.get("Moat Score") not in (
                        "WIDE  ★★★", "NARROW ★★"):
                    show = False
            self._qc_table.setRowHidden(ri, not show)
            if show:
                visible += 1
        self._lbl_count.setText(f"{visible} / {total}  rows")

    def _build_status(self):
        w = QWidget(); w.setFixedHeight(28)
        self._status_w = w
        w.setStyleSheet(f"background:{SURFACE};")
        h = QHBoxLayout(w); h.setContentsMargins(22, 0, 22, 0)
        self._dot = QLabel("●"); self._dot.setFixedWidth(14)
        self._dot.setStyleSheet(f"color:{TEXT3}; font-size:7px;")
        self.status_lbl = QLabel("Ready  ·  Enter market & top N, then click SCAN")
        self.status_lbl.setStyleSheet(
            f"color:{TEXT3}; font-size:10px; font-family:'Segoe UI',sans-serif;")
        self._export_bar = QProgressBar()
        self._export_bar.setFixedSize(180, 10)
        self._export_bar.setTextVisible(False)
        self._export_bar.setRange(0, 100)
        self._export_bar.setValue(0)
        self._export_bar.setVisible(False)
        self._export_bar.setStyleSheet("""
            QProgressBar { background:#333; border-radius:4px; }
            QProgressBar::chunk { background:#4AE06A; border-radius:4px; }
        """)
        h.addWidget(self._dot); h.addWidget(self.status_lbl)
        h.addWidget(self._export_bar)
        h.addStretch()

        self._ver_lbl = QLabel("TradingView  ·  yfinance  ·  PySide6")
        self._ver_lbl.setStyleSheet(f"color:{TEXT3}; font-size:9px; letter-spacing:1px;")
        h.addWidget(self._ver_lbl)
        return w

    # ── Theme ─────────────────────────────────────────────────────────────────

    def _refresh_icon_btns(self):
        _s = f"""
            QPushButton {{
                background:{BORDER2}; color:{TEXT1};
                border:none; border-radius:14px;
                font-size:12px; font-weight:700;
                font-family:'Segoe UI',sans-serif;
            }}
            QPushButton:hover {{ background:{BLUE}; color:#FFFFFF; }}
        """
        self._btn_theme.setStyleSheet(_s)
        self._btn_ai.setStyleSheet(_s)
        self._btn_help.setStyleSheet(_s)

    def _color_header_cols(self):
        dark = self._is_dark
        _cs_bg  = "#0D1F10" if dark else "#E8F5E9"
        _cs_fg  = "#5EC472" if dark else "#276749"
        _mo_bg  = "#1A0D2E" if dark else "#F3E8FF"
        _mo_fg  = "#B07ED4" if dark else "#7C3AED"
        _fu_bg  = "#0D1828" if dark else "#EEF2FB"
        _fu_fg  = "#7EB8D4" if dark else "#1E40AF"
        _keys = [c[1] for c in TABLE_COLS]
        for i, key in enumerate(_keys):
            item = self._table.horizontalHeaderItem(i)
            if not item:
                continue
            if key.startswith("FIN_") or key in ("FIN_Disp", "REAL_Disp", "Signal"):
                item.setBackground(QColor(_cs_bg)); item.setForeground(QColor(_cs_fg))
            elif key == "Moat Score":
                item.setBackground(QColor(_mo_bg)); item.setForeground(QColor(_mo_fg))
            elif key in ("NI Growth%", "Rev Growth%", "NI Qtr%", "Rev Qtr%",
                         "Gross Margin%", "Op Margin%", "ROIC%", "ROE%", "D/E", "P/E", "1Y%"):
                item.setBackground(QColor(_fu_bg)); item.setForeground(QColor(_fu_fg))
        self._table.horizontalHeader().update()

    def _toggle_theme(self):
        self._is_dark = not self._is_dark
        theme = DARK_THEME if self._is_dark else LIGHT_THEME
        globals().update(theme)
        self._btn_theme.setText("☀" if self._is_dark else "🌙")
        self._apply_theme()

    def _apply_theme(self):
        dark = self._is_dark
        # ── ColoredHeader ──
        ColoredHeader._DEF_BG   = QColor("#0A1628" if dark else "#E2E8F0")
        ColoredHeader._DEF_FG   = QColor("#7EB8D4" if dark else "#2D3748")
        ColoredHeader._BORDER_B = QColor(BLUE)
        ColoredHeader._BORDER_R = QColor(BORDER)
        self._color_header_cols()

        # ── Main window stylesheet (scrollbars, tooltip) ──
        self.setStyleSheet(f"""
            QMainWindow {{ background:{BG}; }}
            * {{ background:transparent; color:{TEXT1}; }}
            QScrollBar:vertical {{ background:{SURFACE}; width:8px; border-radius:4px; }}
            QScrollBar::handle:vertical {{ background:{BORDER2}; border-radius:4px; }}
            QScrollBar:horizontal {{ background:{SURFACE}; height:8px; border-radius:4px; }}
            QScrollBar::handle:horizontal {{ background:{BORDER2}; border-radius:4px; }}
            QTableWidget {{ gridline-color:{BORDER}; outline:none; }}
            QTableWidget::item:selected {{ background:{BORDER2}; }}
            QToolTip {{
                background:{SURFACE}; color:{TEXT1};
                border:1px solid {BORDER2}; font-size:10px; padding:4px;
            }}
        """)
        # ── Header bar ──
        self._header_w.setStyleSheet(f"background:{SURFACE};")
        self._tag_lbl.setStyleSheet(f"color:{TEXT3}; font-size:9px; letter-spacing:2px;")
        self._title_lbl.setText(
            f'<span style="color:{TEXT1};font-size:14px;font-weight:700;letter-spacing:4px">FUNDAMENTAL</span>'
            f'<span style="color:{BLUE};font-size:14px;font-weight:700;letter-spacing:4px"> SCREENER</span>'
        )
        self._refresh_icon_btns()

        # ── Controls bar ──
        _lbl_s = f"color:{TEXT1}; font-size:9px; font-weight:600; letter-spacing:1.5px;"
        self._controls_w.setStyleSheet(f"background:{BG};")
        self._lbl_market.setStyleSheet(_lbl_s)
        self._lbl_top.setStyleSheet(_lbl_s)
        self._lbl_tk.setStyleSheet(_lbl_s)
        self._lbl_filter.setStyleSheet(_lbl_s)
        self._market.setStyleSheet(self._combo_style())
        self._top_spin.setStyleSheet(self._spinbox_style())
        self._spin_minus.setStyleSheet(self._spin_btn_style())
        self._spin_plus.setStyleSheet(self._spin_btn_style())
        self._use_yf.setStyleSheet(f"color:{TEXT1}; font-size:10px; font-family:'Segoe UI',sans-serif;")
        self._btn_scan.setStyleSheet(self._btn_style(BLUE, BLUE_HV))
        self._btn_export.setStyleSheet(self._btn_style("#2E5C2E", "#1A4C1A"))
        self._ticker_inp.setStyleSheet(self._input_style(font="Consolas,monospace", size="13px", ls="3px"))
        self._btn_lookup.setStyleSheet(self._btn_style(BLUE, BLUE_HV, size="9px"))
        self._filter_inp.setStyleSheet(self._input_style())
        self._lbl_count.setStyleSheet(f"color:{TEXT3}; font-size:9px; letter-spacing:1px;")
        self._lbl_qf.setStyleSheet(f"color:{TEXT1}; font-size:9px; font-weight:600; letter-spacing:1.5px;")
        for chk in self._chip_checks:
            chk.setStyleSheet(self._chip_style())
        self._btn_clr.setStyleSheet(self._btn_style(PANEL, BORDER2, "9px"))

        # ── Table ──
        self._table.setStyleSheet(f"""
            QTableWidget {{
                background:{BG}; alternate-background-color:{PANEL};
                font-size:11px; font-family:'Segoe UI',sans-serif;
                selection-background-color:{BORDER2};
                border:none;
            }}
        """)
        self._recolor_table()
        # ── Detail card ──
        self._detail.apply_theme()
        # ── QC table ──
        self._qc_table.setStyleSheet(f"""
            QTableWidget {{
                background:{BG}; alternate-background-color:{PANEL};
                font-size:11px; font-family:'Segoe UI',sans-serif;
                selection-background-color:{BORDER2};
                border:none;
            }}
        """)
        self._color_qc_header_cols()
        self._recolor_qc_table()
        # ── QC detail card ──
        self._qc_detail.apply_theme()
        # ── Tab widget ──
        self._tabs_w.setStyleSheet(self._tab_style())
        # ── Tab containers + splitters ──
        if hasattr(self, "_cs_w"):
            self._cs_w.setStyleSheet(f"background:{BG};")
        if hasattr(self, "_qc_w"):
            self._qc_w.setStyleSheet(f"background:{BG};")
        if hasattr(self, "_cs_vsplit"):
            self._cs_vsplit.setStyleSheet(f"QSplitter::handle {{ background:{BORDER}; height:2px; }}")
        if hasattr(self, "_qc_vsplit"):
            self._qc_vsplit.setStyleSheet(f"QSplitter::handle {{ background:{BORDER}; height:2px; }}")
        # ── Chart panel ──
        self._chart_panel.set_theme(dark)
        # ── Portfolio ──
        self._portfolio_panel.apply_theme()
        # ── Dashboard — regenerate HTML theo theme mới ──
        self._dash_browser.setStyleSheet(f"""
            QTextBrowser {{
                background:{BG}; color:{TEXT1};
                border:none; padding:16px 20px;
                font-family:'Segoe UI',sans-serif;
            }}
        """)
        if hasattr(self, "_df") and self._df is not None and len(self._df) > 0:
            self._update_dashboard(self._df)
        # ── Status bar ──
        self._status_w.setStyleSheet(f"background:{SURFACE};")
        self._ver_lbl.setStyleSheet(f"color:{TEXT3}; font-size:9px; letter-spacing:1px;")

    # ── Style helpers ─────────────────────────────────────────────────────────

    def _sep(self):
        f = QFrame(); f.setFrameShape(QFrame.HLine); f.setFixedHeight(1)
        f.setStyleSheet(f"background:{BORDER}; border:none;")
        return f

    def _input_style(self, font="'Segoe UI',sans-serif", size="11px", ls="0"):
        return f"""
            QLineEdit {{
                background:{INPUT_BG}; color:{TEXT1};
                border:1px solid {BORDER}; border-radius:3px;
                padding:0 10px; font-size:{size};
                font-family:{font}; letter-spacing:{ls};
                selection-background-color:{BLUE};
            }}
            QLineEdit:focus {{ border-color:{BLUE}; }}
        """

    def _spin_btn_style(self):
        return f"""
            QPushButton {{
                background:{SURFACE}; color:{TEXT1};
                border:1px solid {BORDER}; border-radius:3px;
                font-size:15px; font-weight:600;
                padding:0;
            }}
            QPushButton:hover   {{ background:{BORDER};  color:{TEXT1}; }}
            QPushButton:pressed {{ background:{BORDER2}; color:{TEXT1}; }}
        """

    def _spinbox_style(self):
        return f"""
            QSpinBox {{
                background:{INPUT_BG}; color:{TEXT1};
                border:1px solid {BORDER}; border-radius:3px;
                padding:0 4px 0 6px; font-size:11px;
                font-family:'Segoe UI',sans-serif;
            }}
            QSpinBox:focus {{ border-color:{BLUE}; }}
            QSpinBox::up-button {{
                subcontrol-origin:border; subcontrol-position:top right;
                background:{SURFACE}; border-left:1px solid {BORDER};
                border-top-right-radius:3px;
                width:18px; height:16px;
            }}
            QSpinBox::up-button:hover   {{ background:{BORDER}; }}
            QSpinBox::up-button:pressed {{ background:{BORDER2}; }}
            QSpinBox::down-button {{
                subcontrol-origin:border; subcontrol-position:bottom right;
                background:{SURFACE}; border-left:1px solid {BORDER};
                border-bottom-right-radius:3px;
                width:18px; height:16px;
            }}
            QSpinBox::down-button:hover   {{ background:{BORDER}; }}
            QSpinBox::down-button:pressed {{ background:{BORDER2}; }}
            QSpinBox::up-arrow   {{ color:{TEXT1}; width:10px; height:10px; }}
            QSpinBox::down-arrow {{ color:{TEXT1}; width:10px; height:10px; }}
        """

    def _combo_style(self):
        return f"""
            QComboBox {{
                background:{INPUT_BG}; color:{TEXT1};
                border:1px solid {BORDER}; border-radius:3px;
                padding:0 10px; font-size:11px; min-width:100px;
            }}
            QComboBox:focus {{ border-color:{BLUE}; }}
            QComboBox::drop-down {{ border:none; width:20px; }}
            QComboBox QAbstractItemView {{
                background:{SURFACE}; color:{TEXT1};
                selection-background-color:{BORDER2};
                border:1px solid {BORDER};
            }}
        """

    def _chip_style(self, checked=False):
        fg = BLUE if checked else TEXT2
        border = BLUE if checked else BORDER
        return (f"QCheckBox {{ color:{fg}; font-size:9px; font-weight:600;"
                f" font-family:'Segoe UI',sans-serif; spacing:4px;"
                f" padding:2px 8px; border:1px solid {border}; border-radius:10px; }}"
                f"QCheckBox:checked {{ color:{BLUE}; border-color:{BLUE}; }}"
                f"QCheckBox:unchecked {{ color:{TEXT2}; border-color:{BORDER}; }}")

    def _btn_style(self, bg, hv, size="10px"):
        return f"""
            QPushButton {{
                background:{bg}; color:#FFFFFF;
                border:none; border-radius:3px;
                font-size:{size}; font-weight:700; letter-spacing:1.5px;
                font-family:'Segoe UI',sans-serif;
            }}
            QPushButton:hover   {{ background:{hv}; }}
            QPushButton:pressed {{ background:{hv}; }}
            QPushButton:disabled {{
                background:{PANEL}; color:{TEXT3};
            }}
        """

    # ── Spinner ───────────────────────────────────────────────────────────────

    def _set_status(self, text, color=TEXT3):
        self.status_lbl.setText(text)
        self.status_lbl.setStyleSheet(
            f"color:{color}; font-size:10px; font-family:'Segoe UI',sans-serif;")
        self._dot.setStyleSheet(f"color:{color}; font-size:7px;")

    def _tick(self):
        self._spin_idx = (self._spin_idx + 1) % len(SPINNER)
        self._set_status(f"{SPINNER[self._spin_idx]}  {self._spin_msg}", AMBER)

    # ── Scan logic ────────────────────────────────────────────────────────────

    def _on_scan(self):
        if self._worker and self._worker.isRunning():
            self._worker.cancel()
            self._btn_scan.setText("▶  SCAN")
            return

        market = self._market.currentText()
        top    = self._top_spin.value()
        use_yf = self._use_yf.isChecked()

        self._btn_scan.setText("⏹  Cancel")
        self._btn_export.setEnabled(False)
        self._spin_msg = f"Scanning {market.upper()} top {top}…"
        self._spin_tmr.start(80)

        self._worker = ScanWorker(market, top, use_yf)
        self._worker.progress.connect(lambda t: setattr(self, "_spin_msg", t))
        self._worker.ticker_update.connect(self._on_ticker_progress)
        self._worker.market_ctx.connect(self._on_market_ctx)
        self._worker.done.connect(self._on_scan_done)
        self._worker.failed.connect(self._on_scan_failed)
        self._worker.start()

    def _on_market_ctx(self, market_ok, index_1y):
        """Bối cảnh thị trường — chỉ để hiển thị trên Dashboard, không chấm điểm."""
        self._market_ok = market_ok
        self._index_1y  = index_1y

    def _on_ticker_progress(self, idx, total, ticker, score):
        self._spin_msg = f"Moat [{idx}/{total}] {ticker} → {score}"

    # ── Scan cache ────────────────────────────────────────────────────────────

    def _cache_path(self):
        base = os.path.dirname(os.path.abspath(sys.argv[0]))
        return os.path.join(base, "scan_cache.pkl")

    def _save_cache(self, df, market: str, top: int):
        import pickle
        try:
            payload = {
                "df": df,
                "market": market,
                "top": top,
                "saved_at": pd.Timestamp.now(),
            }
            with open(self._cache_path(), "wb") as f:
                pickle.dump(payload, f)
        except Exception:
            pass

    def _load_cache(self):
        import pickle
        try:
            path = self._cache_path()
            if not os.path.exists(path):
                return
            with open(path, "rb") as f:
                payload = pickle.load(f)
            df       = payload["df"]
            market   = payload.get("market", "")
            top      = payload.get("top", len(df))
            saved_at = payload.get("saved_at", None)
            age_str  = ""
            if saved_at is not None:
                delta = pd.Timestamp.now() - saved_at
                h = int(delta.total_seconds() // 3600)
                m = int((delta.total_seconds() % 3600) // 60)
                age_str = f"{h}h {m}m" if h else f"{m}m"

            self._populate_table(df)
            self._populate_qc_table(df)
            self._df = df
            self._update_dashboard(df)
            self._btn_export.setEnabled(True)
            # Khôi phục market/top settings nếu có
            idx = self._market.findText(market, Qt.MatchFixedString)
            if idx >= 0:
                self._market.setCurrentIndex(idx)
            self._top_spin.setValue(top)
            self._set_status(
                f"📂  Cache loaded  ·  {len(df)} stocks  ·  "
                f"Lưu {age_str} trước  ·  Nhấn SCAN để cập nhật mới",
                TEXT2
            )
            self._portfolio_panel.refresh_scan(df)
        except Exception:
            pass   # Cache lỗi hoặc không tương thích → bỏ qua

    def _on_scan_done(self, df: pd.DataFrame):
        self._spin_tmr.stop()
        self._btn_scan.setText("▶  SCAN")
        self._btn_export.setEnabled(True)
        self._populate_table(df)
        self._populate_qc_table(df)

        self._df = df
        self._update_dashboard(df)
        self._tabs_w.setCurrentIndex(0)
        _sig = df["Signal"] if "Signal" in df.columns else pd.Series(dtype=str)
        self._set_status(
            f"✓  {len(df)} mã  ·  "
            f"TỐT & THỰC: {(_sig=='🟢 TỐT & THỰC').sum()}  "
            f"KHÁ: {(_sig=='🔵 KHÁ').sum()}  "
            f"CỜ ĐỎ: {(_sig=='🔴 CỜ ĐỎ').sum()}  "
            f"THEO DÕI: {(_sig=='🟡 THEO DÕI').sum()}",
            GREEN
        )
        # Lưu cache để lần sau không cần scan lại
        self._save_cache(df, self._market.currentText(), self._top_spin.value())
        # Cập nhật Portfolio với giá/điểm mới nhất
        self._portfolio_panel.refresh_scan(df)

    def _on_scan_failed(self, msg):
        self._spin_tmr.stop()
        self._btn_scan.setText("▶  SCAN")
        self._set_status(f"✕  {msg}", RED)

    # ── Table population ──────────────────────────────────────────────────────

    def _populate_table(self, df: pd.DataFrame):
        self._ticker_data = {row.get("Ticker", ""): row.to_dict()
                             for _, row in df.iterrows() if row.get("Ticker")}
        self._table.setSortingEnabled(False)
        self._table.setRowCount(0)
        self._table.setRowCount(len(df))

        PCT_KEYS = {"NI Growth%","Rev Growth%","EPS Growth%","FCF_Margin%",
                    "NI Qtr%","Rev Qtr%","NI_Trend_Gap",
                    "Gross Margin%","Net Margin%","ROE%",
                    "1W%","1M%","3M%","6M%","1Y%"}
        is_vn = self._market.currentText().lower() == "vietnam"

        for ri, (_, row) in enumerate(df.iterrows()):
            for ci, (hdr, key, _, align) in enumerate(TABLE_COLS):
                val = row.get(key)
                cell = None

                if key == "_no_":
                    cell = NumItem(ri + 1, "int")
                    cell.setForeground(QColor(TEXT1))

                elif key == "Ticker":
                    cell = QTableWidgetItem(str(val) if val else "")
                    cell.setFont(QFont("Consolas", 9, QFont.Bold))
                    cell.setForeground(QColor("#00BFFF"))

                elif key == "Tên Công Ty":
                    cell = QTableWidgetItem(str(val) if val else "")
                    cell.setForeground(QColor(TEXT1))

                elif key == "Sector":
                    cell = QTableWidgetItem(str(val) if val else "")
                    cell.setForeground(QColor(TEXT1))

                elif key == "Moat Score":
                    cell = QTableWidgetItem(str(val) if val else "—")
                    if val in MOAT_SCORE_STYLE:
                        fg, bg = MOAT_SCORE_STYLE[val]
                        cell.setForeground(QColor(f"#{fg}"))
                        cell.setBackground(QColor(f"#{bg}"))

                elif key == "Signal":
                    cell = QTableWidgetItem(str(val) if val else "")
                    if val in SIGNAL_COLORS:
                        fg, bg = SIGNAL_COLORS[val]
                        cell.setForeground(QColor(fg))
                        cell.setBackground(QColor(bg))

                elif key == "Cảnh báo":
                    txt  = str(val) if val else ""
                    cell = QTableWidgetItem(txt if txt else "—")
                    nflag = len(txt.split()) if txt else 0
                    if nflag >= 2:
                        cell.setForeground(QColor("#9C0006")); cell.setBackground(QColor("#FFC7CE"))
                    elif nflag == 1:
                        cell.setForeground(QColor("#7A3E00")); cell.setBackground(QColor("#FFE0C2"))
                    else:
                        cell.setForeground(QColor(TEXT3))
                    cell.setFont(QFont("Consolas", 8, QFont.Bold))

                elif key == "FIN_Disp":
                    cell = _score_cell(row.get("FIN_Score"), row.get("FIN_App"))
                elif key == "REAL_Disp":
                    cell = _score_cell(row.get("REAL_Score"), row.get("REAL_App"))

                elif key.startswith("FIN_"):
                    if val is True:
                        cell = QTableWidgetItem("✓")
                        cell.setForeground(QColor("#1A5C2B")); cell.setBackground(QColor("#C6EFCE"))
                    elif val is False:
                        cell = QTableWidgetItem("✗")
                        cell.setForeground(QColor("#9C0006")); cell.setBackground(QColor("#FFC7CE"))
                    else:
                        cell = QTableWidgetItem("—"); cell.setForeground(QColor(TEXT3))

                elif key == "PEG":
                    cell = NumItem(val)
                    if isinstance(val, float) and val == val:
                        if val <= 1.0:
                            cell.setForeground(QColor("#1A5C2B"))
                            cell.setBackground(QColor("#C6EFCE"))
                        elif val <= 1.5:
                            cell.setForeground(QColor("#1A3A5C"))
                            cell.setBackground(QColor("#DDEEFF"))
                        elif val <= 2.5:
                            cell.setForeground(QColor("#7D6608"))
                        else:
                            cell.setForeground(QColor(RED))
                    else:
                        cell.setForeground(QColor(TEXT3))


                elif key == "Price ($)":
                    cell = NumItem(val, "vnd" if is_vn else "price")
                elif key == "MCap ($B)":
                    cell = NumItem(val, "mcap_vnd" if is_vn else "mcap")

                elif key in ("Doanh thu ($B)", "Lợi nhuận ($B)", "FCF ($B)",
                             "DThu quý ($B)", "LN quý ($B)"):
                    cell = NumItem(val, "bil_vnd" if is_vn else "bil")
                    if isinstance(val, float) and val == val:
                        # Lợi nhuận hoặc dòng tiền âm là thông tin quan trọng — tô đỏ ngay
                        cell.setForeground(QColor(RED if val < 0 else TEXT1))
                elif key in PCT_KEYS:
                    cell = NumItem(val, "pct")
                    if isinstance(val, float):
                        cell.setForeground(QColor(GREEN) if val > 0 else
                                           QColor(RED)   if val < 0 else
                                           QColor(TEXT3))
                else:
                    cell = NumItem(val)

                if cell:
                    cell.setTextAlignment(int(align))
                    _tip = _tip_for(key)
                    if _tip:
                        cell.setToolTip(_tip)
                    self._table.setItem(ri, ci, cell)

        self._table.setSortingEnabled(True)
        # Đổi header label theo market
        _keys = [c[1] for c in TABLE_COLS]
        if is_vn:
            if "Price ($)" in _keys:
                self._table.horizontalHeaderItem(_keys.index("Price ($)")).setText("Giá (₫)")
            if "MCap ($B)" in _keys:
                self._table.horizontalHeaderItem(_keys.index("MCap ($B)")).setText("MCap (tỷ₫)")
        else:
            if "Price ($)" in _keys:
                self._table.horizontalHeaderItem(_keys.index("Price ($)")).setText("Price($)")
            if "MCap ($B)" in _keys:
                self._table.horizontalHeaderItem(_keys.index("MCap ($B)")).setText("MCap($B)")
        self._apply_filter()

    # ── Dashboard ─────────────────────────────────────────────────────────────

    def _update_dashboard(self, df):
        self._dash_browser.setHtml(self._generate_dashboard_html(df))

    def _generate_dashboard_html(self, df):  # noqa: C901
        from datetime import datetime

        market = self._market.currentText()
        is_vn  = market.lower() == "vietnam"
        dark   = getattr(self, "_is_dark", False)

        # Màu nhấn phải sáng hơn hẳn trên nền tối, nếu không chữ màu bị chìm.
        if dark:
            BG_="#0B0E18"; CARD="#131928"; BRD="#2C3A50"; T1="#E6ECF5"; T2="#9AA9BC"; T3="#6B7C91"
            HDR="#16233A"; ROW_A="#101728"
            G="#3DD97F"; R="#FF6B5E"; A_="#FFB443"; B="#5FA8FF"; O="#FF9448"
        else:
            BG_="#FFFFFF"; CARD="#F7FAFC"; BRD="#D1D9E6"; T1="#1A202C"; T2="#4A5568"; T3="#8494A8"
            HDR="#E4EAF5"; ROW_A="#F0F4FA"
            G="#15803D"; R="#B91C1C"; A_="#B45309"; B="#1D4ED8"; O="#C2410C"

        if df is None or df.empty:
            return f'<body style="background:{BG_};color:{T2};font-family:Segoe UI;padding:40px;">Chưa có dữ liệu.</body>'

        n = len(df)
        sig = df["Signal"] if "Signal" in df.columns else pd.Series([""] * n)

        def cnt(s):
            return int((sig == s).sum())

        n_good  = cnt("🟢 TỐT & THỰC")
        n_flag  = cnt("🔴 CỜ ĐỎ")
        n_ok    = cnt("🔵 KHÁ")
        n_weak  = cnt("🟠 THẬT NHƯNG YẾU")
        _w = df["Cảnh báo"] if "Cảnh báo" in df.columns else pd.Series([""] * n)
        n_warn = int((_w.fillna("") != "").sum())
        # Mã đạt chất lượng nhưng giá đang bị đánh xuống — tình huống đáng chú ý
        # nhất với người mua khi giá giảm.
        n_gw   = int(((sig == "🟢 TỐT & THỰC") & (_w.fillna("") != "")).sum())
        avg_fin  = df["FIN_Pct"].mean() * 100  if "FIN_Pct"  in df.columns else 0
        avg_real = df["REAL_Pct"].mean() * 100 if "REAL_Pct" in df.columns else 0

        def kpi(val, label, color):
            return (f'<td style="background:{CARD};border:1px solid {BRD};border-radius:6px;'
                    f'padding:12px 6px;text-align:center;width:12.5%;">'
                    f'<div style="color:{color};font-size:23px;font-weight:700;'
                    f'font-family:Consolas,monospace;">{val}</div>'
                    f'<div style="color:{T2};font-size:10px;font-weight:600;letter-spacing:.5px;'
                    f'margin-top:3px;">{label}</div></td>')

        kpis = ('<table style="width:100%;border-spacing:7px;border-collapse:separate;"><tr>'
                + kpi(n, "TỔNG SỐ MÃ", T1)
                + kpi(n_good, "TỐT &amp; THỰC", G)
                + kpi(n_flag, "CỜ ĐỎ", R)
                + kpi(n_ok, "KHÁ", B)
                + kpi(f"{avg_fin:.0f}%", "TB TÀI CHÍNH", B)
                + kpi(f"{avg_real:.0f}%", "TB THỰC", O)
                + kpi(n_warn, "CÓ CẢNH BÁO GIÁ", A_)
                + kpi(n_gw, "TỐT NHƯNG GIÁ GIẢM", O)
                + '</tr></table>')

        # ── Bối cảnh thị trường: hiển thị, KHÔNG cộng vào điểm ────────────
        mk = getattr(self, "_market_ok", None)
        ix = getattr(self, "_index_1y", None)
        if mk is None:
            mk_html = ""
        else:
            mc, mt = (G, "THUẬN LỢI — chỉ số trên MA50 &amp; MA200") if mk else \
                     (R, "BẤT LỢI — chỉ số dưới MA50 hoặc MA200")
            ix_s = f"  ·  Chỉ số 1 năm: {ix:+.1f}%" if isinstance(ix, (int, float)) else ""
            mk_html = (f'<div style="background:{CARD};border:1px solid {BRD};border-left:3px solid {mc};'
                       f'border-radius:5px;padding:9px 14px;margin:14px 0;">'
                       f'<span style="color:{mc};font-weight:700;font-size:12px;">Thị trường: {mt}</span>'
                       f'<span style="color:{T2};font-size:11px;">{ix_s}</span><br>'
                       f'<span style="color:{T3};font-size:10px;">Đây là bối cảnh chung, giống nhau cho mọi mã '
                       f'— không cộng vào điểm để tránh làm trôi ngưỡng tín hiệu theo thị trường.</span></div>')

        def sec(title, sub, color):
            return (f'<div style="margin:22px 0 8px 0;">'
                    f'<span style="color:{color};font-size:14px;font-weight:700;">{title}</span>'
                    f'<span style="color:{T3};font-size:11px;margin-left:10px;">{sub}</span></div>')

        def table(headers, rows, widths=None):
            if not rows:
                return f'<div style="color:{T3};font-size:11px;font-style:italic;padding:8px 4px;">Không có mã nào.</div>'
            th = "".join(
                f'<th style="background:{HDR};color:{T2};font-size:10px;font-weight:700;'
                f'padding:7px 8px;text-align:left;border-bottom:1px solid {BRD};'
                f'{f"width:{widths[i]};" if widths else ""}">{h}</th>'
                for i, h in enumerate(headers))
            tr = ""
            for i, row in enumerate(rows):
                bgc = ROW_A if i % 2 else "transparent"
                tds = "".join(
                    f'<td style="padding:6px 8px;font-size:11px;color:{c or T1};'
                    f'{"font-weight:700;" if bold else ""}'
                    f'{f"background:{cbg};" if cbg else ""}">{v}</td>'
                    for v, c, bold, cbg in row)
                tr += f'<tr style="background:{bgc};">{tds}</tr>'
            return (f'<table style="width:100%;border-collapse:collapse;border:1px solid {BRD};'
                    f'border-radius:5px;overflow:hidden;"><tr>{th}</tr>{tr}</table>')

        def cell(v, color=None, bold=False, bg=None):
            return (v, color, bold, bg)

        def money(v):
            if not isinstance(v, (int, float)) or v != v:
                return "—"
            return f"{v:,.0f}₫" if is_vn else f"${v:,.2f}"

        def score_badge(g, a):
            if not a:
                return f'<span style="color:{T3};">—</span>'
            p = g / a
            c = G if p >= .8 else B if p >= .6 else A_ if p >= .4 else R
            return f'<span style="color:{c};font-weight:700;">{int(g)}/{int(a)}</span>'

        def warn_cell(r):
            """Ô cảnh báo giá — cùng bộ dấu với hai bảng chính."""
            w = r.get("Cảnh báo", "") or ""
            if not w:
                return cell('<span style="color:%s;">—</span>' % T3)
            c = R if len(w.split()) >= 2 else O
            # Dấu "<MA250" phải escape, nếu không trình dựng coi là thẻ mở và nuốt mất
            esc = w.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
            return cell('<span style="color:%s;font-weight:700;'
                        'font-family:Consolas,monospace;">%s</span>' % (c, esc))

        def rows_for(sub, extra_cols):
            out = []
            for i, (_, r) in enumerate(sub.iterrows(), 1):
                base = [cell(str(i), T3), cell(r.get("Ticker", ""), "#00BFFF", True),
                        cell(str(r.get("Tên Công Ty", ""))[:34], T1),
                        cell(str(r.get("Sector", ""))[:20], T2),
                        cell(money(r.get("Price ($)")), T1),
                        cell(score_badge(r.get("FIN_Score"), r.get("FIN_App"))),
                        cell(score_badge(r.get("REAL_Score"), r.get("REAL_App")))]
                out.append(base + extra_cols(r))
            return out

        def num(v, fmt="{:.1f}", good=None):
            if not isinstance(v, (int, float)) or v != v:
                return cell("—", T3)
            c = T1
            if good is not None:
                c = G if good(v) else R
            return cell(fmt.format(v), c)

        HDRS = ["#", "Ticker", "Công ty", "Ngành", "Giá", "T.Chính", "Thực"]

        # ── 1. TỐT & THỰC ────────────────────────────────────────────────
        good = df[sig == "🟢 TỐT & THỰC"].sort_values(
            ["REAL_Pct", "FIN_Pct"], ascending=False) if n_good else df.iloc[0:0]
        good_tbl = table(
            HDRS + ["ROIC%", "FCF/NI", "FCF Mgn%", "1Y%", "Cảnh báo"],
            rows_for(good, lambda r: [
                num(r.get("ROIC%"), "{:.1f}%"),
                num(r.get("FCF/NI"), "{:.2f}×", lambda v: v >= 0.8),
                num(r.get("FCF_Margin%"), "{:.1f}%", lambda v: v > 8),
                num(r.get("1Y%"), "{:+.1f}%", lambda v: v > 0),
                warn_cell(r),
            ]))

        # ── 2. CỜ ĐỎ — kết quả đặc trưng của hệ thống ────────────────────
        flag = df[sig == "🔴 CỜ ĐỎ"].sort_values(
            "FIN_Pct", ascending=False) if n_flag else df.iloc[0:0]

        def why(r):
            """Lý do cụ thể vì sao số liệu chưa được chứng minh."""
            bad = []
            if r.get("REAL_CASH") is False:
                v = r.get("FCF/NI")
                bad.append(f"dòng tiền chỉ bằng {v:.2f}× lợi nhuận" if isinstance(v, (int, float)) and v == v
                           else "dòng tiền không theo kịp lợi nhuận")
            if r.get("REAL_FCFM") is False:
                bad.append("biên dòng tiền mỏng")
            if r.get("REAL_BUYB") is False:
                v = r.get("EPS_NI_Gap")
                bad.append(f"EPS tăng nhanh hơn lợi nhuận {v:.0f} điểm % (mua cổ phiếu quỹ)"
                           if isinstance(v, (int, float)) and v == v else "EPS tăng nhờ buyback")
            if r.get("REAL_LEV") is False:
                v = r.get("ROE_ROIC_Gap")
                bad.append(f"ROE vượt ROIC {v:.0f} điểm % (đòn bẩy)"
                           if isinstance(v, (int, float)) and v == v else "ROE do đòn bẩy")
            if r.get("REAL_BAL") is False:
                v = r.get("Nợ ròng/FCF")
                bad.append(f"nợ ròng cần {v:.1f} năm dòng tiền mới trả hết"
                           if isinstance(v, (int, float)) and v == v
                           else "nợ ròng lớn mà dòng tiền không đủ trả")
            if r.get("REAL_TREND") is False:
                v = r.get("NI Qtr%")
                bad.append(f"quý gần nhất lợi nhuận {v:+.0f}% — đã xấu đi"
                           if isinstance(v, (int, float)) and v == v
                           else "quý gần nhất đã xấu đi")
            if r.get("REAL_SUS") is False:
                bad.append("chưa bền vững qua 5 năm")
            return "; ".join(bad) or "—"

        flag_tbl = table(
            HDRS + ["Vì sao chưa chứng minh được", "Cảnh báo"],
            rows_for(flag, lambda r: [cell(why(r), R), warn_cell(r)]),
            widths=["3%", "6%", "16%", "11%", "7%", "6%", "5%", "34%", "12%"])

        # ── 3. Cảnh báo giá — mục riêng ──────────────────────────────────
        # Chỉ báo, không tham gia chấm điểm. Sắp mã còn nguyên chất lượng lên
        # trước, vì đó là nhóm đáng đọc nhất với người mua khi giá giảm.
        _wcol = df["Cảnh báo"].fillna("") if "Cảnh báo" in df.columns else pd.Series([""] * n)
        warn_df = df[_wcol != ""].copy()
        if not warn_df.empty:
            _rank = {"🟢 TỐT & THỰC": 0, "🔵 KHÁ": 1, "🟠 THẬT NHƯNG YẾU": 2,
                     "🟡 THEO DÕI": 3, "🔴 CỜ ĐỎ": 4, "⚫ LOẠI": 5, "⚪ THIẾU DỮ LIỆU": 6}
            warn_df["_r"] = warn_df["Signal"].map(lambda x: _rank.get(x, 9))
            warn_df = warn_df.sort_values(["_r", "% từ đỉnh"], ascending=[True, True])

        def sig_cell(v):
            fg, bg = SIGNAL_COLORS.get(v, (T2, None))
            return cell(v or "—", fg, True, bg)

        warn_tbl = table(
            ["#", "Ticker", "Công ty", "Ngành", "Giá", "T.Chính", "Thực",
             "% từ đỉnh", "Giá/MA250", "Phiên", "Cảnh báo", "Tín hiệu"],
            rows_for(warn_df, lambda r: [
                num(r.get("% từ đỉnh"), "{:+.1f}%", lambda v: v > -20),
                num(r.get("Giá/MA250%"), "{:+.1f}%", lambda v: v > 0),
                num(r.get("Chg%"), "{:+.1f}%", lambda v: v > -3),
                warn_cell(r),
                sig_cell(r.get("Signal", "")),
            ]),
            widths=["3%", "6%", "17%", "11%", "7%", "6%", "5%", "8%", "8%", "7%", "10%", "12%"])

        # ── 4. Tỷ lệ đạt từng tiêu chí ───────────────────────────────────
        hit_rows = []
        for grp, keys, crit in (("A", FIN_KEYS, FIN_CRITERIA), ("B", REAL_KEYS, REAL_CRITERIA)):
            for k in keys:
                col = f"{'FIN' if grp == 'A' else 'REAL'}_{k}"
                if col not in df.columns:
                    continue
                s   = df[col]
                app = int(s.notna().sum())
                got = int((s == True).sum())  # noqa: E712
                pct = (got / app * 100) if app else 0
                # Tiêu chí gần như ai cũng đạt hoặc ai cũng trượt thì không phân loại được gì
                if app == 0:
                    note, nc = "không có dữ liệu", R
                elif pct >= 90:
                    note, nc = "quá lỏng — hầu như mã nào cũng đạt", A_
                elif pct <= 5:
                    note, nc = "quá chặt — gần như không mã nào đạt", A_
                else:
                    note, nc = "phân loại tốt", G
                bar_c = G if pct >= 60 else B if pct >= 30 else A_
                # Thanh vẽ bằng ký tự khối: engine rich-text của Qt không hỗ trợ
                # inline-block nên div lồng nhau bị đè lên chữ. Ký tự thì luôn đúng.
                filled = int(round(pct / 10))
                bar = (f'<span style="color:{bar_c};font-family:Consolas,monospace;'
                       f'font-size:12px;letter-spacing:-1px;">{"█" * filled}</span>'
                       f'<span style="color:{BRD};font-family:Consolas,monospace;'
                       f'font-size:12px;letter-spacing:-1px;">{"█" * (10 - filled)}</span>')
                hit_rows.append([
                    cell("Nhóm A" if grp == "A" else "Nhóm B", B if grp == "A" else O, True),
                    cell(crit[k]["label"], T1, True),
                    cell(f"{got}/{app}", T1),
                    cell(f"{pct:.0f}%", bar_c, True),
                    cell(bar),
                    cell(note, nc),
                ])
        hit_tbl = table(["Nhóm", "Tiêu chí", "Đạt", "Tỷ lệ", "Mức", "Nhận xét"], hit_rows,
                        widths=["8%", "17%", "9%", "7%", "14%", "45%"])

        # ── 5. Theo ngành ────────────────────────────────────────────────
        sec_rows = []
        if "Sector" in df.columns:
            for name, g in df.groupby("Sector"):
                if not name:
                    continue
                gs = g["Signal"] if "Signal" in g.columns else pd.Series(dtype=str)
                sec_rows.append([
                    cell(str(name)[:26], T1, True),
                    cell(str(len(g)), T2),
                    cell(str(int((gs == "🟢 TỐT & THỰC").sum())), G, True),
                    cell(str(int((gs == "🔴 CỜ ĐỎ").sum())), R, True),
                    num(g["FIN_Pct"].mean() * 100 if "FIN_Pct" in g.columns else None, "{:.0f}%"),
                    num(g["REAL_Pct"].mean() * 100 if "REAL_Pct" in g.columns else None, "{:.0f}%"),
                ])
            sec_rows.sort(key=lambda r: -int(r[2][0]))
        sec_tbl = table(["Ngành", "Số mã", "Tốt&Thực", "Cờ đỏ", "TB T.Chính", "TB Thực"], sec_rows)

        stamp = datetime.now().strftime("%d/%m/%Y %H:%M")
        return f"""<body style="background:{BG_};color:{T1};font-family:'Segoe UI',sans-serif;
                                margin:0;padding:18px 22px;">
  <table style="width:100%;border-collapse:collapse;margin-bottom:14px;"><tr>
    <td style="font-size:17px;font-weight:700;color:{T1};letter-spacing:.5px;">
      TÀI CHÍNH TỐT &amp; CÓ CHỨNG MINH</td>
    <td align="right" style="font-size:11px;color:{T3};">{market.upper()}  ·  {n} mã  ·  {stamp}</td>
  </tr></table>
  {kpis}
  {mk_html}

  {sec(f'🟢  TỐT &amp; THỰC — {n_good} mã', 'số liệu tài chính mạnh VÀ vượt được các phép kiểm chứng', G)}
  {good_tbl}

  {sec(f'🔴  CỜ ĐỎ — {n_flag} mã', 'số liệu tài chính ≥80% tiêu chí nhưng kiểm chứng dưới 60% — số đẹp chưa chứng minh được', R)}
  {flag_tbl}

  {sec(f'⚠️  CẢNH BÁO GIÁ — {n_warn} mã',
       f'{n_gw} mã trong số này vẫn đạt 🟢 TỐT &amp; THỰC — chất lượng còn nguyên nhưng giá đã bị đánh xuống. '
       f'Ba dấu này chỉ mô tả trạng thái giá, không tham gia chấm điểm và không dùng để lọc', A_)}
  {warn_tbl}

  {sec('📊  Tỷ lệ đạt của từng tiêu chí', 'tiêu chí quá lỏng hoặc quá chặt thì không phân loại được gì — dùng để hiệu chỉnh ngưỡng', B)}
  {hit_tbl}

  {sec('🏢  Theo ngành', 'ngành nào số liệu đẹp mà khó chứng minh nhất', T2)}
  {sec_tbl}

  <div style="margin-top:20px;color:{T3};font-size:10px;line-height:1.6;">
    Điểm chấm theo tỷ lệ trên số tiêu chí <i>thực sự áp dụng được</i> — mã thiếu dữ liệu và nhóm Tài chính
    không bị phạt bởi mẫu số cố định.<br>
    Nhóm A đo số liệu có đẹp không. Nhóm B đo số liệu đó có thật không. Điểm A cao mà điểm B thấp
    là tín hiệu cảnh báo, không phải điểm trung bình.
  </div>
</body>"""

    def _recolor_table(self):
        """Re-apply theme-sensitive foreground colors to existing cells."""
        PCT_KEYS = {"NI Growth%","Rev Growth%","EPS Growth%","FCF_Margin%",
                    "NI Qtr%","Rev Qtr%","NI_Trend_Gap",
                    "Gross Margin%","Net Margin%","ROE%",
                    "1W%","1M%","3M%","6M%","1Y%"}
        _keys = [c[1] for c in TABLE_COLS]
        for ri in range(self._table.rowCount()):
            for ci, key in enumerate(_keys):
                item = self._table.item(ri, ci)
                if not item:
                    continue
                if key in ("_no_", "Tên Công Ty", "Sector"):
                    item.setForeground(QColor(TEXT1))
                elif key in PCT_KEYS and hasattr(item, "_num"):
                    v = item._num
                    if v != float("-inf"):
                        item.setForeground(QColor(GREEN) if v > 0 else
                                           QColor(RED)   if v < 0 else
                                           QColor(TEXT3))

    def _clear_chips(self):
        for chk in self._chip_checks:
            chk.blockSignals(True)
            chk.setChecked(False)
            chk.blockSignals(False)
        self._apply_filter()

    def _apply_filter(self):
        if hasattr(self, "_tabs_w") and self._tabs_w.currentIndex() == 2:
            self._apply_qc_filter()
            return
        text      = self._filter_inp.text().strip().lower()
        only_sb   = self._chk_sb.isChecked()
        only_moat = self._chk_moat.isChecked()
        only_1y   = self._chk_1y.isChecked()
        only_52w  = self._chk_52w.isChecked()
        any_chip  = only_sb or only_moat or only_1y or only_52w
        td        = getattr(self, "_ticker_data", {})

        visible = 0
        total = self._table.rowCount()
        for ri in range(total):
            show = True
            if text:
                t_item = self._table.item(ri, 1)
                n_item = self._table.item(ri, 2)
                t_txt  = (t_item.text() if t_item else "").lower()
                n_txt  = (n_item.text() if n_item else "").lower()
                show   = (text in t_txt) or (text in n_txt)
            if show and any_chip:
                ticker = (self._table.item(ri, 1).text()
                          if self._table.item(ri, 1) else "")
                rd = td.get(ticker, {})
                if only_sb and rd.get("Signal") != "🟢 TỐT & THỰC":
                    show = False
                if show and only_moat and rd.get("Signal") != "🔴 CỜ ĐỎ":
                    show = False
                if show and only_1y:
                    v = rd.get("FCF/NI")
                    if v is None or (isinstance(v, float) and pd.isna(v)) or v < 0.8:
                        show = False
                if show and only_52w and rd.get("Moat Score") not in (
                        "WIDE  ★★★", "NARROW ★★"):
                    show = False
            self._table.setRowHidden(ri, not show)
            if show:
                visible += 1
        self._lbl_count.setText(f"{visible} / {total}  rows")

    # ── Row selection → detail card ───────────────────────────────────────────

    def _on_row_selected(self):
        rows = self._table.selectedItems()
        if not rows or self._df is None: return
        ri = self._table.currentRow()
        ticker_item = self._table.item(ri, 1)
        if not ticker_item: return
        ticker = ticker_item.text()
        match = self._df[self._df["Ticker"] == ticker]
        if not match.empty:
            self._detail.show_row(match.iloc[0].to_dict())
            self._ticker_inp.setText(ticker)
            self._chart_panel.load_ticker(ticker)

    # ── Ticker lookup ─────────────────────────────────────────────────────────

    def _on_ticker_typed(self, text):
        if len(text) >= 1:
            self._lookup_tmr.start(700)

    def _do_lookup(self):
        self._lookup_tmr.stop()
        ticker = self._ticker_inp.text().strip().upper()
        if not ticker: return

        # Nếu đang ở Dashboard → tìm và scroll đến mã trong nội dung HTML
        if self._tabs_w.currentIndex() == 0:
            self._dash_find(ticker)
            return

        on_qc = (self._tabs_w.currentIndex() == 2)

        # Check if already in scan results
        if self._df is not None:
            match = self._df[self._df["Ticker"] == ticker]
            if not match.empty:
                row = match.iloc[0].to_dict()
                self._detail.show_row(row)
                self._qc_detail.show_row(row)
                self._chart_panel.load_ticker(ticker)
                # Highlight đúng bảng theo tab đang mở
                tbl = self._qc_table if on_qc else self._table
                for ri in range(tbl.rowCount()):
                    item = tbl.item(ri, 1)
                    if item and item.text() == ticker:
                        tbl.selectRow(ri)
                        tbl.scrollToItem(item)
                        break
                return

        # Not in df — fetch from TradingView + yfinance
        if self._ticker_wkr and self._ticker_wkr.isRunning():
            return
        self._set_status(f"Looking up {ticker}…", AMBER)
        market = self._market.currentText()
        self._ticker_wkr = TickerWorker(ticker, market)
        self._ticker_wkr.done.connect(self._on_ticker_done)
        self._ticker_wkr.failed.connect(lambda m: self._set_status(f"✕  {m}", RED))
        self._ticker_wkr.start()

    def _dash_find(self, ticker: str):
        """Scroll và highlight tất cả lần xuất hiện của ticker trong Dashboard."""
        from PySide6.QtGui import QTextCharFormat, QTextCursor
        from PySide6.QtWidgets import QTextEdit

        # Xoá highlight cũ
        self._dash_browser.setExtraSelections([])

        if not ticker:
            return

        # Tìm tất cả vị trí xuất hiện và tạo highlight vàng
        doc = self._dash_browser.document()
        fmt = QTextCharFormat()
        fmt.setBackground(QColor("#FFD700"))
        fmt.setForeground(QColor("#000000"))

        highlights = []
        cur = QTextCursor(doc)
        while True:
            cur = doc.find(ticker, cur)
            if cur.isNull():
                break
            sel = QTextEdit.ExtraSelection()
            sel.format = fmt
            sel.cursor = cur
            highlights.append(sel)

        if highlights:
            self._dash_browser.setExtraSelections(highlights)
            # Scroll đến lần đầu tiên
            self._dash_browser.setTextCursor(highlights[0].cursor)
            self._set_status(
                f"📍  {ticker} — {len(highlights)} vị trí trong Dashboard", GREEN)
        else:
            self._set_status(
                f"✕  {ticker} không có trong Dashboard hiện tại", TEXT3)

    def _on_ticker_done(self, df: pd.DataFrame):
        if df.empty: return
        row = df.iloc[0].to_dict()
        self._detail.show_row(row)
        self._qc_detail.show_row(row)
        ticker = row.get("Ticker", "")
        self._set_status(
            f"✓  {ticker}  ·  {row.get('Tên Công Ty','')}  ·  {row.get('Signal','')}  "
            f"(Tài chính {row.get('FIN_Score',0)}/{row.get('FIN_App',0)}  ·  "
            f"Thực {row.get('REAL_Score',0)}/{row.get('REAL_App',0)})",
            GREEN
        )
        self._chart_panel.load_ticker(ticker)

    # ── Help ──────────────────────────────────────────────────────────────────

    def _show_help(self):
        dlg = HelpDialog(self)
        dlg.exec()

    def _show_ai_settings(self):
        # Dựng mới mỗi lần mở để dialog lấy đúng bảng màu của theme hiện tại
        dlg = AISettingsDialog(self)
        dlg.exec()

    # ── Export ────────────────────────────────────────────────────────────────

    def _on_export(self):
        if self._df is None: return
        from datetime import datetime
        market  = self._market.currentText()
        top     = self._top_spin.value()
        default = f"tv_top{top}_{market}_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
        path, _ = QFileDialog.getSaveFileName(
            self, "Export Excel", default, "Excel (*.xlsx)")
        if not path: return
        if not path.endswith(".xlsx"): path += ".xlsx"

        # Điểm đã được tính sẵn trong score_stocks() — không cần chấm lại
        df_export = self._df.copy()

        # Hiện progress bar, disable nút
        self._export_bar.setValue(0)
        self._export_bar.setVisible(True)
        self._btn_export.setEnabled(False)
        self._set_status("Exporting…", AMBER)

        self._export_wkr = ExportWorker(
            df_export, path, market, top, self._use_yf.isChecked())
        self._export_wkr.progress.connect(self._on_export_progress)
        self._export_wkr.done.connect(self._on_export_done)
        self._export_wkr.failed.connect(self._on_export_failed)
        self._export_wkr.start()

    def _on_export_progress(self, pct, msg):
        self._export_bar.setValue(pct)
        if msg:
            self._set_status(f"💾  {msg}", AMBER)

    def _on_export_done(self, path):
        self._export_bar.setValue(100)
        self._set_status(f"✓  Exported → {path}", GREEN)
        self._btn_export.setEnabled(True)
        QTimer.singleShot(2000, lambda: self._export_bar.setVisible(False))

    def _on_export_failed(self, msg):
        self._export_bar.setVisible(False)
        self._set_status(f"✕  Export failed: {msg}", RED)
        self._btn_export.setEnabled(True)


# ─────────────────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    # Tắt GPU/sandbox của WebEngine — fix lỗi GLES3/DirectComposition trên Windows 10
    os.environ.setdefault("QTWEBENGINE_CHROMIUM_FLAGS",
                          "--disable-gpu --no-sandbox --disable-dev-shm-usage"
                          " --disable-gpu-sandbox --disable-software-rasterizer")
    os.environ.setdefault("QTWEBENGINE_DISABLE_SANDBOX", "1")
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    win = MainWindow()
    win.showMaximized()
    sys.exit(app.exec())
