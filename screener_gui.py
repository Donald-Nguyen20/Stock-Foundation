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
    QDialog, QTextBrowser, QTabWidget, QProgressBar, QInputDialog,
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

import re

import claude_auth
import ai_analysis
import universe_charts
import analysis_charts
import analysis_report
import fund_holders
import fund_chart
import signal_tracking
import valuation
import price_chart
from html import escape as _html_escape

fetch_data             = _m.fetch_data
clean_df               = _m.clean_df
score_stocks           = _m.score_stocks
save_snapshot          = _m.save_snapshot
snapshot_dir           = _m.snapshot_dir
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

# Nhãn tiêu chí để mô tả «đạt thêm / mất» giữa hai lần quét
_CHANGE_LABELS = {**{f"FIN_{k}": FIN_CRITERIA[k]["label"] for k in FIN_KEYS},
                  **{f"REAL_{k}": REAL_CRITERIA[k]["label"] for k in REAL_KEYS}}

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
    ("Ngụ ý%",     "Tăng trưởng ngụ ý%", 62, Qt.AlignRight | Qt.AlignVCenter),
    ("Lợi suất",   "Lợi suất%",     62,  Qt.AlignRight  | Qt.AlignVCenter),
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
    ("Định giá",   "Định giá",      84,  Qt.AlignCenter | Qt.AlignVCenter),
    ("Thay đổi",   "Thay đổi",     140,  Qt.AlignCenter | Qt.AlignVCenter),
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
    ("ΔBiên gộp",  "ΔBiên gộp",     70,  Qt.AlignRight  | Qt.AlignVCenter),
    ("ΔROIC",      "ΔROIC%",        62,  Qt.AlignRight  | Qt.AlignVCenter),
    ("Dốc biên",   "Dốc biên",      64,  Qt.AlignRight  | Qt.AlignVCenter),
    ("Vị trí biên","Vị trí biên",   70,  Qt.AlignRight  | Qt.AlignVCenter),
    ("Bền vững",   "Moat Score",   110,  Qt.AlignCenter | Qt.AlignVCenter),
    # 6 tiêu chí nhóm B
    ("Tiền",       "REAL_CASH",     42,  Qt.AlignCenter | Qt.AlignVCenter),
    ("Biên",       "REAL_FCFM",     42,  Qt.AlignCenter | Qt.AlignVCenter),
    ("K.buyback",  "REAL_BUYB",     62,  Qt.AlignCenter | Qt.AlignVCenter),
    ("ROE thật",   "REAL_LEV",      56,  Qt.AlignCenter | Qt.AlignVCenter),
    ("Cân đối",    "REAL_BAL",      52,  Qt.AlignCenter | Qt.AlignVCenter),
    ("Còn đúng",   "REAL_TREND",    58,  Qt.AlignCenter | Qt.AlignVCenter),
    ("Chưa mòn",   "REAL_EROD",     58,  Qt.AlignCenter | Qt.AlignVCenter),
    ("5 năm",      "REAL_SUS",      46,  Qt.AlignCenter | Qt.AlignVCenter),
    ("Tài chính",  "FIN_Disp",      64,  Qt.AlignCenter | Qt.AlignVCenter),
    ("Thực",       "REAL_Disp",     58,  Qt.AlignCenter | Qt.AlignVCenter),
    ("Tín hiệu",   "Signal",       150,  Qt.AlignCenter | Qt.AlignVCenter),
    ("Định giá",   "Định giá",      84,  Qt.AlignCenter | Qt.AlignVCenter),
    ("Thay đổi",   "Thay đổi",     140,  Qt.AlignCenter | Qt.AlignVCenter),
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
        "<i>Ngưỡng đạt: trên 15%, hoặc nằm trong nhóm 30% dẫn đầu ngành của chính nó — "
        "cửa thứ hai này để một doanh nghiệp giỏi trong ngành biên mỏng không bị loại oan.</i>",
    "Gross Margin%":
        "<b>Biên lợi nhuận gộp</b> = (doanh thu − giá vốn hàng bán) / doanh thu.<br>"
        "Phản ánh khả năng bán cao hơn chi phí sản xuất — tức sức mạnh định giá.<br>"
        "<i>Ngưỡng đạt: trên 40%, <b>hoặc</b> nằm trong nhóm 30% dẫn đầu ngành của chính nó. "
        "Biên gộp lệch rất mạnh theo ngành — phần mềm thường 60–80%, bán lẻ hiếm khi vượt 30% — "
        "nên cửa thứ hai cho phép một mã bán lẻ mạnh vẫn qua được. Chỉ áp dụng khi ngành có ít "
        "nhất 8 mã trong lần quét.</i>",
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
        "<i>Chỉ hiển thị tham khảo, không chấm điểm. Điểm chất lượng trả lời «doanh nghiệp này có tốt "
        "và có thật không»; câu «giá này có rẻ không» nằm ở cột Định giá (thị trường Mỹ) — tách hẳn "
        "khỏi điểm.</i>",
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
    "ΔBiên gộp":
        "<b>Thay đổi biên lợi nhuận gộp</b> = biên gộp TTM − biên gộp năm tài chính gần nhất, "
        "tính bằng điểm phần trăm.<br>"
        "Mọi chỉ số chất lượng khác đều là ảnh chụp một thời điểm. Cột này cho biết mức biên "
        "đó đang giữ, đang cải thiện hay đang bào mòn.<br>"
        "<i>Biên gộp rất ổn định theo năm — trung vị gần như bằng 0 — nên tụt quá 1 điểm đã là "
        "tín hiệu đáng chú ý. Là dữ liệu cho tiêu chí «Chưa xói mòn».</i>",
    "ΔROIC%":
        "<b>Thay đổi ROIC</b> so với năm tài chính gần nhất, tính bằng phần trăm tương đối.<br>"
        "−20% nghĩa là ROIC còn 80% mức năm ngoái.<br>"
        "<i>Hai doanh nghiệp cùng ROIC 16% nhưng một bên đi lên từ 8%, một bên rơi từ 40% — "
        "nếu không nhìn cột này thì chúng được chấm giống hệt nhau. Ngưỡng trượt: dưới −15%.</i>",
    "Dốc biên":
        "<b>Dốc biên lợi nhuận gộp qua các năm</b>, tính bằng điểm phần trăm mỗi năm.<br>"
        "Khớp một đường thẳng qua chuỗi biên gộp 4–5 năm lấy từ báo cáo thường niên. "
        "−3,1 nghĩa là mỗi năm mất khoảng 3 điểm biên.<br>"
        "<i>Mạnh hơn hẳn phép so hai điểm: một đường đi xuống suốt bốn năm nói được nhiều hơn "
        "«năm nay thấp hơn năm ngoái». Ngưỡng trượt: dốc dưới −1 điểm/năm, hoặc giảm 3 năm liên "
        "tiếp. Cần bật yfinance; không có thì tiêu chí lùi về so TTM với năm trước.</i>",
    "Vị trí biên":
        "<b>Vị trí biên gộp hiện tại trong dải lịch sử của chính doanh nghiệp</b>, từ 0 đến 100.<br>"
        "0 nghĩa là đang ở mức thấp nhất trong 4–5 năm qua; 100 là cao nhất.<br>"
        "<i>Chỉ hiển thị, không chấm điểm. Một mã biên 35% có thể là đỉnh lịch sử của nó, "
        "hoặc là đáy — con số tuyệt đối không phân biệt được.</i>",
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
    "Định giá":
        "<b>Định giá ngược (reverse DCF)</b> — chỉ áp dụng cho thị trường Mỹ.<br>"
        "Câu hỏi: giá hiện tại đã tính sẵn dòng tiền tăng bao nhiêu mỗi năm? Lấy vốn hóa, chiết khấu "
        f"{valuation.DISCOUNT:.0%}/năm, {valuation.YEARS} năm rồi tăng "
        f"{valuation.TERMINAL * 100:.1f}".replace(".", ",") + "% mãi mãi, giải ra mức tăng trưởng "
        "ngụ ý. Đem so với mức doanh nghiệp đã chứng minh — doanh thu nhiều năm (khi quét kèm "
        "yfinance), nếu không thì doanh thu 12 tháng; mốc được kẹp trong khoảng "
        f"{valuation.REF_LO:g}% đến {valuation.REF_HI:g}%/năm.<br>"
        f"<b>Rẻ</b>: ngụ ý thấp hơn mốc từ {-valuation.CHEAP_GAP:g} điểm · "
        f"<b>Hợp lý</b>: lệch dưới {valuation.RICH_GAP:g} điểm · "
        f"<b>Đắt</b>: cao hơn {valuation.RICH_GAP:g}–{valuation.VERY_RICH_GAP:g} điểm, hoặc ngụ ý từ "
        f"{valuation.ABS_RICH:g}%/năm · "
        f"<b>Rất đắt</b>: cao hơn từ {valuation.VERY_RICH_GAP:g} điểm, hoặc ngụ ý từ "
        f"{valuation.ABS_VERY_RICH:g}%/năm.<br>"
        "Dấu <b>?</b> = độ tin cậy thấp: phải dùng lợi nhuận thay cho dòng tiền thật, hoặc tăng "
        "trưởng gần đây quá nóng để làm mốc.<br>"
        "<i>Nhãn riêng, không cộng vào điểm và không đổi tín hiệu — chất lượng và giá là hai câu hỏi "
        "khác nhau. Rê chuột vào từng ô để xem cách tính của mã đó.</i>",
    "Tăng trưởng ngụ ý%":
        "<b>Tăng trưởng ngụ ý</b> — mức tăng dòng tiền mỗi năm, suốt "
        f"{valuation.YEARS} năm, mà giá hiện tại đã tính sẵn. Càng cao thì giá càng đòi hỏi nhiều ở "
        "tương lai.<br><i>Chỉ thị trường Mỹ. Không chấm điểm.</i>",
    "Lợi suất%":
        "<b>Lợi suất dòng tiền</b> = dòng tiền dùng để định giá / vốn hóa. Dòng tiền là FCF; nhóm tài "
        "chính, và mã có FCF dưới "
        f"{valuation.LOW_CASH_CONVERSION:g}".replace(".", ",") + " lần lợi nhuận, dùng lợi nhuận "
        "ròng. Lợi suất 5% tương đương P/FCF 20.<br><i>Chỉ thị trường Mỹ. Không chấm điểm.</i>",
    "Thay đổi":
        "<b>Thay đổi so với lần quét trước</b> của cùng thị trường (lần gần nhất trước hôm nay).<br>"
        "▲ lên hạng tín hiệu · ▼ xuống hạng · MỚI: lần trước không có trong danh sách · "
        "• N tiêu chí: tín hiệu giữ nguyên nhưng có tiêu chí đổi kết quả · • giá X: chỉ nhãn định "
        "giá đổi.<br>Rê chuột vào từng ô để xem chi tiết. Mỗi lần quét được lưu một ảnh chụp trong "
        "thư mục snapshots.",
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
from PySide6.QtWebEngineCore import QWebEnginePage


class _FundWorker(QThread):
    """Tải và rút gọn dữ liệu 13F ở luồng nền.

    Dữ liệu dùng chung cho mọi mã nên chỉ cần một luồng mỗi phiên, không
    phụ thuộc mã đang chọn.
    """
    progress    = Signal(str)
    quarter     = Signal(str)
    finished_ok = Signal()
    failed      = Signal(str, bool)     # thông báo, có phải SEC từ chối truy cập

    def __init__(self, parent=None):
        super().__init__(parent)
        self._stop = False

    def stop(self):
        self._stop = True

    def run(self):
        try:
            fund_holders.update(progress=self.progress.emit,
                                stop=lambda: self._stop,
                                on_quarter=self.quarter.emit)
        except fund_holders.Cancelled:
            return
        except fund_holders.SecAccessError as e:
            self.failed.emit(str(e), True)
            return
        except Exception as e:
            self.failed.emit(f"{type(e).__name__}: {e}", False)
            return
        self.finished_ok.emit()


class _PriceWorker(QThread):
    """Tải lịch sử giá một mã từ Yahoo Finance ở luồng nền."""
    done   = Signal(str, object)    # mã, price_chart.PriceHistory
    failed = Signal(str, str)       # mã, thông báo

    def __init__(self, ticker, exchange, market, expect, parent=None):
        super().__init__(parent)
        self._args = (ticker, exchange, market, expect)

    def run(self):
        ticker, exchange, market, expect = self._args
        try:
            hist = price_chart.fetch(ticker, exchange, market, expect=expect,
                                     stop=self.isInterruptionRequested)
        except price_chart.Cancelled:
            return
        except ValueError as e:         # không ký hiệu nào có giá
            self.failed.emit(ticker, str(e))
            return
        except Exception as e:
            self.failed.emit(ticker, f"{type(e).__name__}: {e}")
            return
        self.done.emit(ticker, hist)


class _MapPage(QWebEnginePage):
    """Bắt cú bấm chấm trên bản đồ.

    Trang chuyển tới https://pick.local/<mã>; ở đây chặn lại và báo mã đó
    lên thay vì để trình duyệt thật sự rời trang.
    """
    picked = Signal(str)

    def acceptNavigationRequest(self, url, nav_type, is_main_frame):
        if url.host() == universe_charts.PICK_HOST:
            ticker = url.path().lstrip("/").strip().upper()
            if ticker:
                # Hoãn sang vòng lặp sau: nạp lại trang ngay trong lúc
                # Chromium đang hỏi điều hướng sẽ làm sập tiến trình
                QTimer.singleShot(0, lambda: self.picked.emit(ticker))
            return False
        return super().acceptNavigationRequest(url, nav_type, is_main_frame)


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


class _ElidedLabel(QLabel):
    """Nhãn một dòng tự cắt «…» khi thiếu chỗ thay vì ép khung chứa giãn ra.
    Rê chuột để đọc đủ."""

    def minimumSizeHint(self):
        s = super().minimumSizeHint()
        s.setWidth(min(s.width(), 40))
        return s

    def setText(self, text):
        super().setText(text)
        self.setToolTip(text)

    def paintEvent(self, _event):
        painter = QPainter(self)
        r = self.contentsRect()
        text = self.fontMetrics().elidedText(self.text(), Qt.ElideRight, r.width())
        self.style().drawItemText(painter, r, Qt.AlignRight | Qt.AlignVCenter, self.palette(),
                                  self.isEnabled(), text, self.foregroundRole())


class ChartPanel(QWidget):
    """Khung biểu đồ — QWebEngineView dựng HTML Plotly."""
    ticker_picked = Signal(str)     # bấm một chấm trên bản đồ
    _PX_TTL   = 30 * 60     # giây: giá cũ hơn thì tải lại khi xem
    _PX_KEEP  = 60          # số mã giữ giá trong bộ nhớ
    _PX_DELAY = 300         # ms: chờ người dùng thôi lướt bảng rồi mới tải
    _RANGE_SHORT = {"6m": "6T", "1y": "1N", "3y": "3N", "5y": "5N"}

    def __init__(self, parent=None):
        super().__init__(parent)
        # Khớp với theme mặc định của app (sáng); lệch nhau thì biểu đồ đen
        # nằm giữa giao diện trắng.
        self._dark     = False
        self._html_q   = ""
        self._html_a   = ""
        self._html_map = ""
        self._map_key  = None
        self._map_view = universe_charts.DEFAULT_VIEW
        self._df       = None
        self._mode     = "quarterly"
        self._worker   = None
        self._ticker   = ""
        self._spin_idx = 0
        self._spin_msg = ""
        self._spin_tmr = QTimer(self)
        self._spin_tmr.timeout.connect(self._tick)
        # Chế độ "Quỹ": số quỹ sở hữu theo quý từ SEC 13F
        self._fund_worker = None
        self._fund_done   = False   # đã cập nhật xong trong phiên này
        self._fund_err    = ""
        self._fund_tick   = 0
        self._market_us   = True
        self._market      = ""
        # Chế độ "Giá": nến từ Yahoo Finance, giữ theo mã trong phiên
        # Khoá (thị trường, mã): cùng một mã ở hai thị trường là hai ký hiệu Yahoo
        self._px_cache   = {}       # khoá → PriceHistory
        self._px_time    = {}       # khoá → lúc tải xong (time.monotonic)
        self._px_err     = {}       # khoá → lỗi của lần tải gần nhất
        self._px_workers = {}       # khoá → _PriceWorker đang chạy
        self._px_range   = price_chart.DEFAULT_RANGE
        self._px_timer   = QTimer(self)
        self._px_timer.setSingleShot(True)
        self._px_timer.setInterval(self._PX_DELAY)
        self._px_timer.timeout.connect(lambda: self._ensure_price(self._ticker))
        app = QApplication.instance()
        if app is not None:
            app.aboutToQuit.connect(self.shutdown)
        self.setStyleSheet(f"background:{BG};")

        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)
        self._bar_w    = self._build_bar()
        self._sep_line = self._sep()
        layout.addWidget(self._bar_w)
        layout.addWidget(self._sep_line)

        self._web = QWebEngineView()
        page = _MapPage(self._web)
        page.picked.connect(self.ticker_picked)
        self._web.setPage(page)
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
        h.addWidget(self._chart_lbl)
        # Chiếm phần trống giữa tên mã và các nút, chữ dồn phải sát nút
        self._chart_status = _ElidedLabel("Click a row to load chart")
        self._chart_status.setStyleSheet(f"color:{TEXT3}; font-size:9px; letter-spacing:1px;")
        h.addWidget(self._chart_status, 1); h.addSpacing(10)
        self._map_combo = QComboBox()
        for key, label in universe_charts.VIEWS.items():
            self._map_combo.addItem(label, key)
        self._map_combo.setFixedHeight(22)
        self._map_combo.setCursor(Qt.PointingHandCursor)
        self._map_combo.setStyleSheet(self._combo_style())
        self._map_combo.setVisible(False)
        self._map_combo.currentIndexChanged.connect(self._on_map_view)
        h.addWidget(self._map_combo)
        # Khung thời gian của đồ thị giá — chỉ hiện ở chế độ Giá
        self._rng_btns = {}
        for key, short in self._RANGE_SHORT.items():
            label, _days, freq = price_chart.RANGES[key]
            b = QPushButton(short)
            b.setFixedSize(28, 22); b.setCursor(Qt.PointingHandCursor)
            b.setToolTip(f"{label} · nến {'ngày' if freq == 'D' else 'tuần'}")
            b.setVisible(False)
            b.clicked.connect(lambda _=False, k=key: self._set_range(k))
            h.addWidget(b)
            self._rng_btns[key] = b
        self._btn_q = QPushButton("Theo quý")
        self._btn_a = QPushButton("Theo năm")
        self._btn_p = QPushButton("Giá")
        self._btn_m = QPushButton("Bản đồ")
        self._btn_f = QPushButton("Quỹ")
        for btn in (self._btn_q, self._btn_a, self._btn_p, self._btn_m, self._btn_f):
            btn.setFixedSize(70, 22); btn.setCursor(Qt.PointingHandCursor)
        self._btn_q.clicked.connect(lambda: self._switch_mode("quarterly"))
        self._btn_a.clicked.connect(lambda: self._switch_mode("annual"))
        self._btn_p.clicked.connect(lambda: self._switch_mode("price"))
        self._btn_m.clicked.connect(lambda: self._switch_mode("map"))
        self._btn_f.clicked.connect(lambda: self._switch_mode("funds"))
        h.addWidget(self._btn_q); h.addWidget(self._btn_a); h.addWidget(self._btn_p)
        h.addWidget(self._btn_m); h.addWidget(self._btn_f)
        self._refresh_btns()
        return w

    def _sep(self):
        f = QFrame(); f.setFrameShape(QFrame.HLine); f.setFixedHeight(1)
        f.setStyleSheet(f"background:{BORDER}; border:none;"); return f

    def _combo_style(self):
        return (f"QComboBox {{ background:transparent; color:{TEXT1};"
                f" border:1px solid {BORDER}; border-radius:3px; padding:0 8px;"
                f" font-size:9px; font-weight:600; letter-spacing:1px;"
                f" font-family:'Segoe UI',sans-serif; min-width:130px; }}"
                f"QComboBox:hover {{ border-color:{BORDER2}; }}"
                f"QComboBox::drop-down {{ border:none; width:16px; }}"
                f"QComboBox QAbstractItemView {{ background:{SURFACE}; color:{TEXT1};"
                f" selection-background-color:{BORDER2}; border:1px solid {BORDER}; }}")

    def load_ticker(self, ticker: str, show_chart: bool = True):
        """Nạp biểu đồ của một mã.

        show_chart=False khi chọn mã ngay trên bản đồ: ở lại bản đồ, khoanh
        mã mới, còn biểu đồ của mã vẫn tải ngầm cho nút "Theo quý".
        """
        ticker = ticker.upper()
        # Đang xem số quỹ hay giá thì đổi mã vẫn ở lại đó để so nhanh giữa các mã
        stay = ((not show_chart and self._mode == "map")
                or (self._mode == "funds" and self._is_us(ticker))
                or self._mode == "price")
        if ticker == self._ticker and (self._html_q or self._html_a):
            if self._mode in ("map", "funds") and not stay:   # quay về mã đã nạp
                self._switch_mode("quarterly")
            return
        self._ticker = ticker
        self._html_q = ""; self._html_a = ""
        if stay and self._mode == "funds":
            self._render_funds()
        elif stay and self._mode == "price":
            self._render_price()
            self._px_timer.start()          # lướt nhanh qua bảng thì chỉ tải mã dừng lại
        elif stay:
            self._render_map()
        else:
            self._mode = "quarterly"
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
        if self._mode in ("funds", "price"):        # dòng trạng thái đang báo việc khác
            self._refresh_btns()
            return
        if self._mode == "map":                     # đang xem bản đồ thì giữ nguyên
            self._set_status(f"✓  {self._ticker} sẵn sàng ở “Theo quý”", GREEN)
            self._refresh_btns()
            return
        self._mode   = "quarterly"
        self._web.setHtml(html_q)
        parts = summary.split("·")
        self._set_status(f"✓  {parts[-1].strip()}" if len(parts) > 1 else "✓", GREEN)
        self._refresh_btns()

    def _on_failed(self, msg):
        self._spin_tmr.stop()
        if self._mode in ("funds", "price"):
            return
        if self._mode == "map":
            self._set_status(f"✕  {msg}", RED)
            return
        bg = "#0A1628" if self._dark else "#F0F4F9"
        self._web.setHtml(
            f'<html><body style="background:{bg};margin:0;display:flex;'
            f'align-items:center;justify-content:center;height:100vh;">'
            f'<p style="color:#DC2626;font-family:Segoe UI,sans-serif;'
            f'font-size:11px;">Error: {msg}</p></body></html>')
        self._set_status(f"✕  {msg}", RED)

    def _switch_mode(self, mode):
        if mode == "funds":
            if self._is_us(self._ticker):
                self._mode = "funds"
                self._render_funds()
                self._refresh_btns()
                self._ensure_funds()        # bấm lại khi lỗi để thử lại
            return
        if mode == "price":
            if self._ticker:
                self._mode = "price"
                self._ensure_price(self._ticker)    # bấm lại khi lỗi để thử lại
                self._render_price()
                self._refresh_btns()
            return
        if mode == self._mode:
            return
        leaving_side = self._mode in ("funds", "price")
        if mode == "map":
            if self._df is None:
                return
            self._mode = "map"
            self._render_map()
        else:
            html = self._html_a if mode == "annual" else self._html_q
            if not html:
                return
            self._mode = mode
            if self._ticker:
                self._chart_lbl.setText(self._ticker)
            self._web.setHtml(html)
        self._refresh_btns()
        if leaving_side and not self._spin_tmr.isActive():    # đang tải thì _tick lo
            self._set_status(f"✓  {self._ticker}" if self._html_q else "",
                             GREEN if self._html_q else TEXT3)

    def set_universe(self, df):
        """Nhận kết quả quét để dựng bản đồ so sánh.

        Dùng đúng dữ liệu đã có trong bảng, không gọi mạng thêm.
        """
        self._df = df
        self._map_key = None
        # Chưa chọn mã nào thì mở thẳng bản đồ — hữu ích hơn ô trống
        if not self._ticker:
            self._mode = "map"
        self._render_map()
        self._render_price()
        self._refresh_btns()

    def _on_map_view(self, _idx):
        self._map_view = self._map_combo.currentData() or universe_charts.DEFAULT_VIEW
        self._render_map()

    def _render_map(self):
        """Dựng lại bản đồ theo loại đang chọn, mã đang chọn và theme.

        Chỉ hiện lên khung khi đang ở chế độ bản đồ; dựng trùng thì bỏ qua.
        """
        if self._df is None:
            return
        key = (id(self._df), self._map_view, self._ticker, self._dark)
        if key != self._map_key:
            try:
                self._html_map = universe_charts.build_html(
                    self._df, dark=self._dark, view=self._map_view, focus=self._ticker)
            except Exception as e:
                self._html_map = ""
                self._set_status(f"✕  bản đồ lỗi: {e}", RED)
            self._map_key = key
        if self._mode != "map" or not self._html_map:
            return
        self._chart_lbl.setText(f"BẢN ĐỒ · {self._ticker}" if self._ticker else "TOÀN BỘ MÃ")
        self._set_status("so sánh giữa các mã", TEXT3)
        self._web.setHtml(self._html_map)

    def _set_status(self, text, color=TEXT3):
        self._chart_status.setText(text)
        self._chart_status.setStyleSheet(f"color:{color}; font-size:9px; letter-spacing:1px;")

    def _tick(self):
        self._spin_idx = (self._spin_idx + 1) % len(SPINNER)
        if self._mode in ("funds", "price"):    # nhường dòng trạng thái cho việc đang xem
            return
        self._set_status(f"{SPINNER[self._spin_idx]}  {self._spin_msg}", AMBER)

    # ── Chế độ Quỹ ────────────────────────────────────────────────────────────

    def set_market(self, market: str):
        """Thị trường đang chọn — dùng khi mã không có trong kết quả quét."""
        self._market_us = market.lower() in {"america", "nasdaq", "nyse"}
        changed = market != self._market
        self._market = market
        self._refresh_btns()
        if changed and self._mode == "price":
            self._render_price()
            self._px_timer.start()

    def _is_us(self, ticker: str) -> bool:
        if not ticker:
            return False
        df = self._df
        if df is not None and "ticker" in df.columns and "Ticker" in df.columns:
            hit = df.loc[df["Ticker"] == ticker, "ticker"]
            if not hit.empty:
                return str(hit.iat[0]).split(":")[0].upper() in fund_holders.US_EXCHANGES
        return self._market_us

    def _fund_running(self) -> bool:
        return self._fund_worker is not None and self._fund_worker.isRunning()

    def _render_funds(self):
        if self._mode != "funds" or not self._ticker:
            return
        t = self._ticker
        self._chart_lbl.setText(f"QUỸ · {t}")
        read_err = ""
        try:
            hist = fund_holders.history(t)
        except Exception as e:              # báo thay vì sập; lần vẽ sau đọc lại
            hist = None
            read_err = f"Không đọc được dữ liệu đã lưu: {e}"
        err = self._fund_err or read_err
        running = self._fund_running()
        if hist is not None and hist["holders"].notna().any():
            note = "đang tải thêm các quý cũ…" if running else ""
            self._web.setHtml(fund_chart.build_html(hist, t, self._dark, note))
            if not running and not err:
                if self._fund_done or fund_holders.user_agent():
                    self._set_status(f"✓  {len(hist)} quý · SEC 13F", GREEN)
                else:   # xem bản đã lưu, không cập nhật được
                    self._set_status(f"{len(hist)} quý đã lưu · chưa có liên hệ SEC "
                                     "để cập nhật", TEXT3)
            return
        if err:
            title, tone = "Chưa tải được dữ liệu quỹ", "error"
            body = (f"{_html_escape(err)}<br><br>"
                    "Bấm lại nút <b>Quỹ</b> để thử lại.")
        elif running or not fund_holders.cached_quarters():
            title, tone = "Đang tải dữ liệu 13F lần đầu…", "warn"
            body = ("Mỗi quý khoảng 100 MB, đủ 12 quý mất chừng 10 phút. "
                    "Biểu đồ hiện dần khi từng quý xong; các lần sau chỉ tải quý mới.")
        else:
            title, tone = f"Không thấy {t} trong dữ liệu 13F", "dim"
            body = ("Mã có thể mới niêm yết, là chứng chỉ lưu ký hoặc quỹ ETF ít "
                    "quỹ báo cáo, hoặc chưa có trong bảng ánh xạ CUSIP của SEC.")
        self._web.setHtml(fund_chart.message_html(title, body, self._dark, tone))

    def _ensure_funds(self):
        """Cập nhật dữ liệu 13F một lần mỗi phiên; không có mạng vẫn xem bản đã lưu."""
        if self._fund_running() or self._fund_done:
            return
        if not fund_holders.user_agent():
            if fund_holders.cached_quarters():
                # Đã có dữ liệu thì xem bản đó, không hỏi lại mỗi lần mở app
                return
            if not self._ask_sec_contact():
                self._fund_err = "Cần tên và email liên hệ thì SEC mới cho tải dữ liệu."
                self._set_status("✕  chưa có thông tin liên hệ cho SEC", RED)
                self._render_funds()
                return
        self._fund_err = ""
        w = _FundWorker(self)
        w.progress.connect(self._on_fund_progress)
        w.quarter.connect(lambda _k: self._render_funds())
        w.finished_ok.connect(self._on_fund_done)
        w.failed.connect(self._on_fund_failed)
        self._fund_worker = w
        w.start()
        self._render_funds()

    def _ask_sec_contact(self) -> bool:
        text, ok = QInputDialog.getText(
            self, "Dữ liệu quỹ từ SEC",
            "SEC (Mỹ) chỉ cho tải dữ liệu Form 13F khi mỗi yêu cầu ghi tên và email\n"
            "liên hệ của người tải. Nhập theo dạng:   Tên của bạn  email@domain.com\n\n"
            "Thông tin này chỉ gửi tới sec.gov và lưu trong thư mục fund_cache trên máy.\n"
            "Lần đầu tải 12 quý (khoảng 1,2 GB, chừng 10 phút); sau đó chỉ tải quý mới.")
        if not ok or not text.strip():
            return False
        try:
            fund_holders.save_user_agent(text)
        except (ValueError, OSError) as e:
            self._fund_err = f"Thông tin liên hệ chưa hợp lệ: {e}"
            return False
        return True

    def _on_fund_progress(self, msg: str):
        self._fund_tick += 1
        if self._mode == "funds":
            self._set_status(f"{SPINNER[self._fund_tick % len(SPINNER)]}  {msg}", AMBER)

    def _on_fund_done(self):
        self._fund_done = True
        self._render_funds()

    def _on_fund_failed(self, msg: str, access_denied: bool):
        if access_denied and not fund_holders.ua_from_env():
            fund_holders.forget_user_agent()    # để lần bấm sau hỏi lại
        self._fund_err = msg
        if self._mode == "funds":
            self._set_status(f"✕  {msg}", RED)
            self._render_funds()

    # ── Chế độ Giá ────────────────────────────────────────────────────────────

    def _scan_row(self, ticker: str):
        """Dòng của mã trong kết quả quét (dict), None khi không có."""
        df = self._df
        if df is None or not ticker or "Ticker" not in df.columns:
            return None
        hit = df.loc[df["Ticker"] == ticker]
        return None if hit.empty else hit.iloc[0].to_dict()

    def _price_context(self, ticker: str):
        """(sàn, giá lúc quét, vùng định giá, lý do không có vùng) từ kết quả quét."""
        row = self._scan_row(ticker)
        if row is None:
            return "", None, None, "mã không có trong kết quả quét hiện tại"
        code = _txt(row.get("ticker"))
        exch = code.split(":")[0] if ":" in code else ""
        bands = valuation.price_bands(row)
        note = "" if bands is not None else (_txt(row.get("Định giá_lý do"))
                                             or "mã chưa có nhãn định giá")
        return exch, _fnum(row.get("Price ($)")), bands, note

    def _px_key(self, ticker: str):
        return (self._market, ticker)

    def _ensure_price(self, ticker: str):
        """Tải giá khi chưa có, đã cũ, hoặc lần trước lỗi."""
        key = self._px_key(ticker)
        if self._mode != "price" or not ticker or key in self._px_workers:
            return
        loaded = self._px_time.get(key)
        if key in self._px_cache and loaded is not None \
                and time.monotonic() - loaded < self._PX_TTL:
            return
        had_err = self._px_err.pop(key, None) is not None
        exch, expect, _bands, _note = self._price_context(ticker)
        w = _PriceWorker(ticker, exch, self._market, expect, self)
        w.done.connect(lambda _t, hist, k=key: self._on_price_done(k, hist))
        w.failed.connect(lambda _t, msg, k=key: self._on_price_failed(k, msg))
        w.finished.connect(lambda k=key, w=w: self._on_price_finished(k, w))
        self._px_workers[key] = w
        w.start()
        if had_err:
            self._render_price()                # thôi báo lỗi cũ, hiện «đang tải»

    def _on_price_done(self, key, hist):
        self._px_cache.pop(key, None)           # đưa xuống cuối: mới tải nhất
        self._px_cache[key] = hist
        self._px_time[key] = time.monotonic()
        self._px_err.pop(key, None)
        while len(self._px_cache) > self._PX_KEEP:
            old = next(iter(self._px_cache))
            self._px_cache.pop(old); self._px_time.pop(old, None)
        if key == self._px_key(self._ticker):
            self._render_price()

    def _on_price_failed(self, key, msg: str):
        self._px_err[key] = msg
        if key == self._px_key(self._ticker):
            self._render_price()

    def _on_price_finished(self, key, worker):
        if self._px_workers.get(key) is worker:
            del self._px_workers[key]
        worker.deleteLater()

    def _set_range(self, key: str):
        if key not in price_chart.RANGES or key == self._px_range:
            return
        self._px_range = key
        self._refresh_btns()
        self._render_price()

    def _render_price(self):
        if self._mode != "price" or not self._ticker:
            return
        t = self._ticker
        self._chart_lbl.setText(f"GIÁ · {t}")
        hist = self._px_cache.get(self._px_key(t))
        if hist is not None:
            _exch, _expect, bands, note = self._price_context(t)
            try:
                html = price_chart.build_html(hist, self._px_range, bands, note, self._dark)
            except Exception as e:              # dữ liệu lạ thì báo, không sập app
                self._set_status(f"✕  không vẽ được đồ thị giá: {e}", RED)
                self._web.setHtml(price_chart.message_html(
                    f"Không vẽ được đồ thị giá {t}", _html_escape(f"{type(e).__name__}: {e}"),
                    self._dark, "error"))
                return
            self._web.setHtml(html)
            label = price_chart.RANGES[self._px_range][0]
            self._set_status(f"✓  {hist.symbol} · {label} · Yahoo Finance", GREEN)
            return
        err = self._px_err.get(self._px_key(t))
        if err:
            self._set_status(f"✕  {err}", RED)
            title, tone = f"Chưa tải được giá {t}", "error"
            body = (f"{_html_escape(err)}<br><br>"
                    "Bấm lại nút <b>Giá</b> để thử lại.")
        else:
            self._set_status(f"Đang tải giá {t} từ Yahoo Finance…", AMBER)
            title, tone = f"Đang tải giá {t}…", "warn"
            body = ("Khoảng 5 năm nến ngày từ Yahoo Finance. Mỗi mã chỉ tải một lần, "
                    "đổi khung thời gian không phải tải lại.")
        self._web.setHtml(price_chart.message_html(title, body, self._dark, tone))

    def shutdown(self):
        """Dừng luồng tải 13F và các luồng tải giá trước khi thoát để Qt không
        huỷ luồng đang chạy. Xin dừng tất cả trước rồi mới chờ, để chúng dừng
        song song.

        Gọi lúc cửa sổ đã đóng nên chờ không làm treo giao diện. Luồng 13F xét
        cờ dừng sau mỗi 1 MB tải và mỗi khúc đọc; bước chặn lâu nhất là một lần
        đọc mạng (fund_holders.TIMEOUT), nên 45 giây là đủ. Luồng giá xét cờ
        trước mỗi ký hiệu Yahoo; một ký hiệu là vài yêu cầu, mỗi cái ≤ 10 giây.
        """
        self._px_timer.stop()
        workers = list(self._px_workers.values())
        for w in workers:
            w.requestInterruption()
        deadline = time.monotonic() + 20
        if self._fund_running():
            self._fund_worker.stop()
            self._fund_worker.wait(45_000)
        for w in workers:
            left = max(0, int((deadline - time.monotonic()) * 1000))
            if not w.wait(left):                # mạng treo: đằng nào cũng đang thoát
                w.terminate()
                w.wait(2_000)

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
        has_m = self._df is not None
        has_f = self._is_us(self._ticker)
        has_p = bool(self._ticker)
        self._map_combo.setVisible(self._mode == "map" and has_m)
        for key, b in self._rng_btns.items():
            b.setVisible(self._mode == "price" and has_p)
            b.setStyleSheet(act if key == self._px_range else off)
        self._btn_q.setEnabled(has_q)
        self._btn_a.setEnabled(has_a)
        self._btn_p.setEnabled(has_p)
        self._btn_m.setEnabled(has_m)
        self._btn_f.setEnabled(has_f)
        self._btn_p.setToolTip(
            "Đồ thị nến, đường trung bình và vùng giá theo định giá" if has_p
            else "Chọn một mã trước")
        self._btn_f.setToolTip(
            "Số quỹ sở hữu theo quý (SEC Form 13F)" if has_f
            else "Chỉ có cho cổ phiếu niêm yết ở Mỹ — chọn một mã trước")
        self._btn_q.setStyleSheet(act if (self._mode == "quarterly" and has_q) else off if has_q else dis)
        self._btn_a.setStyleSheet(act if (self._mode == "annual"    and has_a) else off if has_a else dis)
        self._btn_p.setStyleSheet(act if (self._mode == "price"     and has_p) else off if has_p else dis)
        self._btn_m.setStyleSheet(act if (self._mode == "map"       and has_m) else off if has_m else dis)
        self._btn_f.setStyleSheet(act if (self._mode == "funds"     and has_f) else off if has_f else dis)

    def set_theme(self, dark: bool):
        self._dark = dark
        self._render_map()
        self._map_combo.setStyleSheet(self._combo_style())
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


# Nhãn định giá (chữ, nền) — cùng bộ màu với nhãn tín hiệu
_VALUATION_COLORS = {
    "Rẻ":      ("#1A5C2B", "#C6EFCE"),
    "Hợp lý":  ("#1B3A5C", "#DDEEFF"),
    "Đắt":     ("#7D6608", "#FFF2CC"),
    "Rất đắt": ("#9C0006", "#FFC7CE"),
}
_CHANGE_STYLE = (("▲", "#1A5C2B", "#C6EFCE"), ("▼", "#9C0006", "#FFC7CE"),
                 ("MỚI", "#1B3A5C", "#DDEEFF"))


def _fnum(v):
    """Số thực hoặc None — chịu được NaN, chuỗi rỗng và giá trị đọc từ CSV."""
    try:
        f = float(v)
    except (TypeError, ValueError):
        return None
    return None if f != f else f


def _txt(v) -> str:
    """Chuỗi hoặc "" — NaN không được hiện thành chữ 'nan'."""
    return "" if v is None or (isinstance(v, float) and v != v) else str(v)


def _is_reliable(v) -> bool:
    # Đọc lại từ ảnh chụp CSV thì cờ là chuỗi "True"/"False"
    return v is True or str(v) == "True"


def _valuation_cell(row):
    """Ô nhãn định giá. Sắp tăng dần thì Rẻ lên đầu và ô trống xuống cuối
    (sắp giảm dần thì ngược lại)."""
    verdict = _txt(row.get("Định giá"))
    rank    = _fnum(row.get("Định giá_rank"))
    reason  = _txt(row.get("Định giá_lý do"))
    cell = NumItem(rank if verdict and rank is not None else 99.0)
    if not verdict:
        cell.setText("—")
        cell.setForeground(QColor(TEXT3))
        if not reason and "Định giá_lý do" not in row:
            reason = "chưa có số liệu định giá — nhấn SCAN để tính"
        cell.setToolTip("<b>Chưa định giá được</b>"
                        + (f"<br>{_html_escape(reason)}" if reason else ""))
        return cell
    sure = _is_reliable(row.get("Định giá_tin cậy"))
    cell.setText(verdict if sure else f"{verdict} ?")
    fg, bg = _VALUATION_COLORS.get(verdict, (TEXT1, PANEL))
    cell.setForeground(QColor(fg)); cell.setBackground(QColor(bg))
    tip = f"<b>Định giá: {_html_escape(verdict)}</b>"
    if not sure:
        tip += " <i>— độ tin cậy thấp</i>"
    if reason:
        tip += f"<br>{_html_escape(reason)}"
    pctl = _fnum(row.get("Lợi suất_pctl"))
    if pctl is not None:
        basis = _html_escape(_txt(row.get("Cơ sở ĐG")))
        tip += (f"<br><br>Lợi suất cao hơn {pctl:.0f}% số mã cùng ngành, cùng cơ sở {basis} "
                f"trong lần quét này (100 = rẻ nhất).")
    tip += "<br><i>Nhãn riêng — không cộng vào điểm, không đổi tín hiệu.</i>"
    cell.setToolTip(tip)
    return cell


def _change_cell(row):
    """Ô thay đổi so với lần quét trước. Sắp giảm dần thì mã lên hạng nhiều nhất lên đầu."""
    text = _txt(row.get("Thay đổi"))
    cell = NumItem(_fnum(row.get("Thay đổi_rank")) or 0.0)
    cell.setText(text or "—")
    for prefix, fg, bg in _CHANGE_STYLE:
        if text.startswith(prefix):
            cell.setForeground(QColor(fg)); cell.setBackground(QColor(bg))
            break
    else:
        if not text:
            cell.setForeground(QColor(TEXT3))
    detail = _txt(row.get("Thay đổi_chi tiết"))
    if detail:
        lines = [_html_escape(s) for s in detail.split("\n")]
        cell.setToolTip(f"<b>{lines[0]}</b><br>" + "<br>".join(lines[1:]))
    return cell


def _implied_cell(val):
    """Ô tăng trưởng ngụ ý — nghiệm chạm biên hiện thành cận (≥ +100%), không phải số."""
    v = _fnum(val)
    cell = NumItem(v)
    if v is not None:
        cell.setText(valuation.growth_label(v, "."))
    return cell


def _pct_cell(val, fmt):
    """Ô phần trăm không tô màu — giữ màu chữ mặc định để đổi theme không lệch."""
    v = _fnum(val)
    cell = NumItem(v)
    if v is not None:
        cell.setText(fmt.format(v))
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
                    moat_cache[ticker] = fetch_moat_yfinance(ticker, sector)
                    self.ticker_update.emit(i, total, ticker, moat_cache[ticker][1])
                    time.sleep(0.3)

            # Hướng thị trường chỉ dùng để hiển thị bối cảnh trên Dashboard —
            # KHÔNG cộng vào điểm, vì nó giống nhau cho mọi mã nên không phân biệt được mã nào.
            self.progress.emit(f"Fetching market direction ({_MKT_INDEX.get(self.market.lower(), ('','SP:SPX'))[1]})…")
            market_ok, index_1y = fetch_market_direction(self.market)
            self.market_ctx.emit(market_ok, index_1y)
            self.progress.emit("Chấm điểm Tài chính + Kiểm chứng…")
            df = score_stocks(df, moat_cache, self.market)
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
            df = score_stocks(df, moat_cache, self.market)
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

        # Định giá và thay đổi so với lần quét trước — thông tin kèm, không phải tiêu chí
        self._lbl_extra = QLabel("")
        self._lbl_extra.setWordWrap(True)
        self._lbl_extra.setTextFormat(Qt.RichText)
        self._lbl_extra.setStyleSheet(self._extra_style())
        self._lbl_extra.hide()
        layout.addWidget(self._lbl_extra)

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

    @staticmethod
    def _extra_style():
        return f"color:{TEXT2}; font-size:11px; font-family:'Segoe UI',sans-serif;"

    def _show_extra(self, row: dict):
        """Dòng định giá và dòng thay đổi dưới các ô tiêu chí."""
        lines, tips = [], []
        verdict = _txt(row.get("Định giá"))
        if verdict:
            fg, bg = _VALUATION_COLORS.get(verdict, (TEXT1, PANEL))
            sure = _is_reliable(row.get("Định giá_tin cậy"))
            g, ref = _fnum(row.get("Tăng trưởng ngụ ý%")), _fnum(row.get("Mốc tăng trưởng%"))
            yld = _fnum(row.get("Lợi suất%"))
            vn  = lambda v, f: format(v, f).replace(".", ",")    # câu văn: dấu phẩy thập phân
            facts = []
            if g is not None:
                facts.append(f"giá ngụ ý tăng {valuation.growth_label(g)}/năm "
                             f"trong {valuation.YEARS} năm")
            if ref is not None:
                facts.append(f"mốc đã chứng minh {vn(ref, '+.1f')}% "
                             f"({_txt(row.get('Mốc ĐG')) or '—'})")
            if yld is not None:
                facts.append(f"lợi suất {_txt(row.get('Cơ sở ĐG'))} {vn(yld, '.1f')}%")
            lines.append(
                f'<span style="color:{fg}; background-color:{bg}; font-weight:700;">'
                f'&nbsp;💲 {_html_escape(verdict)}{"" if sure else " ?"}&nbsp;</span>&nbsp; '
                + _html_escape(" · ".join(facts)))
            reason = _txt(row.get("Định giá_lý do"))
            if reason:
                tips.append("<b>Định giá</b><br>" + _html_escape(reason))
        detail = _txt(row.get("Thay đổi_chi tiết"))
        if detail:
            parts  = detail.split("\n")
            change = _txt(row.get("Thay đổi")) or "không đổi"
            body   = " · ".join(parts[1:])
            lines.append(f"🔄 <b>{_html_escape(change)}</b> — {_html_escape(parts[0].rstrip(':'))}"
                         + (f": {_html_escape(body)}" if body else ""))
            tips.append("<br>".join(_html_escape(p) for p in parts))
        self._lbl_extra.setText("<br>".join(lines))
        self._lbl_extra.setToolTip("<br><br>".join(tips))
        self._lbl_extra.setVisible(bool(lines))

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
        self._show_extra(row)

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
        self._lbl_extra.clear()
        self._lbl_extra.hide()
        self._lbl_empty.show()

    def apply_theme(self):
        self.setStyleSheet(f"background:{PANEL};")
        self._box.setStyleSheet(f"background:{SURFACE}; border-radius:4px;")
        self._lbl_ticker.setStyleSheet(
            f"color:{TEXT1}; font-size:20px; font-weight:700;"
            f" font-family:'Consolas',monospace; letter-spacing:3px;")
        self._lbl_info.setStyleSheet(
            f"color:{TEXT1}; font-size:13px; font-family:'Segoe UI',sans-serif;")
        self._lbl_extra.setStyleSheet(self._extra_style())
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
        self.resize(880, 720)
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
        foot.setObjectName("helpFoot")      # để nền/viền không lan xuống nút
        foot.setFixedHeight(44)
        foot.setStyleSheet(f"#helpFoot {{ background:{SURFACE}; border-top:1px solid {BORDER}; }}")
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

    @staticmethod
    def _num(x, nd=0):
        """Số kiểu Việt Nam: dấu phẩy thập phân."""
        return f"{x:.{nd}f}".replace(".", ",")

    @classmethod
    def _pct(cls, x):
        return f"{cls._num(x, 1)}%"

    @classmethod
    def _valuation_rows(cls, dim, accent, text):
        """(nhãn định giá, bảng bội số ↔ tăng trưởng, các bước tính) — lấy thẳng
        tham số từ valuation.py để trang trợ giúp không lệch với cách tính."""
        V = valuation
        conds = {
            "Rẻ":      f"chênh ≤ {V.CHEAP_GAP:g}",
            "Hợp lý":  f"{V.CHEAP_GAP:g} &lt; chênh &lt; +{V.RICH_GAP:g}",
            "Đắt":     f"+{V.RICH_GAP:g} ≤ chênh &lt; +{V.VERY_RICH_GAP:g}<br>"
                       f"hoặc ngụ ý ≥ {V.ABS_RICH:g}%/năm",
            "Rất đắt": f"chênh ≥ +{V.VERY_RICH_GAP:g}<br>hoặc ngụ ý ≥ {V.ABS_VERY_RICH:g}%/năm",
        }
        means = {
            "Rẻ":      "Giá đòi hỏi ít hơn những gì doanh nghiệp đã làm được. Nếu tăng trưởng không gãy, "
                       "giá có biên an toàn.",
            "Hợp lý":  "Giá phản ánh đúng quỹ đạo hiện tại — lợi nhuận về sau chủ yếu đến từ chính tăng "
                       "trưởng của doanh nghiệp, không từ việc thị trường định giá lại.",
            "Đắt":     "Giá đòi doanh nghiệp tăng nhanh hơn mức đã chứng minh. Cần một lý do cụ thể vì sao "
                       "tương lai sẽ nhanh hơn quá khứ.",
            "Rất đắt": "Giá đã tính sẵn một kịch bản rất ít doanh nghiệp đạt được. Chỉ cần tăng trưởng "
                       "chậm lại cũng đủ làm giá giảm mạnh.",
        }
        val_rows = ""
        for v in V.VERDICTS:
            fg, bg = _VALUATION_COLORS[v][:2]
            val_rows += f"""
            <tr style="background-color:{bg};">
              <td style="color:{fg};font-weight:700;font-size:13px;padding:9px 13px;">{v}</td>
              <td style="color:{fg};font-weight:700;font-size:12px;padding:9px 13px;font-family:Consolas,monospace;">{conds[v]}</td>
              <td style="color:{fg};font-size:12px;padding:9px 13px;line-height:1.55;">{means[v]}</td>
            </tr>"""

        ex_rows = ""
        for i, mult in enumerate((10, 15, 20, 25, 33.3, 50)):
            bg = "#F0F4FA" if i % 2 == 0 else "#FFFFFF"
            g = V.growth_label(V.implied_growth(mult))
            ex_rows += f"""
            <tr style="background-color:{bg};">
              <td style="color:{text};font-family:Consolas,monospace;padding:6px 13px;">{cls._num(mult, 0 if mult == int(mult) else 1)}×</td>
              <td style="color:{text};font-family:Consolas,monospace;padding:6px 13px;">{cls._pct(100 / mult)}</td>
              <td style="color:{accent};font-weight:700;font-family:Consolas,monospace;padding:6px 13px;">{g}/năm</td>
            </tr>"""

        conv = cls._num(V.LOW_CASH_CONVERSION, 1)
        steps = [
            ("Chọn dòng tiền gốc",
             f"<b>FCF</b> 12 tháng gần nhất. Nhóm <b>Tài chính</b> dùng lợi nhuận ròng — tiền cho vay và tiền "
             f"gửi chạy qua dòng tiền hoạt động nên FCF không phản ánh kinh doanh. <b>REIT</b> dùng FCF khi "
             f"FCF ≥ lợi nhuận ròng (khấu hao bất động sản kéo lợi nhuận xuống dưới tiền thật thu về), ngược "
             f"lại dùng lợi nhuận. Mã thường có FCF dưới {conv} lần lợi nhuận cũng lùi về lợi nhuận và bị "
             f"đánh dấu <b>?</b>. Cả hai đều âm thì không định giá."),
            ("Giải tăng trưởng ngụ ý",
             f"Tìm g sao cho <b>Vốn hóa ÷ Dòng tiền</b> = Σ<sub>t=1…{V.YEARS}</sub> (1+g)<sup>t</sup> ÷ "
             f"(1+r)<sup>t</sup> + giá trị cuối kỳ, với r = {V.DISCOUNT:.0%} và sau năm {V.YEARS} dòng tiền "
             f"tăng {cls._pct(V.TERMINAL * 100)} mãi mãi (quanh mức GDP danh nghĩa). Giải bằng chia đôi trong "
             f"khoảng {V.G_LO:.0%} đến +{V.G_HI:.0%}/năm; chạm biên thì hiện "
             f"«{V.growth_label(V.G_HI * 100)}» hoặc «{V.growth_label(V.G_LO * 100)}» — khi đó chỉ là cận."),
            ("Lấy mốc đã chứng minh",
             f"Tăng trưởng doanh thu bình quân nhiều năm (khi quét kèm yfinance Moat), nếu không có thì doanh "
             f"thu 12 tháng. Mốc được kẹp trong {V.REF_LO:g}% … +{V.REF_HI:g}%/năm: doanh thu tăng 60% một năm "
             f"không có nghĩa là giữ được 60% suốt {V.YEARS} năm."),
            ("So sánh → nhãn",
             "<b>Chênh = ngụ ý − mốc</b>, tính bằng điểm phần trăm, rồi áp ngưỡng ở bảng nhãn bên dưới."),
            ("Xếp trong ngành",
             f"Lợi suất dòng tiền được xếp phân vị trong cùng ngành, cùng cơ sở (FCF với FCF, lợi nhuận với "
             f"lợi nhuận — ngân hàng không xếp chung với REIT), cần ít nhất {V.MIN_PEERS} mã. 100 = rẻ nhất "
             f"ngành. Chỉ để tham khảo, không đổi nhãn."),
        ]
        val_steps = ""
        for i, (name, how) in enumerate(steps, 1):
            bg = "#F0F4FA" if i % 2 else "#FFFFFF"
            val_steps += f"""
            <tr style="background-color:{bg};">
              <td style="color:{dim};font-weight:700;padding:8px 12px;">{i}</td>
              <td style="color:{accent};font-weight:700;font-size:12px;padding:8px 12px;">{name}</td>
              <td style="color:{text};font-size:12px;padding:8px 12px;line-height:1.55;">{how}</td>
            </tr>"""
        return val_rows, ex_rows, val_steps

    @staticmethod
    def _change_rows(accent, text):
        rows = ""
        for i, (lbl, desc) in enumerate([
            ("▲ từ X", "Tín hiệu lên hạng so với lần trước; X là tín hiệu cũ."),
            ("▼ từ X", "Tín hiệu xuống hạng. Thường đáng xem hơn mã đứng yên — xem tooltip để biết tiêu "
                       "chí nào vừa mất."),
            ("MỚI", "Lần quét trước không có mã này trong danh sách."),
            ("• N tiêu chí", "Tín hiệu giữ nguyên nhưng có N tiêu chí đổi kết quả (đạt thêm hoặc mất)."),
            ("• giá X", "Chỉ nhãn định giá đổi, X là nhãn mới."),
        ]):
            bg = "#F0F4FA" if i % 2 == 0 else "#FFFFFF"
            rows += f"""
            <tr style="background-color:{bg};">
              <td style="color:{accent};font-weight:700;font-size:12px;padding:8px 12px;font-family:Consolas,monospace;">{lbl}</td>
              <td style="color:{text};font-size:12px;padding:8px 12px;line-height:1.55;">{desc}</td>
            </tr>"""
        return rows + f"""
            <tr><td colspan="2" style="color:#64748B;font-size:11px;padding:8px 12px;font-style:italic;">
              Rê chuột vào ô để xem chi tiết: tín hiệu cũ → mới, điểm Tài chính và Thực, tiêu chí đạt thêm/mất,
              nhãn định giá và giá thay đổi.</td></tr>"""

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

        val_rows, ex_rows, val_steps = self._valuation_rows(H_DIM, H_BLUE, H_TEXT)
        chg_rows = self._change_rows(H_BLUE, H_TEXT)
        V = valuation
        horizons = "/".join(str(h) for h in signal_tracking.HORIZONS)

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
  hiển thị hai điểm cạnh nhau và gắn cờ đỏ cho trường hợp đó.<br><br>
  Với thị trường Mỹ còn câu thứ ba, cũng tách riêng: <b style="color:#16A34A;">Giá hôm nay có đang đòi hỏi
  quá nhiều không?</b> Câu này trả lời bằng nhãn <b>Định giá</b> (mục ⑦) — không cộng vào điểm, không đổi
  tín hiệu, vì một doanh nghiệp tốt và thật vẫn có thể đang bị mua quá đắt.
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
  <li><b>Định giá</b> — P/E, PEG chỉ để tham khảo; nhãn <b>Định giá</b> (mục ⑦) là một cột riêng, không
      bao giờ cộng vào điểm A/B hay đổi tín hiệu. Điểm trả lời «doanh nghiệp này có tốt và có thật không»;
      nhãn định giá trả lời «giá này đòi hỏi bao nhiêu» — trộn hai câu lại thì một mã kém nhưng rẻ sẽ
      trông giống một mã tốt.</li>
  <li><b>Hệ số nhân Conviction</b> — trước đây nhân điểm với chất lượng moat, mà moat lại suy ra từ ROE
      và biên gộp vốn đã được chấm rồi. Đó là đếm ba lần cùng một thứ.</li>
</ul>

<h2>⑦ ĐỊNH GIÁ NGƯỢC — REVERSE DCF (chỉ thị trường Mỹ)</h2>
<div class="box" style="background:#ECFDF3;border-left-color:#16A34A;">
  <b>Đặt câu hỏi ngược lại.</b> Hỏi «cổ phiếu này đáng giá bao nhiêu» thì kết quả phụ thuộc hoàn toàn vào
  giả định tăng trưởng mình tự đặt — chỉnh vài điểm phần trăm là ra mọi con số mình muốn. Câu hỏi ngược
  khách quan hơn nhiều: <b>với giá hôm nay, thị trường đang ngụ ý dòng tiền phải tăng bao nhiêu mỗi năm
  trong {V.YEARS} năm tới?</b> Rồi đặt con số đó cạnh mức doanh nghiệp đã thực sự làm được.
  Kết quả là một <i>mức tăng trưởng</i>, không phải một <i>giá mục tiêu</i>.
</div>

<p style="color:#64748B;font-size:12px;margin:0 0 6px 0;"><b>Các bước tính</b></p>
<table>
  <tr><th style="width:30px;">#</th><th style="width:150px;">Bước</th><th>Cách làm</th></tr>
  {val_steps}
</table>

<p style="color:#64748B;font-size:12px;margin:14px 0 6px 0;"><b>Đọc nhanh: bội số dòng tiền ↔ tăng trưởng ngụ ý</b>
  (chiết khấu {V.DISCOUNT:.0%}, {V.YEARS} năm, sau đó {self._pct(V.TERMINAL * 100)} mãi mãi)</p>
<table style="width:60%;">
  <tr><th>Vốn hóa / dòng tiền</th><th>Lợi suất dòng tiền</th><th>Giá đang ngụ ý tăng</th></tr>
  {ex_rows}
</table>
<p class="tip">Ví dụ: P/FCF 20 (lợi suất 5%) nghĩa là thị trường đang trả trước cho dòng tiền tăng chừng
  {V.growth_label(V.implied_growth(20))}/năm suốt {V.YEARS} năm. Doanh nghiệp tăng doanh thu 15%/năm thì
  giá đó là rẻ; tăng 3%/năm thì giá đó là đắt.</p>

<p style="color:#64748B;font-size:12px;margin:14px 0 6px 0;"><b>Nhãn định giá</b> — chênh = tăng trưởng
  ngụ ý − mốc đã chứng minh, tính bằng điểm phần trăm</p>
<table>
  <tr><th style="width:90px;">Nhãn</th><th style="width:250px;">Điều kiện</th><th>Ý nghĩa</th></tr>
  {val_rows}
</table>
<ul>
  <li><b>Dấu ?</b> sau nhãn = độ tin cậy thấp: phải lùi về lợi nhuận ròng vì dòng tiền thật đang thấp
      (FCF dưới {self._num(V.LOW_CASH_CONVERSION, 1)} lần lợi nhuận, hoặc REIT đang mua thêm tài sản), hoặc
      tăng trưởng gần đây vượt {V.HOT_GROWTH:g}%/năm — nếu đó là đỉnh chu kỳ thì dòng tiền hiện tại không bền
      và kết luận «rẻ» không đáng tin.</li>
  <li><b>Ngưỡng tuyệt đối</b> {V.ABS_RICH:g}% và {V.ABS_VERY_RICH:g}%: tăng {V.ABS_RICH:g}%/năm suốt
      {V.YEARS} năm nghĩa là dòng tiền gấp {self._num((1 + V.ABS_RICH / 100) ** V.YEARS, 1)} lần — rất ít
      doanh nghiệp cỡ lớn làm được, bất kể quá khứ nhanh đến đâu. Vì thế mã có mức ngụ ý này ít nhất là
      Đắt dù mốc quá khứ cũng cao.</li>
  <li><b>Cột trong bảng:</b> <b>Ngụ ý%</b> — tăng trưởng ngụ ý; <b>Lợi suất</b> — dòng tiền gốc / vốn
      hóa (5% ≈ P/FCF 20); <b>Định giá</b> — nhãn. Rê chuột vào ô Định giá để đọc cách tính của riêng mã
      đó, kèm thứ hạng lợi suất trong ngành.</li>
</ul>

<div class="box" style="background:#EEF4FF;border-left-color:#2563EB;">
  <b style="color:#1E40AF;">Từ nhãn ra mức giá — nút Giá trong khung biểu đồ.</b> Tăng trưởng ngụ ý chỉ phụ
  thuộc bội số vốn hóa / dòng tiền, mà bội số tỉ lệ thuận với giá khi số cổ phiếu không đổi. Nên mỗi ranh
  giới nhãn ứng với đúng một mức giá:
  <b>giá ranh giới = bội số ứng với mức tăng trưởng ở ranh giới × dòng tiền mỗi cổ phiếu</b>.
  Ba ranh giới là mốc − {V.CHEAP_GAP * -1:g} điểm (Rẻ | Hợp lý), mốc + {V.RICH_GAP:g} điểm nhưng không quá
  {V.ABS_RICH:g}% (Hợp lý | Đắt), mốc + {V.VERY_RICH_GAP:g} điểm nhưng không quá {V.ABS_VERY_RICH:g}%
  (Đắt | Rất đắt) — đúng luật của bảng nhãn ở trên. Đồ thị tô bốn dải màu theo các mức này:
  <b style="color:#16A34A;">Rẻ</b> · <b style="color:#2563EB;">Hợp lý</b> ·
  <b style="color:#B7791F;">Đắt</b> · <b style="color:#DC2626;">Rất đắt</b>, và dòng thứ hai trên đồ thị cho
  biết giá phải giảm/tăng bao nhiêu phần trăm thì đổi nhãn.
  <ul>
    <li><b>Chỉ đúng cho hiện tại.</b> Dòng tiền và mốc lấy từ lần quét gần nhất, nên các dải chỉ đọc cho giá
        hôm nay — không dùng để nói giá năm ngoái rẻ hay đắt (khi đó dòng tiền khác). Vì thế dải chỉ tô đậm
        ở mép phải đồ thị.</li>
    <li><b>Dải Hợp lý thường rất rộng</b> — mỗi điểm tăng trưởng suốt {V.YEARS} năm làm giá trị đổi nhiều.
        Nằm trong Hợp lý chưa phải điểm mua; gần mép Rẻ mới đáng chú ý hơn.</li>
    <li><b>Chỉ cổ phiếu Mỹ</b>, như nhãn. Mã thị trường khác vẫn xem được nến nhưng không có dải.</li>
    <li><b>Giá lệch xa giá lúc quét</b> (dưới một nửa hoặc trên gấp đôi — thường do chia tách cổ phiếu) thì
        không vẽ dải; quét lại để có số đúng. Mã có dấu ? thì dòng đầu ghi «kém tin cậy».</li>
    <li><b>Không phải tín hiệu mua bán.</b> Dải cho biết giá đang đòi hỏi tăng trưởng bao nhiêu; doanh nghiệp
        có làm được hay không vẫn là câu hỏi của nhóm A và B.</li>
  </ul>
</div>

<div class="box">
  <b style="color:#1E40AF;">Nguồn gốc phương pháp.</b> Đây là cách tiếp cận <i>Expectations Investing</i>
  của <b>Alfred Rappaport</b> và <b>Michael J. Mauboussin</b> (Harvard Business School Press, 2001; bản
  cập nhật 2021, Columbia Business School Publishing). Ý chính: đừng tự ước giá trị rồi so với giá — hãy
  đọc giá để biết thị trường đang kỳ vọng gì, rồi hỏi kỳ vọng đó có thực tế không. Mauboussin hiện làm
  nghiên cứu tại Counterpoint Global (Morgan Stanley Investment Management).
  Aswath Damodaran (NYU Stern) dùng cùng logic đọc ngược kỳ vọng từ giá, ví dụ khi tính mức bù rủi ro
  ngụ ý của S&amp;P 500.<br><br>
  Ngưỡng tuyệt đối mượn ý từ <i>The Base Rate Book</i> (Mauboussin, Callahan &amp; Majd, Credit Suisse,
  2016): số doanh nghiệp lớn giữ được tăng trưởng cao suốt mười năm là rất ít, nên kỳ vọng đó phải được
  coi là ngoại lệ chứ không phải mặc định.
</div>

<table cellspacing="0" cellpadding="0" style="margin:12px 0;"><tr>
<td style="background-color:#C2410C;width:3px;"></td>
<td style="background-color:#FFF6E5;padding:11px 15px;font-size:12px;line-height:1.6;">
  <b style="color:#C2410C;">Đơn giản hóa so với bản gốc — cần biết khi đọc nhãn.</b>
  <ul>
    <li><b>Chiết khấu cố định {V.DISCOUNT:.0%} cho mọi mã.</b> Bản gốc dùng chi phí vốn riêng từng doanh
        nghiệp (theo beta và cơ cấu nợ). Mã rủi ro cao vì thế được chiết khấu nhẹ tay hơn thực tế và trông
        rẻ hơn; mã rất ổn định thì ngược lại.</li>
    <li><b>Thời kỳ dự phóng cố định {V.YEARS} năm.</b> Bản gốc còn giải ra <i>thời kỳ tạo giá trị ngụ
        ý</i> — số năm thị trường tin doanh nghiệp còn sinh lời vượt chi phí vốn.</li>
    <li><b>Gộp mọi thứ vào một biến tăng trưởng dòng tiền.</b> Bản gốc tách thành các động lực giá trị —
        tăng trưởng doanh thu, biên lợi nhuận hoạt động, vốn đầu tư thêm — để biết kỳ vọng nào đang gánh
        giá.</li>
    <li><b>Dùng vốn hóa và dòng tiền 12 tháng gần nhất.</b> Bản gốc dùng giá trị doanh nghiệp (cộng nợ,
        trừ tiền mặt) với dòng tiền đã chuẩn hóa. Một năm đột biến, tốt hay xấu, kéo lệch kết quả.</li>
    <li><b>FCF chưa trừ thù lao bằng cổ phiếu (SBC).</b> Với nhiều công ty công nghệ, dòng tiền thật cho
        cổ đông thấp hơn đáng kể, nên nhãn có xu hướng nghiêng về rẻ.</li>
    <li><b>Mốc là tăng trưởng của chính doanh nghiệp</b>, không phải phân phối tăng trưởng theo quy mô và
        ngành như trong The Base Rate Book.</li>
  </ul>
  <b>Vì sao chỉ thị trường Mỹ:</b> {V.DISCOUNT:.0%} là lợi suất kỳ vọng dài hạn của cổ phiếu định giá bằng
  USD. Với VND, lãi suất phi rủi ro và lạm phát khác hẳn, còn nhóm Tài chính của TradingView ở Việt Nam
  gộp cả bất động sản — áp cùng tham số sẽ ra nhãn sai, nên ở thị trường khác các cột định giá để trống.
</td></tr></table>

<h2>⑧ THAY ĐỔI GIỮA CÁC LẦN QUÉT</h2>
<p style="color:#64748B;font-size:12px;margin:0 0 10px 0;">
  Một tín hiệu đứng yên không nói được nhiều: mã vừa rơi từ TỐT &amp; THỰC xuống KHÁ đáng xem hơn hẳn mã
  nằm ở KHÁ cả năm. Mỗi lần quét xong, chương trình lưu một ảnh chụp vào thư mục <b>snapshots</b> (cạnh
  file chạy), tên dạng <span style="font-family:Consolas,monospace;">YYYYMMDD_THỊTRƯỜNG_topN.csv.gz</span>,
  rồi đặt kết quả cạnh lần quét gần nhất <i>trước hôm nay</i> của cùng thị trường.
</p>
<table>
  <tr><th style="width:130px;">Cột Thay đổi</th><th>Ý nghĩa</th></tr>
  {chg_rows}
</table>
<ul>
  <li><b>Thứ hạng tín hiệu</b> dùng để xét lên/xuống: TỐT &amp; THỰC &gt; KHÁ &gt; THẬT NHƯNG YẾU &gt;
      THEO DÕI &gt; LOẠI &gt; CỜ ĐỎ.</li>
  <li><b>Dashboard</b> có mục 🔄 tóm tắt: mã lên hạng, xuống hạng, mới vào TỐT &amp; THỰC, mới thành
      CỜ ĐỎ, mã rời TỐT &amp; THỰC, và mã vào/ra danh sách.</li>
  <li><b>Cảnh báo tự động</b> khi hai lần quét khác phạm vi (số mã TOP khác nhau — mã vào/ra danh sách khi
      đó không được liệt kê), hoặc chỉ một lần có bật yfinance Moat — một phần thay đổi khi ấy đến từ việc
      có thêm/mất dữ liệu, không phải từ doanh nghiệp.</li>
</ul>

<h2>⑨ HIỆU QUẢ TÍN HIỆU — nút 📈 Hiệu quả</h2>
<p style="color:#64748B;font-size:12px;margin:0 0 10px 0;">
  Câu hỏi khó chịu nhất: <b>trong quá khứ, các nhãn này có thực sự đi trước lợi nhuận không?</b>
  Chỉ dùng ảnh chụp thật của chính máy này — chưa có lịch sử nghĩa là chưa kiểm được, không có số giả lập
  nào thay thế.
</p>
<ol>
  <li>Lấy ảnh chụp thị trường Mỹ, mỗi tháng một ảnh (ngày sớm nhất trong tháng, bản quét rộng nhất);
      một mã chỉ tính một lần mỗi tháng dù có trong cả ảnh AMERICA lẫn NASDAQ.</li>
  <li>Tải giá đóng cửa đã điều chỉnh (gồm cổ tức, chia tách) từ Yahoo Finance.</li>
  <li>Tính lợi nhuận mỗi mã sau {horizons} tháng kể từ ngày chụp, trừ đi S&amp;P 500 cùng kỳ
      → <b>phần vượt S&amp;P</b>.</li>
  <li>Gom theo tín hiệu và theo nhãn định giá: số quan sát, vượt S&amp;P trung bình và trung vị,
      <b>tỷ lệ thắng</b> (phần trăm quan sát vượt S&amp;P).</li>
</ol>
<p class="tip">Cần ít nhất một ảnh chụp cũ hơn 1 tháng mới có mốc đầu tiên — cứ quét đều mỗi tháng là số
  liệu tự tích lũy. Nhóm dưới {signal_tracking.MIN_N} quan sát in mờ, chỉ để tham khảo. Vài tháng dữ liệu
  vẫn còn là nhiễu: các mã chụp cùng ngày chịu chung một nhịp thị trường nên không phải các phép thử độc
  lập. Mã đã hủy niêm yết không còn giá bị loại ra và được liệt kê cuối báo cáo.</p>

<h2>⑩ CÁCH DÙNG</h2>
<ul>
  <li><b>SCAN</b> — lấy top N cổ phiếu theo vốn hóa từ TradingView và chấm cả hai nhóm.</li>
  <li><b>Tab Tài chính</b> — {N_FIN} tiêu chí nhóm A cùng số liệu gốc.</li>
  <li><b>Tab Kiểm chứng</b> — {N_REAL} tiêu chí nhóm B cùng các tỷ số dùng để kiểm chứng.
      Cả hai tab đều hiện đồng thời hai điểm và tín hiệu gộp.</li>
  <li><b>Dashboard</b> — bảng <b style="color:#B91C1C;">CỜ ĐỎ</b> nêu rõ từng mã chưa chứng minh được
      ở điểm nào; bảng <b>tỷ lệ đạt</b> cho biết tiêu chí nào đang quá lỏng hoặc quá chặt; mục
      <b>💲 ĐỊNH GIÁ</b> liệt kê mã TỐT &amp; THỰC đang Rẻ hoặc Hợp lý; mục <b>🔄 Thay đổi</b> tóm tắt
      khác biệt so với lần quét trước.</li>
  <li><b>Chip lọc</b> (🟢 TỐT &amp; THỰC · 🔴 CỜ ĐỎ · 💵 FCF ≥ lợi nhuận · 🏰 Bền vững 5 năm ·
      💲 Giá Rẻ / Hợp lý) — bật nhiều chip cùng lúc để thu hẹp danh sách. Bật 🟢 cùng 💲 để ra mã vừa
      chất lượng vừa chưa bị giá đòi hỏi quá nhiều.</li>
  <li><b>📈 Hiệu quả</b> — báo cáo kiểm chứng tín hiệu bằng giá thật (mục ⑨). Có thể đóng trong lúc đang
      tải giá; mở lại sẽ thấy kết quả.</li>
  <li><b>Khung biểu đồ</b> phía trên bảng — <b>Theo quý / Theo năm</b>: số liệu tài chính của mã đang
      chọn; <b>Giá</b>: nến giá (6T · 1N nến ngày, 3N · 5N nến tuần; mặc định 3 năm) với đường MA50, MA200,
      khối lượng, đỉnh 52 tuần và bốn dải giá Rẻ / Hợp lý / Đắt / Rất đắt theo định giá ngược (mục ⑦) —
      đang xem Giá thì chọn mã khác vẫn ở lại Giá; <b>Bản đồ</b>: cả danh sách trên bốn góc nhìn (Tài chính × Thực, Sinh lời × Lợi suất, Sinh lời
      × Tăng trưởng, Giá so với đỉnh) — bấm một chấm để chọn mã; <b>Quỹ</b>: số quỹ sở hữu theo quý từ hồ
      sơ SEC 13F (chỉ mã Mỹ; lần đầu phải chờ tải dữ liệu).</li>
  <li><b>Tab 📑 Phân tích</b> — <b>LẬP BÁO CÁO</b> cho mã đang chọn (cần đăng nhập Claude ở nút ✦);
      <b>CẬP NHẬT</b> lập lại với số liệu và tin tức mới nhất. <b>Tab 💼 Portfolio</b> — danh mục của
      bạn, lưu trong portfolio.json.</li>
  <li><b>yfinance Moat</b> — bật trước khi quét để có tiêu chí Bền vững 5 năm. Chậm hơn đáng kể
      (khoảng 0,3 giây mỗi mã) và chỉ có dữ liệu cho thị trường Mỹ.</li>
  <li><b>Click vào một dòng</b> — xem chi tiết từng tiêu chí; rê chuột lên ô ✓/✗ để đọc giải thích
      và giá trị thực tế.</li>
</ul>

<div class="box" style="background:#FFF6E5;border-left-color:#C2410C;">
  <b style="color:#C2410C;">Giới hạn cần biết.</b> Các ngưỡng ở đây được chọn theo chuẩn mực phổ biến,
  chưa được kiểm định trên dữ liệu lịch sử. Từ nay mỗi lần quét đều lưu ảnh chụp nên nút 📈 Hiệu quả sẽ
  kiểm dần được — nhưng phải sau nhiều tháng quét đều số liệu mới đủ để kết luận, và không thể kiểm lùi
  về trước lần quét đầu tiên. Ngưỡng biên lợi nhuận dùng chung cho mọi ngành nên vẫn thiên lệch về phía
  công nghệ và y tế. Và hệ thống chỉ trả lời «mua gì», không có quy tắc bán hay cắt lỗ.
  Hãy đọc điểm số như một thứ tự ưu tiên để nghiên cứu, không phải một khuyến nghị đầu tư.
</div>
</body></html>"""

# ─────────────────────────────────────────────────────────────────────────────
# AI Settings — một nút duy nhất để đăng nhập Claude Code mà Agent SDK dùng.
# Luồng đăng nhập thật nằm trong claude_auth.py.
# ─────────────────────────────────────────────────────────────────────────────
class _AuthWorker(QThread):
    """Chạy thao tác xác thực ngoài luồng giao diện.

    Đăng nhập mở trình duyệt rồi chờ người dùng xác nhận — có thể mất vài phút,
    nên không được gọi trên luồng chính.
    """
    done     = Signal(bool, str, object)    # thành công, thông điệp, trạng thái
    progress = Signal(str)                  # mô tả bước đang chạy

    def __init__(self, action, parent=None):
        super().__init__(parent)
        self._action = action

    def run(self):
        try:
            if self._action == "connect":
                # Lo đủ phụ thuộc rồi mở trình duyệt, trong một bước
                ok, msg = claude_auth.connect(progress=self.progress.emit)
            elif self._action == "logout":
                ok, msg = claude_auth.logout()
            else:
                ok, msg = True, ""
            self.done.emit(ok, msg, claude_auth.read_state())
        except Exception as e:              # luồng nền không được chết lặng
            self.done.emit(False, f"{type(e).__name__}: {e}", None)


class AISettingsDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("AI Settings")
        self.setFixedSize(430, 250)
        self._state  = None
        self._worker = None
        self._busy   = False

        self.setStyleSheet(f"""
            QDialog {{ background:{BG}; }}
            *       {{ background:transparent; color:{TEXT1};
                       font-family:'Segoe UI',sans-serif; }}
        """)

        root = QVBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(0)

        head = QWidget(); head.setFixedHeight(48)
        head.setStyleSheet(f"background:{SURFACE}; border-bottom:1px solid {BORDER};")
        hh = QHBoxLayout(head); hh.setContentsMargins(20, 0, 20, 0)
        t = QLabel("AI  SETTINGS")
        t.setStyleSheet(f"color:{TEXT1}; font-size:12px; font-weight:700; letter-spacing:3px;")
        hh.addWidget(t); hh.addStretch()
        root.addWidget(head)

        body = QWidget(); body.setStyleSheet(f"background:{BG};")
        bv = QVBoxLayout(body)
        bv.setContentsMargins(26, 20, 26, 16)
        bv.setSpacing(14)

        row = QHBoxLayout(); row.setSpacing(9)
        row.addStretch()
        self._dot = QLabel("●"); self._dot.setFixedWidth(12)
        self._lbl_state = QLabel()
        self._lbl_state.setStyleSheet(f"color:{TEXT1}; font-size:13px; font-weight:600;")
        row.addWidget(self._dot); row.addWidget(self._lbl_state)
        row.addStretch()
        bv.addLayout(row)

        self._btn = QPushButton()
        self._btn.setFixedHeight(38)
        self._btn.setCursor(Qt.PointingHandCursor)
        self._btn.clicked.connect(self._on_click)
        bv.addWidget(self._btn)

        self._lbl_note = QLabel()
        self._lbl_note.setWordWrap(True)
        self._lbl_note.setAlignment(Qt.AlignCenter)
        self._lbl_note.setStyleSheet(f"color:{TEXT3}; font-size:10px; line-height:150%;")
        bv.addWidget(self._lbl_note)
        bv.addStretch()
        root.addWidget(body, stretch=1)

        foot = QWidget(); foot.setFixedHeight(44)
        foot.setStyleSheet(f"background:{SURFACE}; border-top:1px solid {BORDER};")
        fh = QHBoxLayout(foot); fh.setContentsMargins(20, 0, 20, 0)
        btn_close = QPushButton("ĐÓNG")
        btn_close.setFixedSize(84, 27)
        btn_close.setCursor(Qt.PointingHandCursor)
        btn_close.setStyleSheet(
            f"QPushButton {{ background:{BORDER2}; color:{TEXT1}; border:none;"
            f" border-radius:3px; font-size:10px; font-weight:700; letter-spacing:1.2px; }}"
            f"QPushButton:hover {{ background:{BLUE}; color:#FFFFFF; }}")
        btn_close.clicked.connect(self.close)
        fh.addStretch(); fh.addWidget(btn_close)
        root.addWidget(foot)

        self._state = claude_auth.read_state()
        self._render()

    # ── Hiển thị ───────────────────────────────────────────────────────────
    def _style(self, bg, hover, fg="#FFFFFF"):
        return (f"QPushButton {{ background:{bg}; color:{fg}; border:none;"
                f" border-radius:4px; font-size:11px; font-weight:700; letter-spacing:1px; }}"
                f"QPushButton:hover {{ background:{hover}; }}"
                f"QPushButton:disabled {{ background:{BORDER2}; color:{TEXT3}; }}")

    def _render(self):
        st = self._state
        self._btn.setEnabled(True)

        if st is not None and st.logged_in:
            # Đăng nhập rồi nhưng chạy thử hỏng (vd. Claude Code quá cũ) thì báo vàng
            self._dot.setStyleSheet(f"color:{GREEN if st.verified else AMBER}; font-size:13px;")
            who = " · ".join(x for x in (st.email, st.plan.capitalize()) if x)
            self._lbl_state.setText(f"Đã liên kết · {who}" if who else "Đã liên kết")
            self._btn.setText("NGẮT LIÊN KẾT")
            self._btn.setStyleSheet(self._style(BORDER2, RED, TEXT1))
            note = ("App dùng tài khoản Claude đã đăng nhập, tiêu dung lượng gói thuê bao "
                    "chứ không tính vào tín dụng API. Ngắt liên kết sẽ đăng xuất Claude "
                    "Code trên cả máy này." if st.verified
                    else f"Đã đăng nhập nhưng chạy thử lỗi: {st.error}")
        else:
            self._dot.setStyleSheet(f"color:{TEXT3}; font-size:13px;")
            self._lbl_state.setText("Chưa liên kết")
            self._btn.setText("ĐĂNG NHẬP CLAUDE")
            self._btn.setStyleSheet(self._style(BLUE, BLUE_HV))
            note = ("Bấm một lần: app tự cài công cụ cần thiết rồi mở trình duyệt "
                    "để bạn xác nhận bằng tài khoản Claude. Tính năng AI sẽ tiêu dung "
                    "lượng gói thuê bao, không cần nạp tín dụng API.")

        # Biến môi trường này che mọi profile OAuth, kể cả khi là chuỗi rỗng
        if st is not None and st.api_key_shadow:
            note = ("ANTHROPIC_API_KEY đang được đặt và sẽ che tài khoản OAuth. "
                    "Xoá biến đó nếu muốn dùng đăng nhập này.")
        self._lbl_note.setText(note)

    # ── Hành động ──────────────────────────────────────────────────────────
    def _on_click(self):
        if self._busy or self._state is None:
            return

        action = "logout" if self._state.logged_in else "connect"
        self._busy = True
        self._btn.setEnabled(False)
        self._btn.setText("ĐANG NGẮT…" if action == "logout" else "ĐANG LIÊN KẾT…")
        self._lbl_note.setText("Đang ngắt liên kết…" if action == "logout"
                               else "Đang chuẩn bị…")
        self._worker = _AuthWorker(action, self)
        self._worker.progress.connect(self._lbl_note.setText)
        self._worker.done.connect(self._on_done)
        self._worker.start()

    def _on_done(self, ok, msg, state):
        self._busy = False
        if state is not None:
            self._state = state
        self._render()
        if not ok and msg:
            self._lbl_note.setText(msg.strip().splitlines()[-1][:160])

    def closeEvent(self, event):
        # Không để luồng nền sống sót sau khi dialog đóng
        if self._worker is not None and self._worker.isRunning():
            self._worker.wait(2000)
        super().closeEvent(event)

    def keyPressEvent(self, event):
        if event.key() == Qt.Key_Escape:
            self.close()
        else:
            super().keyPressEvent(event)

# ─────────────────────────────────────────────────────────────────────────────
# Tab Phân tích — gọi Claude giải thích số liệu và tìm tin tức.
# Logic nằm trong ai_analysis.py; ở đây là giao diện và luồng nền.
# ─────────────────────────────────────────────────────────────────────────────
class _AnalysisWorker(QThread):
    delta  = Signal(str)        # mẩu chữ hiện dần
    status = Signal(str)        # bước đang chạy
    done   = Signal(object)     # dict kết quả
    failed = Signal(str)

    def __init__(self, context, ticker, company, model, parent=None):
        super().__init__(parent)
        self._ctx, self._tk = context, ticker
        self._co, self._model = company, model

    def run(self):
        try:
            res = ai_analysis.analyze(
                self._ctx, self._tk, self._co, model=self._model,
                on_delta=self.delta.emit, on_status=self.status.emit)
            self.done.emit(res)
        except Exception as e:
            self.failed.emit(ai_analysis.friendly_error(e))


class AnalysisPanel(QWidget):
    """Hiển thị phân tích định tính cho mã đang chọn."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self._row    = None
        self._worker = None
        self._busy   = False
        self._dark   = False
        self._res    = None     # báo cáo đang hiện, để vẽ lại khi đổi giao diện
        self._cached = False

        self.setStyleSheet(f"background:{BG};")
        v = QVBoxLayout(self)
        v.setContentsMargins(0, 0, 0, 0)
        v.setSpacing(0)

        # ── Thanh trên ─────────────────────────────────────────────────────
        bar = QWidget(); bar.setFixedHeight(74)
        bar.setStyleSheet(f"background:{SURFACE}; border-bottom:1px solid {BORDER};")
        self._bar = bar
        bv = QVBoxLayout(bar)
        bv.setContentsMargins(22, 9, 20, 9); bv.setSpacing(4)

        # Dòng trên: mã, tên, nút
        r1 = QHBoxLayout(); r1.setSpacing(10)
        self._lbl_tk = QLabel("—")
        self._lbl_co = QLabel("")
        r1.addWidget(self._lbl_tk); r1.addWidget(self._lbl_co); r1.addStretch()

        self._btn_run = QPushButton("LẬP BÁO CÁO")
        self._btn_run.setFixedSize(116, 30)
        self._btn_run.setCursor(Qt.PointingHandCursor)
        self._btn_run.clicked.connect(lambda: self._run(force=False))
        r1.addWidget(self._btn_run)

        self._btn_again = QPushButton("CẬP NHẬT")
        self._btn_again.setFixedSize(92, 30)
        self._btn_again.setToolTip("Lập lại báo cáo với số liệu và tin tức mới nhất")
        self._btn_again.setCursor(Qt.PointingHandCursor)
        self._btn_again.clicked.connect(lambda: self._run(force=True))
        r1.addWidget(self._btn_again)
        bv.addLayout(r1)

        # Dòng dưới: giữ điểm số và tín hiệu trước mắt khi đang đọc phân tích
        r2 = QHBoxLayout(); r2.setSpacing(8)
        self._lbl_sig = QLabel("")
        self._lbl_sig.setStyleSheet(
            f"font-size:10px; font-weight:700; padding:2px 9px; border-radius:3px;")
        self._lbl_score = QLabel("")
        r2.addWidget(self._lbl_sig); r2.addWidget(self._lbl_score)
        r2.addStretch()
        self._lbl_meta = QLabel("")
        r2.addWidget(self._lbl_meta)
        bv.addLayout(r2)
        v.addWidget(bar)

        # ── Nội dung ───────────────────────────────────────────────────────
        self._view = QTextBrowser()
        self._view.setOpenExternalLinks(True)

        # Đồ thị bằng chứng: trả lời "mã này đứng ở đâu so với phần còn lại",
        # tức đúng thứ mà phần chữ đang khẳng định.
        self._eviden = QWebEngineView()
        self._evi_df = None

        split = QSplitter(Qt.Horizontal)
        split.setStyleSheet(f"QSplitter::handle {{ background:{BORDER}; width:1px; }}")
        split.addWidget(self._view)
        split.addWidget(self._eviden)
        split.setSizes([880, 620])
        split.setStretchFactor(0, 1)
        split.setStretchFactor(1, 1)
        self._split = split
        v.addWidget(split, stretch=1)

        self.apply_theme()
        self._render_empty()

    # ── Giao diện ──────────────────────────────────────────────────────────
    def apply_theme(self, dark=None):
        if dark is not None:
            self._dark = dark
        self.setStyleSheet(f"background:{BG};")
        self._split.setStyleSheet(f"QSplitter::handle {{ background:{BORDER}; width:1px; }}")
        if self._row is not None:
            self._render_evidence(self._row)
        self._bar.setStyleSheet(f"background:{SURFACE}; border-bottom:1px solid {BORDER};")
        self._lbl_tk.setStyleSheet(
            f"color:{TEXT1}; font-size:18px; font-weight:700; border:none;"
            f" font-family:'Segoe UI',sans-serif; letter-spacing:1px;")
        self._lbl_co.setStyleSheet(f"color:{TEXT2}; font-size:12px; border:none;")
        self._lbl_score.setStyleSheet(f"color:{TEXT2}; font-size:11px; border:none;")
        self._lbl_meta.setStyleSheet(f"color:{TEXT3}; font-size:10px; border:none;")
        # Báo cáo đặt trên nền "giấy" cùng màu khung đồ thị bên cạnh
        self._view.setStyleSheet(
            f"QTextBrowser {{ background:{SURFACE}; color:{TEXT1}; border:none;"
            f" padding:24px 34px; font-family:'Segoe UI',sans-serif; font-size:13px; }}")
        _b = (f"QPushButton {{ background:{BLUE}; color:#FFFFFF; border:none;"
              f" border-radius:3px; font-size:10px; font-weight:700; letter-spacing:1px; }}"
              f"QPushButton:hover {{ background:{BLUE_HV}; }}"
              f"QPushButton:disabled {{ background:{BORDER2}; color:{TEXT3}; }}")
        _f = (f"QPushButton {{ background:transparent; color:{TEXT1};"
              f" border:1px solid {BORDER2}; border-radius:3px;"
              f" font-size:10px; font-weight:700; letter-spacing:1px; }}"
              f"QPushButton:hover {{ border-color:{BLUE}; color:{BLUE}; }}"
              f"QPushButton:disabled {{ color:{TEXT3}; border-color:{BORDER}; }}")
        self._btn_run.setStyleSheet(_b)
        self._btn_again.setStyleSheet(_f)
        if not self._busy and hasattr(self, "_view"):
            if self._res:
                self._show(self._res, self._cached)
            elif self._row is not None:
                self._render_empty()

    def _pal(self) -> dict:
        """Bảng màu của giao diện hiện hành — các hằng màu đổi theo theme."""
        g = globals()
        return {k: g[k] for k in ("BG", "SURFACE", "BORDER", "BORDER2", "BLUE",
                                  "TEXT1", "TEXT2", "TEXT3", "GREEN", "RED", "AMBER")}

    def _render_empty(self):
        self._res = None
        self._view.setHtml(analysis_report.empty(self._pal(), self._row))

    # ── Dữ liệu ────────────────────────────────────────────────────────────
    def _fill_context(self, row):
        """Giữ tín hiệu và hai điểm trước mắt trong lúc đọc phân tích."""
        if not row:
            self._lbl_sig.setText(""); self._lbl_sig.setStyleSheet("")
            self._lbl_score.setText("")
            return
        sig = row.get("Signal", "") or ""
        if sig in SIGNAL_COLORS:
            fg, bg = SIGNAL_COLORS[sig]
            self._lbl_sig.setText(f"  {sig}  ")
            self._lbl_sig.setStyleSheet(
                f"color:{fg}; background:{bg}; font-size:10px; font-weight:700;"
                f" padding:2px 9px; border-radius:3px;")
        else:
            self._lbl_sig.setText(""); self._lbl_sig.setStyleSheet("")
        self._lbl_score.setText(
            f"Tài chính {row.get('FIN_Score','?')}/{row.get('FIN_App','?')}"
            f"   ·   Thực {row.get('REAL_Score','?')}/{row.get('REAL_App','?')}"
            f"   ·   {row.get('Sector','') or ''}")

    def set_row(self, row: dict):
        """Đổi mã đang xem. Nếu có bản lưu trong ngày thì hiện luôn."""
        if self._busy:
            return
        self._row = row
        tk = row.get("Ticker", "—") if row else "—"
        self._lbl_tk.setText(tk)
        self._lbl_co.setText(str(row.get("Tên Công Ty", ""))[:60] if row else "")
        self._fill_context(row)
        self._render_evidence(row)
        cached = ai_analysis.load_cached(tk) if row else None
        if cached:
            self._show(cached, from_cache=True)
        else:
            self._lbl_meta.setText("")
            self._render_empty()
        self._btn_run.setEnabled(bool(row))
        self._btn_again.setEnabled(bool(cached))

    def set_universe(self, df):
        """Nhận kết quả quét để xếp hạng mã trong nhóm."""
        self._evi_df = df
        if self._row is not None:
            self._render_evidence(self._row)

    def _render_evidence(self, row):
        try:
            self._eviden.setHtml(analysis_charts.build_html(
                row, self._evi_df, dark=getattr(self, "_dark", False)))
        except Exception:
            pass

    def _show(self, res: dict, from_cache=False):
        from datetime import datetime
        self._res, self._cached = res, from_cache
        when = (res.get("created", "") or "")[:16]
        try:
            when = datetime.fromisoformat(when).strftime("%H:%M  %d/%m/%Y")
        except ValueError:
            when = when.replace("T", " ")
        bits = [f"Lập lúc {when}"]
        if res.get("searches"):
            bits.append(f"{res['searches']} lượt tìm kiếm")
        self._lbl_meta.setText("   ·   ".join(bits))
        # Thông tin kỹ thuật để trong chú thích, không chiếm chỗ trên báo cáo
        tech = [("Bản lưu trong ngày" if from_cache else "Vừa lập")]
        if res.get("model"):
            tech.append(f"Mô hình: {res['model']}")
        if res.get("cost_usd"):
            tech.append(f"Chi phí ước tính: ${res['cost_usd']:.3f}")
        self._lbl_meta.setToolTip("\n".join(tech))
        note = ("Báo cáo bị cắt vì đạt giới hạn độ dài — bấm CẬP NHẬT để lập lại."
                if res.get("truncated") else "")
        self._view.setHtml(analysis_report.build(
            res.get("markdown", ""), self._pal(), self._row, res, note))

    # ── Chạy ───────────────────────────────────────────────────────────────
    def _run(self, force=False):
        if self._busy or not self._row:
            return
        tk = self._row.get("Ticker", "")
        if not force:
            cached = ai_analysis.load_cached(tk)
            if cached:
                self._show(cached, from_cache=True)
                return

        try:
            ctx = ai_analysis.build_context(
                self._row, FIN_CRITERIA, FIN_KEYS, REAL_CRITERIA, REAL_KEYS)
        except Exception as e:
            self._fail(f"Không dựng được dữ liệu: {e}")
            return

        self._busy = True
        self._btn_run.setEnabled(False)
        self._btn_again.setEnabled(False)
        self._buf = []
        self._res = None
        self._lbl_meta.setText("Đang lập báo cáo…")
        self._lbl_meta.setToolTip("")
        self._view.setHtml(analysis_report.running(self._pal(), self._row))

        self._worker = _AnalysisWorker(
            ctx, tk, str(self._row.get("Tên Công Ty", "")),
            ai_analysis.MODEL_DEFAULT, self)
        self._worker.status.connect(self._on_status)
        self._worker.delta.connect(self._on_delta)
        self._worker.done.connect(self._on_done)
        self._worker.failed.connect(self._fail)
        self._worker.start()

    def _on_status(self, msg):
        self._lbl_meta.setText(msg)

    def _on_delta(self, piece):
        self._buf.append(piece)
        # Dựng lại từ đầu mỗi mẩu thì chậm; chỉ vẽ khi vừa xong một đoạn
        if piece.endswith("\n") or len(self._buf) % 40 == 0:
            self._view.setHtml(analysis_report.build(
                "".join(self._buf), self._pal(), self._row, streaming=True))
            sb = self._view.verticalScrollBar()
            sb.setValue(sb.maximum())

    def _on_done(self, res):
        self._busy = False
        self._btn_run.setEnabled(True)
        self._btn_again.setEnabled(True)
        ai_analysis.save_cache(res.get("ticker", ""), res)
        self._show(res)

    def _fail(self, msg):
        self._busy = False
        self._btn_run.setEnabled(bool(self._row))
        self._btn_again.setEnabled(False)
        self._res = None
        self._lbl_meta.setText("")
        self._lbl_meta.setToolTip("")
        self._view.setHtml(analysis_report.error(self._pal(), msg))


class ValidationWorker(QThread):
    """Đối chiếu ảnh chụp cũ với giá thật — tải giá từ Yahoo nên chạy nền."""
    done   = Signal(object)    # signal_tracking.Report
    failed = Signal(str)

    def __init__(self, folder, parent=None):
        super().__init__(parent)
        self.folder = folder

    def run(self):
        try:
            # Tải giá theo lô và dừng được giữa các lô khi ứng dụng thoát
            fetch = lambda tickers, start: signal_tracking.download_prices(
                tickers, start, stop=self.isInterruptionRequested)
            self.done.emit(signal_tracking.run_validation(self.folder, downloader=fetch))
        except Exception as e:
            self.failed.emit(str(e))


class ValidationDialog(QDialog):
    """Các nhãn tín hiệu trong quá khứ có thật sự đi trước lợi nhuận không?

    Không xoá khi đóng (chỉ ẩn): luồng tải giá có thể còn chạy và Qt không
    được huỷ một luồng đang chạy."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Hiệu quả tín hiệu")
        self.resize(920, 640)
        self._worker  = None
        self._report  = None
        self._error   = ""
        self._loading = False

        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)
        self._browser = QTextBrowser()
        self._browser.setOpenExternalLinks(False)
        layout.addWidget(self._browser, stretch=1)

        self._foot = QWidget()
        self._foot.setObjectName("validFoot")   # để nền/viền không lan xuống nút
        self._foot.setFixedHeight(44)
        fh = QHBoxLayout(self._foot)
        fh.setContentsMargins(18, 0, 18, 0)
        self._hint = QLabel("")
        self._btn_refresh = QPushButton("LÀM MỚI")
        self._btn_refresh.setFixedSize(90, 28)
        self._btn_refresh.setCursor(Qt.PointingHandCursor)
        self._btn_refresh.clicked.connect(self.refresh)
        btn_close = QPushButton("ĐÓNG")
        btn_close.setFixedSize(80, 28)
        btn_close.setCursor(Qt.PointingHandCursor)
        btn_close.clicked.connect(self.close)
        fh.addWidget(self._hint, stretch=1)
        fh.addWidget(self._btn_refresh)
        fh.addSpacing(8)
        fh.addWidget(btn_close)
        layout.addWidget(self._foot)

        app = QApplication.instance()
        if app is not None:
            app.aboutToQuit.connect(self.shutdown)
        self.apply_theme()
        self.refresh()

    def _running(self) -> bool:
        return self._worker is not None and self._worker.isRunning()

    def refresh(self):
        if self._running():
            return
        self._report, self._error = None, ""
        try:
            folder = snapshot_dir()
        except OSError as e:
            self._error = f"không mở được thư mục snapshots ({e})"
            self._render()
            return
        self._hint.setText(f"Ảnh chụp: {folder}")
        self._hint.setToolTip(str(folder))
        self._loading = True
        self._btn_refresh.setEnabled(False)
        self._render()
        if self._worker is not None:
            self._worker.deleteLater()
        self._worker = ValidationWorker(folder, self)
        self._worker.done.connect(self._on_done)
        self._worker.failed.connect(self._on_failed)
        self._worker.start()

    def _on_done(self, rep):
        self._report, self._loading = rep, False
        self._btn_refresh.setEnabled(True)
        self._render()

    def _on_failed(self, msg):
        self._error, self._loading = msg, False
        self._btn_refresh.setEnabled(True)
        self._render()

    def _render(self):
        if self._loading:
            html = (f'<p style="color:{TEXT2};font-family:Segoe UI;">Đang đọc ảnh chụp và tải giá '
                    f'từ Yahoo Finance… (vài chục giây với vài trăm mã)</p>')
        elif self._error:
            html = (f'<p style="color:{RED};font-family:Segoe UI;">Không chạy được kiểm chứng: '
                    f'{_html_escape(self._error)}</p>')
        elif self._report is not None:
            html = signal_tracking.report_html(self._report, {
                "text": TEXT1, "dim": TEXT2, "faint": TEXT3, "green": GREEN, "red": RED,
                "amber": AMBER, "border": BORDER, "head": PANEL})
        else:
            html = ""
        self._browser.setHtml(html)

    def apply_theme(self):
        self.setStyleSheet(f"""
            QDialog      {{ background:{BG}; }}
            QTextBrowser {{
                background:{SURFACE}; color:{TEXT1};
                border:none; padding:18px 24px;
                font-family:'Segoe UI',sans-serif;
            }}
            QPushButton  {{
                background:{BLUE}; color:#FFFFFF;
                border:none; border-radius:3px;
                font-size:10px; font-weight:700; letter-spacing:1.5px;
                font-family:'Segoe UI',sans-serif;
            }}
            QPushButton:hover    {{ background:{BLUE_HV}; }}
            QPushButton:disabled {{ background:{BORDER2}; color:{TEXT2}; }}
        """)
        self._foot.setStyleSheet(
            f"#validFoot {{ background:{SURFACE}; border-top:1px solid {BORDER}; }}")
        self._hint.setStyleSheet(f"color:{TEXT3}; font-size:9px;")
        self._render()

    def shutdown(self):
        """Xin dừng luồng tải giá rồi chờ có giới hạn — Qt không được huỷ một luồng
        đang chạy. Mạng treo thì cắt hẳn: đằng nào ứng dụng cũng đang thoát."""
        if not self._running():
            return
        self._worker.requestInterruption()
        if not self._worker.wait(15_000):
            self._worker.terminate()
            self._worker.wait(2_000)


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
        self._changes     = None   # tóm tắt so với lần quét trước (signal_tracking)
        self._valid_dlg   = None
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

        # Biểu đồ: dựng trước vì tab Dashboard sẽ gắn nó vào
        self._chart_panel = ChartPanel()   # được gắn vào trong tab Dashboard
        self._chart_panel.ticker_picked.connect(self._on_map_pick)
        self._chart_panel.set_market(self._market.currentText())
        self._market.currentTextChanged.connect(self._chart_panel.set_market)

        # Các tab
        self._tabs_w = self._build_tabs()
        self._tabs_w.currentChanged.connect(self._on_tab_changed)


        # Biểu đồ nay nằm trong tab Dashboard, không còn chiếm chỗ cố định
        # bên phải nữa — các tab bảng vì thế có trọn chiều ngang.
        h_split = QSplitter(Qt.Horizontal)
        h_split.setStyleSheet(
            f"QSplitter::handle {{ background:{BORDER}; width:2px; }}")
        h_split.addWidget(self._tabs_w)
        h_split.setSizes([1500])
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

        self._btn_valid = QPushButton("📈  Hiệu quả")
        self._btn_valid.setFixedSize(105, 32)
        self._btn_valid.setCursor(Qt.PointingHandCursor)
        self._btn_valid.setToolTip(
            "<b>Hiệu quả tín hiệu</b><br>Đối chiếu các lần quét cũ (thị trường Mỹ) với giá thật "
            "sau 1/3/6/12 tháng, so với S&amp;P 500 cùng kỳ. Cần ít nhất một ảnh chụp cũ hơn 1 tháng.")
        self._btn_valid.setStyleSheet(self._btn_style("#5B3A8C", "#472C70"))
        self._btn_valid.clicked.connect(self._show_validation)

        row1.addWidget(lbl_market); row1.addWidget(self._market)
        row1.addWidget(lbl_top)
        row1.addWidget(self._spin_minus)
        row1.addWidget(self._top_spin)
        row1.addWidget(self._spin_plus)
        row1.addWidget(self._use_yf)
        row1.addStretch()
        row1.addWidget(self._btn_scan)
        row1.addWidget(self._btn_export)
        row1.addWidget(self._btn_valid)
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
        self._chk_val  = QCheckBox("💲  Giá Rẻ / Hợp lý")
        self._chk_val.setToolTip(
            "<b>Định giá Rẻ hoặc Hợp lý</b> (thị trường Mỹ).<br>Bật cùng 🟢 TỐT &amp; THỰC để ra "
            "mã vừa chất lượng vừa chưa bị giá đòi hỏi quá nhiều.")
        self._chip_checks = (self._chk_sb, self._chk_moat, self._chk_1y, self._chk_52w,
                             self._chk_val)

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
        tabs.addTab(self._cs_w, "  Tài chính  ")

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
        tabs.addTab(self._qc_w, "  Kiểm chứng  ")

        # Tab 3 — My Portfolio
        self._portfolio_panel = PortfolioPanel()
        tabs.addTab(self._portfolio_panel, "  💼 Portfolio  ")
        self._ai_panel = AnalysisPanel()
        tabs.addTab(self._ai_panel, "  📑 Phân tích  ")

        return tabs

    def _build_dashboard_panel(self):
        w = QWidget(); w.setStyleSheet(f"background:{BG};")
        lay = QVBoxLayout(w); lay.setContentsMargins(0,0,0,0); lay.setSpacing(0)
        self._dash_browser = QTextBrowser()
        self._dash_browser.setOpenExternalLinks(False)
        self._dash_browser.setOpenLinks(False)       # tự xử lý thay vì mở trình duyệt
        self._dash_browser.anchorClicked.connect(self._on_dash_link)
        self._dash_browser.setStyleSheet(f"""
            QTextBrowser {{
                background:{BG}; color:#DCE4EE;
                border:none; padding:16px 20px;
                font-family:'Segoe UI',sans-serif;
            }}
        """)
        self._dash_browser.setHtml(self._dash_placeholder())

        split = QSplitter(Qt.Horizontal)
        split.setStyleSheet(f"QSplitter::handle {{ background:{BORDER}; width:1px; }}")
        split.addWidget(self._dash_browser)
        split.addWidget(self._chart_panel)
        split.setSizes([700, 800])   # biểu đồ cần bề ngang cho hai cột ô
        # Cả hai cùng co giãn, nếu không thì phóng to cửa sổ chỉ nới bảng
        split.setStretchFactor(0, 1)
        split.setStretchFactor(1, 1)
        self._dash_split = split
        lay.addWidget(split)
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
                         "NI_Trend_Gap", "LN quý ($B)", "Nợ ròng/FCF",
                         "ΔBiên gộp", "ΔROIC%", "Dốc biên", "Vị trí biên"):
                item.setBackground(QColor(_fu_bg)); item.setForeground(QColor(_fu_fg))
        self._qc_table.horizontalHeader().update()

    def _populate_qc_table(self, df: pd.DataFrame):
        self._qc_data = {}
        self._qc_table.setSortingEnabled(False)
        self._qc_table.setRowCount(0)
        self._qc_table.setRowCount(len(df))

        # Columns that are %-type and get green/red coloring
        PCT_GREEN_RED = {"Gross Margin%", "Op Margin%", "ROIC%", "1Y%", "FCF_Margin%",
                         "NI Qtr%", "Rev Qtr%", "NI_Trend_Gap",
                         "ΔBiên gộp", "ΔROIC%", "Dốc biên"}
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
                elif key == "Định giá":
                    cell = _valuation_cell(rd)
                elif key == "Thay đổi":
                    cell = _change_cell(rd)
                elif key == "Tăng trưởng ngụ ý%":
                    cell = _implied_cell(val)
                elif key == "Lợi suất%":
                    cell = _pct_cell(val, "{:.1f}%")
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
                    if _tip and not cell.toolTip():
                        cell.setToolTip(_tip)
                    self._qc_table.setItem(ri, ci, cell)

        self._qc_table.setSortingEnabled(True)
        self._apply_qc_filter()

    def _recolor_qc_table(self):
        PCT_GREEN_RED = {"Gross Margin%", "Op Margin%", "ROIC%", "1Y%", "FCF_Margin%",
                         "NI Qtr%", "Rev Qtr%", "NI_Trend_Gap",
                         "ΔBiên gộp", "ΔROIC%", "Dốc biên"}
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
            self._ai_panel.set_row(rd)

    def _on_tab_changed(self, idx: int):
        self._apply_filter()

    def _apply_qc_filter(self):
        text      = self._filter_inp.text().strip().lower()
        only_sb   = self._chk_sb.isChecked()
        only_moat = self._chk_moat.isChecked()
        only_1y   = self._chk_1y.isChecked()
        only_52w  = self._chk_52w.isChecked()
        only_val  = self._chk_val.isChecked()
        any_chip  = only_sb or only_moat or only_1y or only_52w or only_val
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
                if show and only_val and rd.get("Định giá") not in ("Rẻ", "Hợp lý"):
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
                         "Gross Margin%", "Op Margin%", "ROIC%", "ROE%", "D/E", "P/E", "1Y%",
                         "Tăng trưởng ngụ ý%", "Lợi suất%", "Định giá"):
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
        self._btn_valid.setStyleSheet(self._btn_style("#5B3A8C", "#472C70"))
        if self._valid_dlg is not None:
            self._valid_dlg.apply_theme()
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
        if hasattr(self, "_ai_panel"):
            self._ai_panel.apply_theme(dark)
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
                "changes": self._changes,
            }
            with open(self._cache_path(), "wb") as f:
                pickle.dump(payload, f)
        except Exception:
            pass

    def _cache_fallback(self):
        """(df, saved_at, market) của lần quét trong bộ nhớ đệm — phải đọc
        trước khi lần quét mới ghi đè lên nó."""
        import pickle
        try:
            with open(self._cache_path(), "rb") as f:
                payload = pickle.load(f)
            return payload.get("df"), payload.get("saved_at"), payload.get("market", "")
        except Exception:
            return None

    def _track_changes(self, df, market: str):
        """(df có cột Thay đổi, tóm tắt, lỗi). Lỗi ở đây không được chặn kết
        quả quét — khi đó trả lại df nguyên vẹn kèm lời nhắn."""
        try:
            prev, prev_day = signal_tracking.load_previous(
                snapshot_dir(), market, fallback=self._cache_fallback())
            out = signal_tracking.annotate_changes(df, prev, prev_day, _CHANGE_LABELS)
            return out, signal_tracking.change_summary(out, prev, prev_day), ""
        except Exception as e:
            return df, None, f"không so được với lần quét trước ({str(e)[:120]})"

    def _changes_suffix(self, err: str = "") -> str:
        """Đuôi dòng trạng thái sau khi quét."""
        if err:
            return f"  ·  ⚠ {err}"
        ch = self._changes
        if not ch:
            return "  ·  chưa có lần quét trước để so"
        d = pd.Timestamp(ch["prev_date"]).strftime("%d/%m")
        return (f"  ·  so với {d}: ▲{len(ch['up'])} ▼{len(ch['down'])}  "
                f"mới TỐT & THỰC {len(ch['new_good'])}")

    def _set_dataframe(self, df):
        """Một chỗ duy nhất gán kết quả quét.

        Trước đây đường nạp từ cache chỉ gán self._df mà quên báo cho hai
        panel biểu đồ, nên bản đồ so sánh và đồ thị bằng chứng trống trơn
        dù đã có dữ liệu.
        """
        self._df = df
        self._chart_panel.set_universe(df)
        self._ai_panel.set_universe(df)

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
            self._changes = payload.get("changes")
            if "Định giá" not in df.columns:
                # Bộ nhớ đệm từ trước khi có định giá — tính bù, không cần quét lại
                try:
                    df = valuation.add_valuation(df, market)
                except Exception:
                    pass
            age_str  = ""
            if saved_at is not None:
                delta = pd.Timestamp.now() - saved_at
                h = int(delta.total_seconds() // 3600)
                m = int((delta.total_seconds() % 3600) // 60)
                age_str = f"{h}h {m}m" if h else f"{m}m"

            self._populate_table(df)
            self._populate_qc_table(df)
            self._set_dataframe(df)
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
        # Thị trường và số mã của chính lần quét — ô chọn có thể đã bị đổi trong lúc chờ
        market = getattr(self._worker, "market", None) or self._market.currentText()
        top    = getattr(self._worker, "top", None) or self._top_spin.value()
        # Phải so trước _save_cache: bộ nhớ đệm cũ là một nguồn của lần quét trước
        df, self._changes, track_err = self._track_changes(df, market)
        self._populate_table(df)
        self._populate_qc_table(df)

        self._set_dataframe(df)
        self._update_dashboard(df)
        self._tabs_w.setCurrentIndex(0)
        _sig = df["Signal"] if "Signal" in df.columns else pd.Series(dtype=str)
        self._set_status(
            f"✓  {len(df)} mã  ·  "
            f"TỐT & THỰC: {(_sig=='🟢 TỐT & THỰC').sum()}  "
            f"KHÁ: {(_sig=='🔵 KHÁ').sum()}  "
            f"CỜ ĐỎ: {(_sig=='🔴 CỜ ĐỎ').sum()}  "
            f"THEO DÕI: {(_sig=='🟡 THEO DÕI').sum()}"
            + self._changes_suffix(track_err),
            GREEN
        )
        # Ảnh chụp theo ngày — tích luỹ dữ liệu để sau này kiểm chứng tín hiệu.
        # Cột Thay đổi chỉ có nghĩa với lần quét trước nên không ghi vào ảnh chụp.
        save_snapshot(df.drop(columns=list(signal_tracking.CHANGE_COLS), errors="ignore"),
                      market, top)
        # Lưu cache để lần sau không cần scan lại
        self._save_cache(df, market, top)
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
                elif key == "Định giá":
                    cell = _valuation_cell(row)
                elif key == "Thay đổi":
                    cell = _change_cell(row)
                elif key == "Tăng trưởng ngụ ý%":
                    cell = _implied_cell(val)
                elif key == "Lợi suất%":
                    cell = _pct_cell(val, "{:.1f}%")

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
                    if _tip and not cell.toolTip():
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
                _tk = str(r.get("Ticker", "") or "")
                base = [cell(str(i), T3),
                        cell(f'<a href="tk:{_tk}" style="color:#4A9EFF;'
                             f'text-decoration:none;font-weight:700;">{_tk}</a>'
                             if _tk else "", None, True),
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

        # Định giá là nhãn riêng; chỉ hiện khi dữ liệu có (thị trường Mỹ)
        val_col = (df["Định giá"].fillna("").astype(str) if "Định giá" in df.columns
                   else pd.Series([""] * n, index=df.index))
        has_val = bool((val_col != "").any())

        def val_badge(r):
            v = _txt(r.get("Định giá"))
            if not v:
                return cell("—", T3)
            fg, bg = _VALUATION_COLORS.get(v, (T1, None))
            return cell(v + ("" if _is_reliable(r.get("Định giá_tin cậy")) else " ?"), fg, True, bg)

        def implied(r):
            s_ = valuation.growth_label(r.get("Tăng trưởng ngụ ý%"), ".")
            return cell(s_, T1) if s_ else cell("—", T3)

        # ── 1. TỐT & THỰC ────────────────────────────────────────────────
        good = df[sig == "🟢 TỐT & THỰC"].sort_values(
            ["REAL_Pct", "FIN_Pct"], ascending=False) if n_good else df.iloc[0:0]
        good_tbl = table(
            HDRS + ["ROIC%", "FCF/NI", "FCF Mgn%", "1Y%"]
            + (["Định giá"] if has_val else []) + ["Cảnh báo"],
            rows_for(good, lambda r: [
                num(r.get("ROIC%"), "{:.1f}%"),
                num(r.get("FCF/NI"), "{:.2f}×", lambda v: v >= 0.8),
                num(r.get("FCF_Margin%"), "{:.1f}%", lambda v: v > 8),
                num(r.get("1Y%"), "{:+.1f}%", lambda v: v > 0),
            ] + ([val_badge(r)] if has_val else []) + [warn_cell(r)]))

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
            if r.get("REAL_EROD") is False:
                v = r.get("Xói mòn")
                bad.append(f"chất lượng đang bào mòn ({v})" if v and v != "giữ vững"
                           else "chất lượng đang bào mòn")
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

        def tk_link(t):
            t = _html_escape(str(t or ""))
            return (f'<a href="tk:{t}" style="color:#4A9EFF;text-decoration:none;'
                    f'font-weight:700;">{t}</a>') if t else ""

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

        # ── 6. Định giá: mã TỐT & THỰC chưa bị giá đòi hỏi quá nhiều ──────
        val_sec = ""
        if has_val:
            vcolors = dict(zip(valuation.VERDICTS, (G, B, A_, R)))
            dist = "  ·  ".join(
                f'<span style="color:{vcolors[v]};font-weight:700;">{v} {int((val_col == v).sum())}</span>'
                for v in valuation.VERDICTS)
            n_none = int((val_col == "").sum())
            if n_none:
                dist += f'  ·  <span style="color:{T3};">không định giá được {n_none}</span>'
            gv = df[(sig == "🟢 TỐT & THỰC") & val_col.isin(["Rẻ", "Hợp lý"])]
            if not gv.empty:
                gv = gv.sort_values(["Định giá_rank", "Lợi suất%"], ascending=[True, False])
            val_tbl = table(
                HDRS + ["Định giá", "Cơ sở", "Lợi suất", "Ngụ ý/năm", "Mốc/năm", "Mốc theo"],
                rows_for(gv, lambda r: [
                    val_badge(r),
                    cell(_html_escape(_txt(r.get("Cơ sở ĐG"))) or "—", T2),
                    num(r.get("Lợi suất%"), "{:.1f}%"),
                    implied(r),
                    num(r.get("Mốc tăng trưởng%"), "{:+.1f}%"),
                    cell(_html_escape(_txt(r.get("Mốc ĐG"))) or "—", T2),
                ]))
            val_sec = (sec(f'💲  ĐỊNH GIÁ — {len(gv)} mã TỐT &amp; THỰC đang Rẻ hoặc Hợp lý',
                           'giá hiện tại ngụ ý dòng tiền tăng bao nhiêu mỗi năm, so với mức doanh nghiệp '
                           'đã chứng minh · dấu ? = độ tin cậy thấp · nhãn riêng, không cộng vào điểm', A_)
                       + f'<div style="font-size:11px;color:{T2};margin:0 2px 8px 2px;">'
                         f'Toàn bộ {n} mã: {dist}</div>'
                       + val_tbl)

        # ── 7. Thay đổi so với lần quét trước ────────────────────────────
        ch = getattr(self, "_changes", None)
        short = signal_tracking.short_signal

        def tk_list(items, cap=40):
            """Danh sách mã dạng liên kết, kèm tín hiệu cũ → mới khi có."""
            out = []
            for it in items[:cap]:
                if isinstance(it, (tuple, list)):
                    t, ps, cs = it
                    to = _html_escape(short(cs)) if cs else "rời danh sách"
                    out.append(f'{tk_link(t)} <span style="color:{T3};">'
                               f'{_html_escape(short(ps))} → {to}</span>')
                else:
                    out.append(tk_link(it))
            more = (f'  <span style="color:{T3};">+{len(items) - cap} mã khác</span>'
                    if len(items) > cap else "")
            return " · ".join(out) + more

        if ch is None:
            chg_sec = (sec('🔄  Thay đổi so với lần quét trước',
                           'chưa có lần quét trước của thị trường này', B)
                       + f'<div style="color:{T3};font-size:11px;font-style:italic;padding:4px;">'
                         'Mỗi lần quét được lưu thành một ảnh chụp trong thư mục snapshots. Từ lần '
                         'quét vào ngày khác, mục này và cột «Thay đổi» trong bảng sẽ cho biết mã nào '
                         'lên hạng, xuống hạng, mới đạt hoặc mất TỐT &amp; THỰC.</div>')
        else:
            cur_sig = dict(zip(df["Ticker"].astype(str), sig.astype(str)))
            groups = [
                ("Mới vào TỐT &amp; THỰC", ch.get("new_good") or [], G),
                ("Rời TỐT &amp; THỰC", [(t, "🟢 TỐT & THỰC", cur_sig.get(t, ""))
                                        for t in ch.get("lost_good") or []], R),
                ("Cờ đỏ mới", ch.get("new_flags") or [], R),
                ("Lên hạng", ch.get("up") or [], G),
                ("Xuống hạng", ch.get("down") or [], O),
            ]
            if ch.get("entered") is not None:
                groups += [("Vào danh sách", ch.get("entered") or [], B),
                           ("Ra khỏi danh sách", ch.get("exited") or [], T2)]
            chg_rows = [[cell(f"{lbl} · {len(items)}", c, True), cell(tk_list(items))]
                        for lbl, items, c in groups if items]
            warns = "".join(
                f'<div style="color:{A_};font-size:11px;margin:2px 2px 6px 2px;">⚠ {_html_escape(w)}</div>'
                for w in ch.get("warnings") or [])
            body = (table(["Nhóm", "Mã  (tín hiệu cũ → mới)"], chg_rows, widths=["17%", "83%"])
                    if chg_rows else
                    f'<div style="color:{T3};font-size:11px;font-style:italic;padding:4px;">'
                    f'Không mã nào đổi tín hiệu.</div>')
            prev_d = pd.Timestamp(ch["prev_date"]).strftime("%d/%m/%Y")
            n_up, n_down = len(ch.get("up") or []), len(ch.get("down") or [])
            chg_sec = (sec(f'🔄  Thay đổi so với lần quét {prev_d}',
                           f'▲ {n_up} mã lên hạng · ▼ {n_down} mã xuống hạng · rê chuột vào cột '
                           f'«Thay đổi» trong bảng để xem từng tiêu chí', B)
                       + warns + body)

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
  {chg_sec}

  {sec(f'🟢  TỐT &amp; THỰC — {n_good} mã', 'số liệu tài chính mạnh VÀ vượt được các phép kiểm chứng', G)}
  {good_tbl}
  {val_sec}

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
        only_val  = self._chk_val.isChecked()
        any_chip  = only_sb or only_moat or only_1y or only_52w or only_val
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
                if show and only_val and rd.get("Định giá") not in ("Rẻ", "Hợp lý"):
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
            _rd = match.iloc[0].to_dict()
            self._detail.show_row(_rd)
            self._ticker_inp.setText(ticker)
            self._chart_panel.load_ticker(ticker)
            self._ai_panel.set_row(_rd)

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

    def _on_dash_link(self, url):
        """Bấm một mã trên Dashboard thì nạp biểu đồ và chi tiết cho mã đó.

        Dashboard không có bảng để chọn dòng, nên đây là cách chọn mã khi
        biểu đồ nằm trong chính tab này.
        """
        href = url.toString()
        if not href.startswith("tk:"):
            return
        ticker = href[3:].strip().upper()
        if ticker:
            self._select_ticker(ticker)

    def _on_map_pick(self, ticker: str):
        """Bấm một chấm trên bản đồ: chọn mã đó nhưng ở lại bản đồ để so tiếp."""
        self._select_ticker(ticker, show_chart=False)

    def _select_ticker(self, ticker: str, show_chart: bool = True):
        """Nạp biểu đồ và chi tiết cho một mã."""
        self._dash_ticker = ticker
        self._ticker_inp.setText(ticker)
        self._chart_panel.load_ticker(ticker, show_chart=show_chart)
        if self._df is not None:
            m = self._df[self._df["Ticker"] == ticker]
            if not m.empty:
                rd = m.iloc[0].to_dict()
                self._detail.show_row(rd)
                self._qc_detail.show_row(rd)
                self._ai_panel.set_row(rd)
        self._set_status(f"📈  {ticker}  ·  đã nạp biểu đồ", BLUE)

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

    def _show_validation(self):
        """Giữ một hộp thoại duy nhất: đóng giữa chừng thì luồng tải giá vẫn
        chạy tiếp, mở lại là thấy kết quả."""
        if self._valid_dlg is None:
            self._valid_dlg = ValidationDialog(self)
        self._valid_dlg.show()
        self._valid_dlg.raise_()
        self._valid_dlg.activateWindow()

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
