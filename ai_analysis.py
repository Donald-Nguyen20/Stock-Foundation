# -*- coding: utf-8 -*-
"""Phân tích định tính một cổ phiếu bằng Claude.

Hai nguyên tắc chi phối thiết kế ở đây:

1. Claude **giải thích số liệu app đã tính**, không tự nhớ lại số. Mọi chỉ số
   của mã được nhét sẵn vào prompt. Để model tự nhớ doanh thu hay biên lợi
   nhuận thì nó sẽ lấy từ dữ liệu huấn luyện đã cũ, và có thể bịa.

2. Tin tức **bắt buộc đi qua công cụ tìm kiếm**, kèm nguồn và ngày. Không có
   tìm kiếm thì phần tin tức chỉ là hư cấu từ trí nhớ cũ.

Dùng **claude_agent_sdk** chứ không phải thư viện `anthropic`: Agent SDK đóng gói
sẵn Claude Code nên chạy bằng tài khoản đã đăng nhập và tiêu dung lượng của gói
thuê bao, thay vì tính vào tín dụng API phải nạp riêng.
"""
from __future__ import annotations

import json
import os
import re
from datetime import date, datetime
from pathlib import Path

MODEL_DEFAULT = "claude-opus-5-5"

# Agent SDK tự trả về total_cost_usd nên không cần tự tính; giữ bảng giá để
# ước lượng khi cần hiển thị trước lúc chạy.
PRICES = {
    "claude-opus-5-5":  (4.00, 20.00),
    "claude-opus-5":    (5.00, 25.00),
    "claude-sonnet-5-5": (2.00, 10.00),
    "claude-haiku-5-5": (0.10, 0.50),
}

SECTIONS = [
    ("nganh",    "NGÀNH KINH DOANH"),
    ("loithe",   "LỢI THẾ CẠNH TRANH"),
    ("tiemnang", "TIỀM NĂNG PHÁT TRIỂN"),
    ("tintuc",   "TIN TỨC QUAN TRỌNG"),
    ("doichieu", "ĐỐI CHIẾU VỚI SỐ LIỆU"),
]

SYSTEM = """Bạn là nhà phân tích cổ phiếu, viết bằng tiếng Việt cho một nhà đầu tư cá nhân.

Người dùng đã có sẵn một bộ lọc định lượng chấm điểm doanh nghiệp theo hai nhóm:
- Nhóm A "Tài chính": số liệu tài chính có tốt không (ROIC, biên, tăng trưởng, nợ).
- Nhóm B "Kiểm chứng": số liệu đó có thật không (dòng tiền có khớp lợi nhuận không,
  ROE có phải do đòn bẩy không, EPS có tăng nhờ mua cổ phiếu quỹ không, quý gần
  nhất đã xấu đi chưa, có bền vững qua 5 năm không).

Việc của bạn là bổ sung đúng phần mà con số không nói được, và giải thích vì sao
các con số đó lại như vậy.

QUY TẮC BẮT BUỘC:

1. KHÔNG tự nhớ lại số liệu tài chính. Mọi chỉ số cần dùng đã được cung cấp trong
   phần dữ liệu. Nếu cần một con số không có ở đó, hãy nói là không có, đừng đoán.

2. Phần TIN TỨC QUAN TRỌNG phải dựa trên kết quả tìm kiếm web. Mỗi tin ghi rõ ngày
   và nguồn. Không tìm được tin nào đáng kể thì viết "Không tìm thấy tin đáng chú ý
   trong thời gian gần đây" — tuyệt đối không bịa tin, không viết tin chung chung
   kiểu "công ty tiếp tục tăng trưởng".

3. KHÔNG đưa khuyến nghị mua, bán hay nắm giữ, và không nêu giá mục tiêu. Bộ lọc
   này cố ý không chấm định giá và không có quy tắc bán; bạn cũng không lấp chỗ đó.

4. Nếu câu chuyện định tính MÂU THUẪN với số liệu, phải nói thẳng ra. Đừng làm mượt
   cho dễ nghe. Mâu thuẫn là thông tin có giá trị nhất đối với người đọc.

5. Viết cụ thể, không viết sáo rỗng. "Biên gộp 70% đến từ việc phần mềm gần như
   không có chi phí nhân bản" là cụ thể; "công ty có lợi thế cạnh tranh mạnh" là
   sáo rỗng và vô dụng.

ĐỊNH DẠNG: trả lời bằng Markdown, dùng đúng năm tiêu đề cấp hai sau đây, đúng thứ tự,
không thêm tiêu đề nào khác:

## NGÀNH KINH DOANH
Doanh nghiệp kiếm tiền bằng cách nào, cấu trúc ngành ra sao, đối thủ chính là ai.

## LỢI THẾ CẠNH TRANH
Có hào kinh tế thật không, loại gì, và nó có giải thích được mức ROIC cùng biên lợi
nhuận trong phần dữ liệu không. Hào đang rộng ra hay đang bị bào mòn.

## TIỀM NĂNG PHÁT TRIỂN
Động lực tăng trưởng cụ thể trong 2–3 năm tới và rủi ro cụ thể. Không nói chung chung.

## TIN TỨC QUAN TRỌNG
Gạch đầu dòng, mỗi dòng bắt đầu bằng ngày và nguồn, rồi tới nội dung và vì sao nó
ảnh hưởng. Tối đa 5 tin.

## ĐỐI CHIẾU VỚI SỐ LIỆU
Chỗ nào câu chuyện khớp với điểm bộ lọc chấm, và chỗ nào lệch. Nêu rõ từng tiêu chí
bị trượt có lý do đã biết hay không — đây là phần quan trọng nhất."""


# ─────────────────────────────────────────────────────────────────────────────
# Dữ liệu đưa vào prompt
# ─────────────────────────────────────────────────────────────────────────────
def _fmt(v, suffix="", dec=1):
    if v is None:
        return "không có"
    try:
        f = float(v)
        if f != f:
            return "không có"
        return f"{f:,.{dec}f}{suffix}"
    except (TypeError, ValueError):
        return str(v)


def build_context(row: dict, fin_criteria: dict, fin_keys, real_criteria: dict, real_keys) -> str:
    """Gom số liệu của một mã thành khối văn bản cho prompt."""
    mark = {True: "ĐẠT", False: "TRƯỢT", None: "không áp dụng / thiếu dữ liệu"}

    def crit_lines(criteria, keys, prefix):
        out = []
        for k in keys:
            cfg = criteria[k]
            got = row.get(f"{prefix}_{k}")
            val = _fmt(row.get(cfg["field"]), dec=2)
            out.append(f"  - {cfg['label']}: {mark.get(got, '?')}  (giá trị: {val})")
        return "\n".join(out)

    warn = row.get("Cảnh báo") or ""
    return f"""DỮ LIỆU BỘ LỌC — {row.get('Ticker','?')} ({row.get('Tên Công Ty','')})
Ngành (theo TradingView): {row.get('Sector','không rõ')}
Vốn hóa: {_fmt(row.get('MCap ($B)'),' tỷ')}   Giá: {_fmt(row.get('Price ($)'),'',2)}

QUY MÔ (bốn quý gần nhất, đơn vị tỷ)
  Doanh thu: {_fmt(row.get('Doanh thu ($B)'))}
  Lợi nhuận ròng: {_fmt(row.get('Lợi nhuận ($B)'))}
  Dòng tiền tự do: {_fmt(row.get('FCF ($B)'))}

SINH LỜI & TĂNG TRƯỞNG
  ROIC: {_fmt(row.get('ROIC%'),'%')}   ROE: {_fmt(row.get('ROE%'),'%')}
  Biên gộp: {_fmt(row.get('Gross Margin%'),'%')}   Biên hoạt động: {_fmt(row.get('Op Margin%'),'%')}
  Lợi nhuận TTM tăng: {_fmt(row.get('NI Growth%'),'%')}   quý gần nhất: {_fmt(row.get('NI Qtr%'),'%')}
  Doanh thu TTM tăng: {_fmt(row.get('Rev Growth%'),'%')}   quý gần nhất: {_fmt(row.get('Rev Qtr%'),'%')}

KIỂM CHỨNG
  FCF / Lợi nhuận ròng: {_fmt(row.get('FCF/NI'),'×',2)}   Biên FCF: {_fmt(row.get('FCF_Margin%'),'%')}
  EPS tăng nhanh hơn lợi nhuận: {_fmt(row.get('EPS_NI_Gap'),' điểm %')}
  ROE vượt ROIC: {_fmt(row.get('ROE_ROIC_Gap'),' điểm %')}
  Tiền ròng: {_fmt(row.get('Net Cash ($B)'),' tỷ')}   Thanh toán hiện hành: {_fmt(row.get('Current Ratio'),'',2)}
  Bền vững 5 năm: {row.get('Moat Score') or 'không có dữ liệu'}

ĐIỂM
  Tài chính: {row.get('FIN_Score','?')}/{row.get('FIN_App','?')}
  Thực:      {row.get('REAL_Score','?')}/{row.get('REAL_App','?')}
  Tín hiệu:  {row.get('Signal','?')}

TỪNG TIÊU CHÍ — NHÓM A (tài chính tốt)
{crit_lines(fin_criteria, fin_keys, 'FIN')}

TỪNG TIÊU CHÍ — NHÓM B (chứng minh là thực)
{crit_lines(real_criteria, real_keys, 'REAL')}

TRẠNG THÁI GIÁ
  So với đỉnh 52 tuần: {_fmt(row.get('% từ đỉnh'),'%')}
  So với trung bình 250 phiên: {_fmt(row.get('Giá/MA250%'),'%')}
  Cảnh báo: {warn if warn else 'không có'}

Hôm nay là {date.today():%d/%m/%Y}."""


# ─────────────────────────────────────────────────────────────────────────────
# Bộ nhớ đệm — mỗi mã một lần mỗi ngày
# ─────────────────────────────────────────────────────────────────────────────
def cache_dir() -> Path:
    d = Path(__file__).resolve().parent / "ai_cache"
    d.mkdir(exist_ok=True)
    return d


def _cache_file(ticker: str, day: str | None = None) -> Path:
    day = day or f"{date.today():%Y%m%d}"
    safe = re.sub(r"[^A-Za-z0-9._-]", "_", ticker)
    return cache_dir() / f"{safe}_{day}.json"


def load_cached(ticker: str) -> dict | None:
    f = _cache_file(ticker)
    try:
        if f.is_file():
            return json.loads(f.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        pass
    return None


def save_cache(ticker: str, data: dict) -> None:
    try:
        _cache_file(ticker).write_text(
            json.dumps(data, ensure_ascii=False, indent=1), encoding="utf-8")
    except OSError:
        pass


def cost_of(model: str, usage) -> float:
    inp, outp = PRICES.get(model, (0.0, 0.0))
    try:
        i = (getattr(usage, "input_tokens", 0) or 0)
        i += (getattr(usage, "cache_read_input_tokens", 0) or 0)
        i += (getattr(usage, "cache_creation_input_tokens", 0) or 0)
        o = getattr(usage, "output_tokens", 0) or 0
        return i / 1e6 * inp + o / 1e6 * outp
    except Exception:
        return 0.0


# ─────────────────────────────────────────────────────────────────────────────
# Gọi Claude qua Claude Agent SDK
#
# Dùng claude_agent_sdk chứ không phải thư viện anthropic. Khác biệt quan trọng:
# Agent SDK đóng gói sẵn Claude Code nên xác thực bằng chính tài khoản bạn đã
# đăng nhập, dùng dung lượng của gói thuê bao. Thư viện anthropic gọi thẳng API
# và luôn tính vào tín dụng API — thứ phải nạp riêng.
#
# Lưu ý về điều khoản: cách này dành cho việc dùng cá nhân với tài khoản của
# chính mình. Nếu phát hành app cho người khác dùng thì Anthropic yêu cầu
# chuyển sang xác thực bằng API key.
# ─────────────────────────────────────────────────────────────────────────────
BILLING_URL = "https://console.anthropic.com/settings/billing"


def friendly_error(e: Exception) -> str:
    """Đổi lỗi của SDK thành câu người dùng hiểu được."""
    msg = str(e)
    low = msg.lower()
    if "not logged in" in low or "invalid api key" in low or "authentication" in low:
        return ("Chưa đăng nhập tài khoản Claude. Mở AI Settings và bấm nút liên kết, "
                "hoặc chạy `claude` trong terminal rồi đăng nhập.")
    if "credit balance is too low" in low or "purchase credits" in low:
        return ("Đang bị tính vào tín dụng API thay vì gói thuê bao. Kiểm tra xem biến "
                "môi trường ANTHROPIC_API_KEY có đang được đặt không — nếu có, xoá nó đi "
                "để quay về dùng tài khoản đã đăng nhập.")
    if "rate limit" in low or "429" in low or "usage limit" in low:
        return ("Đã chạm giới hạn dung lượng của gói thuê bao. Chờ cửa sổ giới hạn "
                "được làm mới rồi thử lại.")
    if "cli" in low and ("not found" in low or "no such" in low):
        return ("Không tìm thấy Claude Code đi kèm SDK. Cài lại bằng: "
                "pip install --force-reinstall claude-agent-sdk")
    if "overloaded" in low or "529" in low:
        return "Máy chủ đang quá tải. Thử lại sau ít phút."
    if "connection" in low or "timeout" in low:
        return "Không kết nối được. Kiểm tra mạng rồi thử lại."
    return f"{type(e).__name__}: {msg[:300]}"


def analyze(context: str, ticker: str, company: str = "",
            model: str = MODEL_DEFAULT,
            on_delta=None, on_status=None) -> dict:
    """Sinh phân tích. Trả về dict đã sẵn sàng lưu cache.

    `on_delta(text)` nhận từng mẩu chữ để hiện dần; `on_status(text)` nhận mô tả
    bước đang chạy.
    """
    import asyncio

    from claude_agent_sdk import (AssistantMessage, ClaudeAgentOptions,
                                  ResultMessage, query)

    def say(m):
        if on_status:
            on_status(m)

    prompt = (f"Phân tích cổ phiếu {ticker}" + (f" ({company})" if company else "") +
              ".\n\nDùng WebSearch tìm tin tức quan trọng gần đây về doanh nghiệp này "
              "trước khi viết phần TIN TỨC QUAN TRỌNG.\n\n" + context)

    opts = ClaudeAgentOptions(
        system_prompt=SYSTEM,
        # Chỉ cho phép tìm web. Agent SDK mang theo cả bộ công cụ đọc ghi tệp
        # của Claude Code, không giới hạn thì nó có quyền đụng vào máy.
        allowed_tools=["WebSearch"],
        disallowed_tools=["Read", "Write", "Edit", "Bash", "Glob", "Grep",
                          "NotebookEdit", "WebFetch"],
        permission_mode="bypassPermissions",
        model=model,
        effort="high",
        max_turns=12,
        # Không nạp CLAUDE.md hay cấu hình của dự án vào ngữ cảnh
        setting_sources=[],
    )

    state = {"text": [], "cost": 0.0, "searches": 0, "subtype": "", "usage": None}

    async def run():
        say("Đang tìm tin tức và phân tích…")
        async for msg in query(prompt=prompt, options=opts):
            if isinstance(msg, AssistantMessage):
                for b in msg.content:
                    if hasattr(b, "text") and b.text:
                        state["text"].append(b.text)
                        if on_delta:
                            on_delta(b.text)
                    elif getattr(b, "name", "") == "WebSearch":
                        state["searches"] += 1
                        say(f"Đang tìm web (lần {state['searches']})…")
            elif isinstance(msg, ResultMessage):
                state["subtype"] = msg.subtype or ""
                state["cost"] = getattr(msg, "total_cost_usd", 0.0) or 0.0
                state["usage"] = getattr(msg, "usage", None)

    # query() là coroutine nên cần vòng lặp riêng — hàm này luôn chạy trong
    # luồng nền của giao diện, không có sẵn event loop.
    loop = asyncio.new_event_loop()
    try:
        asyncio.set_event_loop(loop)
        loop.run_until_complete(run())
    finally:
        try:
            loop.close()
        finally:
            asyncio.set_event_loop(None)

    full = "".join(state["text"]).strip()
    if state["subtype"] and state["subtype"] != "success":
        raise RuntimeError(f"Phiên kết thúc với trạng thái: {state['subtype']}")
    if not full:
        raise RuntimeError("Không nhận được nội dung nào.")

    return {
        "ticker": ticker,
        "company": company,
        "model": model,
        "created": datetime.now().isoformat(timespec="seconds"),
        "markdown": full,
        "sections": split_sections(full),
        "cost_usd": round(float(state["cost"]), 4),
        "searches": state["searches"],
        "truncated": False,
    }


def split_sections(md: str) -> dict:
    """Tách Markdown thành từng mục theo năm tiêu đề đã quy định."""
    out = {}
    titles = {t: k for k, t in SECTIONS}
    cur = None
    buf: list[str] = []
    for line in (md or "").splitlines():
        m = re.match(r"^\s*#{1,3}\s*(.+?)\s*$", line)
        if m:
            name = m.group(1).strip().upper()
            key = titles.get(name)
            if key:
                if cur:
                    out[cur] = "\n".join(buf).strip()
                cur, buf = key, []
                continue
        if cur:
            buf.append(line)
    if cur:
        out[cur] = "\n".join(buf).strip()
    return out
