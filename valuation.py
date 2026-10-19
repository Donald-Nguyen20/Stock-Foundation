"""
Định giá ngược bằng DCF — chỉ cho thị trường Mỹ
================================================
Câu hỏi «cổ phiếu này đáng giá bao nhiêu» phụ thuộc hoàn toàn vào giả định
tăng trưởng người dùng tự đặt: chỉnh vài điểm phần trăm là ra mọi con số
mình muốn. Câu hỏi ngược lại khách quan hơn nhiều: với giá hôm nay, thị
trường đang NGỤ Ý dòng tiền phải tăng bao nhiêu mỗi năm trong 10 năm tới?
Rồi đặt con số đó cạnh mức tăng trưởng doanh nghiệp đã thực sự làm được.

Tách hẳn khỏi điểm chất lượng: không cộng vào FIN/REAL, không đổi tín hiệu.
Một doanh nghiệp tốt và thật vẫn có thể đang bị mua quá đắt — hai câu hỏi
đó phải được trả lời riêng.

Chỉ áp dụng cho Mỹ: chi phí vốn 10% là mức kỳ vọng dài hạn của cổ phiếu
định giá bằng USD. Với VND lãi suất phi rủi ro đã khác, và nhóm Tài chính
của TradingView ở Việt Nam gộp cả bất động sản — áp cùng tham số sẽ sai.
"""
from __future__ import annotations

import math
from dataclasses import dataclass

import pandas as pd

US_MARKETS = frozenset({"america", "nasdaq", "nyse"})

DISCOUNT = 0.10      # chi phí vốn chủ: lợi suất kỳ vọng dài hạn của cổ phiếu Mỹ
TERMINAL = 0.025     # tăng trưởng mãi mãi sau năm thứ 10, quanh mức GDP danh nghĩa
YEARS    = 10
G_LO, G_HI = -0.50, 1.00          # biên tìm nghiệm, theo năm

# Mốc tăng trưởng đã chứng minh được kẹp lại: doanh thu tăng 60% một năm
# không có nghĩa là duy trì được 60% suốt 10 năm.
REF_LO, REF_HI = -5.0, 20.0
CHEAP_GAP      = -5.0     # ngụ ý thấp hơn mốc từ 5 điểm trở lên → Rẻ
RICH_GAP       = 5.0      # cao hơn từ 5 điểm → Đắt
VERY_RICH_GAP  = 12.0     # cao hơn từ 12 điểm → Rất đắt
# Ngưỡng tuyệt đối: 10 năm tăng 20%/năm nghĩa là dòng tiền gấp 6,2 lần —
# rất ít doanh nghiệp cỡ lớn làm được, dù quá khứ có nhanh đến đâu.
ABS_RICH, ABS_VERY_RICH = 20.0, 30.0
LOW_CASH_CONVERSION = 0.6   # FCF dưới 0,6 lần lợi nhuận → định giá theo lợi nhuận
HOT_GROWTH = 40.0           # tăng trưởng gần đây quá cao để làm mốc tin cậy
MIN_PEERS  = 8

FIN_SECTORS = ("Financial Services", "Finance")
REIT_INDUSTRY = "Real Estate Investment Trusts"    # TradingView xếp REIT vào Finance

VERDICTS     = ("Rẻ", "Hợp lý", "Đắt", "Rất đắt")
VERDICT_RANK = {v: i + 1 for i, v in enumerate(VERDICTS)}

# Các cột thêm vào bảng kết quả
COLS = ("Cơ sở ĐG", "Lợi suất%", "Tăng trưởng ngụ ý%", "Mốc tăng trưởng%",
        "Mốc ĐG", "Định giá", "Định giá_rank", "Định giá_tin cậy",
        "Định giá_lý do", "Lợi suất_pctl")


def _num(v) -> float | None:
    if v is None:
        return None
    try:
        f = float(v)
    except (TypeError, ValueError):
        return None
    return f if math.isfinite(f) else None


def _vn(x: float, nd: int = 1) -> str:
    """Số kiểu Việt Nam: dấu phẩy thập phân."""
    return f"{x:.{nd}f}".replace(".", ",")


def pv_multiple(g: float, r: float = DISCOUNT, gt: float = TERMINAL,
                years: int = YEARS) -> float:
    """Giá trị hiện tại của dòng tiền hôm nay bằng 1, tăng g mỗi năm trong
    `years` năm rồi tăng gt mãi mãi. Tăng đơn điệu theo g."""
    total, cf = 0.0, 1.0
    for t in range(1, years + 1):
        cf *= 1 + g
        total += cf / (1 + r) ** t
    return total + cf * (1 + gt) / (r - gt) / (1 + r) ** years


def implied_growth(multiple: float | None) -> float | None:
    """Tốc độ tăng g (%/năm) sao cho pv_multiple(g) bằng `multiple`
    — tức giá trị thị trường chia dòng tiền hiện tại. Kẹp ở biên tìm kiếm."""
    m = _num(multiple)
    if m is None or m <= 0:
        return None
    lo, hi = G_LO, G_HI
    if pv_multiple(lo) >= m:
        return G_LO * 100
    if pv_multiple(hi) <= m:
        return G_HI * 100
    for _ in range(80):
        mid = (lo + hi) / 2
        if pv_multiple(mid) < m:
            lo = mid
        else:
            hi = mid
    return round((lo + hi) / 2 * 100, 1)


def growth_label(g: float | None, decimal: str = ",") -> str:
    """'+12,3%' — hoặc '≥ +100%' / '≤ -50%' khi nghiệm chạm biên tìm kiếm:
    khi đó con số chỉ là cận, không phải mức ngụ ý thật."""
    v = _num(g)
    if v is None:
        return ""
    if v >= G_HI * 100:
        return f"≥ +{G_HI * 100:.0f}%"
    if v <= G_LO * 100:
        return f"≤ {G_LO * 100:.0f}%"
    return f"{v:+.1f}%".replace(".", decimal)


def _empty(reason: str) -> dict:
    out = {c: None for c in COLS}
    out["Định giá"] = ""
    out["Định giá_lý do"] = reason
    return out


def _basis(row) -> tuple[float | None, str, str, bool]:
    """(dòng tiền dùng để định giá, tên cơ sở, ghi chú, cơ sở có chuẩn không).
    Dòng tiền None khi không định giá được — ghi chú khi đó là lý do. «Không
    chuẩn» là khi phải lùi về lợi nhuận vì dòng tiền thật đang thấp."""
    fcf = _num(row.get("FCF_abs"))
    ni  = _num(row.get("NI_abs"))
    if row.get("Industry", "") == REIT_INDUSTRY:
        # Khấu hao bất động sản kéo lợi nhuận ròng xuống dưới tiền thật thu về;
        # FCF gần với AFFO hơn — trừ năm mua tài sản lớn khiến FCF âm.
        if fcf is not None and fcf > 0 and (ni is None or fcf >= ni):
            return fcf, "FCF", "", True
        if ni is not None and ni > 0:
            return ni, "LN", ("REIT đang mua thêm tài sản nên FCF thấp hơn lợi nhuận ròng — "
                              "dùng lợi nhuận ròng, vốn bị khấu hao kéo thấp nên kết luận thiên về đắt"), False
        return None, "", "REIT có FCF và lợi nhuận đều âm — không định giá được", False
    if row.get("Sector", "") in FIN_SECTORS:
        if ni is None or ni <= 0:
            return None, "", "nhóm tài chính đang lỗ — không định giá được", False
        return ni, "LN", ("nhóm tài chính dùng lợi nhuận ròng: tiền cho vay và tiền gửi "
                          "chạy qua dòng tiền hoạt động nên FCF không phản ánh kinh doanh"), True
    if fcf is not None and fcf > 0 and (ni is None or ni <= 0 or fcf / ni >= LOW_CASH_CONVERSION):
        return fcf, "FCF", "", True
    if ni is not None and ni > 0:
        ratio = f"{_vn(fcf / ni, 2)} lần" if fcf is not None else "không rõ"
        return ni, "LN", (f"FCF chỉ bằng {ratio} lợi nhuận (đầu tư lớn hoặc vốn lưu động "
                          f"phình ra) nên dùng lợi nhuận ròng — mức này dễ dãi hơn dòng tiền thật, "
                          f"xem kỹ tiêu chí Tiền ở tab Kiểm chứng"), False
    return None, "", "dòng tiền tự do và lợi nhuận đều âm — không định giá bằng dòng tiền được", False


def _reference(row) -> tuple[float | None, str]:
    """Mức tăng trưởng đã chứng minh: ưu tiên doanh thu nhiều năm (yfinance),
    lùi về doanh thu 12 tháng. Trả về giá trị chưa kẹp và nhãn."""
    cagr = _num(row.get("DT CAGR%"))
    yrs  = _num(row.get("DT số năm"))
    if cagr is not None:
        return cagr, f"doanh thu {int(yrs)} năm" if yrs else "doanh thu nhiều năm"
    g = _num(row.get("Rev Growth%"))
    if g is not None:
        return g, "doanh thu 12 tháng"
    return None, ""


def zone_growths(ref: float) -> tuple[float, float, float]:
    """Mức tăng trưởng ngụ ý tại ba ranh giới Rẻ|Hợp lý, Hợp lý|Đắt, Đắt|Rất đắt.

    Gộp luật chênh so với mốc và ngưỡng tuyệt đối: mốc ≤ REF_HI nên ranh giới
    Rẻ luôn dưới ABS_RICH, còn hai ranh giới trên bị ngưỡng tuyệt đối chặn."""
    return (ref + CHEAP_GAP,
            min(ref + RICH_GAP, ABS_RICH),
            min(ref + VERY_RICH_GAP, ABS_VERY_RICH))


def _verdict(implied: float, ref: float) -> str:
    # Làm tròn hai vế: mốc −4,8 cộng 5 ra 0,20000000000000018 chứ không phải 0,2,
    # khiến mức ngụ ý nằm đúng ranh giới bị xếp nhầm sang vùng bên dưới
    implied = round(implied, 6)
    cheap, rich, very_rich = (round(g, 6) for g in zone_growths(ref))
    if implied <= cheap:
        return "Rẻ"
    if implied < rich:
        return "Hợp lý"
    if implied < very_rich:
        return "Đắt"
    return "Rất đắt"


@dataclass(frozen=True)
class PriceBands:
    """Mức giá mà nhãn định giá đổi sang nhãn kế tiếp, giữ nguyên dòng tiền và
    số cổ phiếu lúc quét. Giá ≤ cheap → Rẻ; < rich → Hợp lý; < very_rich → Đắt."""
    cheap: float
    rich: float
    very_rich: float
    growth: tuple[float, float, float]   # tăng trưởng ngụ ý (%/năm) tại ba mức giá trên
    ref: float                           # mốc đã kẹp
    ref_label: str
    basis: str                           # "FCF" hoặc "LN"
    scan_price: float
    reliable: bool

    @property
    def edges(self) -> tuple[float, float, float]:
        return self.cheap, self.rich, self.very_rich

    def zone(self, price: float) -> str:
        if price <= self.cheap:
            return VERDICTS[0]
        if price < self.rich:
            return VERDICTS[1]
        if price < self.very_rich:
            return VERDICTS[2]
        return VERDICTS[3]


def price_bands(row) -> PriceBands | None:
    """Lật ngược định giá: giá nào thì nhãn đổi. None khi mã không có nhãn
    (ngoài Mỹ, dòng tiền âm, thiếu mốc tăng trưởng…).

    Tăng trưởng ngụ ý chỉ phụ thuộc bội số giá trị / dòng tiền, mà bội số tỉ
    lệ thuận với giá khi số cổ phiếu không đổi — nên mỗi ranh giới tăng
    trưởng ứng với đúng một mức giá: P = pv_multiple(g) × dòng tiền mỗi cổ phiếu.
    """
    if row.get("Định giá") not in VERDICTS:
        return None
    price = _num(row.get("Price ($)"))
    mcap  = _num(row.get("MCap ($B)"))
    if price is None or price <= 0 or mcap is None or mcap <= 0:
        return None
    base, basis, _note, _sound = _basis(row)
    raw_ref, ref_label = _reference(row)
    if base is None or raw_ref is None:
        return None
    ref = max(REF_LO, min(REF_HI, raw_ref))
    per_share = base * price / (mcap * 1e9)
    growth = zone_growths(ref)
    edges = [pv_multiple(g / 100) * per_share for g in growth]
    flag = row.get("Định giá_tin cậy")          # đọc lại từ CSV thì là chuỗi "True"
    return PriceBands(*edges, growth=growth, ref=ref, ref_label=ref_label, basis=basis,
                      scan_price=price, reliable=flag is True or str(flag) == "True")


def value_row(row) -> dict:
    """Định giá một mã. `row` là dict hoặc Series có các trường sau clean_df."""
    mcap = _num(row.get("MCap ($B)"))
    if mcap is None or mcap <= 0:
        return _empty("thiếu vốn hóa")
    base, basis, note, sound = _basis(row)
    if base is None:
        return _empty(note)

    value   = mcap * 1e9
    implied = implied_growth(value / base)
    yld     = round(base / value * 100, 2)
    raw_ref, ref_label = _reference(row)

    out = _empty("")
    out.update({"Cơ sở ĐG": basis, "Lợi suất%": yld, "Tăng trưởng ngụ ý%": implied,
                "Mốc ĐG": ref_label})
    flow = "dòng tiền tự do" if basis == "FCF" else "lợi nhuận"
    head = (f"Giá hiện tại ngụ ý {flow} tăng {growth_label(implied)}/năm trong {YEARS} năm "
            f"(chiết khấu {DISCOUNT:.0%}, sau đó tăng {_vn(TERMINAL * 100)}% mãi mãi).")
    if raw_ref is None:
        out["Định giá_lý do"] = head + " Không có số tăng trưởng doanh thu để so."
        return out

    ref = max(REF_LO, min(REF_HI, raw_ref))
    verdict = _verdict(implied, ref)
    reliable = sound and abs(raw_ref) <= HOT_GROWTH
    gap = implied - ref
    # Nghiệm chạm biên tìm kiếm thì mức ngụ ý và độ chênh chỉ là cận
    at_bound = implied >= G_HI * 100 or implied <= G_LO * 100
    clip = f" (kẹp về {_vn(ref)}% để so)" if ref != raw_ref else ""
    parts = [head,
             f"Mốc đã chứng minh: {ref_label} tăng {_vn(raw_ref)}%/năm{clip}.",
             f"Chênh {'ít nhất ' if at_bound else ''}{'+' if gap >= 0 else '−'}"
             f"{_vn(abs(gap))} điểm → {verdict}."]
    if implied >= ABS_RICH:
        g_txt = f"từ {_vn(implied, 0)}%" if at_bound else f"{_vn(implied)}%"
        parts.append(f"Tăng {g_txt}/năm suốt {YEARS} năm là mức rất ít doanh nghiệp "
                     f"cỡ lớn đạt được, bất kể quá khứ.")
    if abs(raw_ref) > HOT_GROWTH:
        parts.append("Tăng trưởng gần đây quá cao để làm mốc — nếu đây là đỉnh chu kỳ thì "
                     "dòng tiền hiện tại không bền và kết luận «rẻ» không đáng tin.")
    if note:
        parts.append(note[0].upper() + note[1:] + ".")
    out.update({"Mốc tăng trưởng%": round(ref, 1), "Định giá": verdict,
                "Định giá_rank": VERDICT_RANK[verdict], "Định giá_tin cậy": reliable,
                "Định giá_lý do": " ".join(parts)})
    return out


def _sector_percentile(df: pd.DataFrame) -> pd.Series:
    """Lợi suất cao hơn bao nhiêu phần trăm số mã cùng ngành, cùng cơ sở
    (FCF với FCF, lợi nhuận với lợi nhuận — ngân hàng không xếp chung với REIT).
    100 = rẻ nhất. Để trống khi không rõ ngành hoặc nhóm không đủ mã để so."""
    y = pd.to_numeric(df["Lợi suất%"], errors="coerce")
    if "Sector" not in df.columns or y.notna().sum() == 0:
        return pd.Series([None] * len(df), index=df.index, dtype=object)
    sec = df["Sector"].fillna("").astype(str)
    key = sec + "|" + df["Cơ sở ĐG"].fillna("").astype(str)
    counts = y.groupby(key).transform("count")
    pct = y.groupby(key).rank(pct=True) * 100
    return pct.where((counts >= MIN_PEERS) & (sec != "")).round(0)


def add_valuation(df: pd.DataFrame, market: str | None) -> pd.DataFrame:
    """Bản sao của df có thêm các cột định giá. Ngoài thị trường Mỹ thì các
    cột vẫn có mặt (để bảng và ảnh chụp đồng nhất) nhưng để trống."""
    out = df.copy()
    if str(market or "").lower() not in US_MARKETS:
        blank = _empty("chỉ áp dụng cho thị trường Mỹ")
        for c in COLS:
            out[c] = [blank[c]] * len(out)
        return out
    rows = [value_row(r) for r in out.to_dict("records")]
    for c in COLS:
        out[c] = [r[c] for r in rows]
    out["Lợi suất_pctl"] = _sector_percentile(out)
    return out
