"""
Số quỹ sở hữu theo quý — từ bộ dữ liệu Form 13F của SEC
=======================================================
Mọi quỹ quản lý từ 100 triệu USD cổ phiếu Mỹ trở lên phải nộp Form 13F
cho SEC sau mỗi quý, chậm nhất 45 ngày. SEC gom các báo cáo đó thành
một tệp zip mỗi kỳ (~100 MB). Ở đây mỗi kỳ được tải một lần, rút gọn
thành bảng "mỗi CUSIP có bao nhiêu quỹ, tổng bao nhiêu cổ phiếu" rồi
lưu vào fund_cache/ — các lần sau chỉ đọc lại bảng nhỏ đó.

13F ghi mã CUSIP chứ không ghi ticker. Ánh xạ ticker → CUSIP lấy từ
tệp "fails-to-deliver" của SEC, nơi có cả hai cột.

SEC từ chối mọi yêu cầu không khai tên và email liên hệ trong
User-Agent. Thông tin đó do người dùng tự nhập (hoặc đặt biến môi
trường SEC_USER_AGENT); mô-đun này không tự điền gì thay họ.
"""
from __future__ import annotations

import io
import json
import os
import re
import unicodedata
import zipfile
from dataclasses import dataclass
from datetime import date, datetime, timedelta
from functools import lru_cache
from pathlib import Path
from typing import Callable

import pandas as pd
import requests

QUARTERS          = 12          # số quý giữ lại — ba năm
FTD_FILES         = 6           # ~3 tháng tệp fails-to-deliver là đủ phủ gần hết mã
MAP_MAX_AGE_DAYS  = 30
LIST_MAX_AGE_HRS  = 24
CHUNK_ROWS        = 200_000     # đỉnh bộ nhớ ~0,6 GB cho một quý, không chậm hơn
TIMEOUT           = (10, 30)    # (kết nối, mỗi lần đọc) — đủ ngắn để thoát app không phải chờ lâu

SEC      = "https://www.sec.gov"
F13_PAGE = SEC + "/data-research/sec-markets-data/form-13f-data-sets"
FTD_PAGE = SEC + "/data-research/sec-markets-data/fails-deliver-data"

# Tiền tố sàn của TradingView ứng với thị trường Mỹ
US_EXCHANGES = {"NASDAQ", "NYSE", "AMEX", "NYSE ARCA", "ARCA", "BATS", "CBOE", "OTC"}

_UA_ENV  = "SEC_USER_AGENT"
_UA_FILE = "sec_user_agent.txt"
_EMAIL   = re.compile(r"[\w.+-]+@[\w-]+(\.[\w-]+)+")
_MONTHS  = {m: i for i, m in enumerate(
    ["jan", "feb", "mar", "apr", "may", "jun",
     "jul", "aug", "sep", "oct", "nov", "dec"], start=1)}

ProgressFn = Callable[[str], None]
StopFn     = Callable[[], bool]


class SecAccessError(RuntimeError):
    """SEC trả 403 — thiếu hoặc sai tên/email trong User-Agent."""


class Cancelled(RuntimeError):
    """Người dùng đóng app hoặc dừng giữa chừng."""


# ─────────────────────────────────────────────────────────────────────────────
# Thư mục đệm và User-Agent
# ─────────────────────────────────────────────────────────────────────────────
def cache_dir() -> Path:
    d = Path(__file__).resolve().parent / "fund_cache"
    d.mkdir(exist_ok=True)
    return d


def _ascii(text: str) -> str:
    """Header HTTP không nhận chữ có dấu: 'Nguyễn Hiếu' → 'Nguyen Hieu'."""
    text = (text or "").replace("đ", "d").replace("Đ", "D")
    folded = unicodedata.normalize("NFKD", text).encode("ascii", "ignore").decode("ascii")
    return " ".join(folded.split())


def valid_user_agent(text: str) -> bool:
    """Cần có ít nhất một chữ cái ngoài email và một email hợp lệ."""
    text = _ascii(text)
    m = _EMAIL.search(text)
    return bool(m) and bool(re.search(r"[A-Za-z]", _EMAIL.sub("", text)))


def user_agent() -> str | None:
    env = _ascii(os.environ.get(_UA_ENV, ""))
    if env:
        return env
    try:
        saved = _ascii((cache_dir() / _UA_FILE).read_text(encoding="utf-8"))
    except OSError:
        return None
    return saved or None


def save_user_agent(text: str) -> None:
    if not valid_user_agent(text):
        raise ValueError("Cần dạng: Tên email@domain.com")
    (cache_dir() / _UA_FILE).write_text(_ascii(text), encoding="utf-8")


def forget_user_agent() -> None:
    """Xoá thông tin đã lưu (không đụng tới biến môi trường)."""
    try:
        (cache_dir() / _UA_FILE).unlink()
    except OSError:
        pass


def ua_from_env() -> bool:
    return bool(os.environ.get(_UA_ENV, "").strip())


def _session(ua: str) -> requests.Session:
    s = requests.Session()
    s.headers.update({"User-Agent": ua, "Accept-Encoding": "gzip, deflate"})
    return s


def _get(sess: requests.Session, url: str, **kw) -> requests.Response:
    r = sess.get(url, timeout=TIMEOUT, **kw)
    if r.status_code == 403:
        raise SecAccessError("SEC từ chối truy cập (403) — kiểm tra lại tên và email liên hệ")
    r.raise_for_status()
    return r


# ─────────────────────────────────────────────────────────────────────────────
# Sổ ghi chép: kỳ zip nào đã xử lý, danh sách kỳ lần cuối lấy lúc nào
# ─────────────────────────────────────────────────────────────────────────────
def _index_path() -> Path:
    return cache_dir() / "index.json"


def _load_index() -> dict:
    try:
        return json.loads(_index_path().read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return {}


def _save_index(idx: dict) -> None:
    tmp = _index_path().with_suffix(".tmp")
    tmp.write_text(json.dumps(idx, ensure_ascii=False, indent=1), encoding="utf-8")
    tmp.replace(_index_path())


def _fresh(stamp: str | None, max_age: timedelta) -> bool:
    if not stamp:
        return False
    try:
        return datetime.now() - datetime.fromisoformat(stamp) < max_age
    except ValueError:
        return False


# ─────────────────────────────────────────────────────────────────────────────
# Ticker → CUSIP
# ─────────────────────────────────────────────────────────────────────────────
def _sec_url(href: str) -> str | None:
    """Chỉ nhận đường dẫn về chính sec.gov — User-Agent có email không được gửi đi nơi khác."""
    if href.startswith("/") and not href.startswith("//"):
        return SEC + href
    return href if href.startswith(SEC + "/") else None


def parse_ftd_links(html: str) -> list[str]:
    """Đường dẫn tệp fails-to-deliver, mới nhất trước (202609b > 202609a)."""
    links = set(re.findall(r'href="([^"]*cnsfails(\d{6})([ab])\.zip)"', html))
    ordered = sorted(links, key=lambda t: (t[1], t[2]), reverse=True)
    return [u for u in (_sec_url(h) for h, _, _ in ordered) if u]


def read_ftd(raw: bytes) -> pd.DataFrame:
    """Một tệp fails-to-deliver → bảng SYMBOL, CUSIP (đã bỏ trùng)."""
    with zipfile.ZipFile(io.BytesIO(raw)) as z:
        name = next(n for n in z.namelist() if not n.endswith("/"))
        with z.open(name) as f:
            df = pd.read_csv(f, sep="|", dtype=str, encoding="latin-1",
                             usecols=["CUSIP", "SYMBOL"], quoting=3,
                             on_bad_lines="skip")
    df = df.dropna()
    df = df.assign(SYMBOL=df["SYMBOL"].str.strip().str.upper(),
                   CUSIP=df["CUSIP"].str.strip().str.upper())
    df = df[(df["SYMBOL"] != "") & (df["CUSIP"].str.len() == 9)]
    return df.drop_duplicates("SYMBOL", keep="last")


def _build_cusip_map(sess: requests.Session, progress: ProgressFn,
                     stop: StopFn) -> dict[str, str]:
    links = parse_ftd_links(_get(sess, FTD_PAGE).text)[:FTD_FILES]
    if not links:
        raise RuntimeError("không tìm thấy tệp fails-to-deliver trên trang SEC")
    out: dict[str, str] = {}
    # Cũ trước mới sau để CUSIP mới nhất của mỗi mã ghi đè lên
    for i, url in enumerate(reversed(links), start=1):
        if stop():
            raise Cancelled()
        progress(f"Tải bảng mã CUSIP {i}/{len(links)}…")
        df = read_ftd(_get(sess, url).content)
        out.update(zip(df["SYMBOL"], df["CUSIP"]))
    return out


def cusip_map(sess: requests.Session | None = None,
              progress: ProgressFn = lambda _m: None,
              stop: StopFn = lambda: False) -> dict[str, str]:
    """Bảng ticker → CUSIP; dựng lại khi quá 30 ngày và có kết nối.

    Không truyền sess thì chỉ đọc bản đã lưu (rỗng nếu chưa có). Dựng lại
    hỏng mà đã có bản cũ thì dùng tạm bản cũ, không chặn việc tải 13F.
    """
    f = cache_dir() / "cusip_map.json"
    try:
        saved = json.loads(f.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        saved = {}
    if sess is None or _fresh(saved.get("built_at"), timedelta(days=MAP_MAX_AGE_DAYS)):
        return saved.get("map", {})
    try:
        m = _build_cusip_map(sess, progress, stop)
    except (SecAccessError, Cancelled):
        raise
    except Exception:
        if saved.get("map"):
            return saved["map"]
        raise
    tmp = f.with_suffix(".tmp")
    tmp.write_text(json.dumps({"built_at": datetime.now().isoformat(timespec="seconds"),
                               "map": m}), encoding="utf-8")
    tmp.replace(f)
    _cached_map.cache_clear()
    return m


@lru_cache(maxsize=1)
def _cached_map() -> dict[str, str]:
    return cusip_map()


def symbol_variants(ticker: str) -> list[str]:
    """BRK.B ở TradingView là BRKB ở SEC; thử vài cách viết hạng cổ phiếu."""
    t = ticker.strip().upper().split(":")[-1]
    seen: list[str] = []
    for v in (t, t.replace(".", ""), t.replace(".", "/"), t.replace(".", "-"),
              t.replace("-", ""), t.replace("/", "")):
        if v and v not in seen:
            seen.append(v)
    return seen


def cusip_for(ticker: str, cmap: dict[str, str] | None = None) -> str | None:
    cmap = _cached_map() if cmap is None else cmap
    return next((cmap[v] for v in symbol_variants(ticker) if v in cmap), None)


# ─────────────────────────────────────────────────────────────────────────────
# Danh sách kỳ 13F
# ─────────────────────────────────────────────────────────────────────────────
@dataclass(frozen=True)
class Dataset:
    name: str          # vd 01jun2026-31aug2026_form13f.zip
    url: str
    end: date          # ngày cuối của khoảng nộp hồ sơ


def _window_end(name: str) -> date | None:
    m = re.search(r"\d{2}[a-z]{3}\d{4}-(\d{2})([a-z]{3})(\d{4})_form13f", name, re.I)
    if m:
        return date(int(m.group(3)), _MONTHS[m.group(2).lower()], int(m.group(1)))
    m = re.search(r"(\d{4})q([1-4])_form13f", name, re.I)
    if m:
        y, q = int(m.group(1)), int(m.group(2))
        return (date(y + 1, 1, 1) if q == 4 else date(y, 3 * q + 1, 1)) - timedelta(days=1)
    return None


def parse_datasets(html: str) -> list[Dataset]:
    out: dict[str, Dataset] = {}
    for href in re.findall(r'href="([^"]*_form13f\.zip)"', html, re.I):
        url = _sec_url(href)
        name = href.rsplit("/", 1)[-1]
        # Tên này thành tên tệp trong fund_cache nên chỉ nhận ký tự an toàn
        if not url or not re.fullmatch(r"[\w.-]+", name):
            continue
        end = _window_end(name)
        if end:
            out[name] = Dataset(name, url, end)
    return sorted(out.values(), key=lambda d: d.end, reverse=True)


def _list_datasets(sess: requests.Session, idx: dict) -> list[Dataset]:
    if _fresh(idx.get("listed_at"), timedelta(hours=LIST_MAX_AGE_HRS)) and idx.get("list"):
        return [Dataset(n, u, date.fromisoformat(e)) for n, u, e in idx["list"]]
    found = parse_datasets(_get(sess, F13_PAGE).text)
    if not found:
        raise RuntimeError("không đọc được danh sách bộ dữ liệu 13F trên trang SEC")
    idx["listed_at"] = datetime.now().isoformat(timespec="seconds")
    idx["list"] = [[d.name, d.url, d.end.isoformat()] for d in found]
    return found


# ─────────────────────────────────────────────────────────────────────────────
# Rút gọn một kỳ: mỗi CUSIP có bao nhiêu quỹ, tổng bao nhiêu cổ phiếu
# ─────────────────────────────────────────────────────────────────────────────
def period_key(text: str) -> str:
    """'30-JUN-2026' → '2026Q2'."""
    d = datetime.strptime(text.strip(), "%d-%b-%Y")
    return f"{d.year}Q{(d.month - 1) // 3 + 1}"


def quarter_label(key: str) -> str:
    """'2026Q2' → 'Q2/2026'."""
    return f"Q{key[-1]}/{key[:4]}"


def quarter_end(key: str) -> date:
    y, q = int(key[:4]), int(key[-1])
    return (date(y + 1, 1, 1) if q == 4 else date(y, 3 * q + 1, 1)) - timedelta(days=1)


def _member(z: zipfile.ZipFile, name: str) -> str:
    """Tên tệp trong zip, dù nằm ở gốc hay trong thư mục con.

    Phần lớn các kỳ để tệp ở gốc, nhưng có kỳ (01jun2025-31aug2025) bọc
    tất cả trong thư mục 01JUN2025-31AUG2025_form13f/.
    """
    hits = [n for n in z.namelist() if n.rsplit("/", 1)[-1].upper() == name.upper()]
    if not hits:
        raise KeyError(f"không thấy {name} trong tệp zip")
    return min(hits, key=len)


def _read_submissions(z: zipfile.ZipFile) -> tuple[str, pd.DataFrame]:
    """Báo cáo nắm giữ của kỳ chính trong zip, theo số hồ sơ."""
    with z.open(_member(z, "SUBMISSION.tsv")) as f:
        sub = pd.read_csv(f, sep="\t", dtype=str, quoting=3,
                          usecols=["ACCESSION_NUMBER", "FILING_DATE",
                                   "SUBMISSIONTYPE", "CIK", "PERIODOFREPORT"])
    sub = sub[sub["SUBMISSIONTYPE"].isin(("13F-HR", "13F-HR/A"))].dropna(
        subset=["PERIODOFREPORT", "CIK"])
    if sub.empty:
        raise RuntimeError("tệp 13F không có báo cáo nắm giữ nào")
    # Mỗi zip gom hồ sơ nộp trong một khoảng thời gian; gần hết là của
    # một quý — phần còn lại là hồ sơ nộp trễ hoặc sửa đổi cho quý cũ.
    main = sub["PERIODOFREPORT"].mode().iat[0]
    sub = sub[sub["PERIODOFREPORT"] == main].drop_duplicates("ACCESSION_NUMBER")
    # Ngày nộp hỏng thì xếp đầu, để không lấn hồ sơ có ngày thật nộp sau
    fd = pd.to_datetime(sub["FILING_DATE"], format="%d-%b-%Y", errors="coerce")
    sub = sub.assign(FD=fd.fillna(pd.Timestamp.min))
    # Xếp theo số hồ sơ: số thứ tự dùng làm khoá phụ khi hai hồ sơ cùng ngày
    return period_key(main), sub.set_index("ACCESSION_NUMBER")[["CIK", "FD"]].sort_index()


def _read_positions(z: zipfile.ZipFile, accessions: pd.Index,
                    stop: StopFn) -> pd.DataFrame:
    """Cổ phiếu thường (không quyền chọn) theo hồ sơ × CUSIP.

    Hồ sơ ghi bằng số thứ tự trong accessions (cột ACC) thay cho chuỗi
    số hồ sơ, cho bảng vài triệu dòng đỡ tốn bộ nhớ.
    """
    parts = []
    with z.open(_member(z, "INFOTABLE.tsv")) as f:
        reader = pd.read_csv(
            f, sep="\t", dtype=str, quoting=3, on_bad_lines="skip",
            chunksize=CHUNK_ROWS,
            usecols=["ACCESSION_NUMBER", "CUSIP", "SSHPRNAMT",
                     "SSHPRNAMTTYPE", "PUTCALL"])
        for ch in reader:
            if stop():
                raise Cancelled()
            ch = ch[ch["PUTCALL"].isna() & (ch["SSHPRNAMTTYPE"] == "SH")]
            acc = accessions.get_indexer(ch["ACCESSION_NUMBER"])
            ch, acc = ch[acc >= 0], acc[acc >= 0]
            cusip = ch["CUSIP"].str.strip().str.upper()
            # Không ít quỹ gõ thiếu số 0 đầu (037833100 → 37833100)
            cusip = cusip.where(cusip.str.len() >= 9, cusip.str.zfill(9))
            part = pd.DataFrame({
                "ACC": acc.astype("int32"), "CUSIP": cusip.to_numpy(),
                "SH": pd.to_numeric(ch["SSHPRNAMT"], errors="coerce").fillna(0).to_numpy()})
            parts.append(part.groupby(["ACC", "CUSIP"], as_index=False)["SH"].sum())
    if not parts:
        return pd.DataFrame({"ACC": pd.Series(dtype="int32"),
                             "CUSIP": pd.Series(dtype=object), "SH": pd.Series(dtype=float)})
    # Một hồ sơ có thể vắt qua hai khúc đọc nên phải cộng lại lần nữa
    return pd.concat(parts).groupby(["ACC", "CUSIP"], as_index=False)["SH"].sum()


def aggregate(zip_path: Path | str, stop: StopFn = lambda: False) -> tuple[str, pd.DataFrame]:
    """Một zip 13F → (kỳ, bảng CUSIP × [holders, shares]).

    holders = số quỹ (CIK) khác nhau báo cáo có giữ; shares = tổng cổ phiếu.
    Quỹ nộp cả bản gốc lẫn bản sửa đổi thì mỗi mã lấy hồ sơ nộp sau cùng có
    mã đó — bản sửa đổi kiểu "thêm mã" chỉ liệt kê phần thêm nên không thể
    lấy nguyên hồ sơ cuối cùng của quỹ.
    """
    with zipfile.ZipFile(zip_path) as z:
        key, sub = _read_submissions(z)
        pos = _read_positions(z, sub.index, stop)
    if stop():
        raise Cancelled()
    pos = pos.assign(CIK=sub["CIK"].to_numpy()[pos["ACC"]],
                     FD=sub["FD"].to_numpy()[pos["ACC"]])
    pos = (pos.sort_values(["FD", "ACC"])
              .drop_duplicates(["CIK", "CUSIP"], keep="last"))
    out = pos.groupby("CUSIP").agg(holders=("CIK", "nunique"), shares=("SH", "sum"))
    return key, out.astype({"holders": "int32", "shares": "float64"})


# ─────────────────────────────────────────────────────────────────────────────
# Tải và cập nhật
# ─────────────────────────────────────────────────────────────────────────────
def _download(sess: requests.Session, url: str, dest: Path,
              progress: Callable[[int, int], None], stop: StopFn) -> None:
    tmp = dest.with_suffix(".part")
    try:
        with sess.get(url, stream=True, timeout=TIMEOUT) as r:
            if r.status_code == 403:
                raise SecAccessError("SEC từ chối truy cập (403) — kiểm tra lại tên và email liên hệ")
            r.raise_for_status()
            total = int(r.headers.get("Content-Length") or 0)
            done = 0
            with open(tmp, "wb") as f:
                for block in r.iter_content(1 << 20):
                    if stop():
                        raise Cancelled()
                    f.write(block)
                    done += len(block)
                    progress(done, total)
        tmp.replace(dest)
    finally:
        tmp.unlink(missing_ok=True)


def _quarter_file(key: str) -> Path:
    return cache_dir() / f"{key}.pkl"


def _write_quarter(key: str, table: pd.DataFrame) -> None:
    """Ghi qua tệp tạm rồi đổi tên: tắt ngang không để lại bảng cụt."""
    tmp = cache_dir() / f"{key}.pkl.tmp"
    try:
        table.to_pickle(tmp)
        tmp.replace(_quarter_file(key))
    finally:
        tmp.unlink(missing_ok=True)


def _clean_leftovers() -> None:
    """Dọn zip và tệp dở dang do lần trước bị tắt ngang (mỗi zip ~100 MB)."""
    for pat in ("*.part", "*_form13f.zip", "*.tmp"):
        for p in cache_dir().glob(pat):
            try:
                p.unlink()
            except OSError:
                pass


def cached_quarters() -> list[str]:
    """Các quý đã có, cũ trước mới sau."""
    keys = [p.stem for p in cache_dir().glob("*.pkl") if re.fullmatch(r"\d{4}Q[1-4]", p.stem)]
    return sorted(keys)[-QUARTERS:]


def update(progress: ProgressFn = lambda _m: None,
           stop: StopFn = lambda: False,
           on_quarter: Callable[[str], None] = lambda _k: None) -> list[str]:
    """Tải những quý còn thiếu (mới nhất trước) cho đủ QUARTERS quý.

    Gọi lại bao nhiêu lần cũng được: kỳ đã xử lý thì bỏ qua, danh sách kỳ
    chỉ hỏi SEC tối đa mỗi ngày một lần.
    """
    ua = user_agent()
    if not ua:
        raise SecAccessError("chưa có tên và email liên hệ cho SEC")
    _clean_leftovers()
    sess = _session(ua)
    cusip_map(sess, progress, stop)
    _cached_map.cache_clear()

    idx = _load_index()
    done: dict[str, str] = idx.setdefault("done", {})
    datasets = _list_datasets(sess, idx)
    _save_index(idx)

    have = {p.stem for p in cache_dir().glob("*.pkl")}
    # Mỗi kỳ zip ứng với khoảng một quý; lấy dư hai kỳ phòng kỳ chuyển tiếp
    window = datasets[:QUARTERS + 2]
    got = {done[d.name] for d in window if done.get(d.name) in have}
    todo_n = min(QUARTERS - len(got), sum(done.get(d.name) not in have for d in window))
    step = 0
    bad: list[str] = []
    for ds in window:
        if len(got) >= QUARTERS:
            break
        if done.get(ds.name) in have:
            got.add(done[ds.name])
            continue
        step += 1
        tag = f"kỳ {step}/{max(todo_n, step)}"
        dest = cache_dir() / ds.name

        def _p(n: int, total: int, tag: str = tag) -> None:
            pct = f" {n * 100 // total}%" if total else f" {n >> 20} MB"
            progress(f"Tải dữ liệu 13F {tag}{pct}")

        _download(sess, ds.url, dest, _p, stop)
        try:
            progress(f"Đọc dữ liệu 13F {tag}…")
            key, table = aggregate(dest, stop)
        except (KeyError, ValueError, zipfile.BadZipFile) as e:
            # Một kỳ lạ định dạng thì bỏ qua, các kỳ khác vẫn tải tiếp;
            # không ghi vào done nên lần cập nhật sau sẽ thử lại kỳ đó
            bad.append(f"{ds.name.removesuffix('_form13f.zip')}: {e}")
            continue
        finally:
            dest.unlink(missing_ok=True)
        if key not in got:              # kỳ mới hơn đã cho quý này thì giữ bản đó
            _write_quarter(key, table)
            have.add(key)
            got.add(key)
            _load_quarter.cache_clear()
            _cached_map.cache_clear()   # phòng giao diện đã kịp nhớ bảng mã rỗng
            on_quarter(key)
        done[ds.name] = key
        _save_index(idx)
    _prune(keep=sorted(got)[-QUARTERS:])
    if bad:
        raise RuntimeError(f"bỏ qua {len(bad)} kỳ không đọc được ({'; '.join(bad)})")
    return cached_quarters()


def _prune(keep: list[str]) -> None:
    """Bỏ bảng của những quý đã trượt ra ngoài khung QUARTERS quý."""
    if not keep:
        return
    for p in cache_dir().glob("*.pkl"):
        if re.fullmatch(r"\d{4}Q[1-4]", p.stem) and p.stem < keep[0]:
            try:
                p.unlink()
            except OSError:             # đang được đọc: để lần sau dọn
                pass
    _load_quarter.cache_clear()


# ─────────────────────────────────────────────────────────────────────────────
# Đọc lịch sử một mã
# ─────────────────────────────────────────────────────────────────────────────
@lru_cache(maxsize=QUARTERS + 4)
def _load_quarter(key: str) -> pd.DataFrame:
    return pd.read_pickle(_quarter_file(key))


def history(ticker: str, cmap: dict[str, str] | None = None) -> pd.DataFrame:
    """Bảng theo quý (cũ → mới): quarter, label, holders, shares.

    Quý không thấy mã (CUSIP đổi, chưa niêm yết…) để trống. Không tìm ra
    CUSIP thì trả bảng rỗng.
    """
    cols = ["quarter", "label", "holders", "shares"]
    cusip = cusip_for(ticker, cmap)
    if not cusip:
        return pd.DataFrame(columns=cols)
    rows = []
    for key in cached_quarters():
        try:
            t = _load_quarter(key)
        except OSError:                 # vừa bị dọn khỏi khung 12 quý
            t = None
        except Exception:               # tệp hỏng: xoá để lần cập nhật sau tải lại
            t = None
            try:
                _quarter_file(key).unlink()
            except OSError:
                pass
        hit = t is not None and cusip in t.index
        rows.append((key, quarter_label(key),
                     float(t.at[cusip, "holders"]) if hit else float("nan"),
                     float(t.at[cusip, "shares"]) if hit else float("nan")))
    return pd.DataFrame(rows, columns=cols)


def has_any_data() -> bool:
    return bool(cached_quarters()) and bool(_cached_map())
