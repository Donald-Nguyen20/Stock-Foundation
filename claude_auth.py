# -*- coding: utf-8 -*-
"""Đăng nhập Claude bằng OAuth.

Anthropic không mở luồng OAuth để ứng dụng bên thứ ba tự đăng ký client_id.
Cách được hỗ trợ là CLI `ant`: `ant auth login` mở trình duyệt, đổi lấy token
ngắn hạn và lưu thành một profile dưới thư mục cấu hình của hệ điều hành
(trên Windows là %APPDATA%\\Anthropic). Sau đó mọi client `Anthropic()` không
tham số — kể cả trong app này — tự đọc profile đó. Nghĩa là app không bao giờ
chạm vào token, và cũng không cần bạn dán API key vào đâu cả.

Module này chỉ bọc CLI và kiểm chứng kết quả, không tự lưu bí mật ở đâu.
"""
from __future__ import annotations

import json
import os
import platform
import shutil
import subprocess
import sys
from dataclasses import dataclass, field
from pathlib import Path

RELEASES_URL = "https://github.com/anthropics/anthropic-cli/releases"

# Ẩn cửa sổ console khi gọi CLI từ app GUI trên Windows
_NO_WINDOW = subprocess.CREATE_NO_WINDOW if sys.platform == "win32" else 0


# ─────────────────────────────────────────────────────────────────────────────
# Định vị CLI
# ─────────────────────────────────────────────────────────────────────────────
def _candidate_paths():
    exe = "ant.exe" if sys.platform == "win32" else "ant"
    home = Path.home()
    out = [
        home / "go" / "bin" / exe,
        home / ".local" / "bin" / exe,
        home / "bin" / exe,
    ]
    if sys.platform == "win32":
        la = os.environ.get("LOCALAPPDATA")
        if la:
            out += [Path(la) / "Programs" / "ant" / exe,
                    Path(la) / "Microsoft" / "WinGet" / "Links" / exe]
        out.append(Path(home) / "scoop" / "shims" / exe)
        out.append(Path("C:/ProgramData/chocolatey/bin") / exe)
    else:
        out += [Path("/usr/local/bin") / exe, Path("/opt/homebrew/bin") / exe]
    return out


def find_ant() -> str | None:
    """Đường dẫn tới CLI `ant`, hoặc None nếu chưa cài.

    Ưu tiên bản app tự cài để kết quả không đổi khi PATH thay đổi.
    """
    own = managed_bin_dir() / ("ant.exe" if sys.platform == "win32" else "ant")
    try:
        if own.is_file():
            return str(own)
    except OSError:
        pass
    found = shutil.which("ant")
    if found:
        return found
    for p in _candidate_paths():
        try:
            if p.is_file():
                return str(p)
        except OSError:
            continue
    return None


def config_dir() -> Path:
    """Thư mục cấu hình mà CLI dùng để lưu profile và token."""
    custom = os.environ.get("ANTHROPIC_CONFIG_DIR")
    if custom:
        return Path(custom)
    if sys.platform == "win32":
        return Path(os.environ.get("APPDATA", Path.home())) / "Anthropic"
    return Path.home() / ".config" / "anthropic"


def stored_profiles() -> list[str]:
    """Tên các profile đã có token, đọc thẳng từ thư mục credentials."""
    d = config_dir() / "credentials"
    try:
        return sorted(f.stem for f in d.glob("*.json"))
    except OSError:
        return []


# ─────────────────────────────────────────────────────────────────────────────
# Gọi CLI
# ─────────────────────────────────────────────────────────────────────────────
def _run(args: list[str], timeout: int = 30) -> tuple[int, str]:
    """Chạy `ant ...`, trả về (mã thoát, stdout+stderr đã gộp)."""
    ant = find_ant()
    if not ant:
        return 127, "Chưa cài CLI `ant`."
    try:
        p = subprocess.run(
            [ant, *args],
            capture_output=True, text=True, encoding="utf-8", errors="replace",
            timeout=timeout, creationflags=_NO_WINDOW,
        )
        return p.returncode, ((p.stdout or "") + (p.stderr or "")).strip()
    except subprocess.TimeoutExpired:
        return 124, f"Quá thời gian chờ ({timeout} giây)."
    except OSError as e:
        return 1, f"Không chạy được CLI: {e}"


# ─────────────────────────────────────────────────────────────────────────────
# Trạng thái
# ─────────────────────────────────────────────────────────────────────────────
@dataclass
class AuthState:
    sdk_ok: bool = False            # thư viện anthropic đã cài chưa
    sdk_version: str = ""
    cli_path: str | None = None     # đường dẫn CLI ant
    profiles: list[str] = field(default_factory=list)
    status_text: str = ""           # nguyên văn `ant auth status`
    verified: bool = False          # đã gọi API thật thành công chưa
    account: str = ""               # mô tả ngắn để hiển thị
    error: str = ""
    api_key_shadow: bool = False    # ANTHROPIC_API_KEY đang che profile

    @property
    def cli_ok(self) -> bool:
        return bool(self.cli_path)

    @property
    def logged_in(self) -> bool:
        return self.verified


def _sdk_info() -> tuple[bool, str]:
    try:
        import anthropic
        return True, getattr(anthropic, "__version__", "?")
    except Exception:
        return False, ""


def verify_credentials() -> tuple[bool, str]:
    """Chứng thực bằng một lệnh gọi API thật thay vì đoán từ output CLI.

    `ant auth status` chỉ báo cáo trạng thái và tài liệu dặn không dùng mã
    thoát của nó làm kiểm tra sức khoẻ, nên ở đây dùng models.list() — rẻ,
    không sinh token, và chỉ thành công khi thông tin đăng nhập dùng được thật.
    """
    try:
        import anthropic
    except Exception as e:
        return False, f"Chưa cài thư viện anthropic: {e}"
    try:
        client = anthropic.Anthropic()          # tự đọc profile OAuth
        page = client.models.list(limit=1)
        n = len(getattr(page, "data", []) or [])
        return True, f"Gọi API thành công (đọc được {n} model)."
    except Exception as e:
        return False, f"{type(e).__name__}: {e}"


def read_state(verify: bool = True) -> AuthState:
    st = AuthState()
    st.sdk_ok, st.sdk_version = _sdk_info()
    st.cli_path = find_ant()
    st.profiles = stored_profiles()

    # Bẫy số một theo tài liệu: API key đặt sẵn sẽ che mọi profile OAuth,
    # và chuỗi rỗng vẫn chiếm chỗ ưu tiên đó.
    st.api_key_shadow = "ANTHROPIC_API_KEY" in os.environ

    if st.cli_ok:
        _, out = _run(["auth", "status"], timeout=20)
        st.status_text = out

    if verify and st.sdk_ok:
        ok, msg = verify_credentials()
        st.verified = ok
        if ok:
            st.account = msg
        else:
            st.error = msg
    return st


# ─────────────────────────────────────────────────────────────────────────────
# Hành động
# ─────────────────────────────────────────────────────────────────────────────
def login(profile: str | None = None, timeout: int = 300) -> tuple[bool, str]:
    """Mở trình duyệt để uỷ quyền. Chặn cho tới khi xong hoặc hết giờ."""
    if not find_ant():
        return False, "Chưa cài CLI `ant`."
    args = ["auth", "login"]
    if profile:
        args += ["--profile", profile]
    code, out = _run(args, timeout=timeout)
    if code != 0:
        return False, out or f"Đăng nhập thất bại (mã {code})."
    ok, msg = verify_credentials()
    return ok, (out + ("\n" + msg if msg else "")).strip()


def logout(all_profiles: bool = False) -> tuple[bool, str]:
    if not find_ant():
        return False, "Chưa cài CLI `ant`."
    args = ["auth", "logout"] + (["--all"] if all_profiles else [])
    code, out = _run(args, timeout=60)
    return code == 0, out or ("Đã đăng xuất." if code == 0 else "Đăng xuất thất bại.")


# ─────────────────────────────────────────────────────────────────────────────
# Tự cài đặt
# ─────────────────────────────────────────────────────────────────────────────
GITHUB_API = "https://api.github.com/repos/anthropics/anthropic-cli/releases/latest"
_DOWNLOAD_PREFIX = "https://github.com/anthropics/anthropic-cli/releases/download/"


def managed_bin_dir() -> Path:
    """Nơi app tự cài CLI, tách khỏi mọi bản cài sẵn của hệ thống."""
    if sys.platform == "win32":
        base = Path(os.environ.get("LOCALAPPDATA", Path.home()))
    elif sys.platform == "darwin":
        base = Path.home() / "Library" / "Application Support"
    else:
        base = Path(os.environ.get("XDG_DATA_HOME", Path.home() / ".local" / "share"))
    return base / "StockFoundation" / "bin"


def _asset_suffix() -> str:
    """Hậu tố tên file phát hành khớp với máy đang chạy."""
    m = platform.machine().lower()
    arch = ("arm64" if m in ("arm64", "aarch64")
            else "386" if m in ("i386", "i686", "x86")
            else "amd64")
    if sys.platform == "win32":
        return f"windows_{arch}.zip"
    if sys.platform == "darwin":
        return f"macos_{arch}.zip"
    return f"linux_{arch}.tar.gz"


def install_cli(progress=None) -> tuple[bool, str]:
    """Tải CLI `ant` từ trang phát hành chính thức và cài vào thư mục của app.

    Trả về (thành công, thông điệp). `progress` nhận một chuỗi mô tả bước đang chạy.
    """
    import io
    import json
    import tarfile
    import urllib.request
    import zipfile

    def say(msg):
        if progress:
            progress(msg)

    try:
        say("Đang tìm bản phát hành mới nhất…")
        req = urllib.request.Request(
            GITHUB_API, headers={"User-Agent": "stock-foundation",
                                 "Accept": "application/vnd.github+json"})
        with urllib.request.urlopen(req, timeout=30) as r:
            rel = json.load(r)

        want = _asset_suffix()
        asset = next((a for a in rel.get("assets", [])
                      if a.get("name", "").endswith(want)), None)
        if asset is None:
            return False, f"Bản phát hành {rel.get('tag_name','?')} không có gói cho {want}."

        url = asset["browser_download_url"]
        # Chỉ nhận file từ đúng trang phát hành của Anthropic
        if not url.startswith(_DOWNLOAD_PREFIX):
            return False, "Đường dẫn tải về không thuộc kho chính thức — đã dừng."

        size_mb = asset.get("size", 0) / 1e6
        say(f"Đang tải {rel.get('tag_name','')} ({size_mb:.0f} MB)…")
        req = urllib.request.Request(url, headers={"User-Agent": "stock-foundation"})
        with urllib.request.urlopen(req, timeout=300) as r:
            blob = r.read()

        say("Đang giải nén…")
        exe_name = "ant.exe" if sys.platform == "win32" else "ant"
        dest_dir = managed_bin_dir()
        dest_dir.mkdir(parents=True, exist_ok=True)
        dest = dest_dir / exe_name

        data = None
        if url.endswith(".zip"):
            with zipfile.ZipFile(io.BytesIO(blob)) as z:
                name = next((n for n in z.namelist()
                             if Path(n).name.lower() == exe_name.lower()), None)
                if name is None:
                    return False, f"Gói tải về không chứa {exe_name}."
                data = z.read(name)
        else:
            with tarfile.open(fileobj=io.BytesIO(blob), mode="r:gz") as t:
                member = next((m for m in t.getmembers()
                               if Path(m.name).name == exe_name), None)
                if member is None:
                    return False, f"Gói tải về không chứa {exe_name}."
                f = t.extractfile(member)
                data = f.read() if f else None
        if not data:
            return False, "Không đọc được tệp thực thi trong gói tải về."

        dest.write_bytes(data)
        if sys.platform != "win32":
            dest.chmod(0o755)

        code, out = _run(["--version"], timeout=30)
        if code != 0:
            return False, f"Cài xong nhưng chưa chạy được: {out}"
        say("Đã cài xong.")
        return True, f"Đã cài {out.strip().splitlines()[0] if out else 'ant'} vào {dest}"

    except Exception as e:
        return False, f"{type(e).__name__}: {e}"


def ensure_sdk(progress=None) -> tuple[bool, str]:
    """Cài thư viện anthropic nếu thiếu."""
    ok, ver = _sdk_info()
    if ok:
        return True, ver
    if progress:
        progress("Đang cài thư viện anthropic…")
    try:
        p = subprocess.run(
            [sys.executable, "-m", "pip", "install", "--quiet", "anthropic"],
            capture_output=True, text=True, encoding="utf-8", errors="replace",
            timeout=600, creationflags=_NO_WINDOW,
        )
        if p.returncode != 0:
            return False, (p.stderr or p.stdout or "").strip()[:300]
    except Exception as e:
        return False, f"{type(e).__name__}: {e}"
    ok, ver = _sdk_info()
    return ok, ver if ok else "Cài xong nhưng vẫn chưa import được."


def connect(progress=None) -> tuple[bool, str]:
    """Một bước duy nhất: lo đủ phụ thuộc rồi mở trình duyệt để đăng nhập."""
    ok, msg = ensure_sdk(progress)
    if not ok:
        return False, f"Không cài được thư viện anthropic: {msg}"

    if not find_ant():
        ok, msg = install_cli(progress)
        if not ok:
            return False, msg

    if progress:
        progress("Đang mở trình duyệt để đăng nhập…")
    return login()
