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
    """Đường dẫn tới CLI `ant`, hoặc None nếu chưa cài."""
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
