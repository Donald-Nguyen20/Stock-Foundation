# -*- coding: utf-8 -*-
"""Liên kết tài khoản Claude cho các tính năng AI của app.

Phần phân tích chạy qua **claude_agent_sdk**, vốn đóng gói sẵn Claude Code và
dùng chính tài khoản Claude đã đăng nhập — tiêu dung lượng của gói thuê bao
chứ không tính vào tín dụng API.

Vì vậy "liên kết" ở đây là đăng nhập chính Claude Code mà SDK sẽ chạy, bằng
`claude auth login`. Trạng thái đọc từ `claude auth status`, còn việc dùng
được hay không thì kiểm bằng một truy vấn thật với đúng model phân tích: chỉ
cách đó mới bắt được token hết hạn hay bản Claude Code quá cũ so với model.

Claude Code lưu đăng nhập chung cho cả máy, nên đăng xuất ở đây cũng đăng
xuất Claude Code ở mọi nơi khác trên máy này.

Lưu ý điều khoản: cách dùng này dành cho cá nhân với tài khoản của chính mình.
Nếu phát hành app cho người khác dùng thì Anthropic yêu cầu chuyển sang xác
thực bằng API key.
"""
from __future__ import annotations

import json
import os
import shutil
import subprocess
import sys
from dataclasses import dataclass
from pathlib import Path

# Ẩn cửa sổ console khi gọi CLI từ app GUI trên Windows
_NO_WINDOW = subprocess.CREATE_NO_WINDOW if sys.platform == "win32" else 0


# ─────────────────────────────────────────────────────────────────────────────
# Định vị Claude Code
# ─────────────────────────────────────────────────────────────────────────────
def find_claude() -> str | None:
    """Claude Code mà Agent SDK sẽ chạy, hoặc None nếu không có.

    Cùng thứ tự với SDK: bản đi kèm gói trước, không có thì bản trên PATH
    (bản .exe đóng gói bằng PyInstaller không mang theo bản đi kèm).
    """
    exe = "claude.exe" if sys.platform == "win32" else "claude"
    try:
        import claude_agent_sdk
        bundled = Path(claude_agent_sdk.__file__).parent / "_bundled" / exe
        if bundled.is_file():
            return str(bundled)
    except Exception:
        pass
    return shutil.which("claude")


def _run(args: list[str], timeout: int = 30) -> tuple[int, str]:
    """Chạy `claude ...`, trả về (mã thoát, stdout+stderr đã gộp)."""
    cli = find_claude()
    if not cli:
        return 127, ("Không tìm thấy Claude Code. Cài lại bằng: "
                     "pip install --force-reinstall claude-agent-sdk")
    try:
        p = subprocess.run(
            [cli, *args],
            capture_output=True, text=True, encoding="utf-8", errors="replace",
            stdin=subprocess.DEVNULL, timeout=timeout, creationflags=_NO_WINDOW,
        )
        return p.returncode, ((p.stdout or "") + (p.stderr or "")).strip()
    except subprocess.TimeoutExpired:
        return 124, f"Quá thời gian chờ ({timeout} giây)."
    except OSError as e:
        return 1, f"Không chạy được Claude Code: {e}"


# ─────────────────────────────────────────────────────────────────────────────
# Trạng thái
# ─────────────────────────────────────────────────────────────────────────────
@dataclass
class AuthState:
    sdk_ok: bool = False            # claude-agent-sdk đã cài chưa
    sdk_version: str = ""
    cli_path: str | None = None     # Claude Code mà SDK sẽ chạy
    signed_in: bool = False         # `claude auth status` báo đã đăng nhập
    email: str = ""
    plan: str = ""                  # pro, max, …
    verified: bool = False          # chạy thử với model phân tích thành công
    account: str = ""               # mô tả ngắn để hiển thị
    error: str = ""
    api_key_shadow: bool = False    # ANTHROPIC_API_KEY đang che tài khoản OAuth

    @property
    def cli_ok(self) -> bool:
        return bool(self.cli_path)

    @property
    def logged_in(self) -> bool:
        return self.signed_in


def _sdk_info() -> tuple[bool, str]:
    try:
        import claude_agent_sdk
        return True, getattr(claude_agent_sdk, "__version__", "?")
    except Exception:
        return False, ""


def auth_status() -> dict:
    """Kết quả `claude auth status` dạng dict; rỗng nếu không đọc được."""
    _, out = _run(["auth", "status", "--json"], timeout=30)
    start = out.find("{")
    if start < 0:
        return {}
    try:
        data, _ = json.JSONDecoder().raw_decode(out[start:])
    except ValueError:
        return {}
    return data if isinstance(data, dict) else {}


def verify_credentials(model: str | None = None) -> tuple[bool, str]:
    """Chạy thử một truy vấn cực nhỏ với đúng model phân tích.

    Không suy đoán từ tệp cấu hình: chỉ một lần chạy thật mới phân biệt được
    "đã đăng nhập" với "token hết hạn" hay "Claude Code quá cũ cho model này".
    """
    try:
        import asyncio

        from claude_agent_sdk import ClaudeAgentOptions, ResultMessage, query

        import ai_analysis
    except Exception as e:
        return False, f"Chưa cài claude-agent-sdk: {e}"

    model = model or ai_analysis.MODEL_DEFAULT
    out = {"ok": False, "detail": ""}

    async def probe():
        async for m in query(
            prompt="ok",
            options=ClaudeAgentOptions(
                allowed_tools=[], max_turns=1, setting_sources=[], model=model,
                permission_mode="bypassPermissions"),
        ):
            if isinstance(m, ResultMessage):
                out["ok"] = m.subtype == "success" and not m.is_error
                out["detail"] = (m.result or "") if m.is_error else (m.subtype or "")

    loop = asyncio.new_event_loop()
    try:
        asyncio.set_event_loop(loop)
        loop.run_until_complete(probe())
    except Exception as e:
        # SDK chỉ báo "error result: success"; nguyên văn lỗi nằm ở ResultMessage
        err = RuntimeError(out["detail"]) if out["detail"] else e
        return False, ai_analysis.friendly_error(err)
    finally:
        try:
            loop.close()
        finally:
            asyncio.set_event_loop(None)

    if out["ok"]:
        return True, f"Chạy thử {model} thành công — dùng dung lượng gói thuê bao."
    return False, ai_analysis.friendly_error(
        RuntimeError(out["detail"] or "Phiên thử kết thúc mà không rõ lý do."))


def read_state(verify: bool = True) -> AuthState:
    st = AuthState()
    st.sdk_ok, st.sdk_version = _sdk_info()
    st.cli_path = find_claude()

    # Bẫy số một theo tài liệu: API key đặt sẵn sẽ che tài khoản OAuth,
    # và chuỗi rỗng vẫn chiếm chỗ ưu tiên đó.
    st.api_key_shadow = "ANTHROPIC_API_KEY" in os.environ

    if st.cli_ok:
        info = auth_status()
        st.signed_in = bool(info.get("loggedIn"))
        st.email = info.get("email") or ""
        st.plan = info.get("subscriptionType") or ""

    # Chưa đăng nhập thì chạy thử chỉ tốn thời gian mà chắc chắn hỏng
    if verify and st.sdk_ok and st.signed_in:
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
def login(timeout: int = 300) -> tuple[bool, str]:
    """Mở trình duyệt để đăng nhập bằng gói thuê bao. Chặn tới khi xong hoặc hết giờ."""
    code, out = _run(["auth", "login", "--claudeai"], timeout=timeout)
    if code != 0:
        return False, out or f"Đăng nhập thất bại (mã {code})."
    ok, msg = verify_credentials()
    return ok, msg


def logout() -> tuple[bool, str]:
    """Đăng xuất Claude Code — áp dụng cho mọi nơi dùng Claude Code trên máy."""
    code, out = _run(["auth", "logout"], timeout=60)
    return code == 0, out or ("Đã đăng xuất." if code == 0 else "Đăng xuất thất bại.")


# ─────────────────────────────────────────────────────────────────────────────
# Tự cài đặt
# ─────────────────────────────────────────────────────────────────────────────
def ensure_sdk(progress=None) -> tuple[bool, str]:
    """Cài claude-agent-sdk nếu thiếu."""
    ok, ver = _sdk_info()
    if ok:
        return True, ver
    if progress:
        progress("Đang cài claude-agent-sdk…")
    try:
        p = subprocess.run(
            [sys.executable, "-m", "pip", "install", "--quiet", "claude-agent-sdk"],
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
        return False, f"Không cài được claude-agent-sdk: {msg}"

    if progress:
        progress("Đang mở trình duyệt để đăng nhập…")
    return login()
