import os
from dataclasses import dataclass, field
from pathlib import Path

from dotenv import load_dotenv

import sys

# Bản đóng gói (.exe): dữ liệu người dùng (.env, data/) nằm CẠNH file exe; tài nguyên đi kèm (giao diện,
# danh sách nhà cung cấp) nằm trong thư mục giải nén của chương trình.
FROZEN = bool(getattr(sys, "frozen", False))
ROOT = Path(sys.executable).resolve().parent if FROZEN else Path(__file__).resolve().parent.parent
RES = Path(getattr(sys, "_MEIPASS", ROOT)) if FROZEN else ROOT
load_dotenv(ROOT / ".env")
# máy có proxy (công ty / VPN): mọi kết nối tới chính máy này (giao diện, cổng debug Chrome) không đi qua proxy
os.environ["NO_PROXY"] = ",".join(x for x in (os.environ.get("NO_PROXY", ""), "127.0.0.1,localhost,::1") if x)
os.environ["no_proxy"] = os.environ["NO_PROXY"]


def _documents_dir() -> Path:
    """Thư mục Documents thật của máy này (kể cả khi bị chuyển sang OneDrive hoặc ổ khác)."""
    try:
        import ctypes
        from ctypes import wintypes

        class GUID(ctypes.Structure):
            _fields_ = [("a", wintypes.DWORD), ("b", wintypes.WORD), ("c", wintypes.WORD), ("d", ctypes.c_ubyte * 8)]

        fid = GUID(0xFDD39AD0, 0x238F, 0x46AF, (ctypes.c_ubyte * 8)(0xAD, 0xB4, 0x6C, 0x85, 0x48, 0x03, 0x69, 0xC7))
        out = ctypes.c_wchar_p()
        if ctypes.windll.shell32.SHGetKnownFolderPath(ctypes.byref(fid), 0, None, ctypes.byref(out)) == 0:
            p = Path(out.value)
            ctypes.windll.ole32.CoTaskMemFree(out)
            return p
    except Exception:  # noqa: BLE001
        pass
    return Path.home() / "Documents"


def default_output_dir() -> str:
    """Nơi lưu mặc định: <Documents>\\HoaDonDienTu của máy này. Máy đã có sẵn thư mục cũ
    (~\\Documents\\HoaDonDienTu) thì giữ nguyên để không mất dữ liệu đã tải."""
    legacy = Path.home() / "Documents" / "HoaDonDienTu"
    return str(legacy if legacy.exists() else _documents_dir() / "HoaDonDienTu")


USER_SETTINGS = ROOT / "data" / "settings.json"


def _load_user_settings() -> dict:
    try:
        import json
        return json.loads(USER_SETTINGS.read_text(encoding="utf-8"))
    except (FileNotFoundError, ValueError, OSError):
        return {}


def _bool(name: str, default: bool) -> bool:
    v = os.getenv(name)
    if v is None:
        return default
    return v.strip().lower() in ("1", "true", "yes", "on")


@dataclass(frozen=True)
class Settings:
    supabase_url: str = os.getenv("SUPABASE_URL", "")
    supabase_key: str = os.getenv("SUPABASE_SERVICE_KEY", "")
    bucket: str = os.getenv("SUPABASE_BUCKET", "invoices")
    gdt_base_url: str = os.getenv("GDT_BASE_URL", "https://hoadondientu.gdt.gov.vn/api")
    gdt_verify_ssl: bool = _bool("GDT_VERIFY_SSL", True)
    gdt_delay: float = float(os.getenv("GDT_REQUEST_DELAY", "0.3"))
    render_fallback: bool = _bool("RENDER_FALLBACK_PDF", True)
    # Lấy PDF gốc bằng Chrome chạy ẩn trên trang tra cứu của NCC (không cần đăng nhập cổng thuế)
    headless_pdf: bool = _bool("HEADLESS_PDF", True)
    fallback_font: str = os.getenv("FALLBACK_FONT", "C:/Windows/Fonts/arial.ttf")
    host: str = os.getenv("APP_HOST", "127.0.0.1")
    port: int = int(os.getenv("APP_PORT", "8000"))
    providers_file: Path = field(default=RES / "app" / "providers" / "providers.json")

    # --- Tự động tải qua Chrome (người dùng tự đăng nhập) ---
    output_dir: str = os.getenv("OUTPUT_DIR") or default_output_dir()
    chrome_path: str = os.getenv("CHROME_PATH", "")
    cdp_port: int = int(os.getenv("CHROME_DEBUG_PORT") or "9333")  # dòng trống trong .env vẫn lấy mặc định

    # --- Tự nhập file từ thư mục tải về ---
    watch_dir: str = os.getenv("WATCH_DIR") or str(Path.home() / "Downloads")
    watch_interval: float = float(os.getenv("WATCH_INTERVAL", "5"))

    # --- TConnect Open API (lấy hộ hóa đơn từ cổng thuế) ---
    tconnect_auth_base: str = os.getenv("TCONNECT_AUTH_BASE", "https://sme-open-api-sandbox.tconnect.vn")
    tconnect_partner_code: str = os.getenv("TCONNECT_PARTNER_CODE", "")
    tconnect_client_id: str = os.getenv("TCONNECT_CLIENT_ID", "")
    tconnect_client_secret: str = os.getenv("TCONNECT_CLIENT_SECRET", "")
    tconnect_aes_key: str = os.getenv("TCONNECT_AES_KEY", "")
    tconnect_username: str = os.getenv("TCONNECT_USERNAME", "")
    tconnect_password: str = os.getenv("TCONNECT_PASSWORD", "")


settings = Settings()

# Thứ tự ưu tiên nơi lưu: người dùng chọn trong phần mềm (data/settings.json) > OUTPUT_DIR trong .env > mặc định (Documents)
_DEFAULT_OUT = settings.output_dir
if _load_user_settings().get("output_dir"):
    object.__setattr__(settings, "output_dir", _load_user_settings()["output_dir"])


def custom_output_dir() -> bool:
    return bool(_load_user_settings().get("output_dir"))


def set_output_dir(path: str | None) -> str:
    """Đổi nơi lưu (path=None: về mặc định). Kiểm tra tạo được thư mục và ghi được file; nhớ cho lần sau."""
    import json
    s = _load_user_settings()
    if path:
        p = Path(path).expanduser()
        if not p.is_absolute():
            raise ValueError("Hãy nhập đường dẫn đầy đủ, ví dụ D:\\HoaDon")
        try:
            p.mkdir(parents=True, exist_ok=True)
            probe = p / ".awbook_test"
            probe.write_text("ok", encoding="utf-8")
            probe.unlink()
        except OSError as e:
            raise ValueError(f"Không ghi được vào thư mục này: {e}") from e
        s["output_dir"] = str(p)
    else:
        s.pop("output_dir", None)
    USER_SETTINGS.parent.mkdir(parents=True, exist_ok=True)
    USER_SETTINGS.write_text(json.dumps(s, ensure_ascii=False), encoding="utf-8")
    object.__setattr__(settings, "output_dir", s.get("output_dir") or _DEFAULT_OUT)
    return settings.output_dir
