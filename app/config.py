import os
from dataclasses import dataclass, field
from pathlib import Path

from dotenv import load_dotenv

ROOT = Path(__file__).resolve().parent.parent
load_dotenv(ROOT / ".env")


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
    gdt_delay: float = float(os.getenv("GDT_REQUEST_DELAY", "0.6"))
    render_fallback: bool = _bool("RENDER_FALLBACK_PDF", True)
    fallback_font: str = os.getenv("FALLBACK_FONT", "C:/Windows/Fonts/arial.ttf")
    host: str = os.getenv("APP_HOST", "127.0.0.1")
    port: int = int(os.getenv("APP_PORT", "8000"))
    providers_file: Path = field(default=ROOT / "app" / "providers" / "providers.json")

    # --- Tự động tải qua Chrome (người dùng tự đăng nhập) ---
    output_dir: str = os.getenv("OUTPUT_DIR", str(Path.home() / "Documents" / "HoaDonDienTu"))
    chrome_path: str = os.getenv("CHROME_PATH", "")
    cdp_port: int = int(os.getenv("CHROME_DEBUG_PORT", "9333"))

    # --- Tự nhập file từ thư mục tải về ---
    watch_dir: str = os.getenv("WATCH_DIR", str(Path.home() / "Downloads"))
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
