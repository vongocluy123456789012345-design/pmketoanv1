"""Khởi động bản desktop: chạy máy chủ nội bộ rồi mở một cửa sổ ứng dụng (Chrome/Edge chế độ --app).

Bấm đúp AWBook.exe -> cửa sổ phần mềm hiện ra. Đóng cửa sổ -> phần mềm tự thoát.
Dữ liệu (cấu hình .env, hồ sơ Chrome đăng nhập, lịch, log) nằm trong thư mục "data" cạnh file exe.
"""
from __future__ import annotations

import multiprocessing
import os
import socket
import subprocess
import sys
import threading
import time
import urllib.request
import webbrowser
from pathlib import Path


def _setup_logging(root: Path) -> None:
    """Bản không có cửa sổ console: ghi log ra data/app.log (và tránh lỗi khi stdout là None)."""
    (root / "data").mkdir(parents=True, exist_ok=True)
    log = open(root / "data" / "app.log", "a", encoding="utf-8", buffering=1)
    if sys.stdout is None or getattr(sys, "frozen", False):
        sys.stdout = log
    if sys.stderr is None or getattr(sys, "frozen", False):
        sys.stderr = log


def _port_free(port: int) -> bool:
    with socket.socket() as s:
        s.settimeout(0.5)
        return s.connect_ex(("127.0.0.1", port)) != 0


def _is_our_app(port: int) -> bool:
    try:
        with urllib.request.urlopen(f"http://127.0.0.1:{port}/api/ping", timeout=2) as r:
            return b'"hoadon"' in r.read()
    except Exception:  # noqa: BLE001
        return False


def _find_browser() -> str | None:
    cands = []
    for var in ("PROGRAMFILES", "PROGRAMFILES(X86)", "LOCALAPPDATA"):
        base = os.environ.get(var)
        if base:
            cands += [Path(base) / "Google/Chrome/Application/chrome.exe",
                      Path(base) / "Microsoft/Edge/Application/msedge.exe"]
    return next((str(c) for c in cands if c.is_file()), None)


def _ui_alive(prof_key: str) -> bool:
    """Còn tiến trình trình duyệt nào đang dùng hồ sơ của cửa sổ phần mềm không."""
    import psutil
    for p in psutil.process_iter(["cmdline"]):
        try:
            cl = " ".join(p.info["cmdline"] or []).lower()
        except Exception:  # noqa: BLE001
            continue
        if prof_key in cl:
            return True
    return False


def _window_geometry() -> tuple[int, int, int, int]:
    """Kích thước + vị trí cửa sổ (đơn vị Chrome dùng = điểm đã tính tỉ lệ phóng to của Windows), vừa trong
    vùng làm việc của màn hình. Màn hình nhỏ / phóng to 125-150% (vd 1920x1080 @150% = 1280x672) mà mở cố định
    1280x860 thì đáy cửa sổ (nút Trợ lý AI, nút bấm cuối trang) bị tràn ra ngoài màn hình."""
    w, h = 1280, 860
    try:
        import ctypes
        from ctypes import wintypes

        user32 = ctypes.windll.user32
        try:
            user32.SetProcessDPIAware()  # lấy số điểm ảnh thật rồi tự quy đổi theo DPI
        except Exception:  # noqa: BLE001
            pass
        rc = wintypes.RECT()
        user32.SystemParametersInfoW(0x0030, 0, ctypes.byref(rc), 0)  # SPI_GETWORKAREA (trừ thanh taskbar)
        try:
            dpi = user32.GetDpiForSystem() or 96
        except Exception:  # noqa: BLE001
            dpi = 96
        k = 96 / dpi
        aw, ah = int((rc.right - rc.left) * k), int((rc.bottom - rc.top) * k)
        ax, ay = int(rc.left * k), int(rc.top * k)
        if aw > 200 and ah > 200:
            w, h = min(w, aw), min(h, ah)
            return w, h, ax + (aw - w) // 2, ay + (ah - h) // 2
    except Exception:  # noqa: BLE001
        pass
    return w, h, 40, 20


def _forget_window_placement(prof: Path) -> None:
    """Chrome nhớ kích thước cửa sổ app lần trước (có thể đã tràn màn hình) -> xóa để dùng kích thước mới."""
    import json

    f = prof / "Default" / "Preferences"
    try:
        data = json.loads(f.read_text(encoding="utf-8"))
        browser = data.get("browser") or {}
        if "app_window_placement" in browser or "window_placement" in browser:
            browser.pop("app_window_placement", None)
            browser.pop("window_placement", None)
            f.write_text(json.dumps(data), encoding="utf-8")
    except Exception:  # noqa: BLE001 - chưa có hồ sơ / file đang bị khóa -> bỏ qua
        pass


def _open_window(url: str, root: Path):
    exe = _find_browser()
    if not exe:
        webbrowser.open(url)
        return None
    prof = root / "data" / "ui-profile"  # hồ sơ riêng cho cửa sổ phần mềm (khác hồ sơ đăng nhập cổng thuế)
    prof.mkdir(parents=True, exist_ok=True)
    _forget_window_placement(prof)
    w, h, x, y = _window_geometry()
    return subprocess.Popen([exe, f"--app={url}", f"--user-data-dir={prof}", "--no-first-run",
                             "--no-default-browser-check", f"--window-size={w},{h}", f"--window-position={x},{y}"])


def main() -> None:
    multiprocessing.freeze_support()  # bắt buộc với bản đóng gói (đọc PDF dùng nhiều tiến trình)
    from app.config import ROOT, settings

    if getattr(sys, "frozen", False):
        _setup_logging(ROOT)

    port = settings.port
    if not _port_free(port):
        if _is_our_app(port):  # đã có một cửa sổ phần mềm đang chạy -> chỉ mở thêm cửa sổ
            _open_window(f"http://127.0.0.1:{port}/", ROOT)
            return
        port = next(p for p in range(port + 1, port + 50) if _port_free(p))

    import uvicorn

    from app import main as appmain

    server = uvicorn.Server(uvicorn.Config(appmain.app, host="127.0.0.1", port=port, log_level="warning",
                                           log_config=None))
    threading.Thread(target=server.run, daemon=True).start()
    for _ in range(100):
        if server.started:
            break
        time.sleep(0.1)
    url = f"http://127.0.0.1:{port}/"
    proc = _open_window(url, ROOT)

    t0 = time.time()
    prof_key = str(ROOT / "data" / "ui-profile").lower()
    seen_ui, gone = False, 0
    # Thoát khi cửa sổ phần mềm bị đóng = không còn tiến trình trình duyệt nào dùng hồ sơ "ui-profile".
    # Cách này đúng kể cả khi trình duyệt chuyển giao cho tiến trình khác. KHÔNG dùng nhịp tim làm căn cứ
    # chính: cửa sổ bị che/thu nhỏ sau cửa sổ đăng nhập cổng thuế thì Chrome làm chậm bộ đếm giờ của trang
    # (có thể hàng phút) và phần mềm sẽ tưởng cửa sổ đã đóng, tự tắt giữa chừng.
    # Nhịp tim chỉ là dự phòng khi không tìm thấy tiến trình cửa sổ, và chờ rất lâu (30 phút).
    while not server.should_exit:
        time.sleep(3)
        age = time.time() - t0
        alive = _ui_alive(prof_key)
        if alive:
            seen_ui, gone = True, 0
        elif seen_ui:
            gone += 1
            if gone >= 2:     # 2 lần liên tiếp (~6 giây) không còn cửa sổ -> đã đóng
                break
        else:                 # chưa từng thấy tiến trình cửa sổ (vd mở bằng trình duyệt mặc định) -> dự phòng
            last = appmain.PING["last"]
            if last and time.time() - last > 1800:
                break
            if not last and age > 300:
                break
    # Dừng Chrome đăng nhập cổng thuế do phần mềm mở (nếu còn) để không để lại tiến trình mồ côi
    try:
        from app.gdt.browser import BROWSER
        if BROWSER._browser is not None:
            BROWSER.call(lambda: BROWSER._browser.close(), 10)
    except Exception:  # noqa: BLE001
        pass
    server.should_exit = True
    if proc and proc.poll() is None:
        proc.terminate()
    time.sleep(1)
    os._exit(0)


if __name__ == "__main__":
    main()
