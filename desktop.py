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


# Bản exe không có cửa sổ console: stdout/stderr = None, kể cả trong các TIẾN TRÌNH PHỤ đọc hóa đơn song song.
# Tiến trình phụ gặp lỗi mà ghi log vào None -> bật hộp "Unhandled exception in script". Gắn sẵn chỗ ghi rỗng.
if sys.stdout is None:
    sys.stdout = open(os.devnull, "w", encoding="utf-8")
if sys.stderr is None:
    sys.stderr = open(os.devnull, "w", encoding="utf-8")


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


_BRANDED: set[int] = set()
_ICON_HANDLES: dict[str, int] = {}


def _brand_windows(icon: Path, relaunch: str) -> int:
    """Cửa sổ phần mềm là cửa sổ Chrome/Edge chế độ app -> mặc định thanh taskbar hiện biểu tượng trình duyệt.
    Gắn cho cửa sổ (tiêu đề "AWBook") 1 mã ứng dụng riêng (AppUserModelID) + icon AWBook -> taskbar hiện logo
    AWBook, tách nhóm khỏi Chrome; ghim vào taskbar sẽ mở lại AWBook.exe. Trả về số cửa sổ vừa gắn."""
    if os.name != "nt" or not icon.is_file():
        return 0
    try:
        import ctypes
        from ctypes import POINTER, Structure, byref, c_ubyte, c_ulong, c_ushort, c_void_p, wintypes

        user32, shell32 = ctypes.windll.user32, ctypes.windll.shell32
        user32.LoadImageW.restype = wintypes.HANDLE
        user32.SendMessageW.argtypes = [wintypes.HWND, wintypes.UINT, wintypes.WPARAM, wintypes.LPARAM]
        hwnds: list[int] = []

        @ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)
        def _cb(h, _l):
            if user32.IsWindowVisible(h):
                t = ctypes.create_unicode_buffer(256)
                c = ctypes.create_unicode_buffer(64)
                user32.GetWindowTextW(h, t, 256)
                user32.GetClassNameW(h, c, 64)
                if c.value == "Chrome_WidgetWin_1" and t.value.strip() == "AWBook":
                    hwnds.append(int(h))
            return True

        user32.EnumWindows(_cb, 0)
        if not hwnds:
            return 0
        # icon cửa sổ (góc trái + taskbar)
        for size, flag in ((11, "big"), (49, "small")):   # SM_CXICON / SM_CXSMICON
            if flag not in _ICON_HANDLES:
                px = user32.GetSystemMetrics(size) or (32 if flag == "big" else 16)
                _ICON_HANDLES[flag] = user32.LoadImageW(None, str(icon), 1, px, px, 0x10) or 0  # IMAGE_ICON, LR_LOADFROMFILE

        class GUID(Structure):
            _fields_ = [("d1", c_ulong), ("d2", c_ushort), ("d3", c_ushort), ("d4", c_ubyte * 8)]

        class PKEY(Structure):
            _fields_ = [("fmtid", GUID), ("pid", c_ulong)]

        class PROPVARIANT(Structure):
            _fields_ = [("vt", c_ushort), ("r1", c_ushort), ("r2", c_ushort), ("r3", c_ushort),
                        ("p", c_void_p), ("p2", c_void_p)]

        iid = GUID(0x886D8EEB, 0x8CF2, 0x4446, (c_ubyte * 8)(0x8D, 0x02, 0xCD, 0xBA, 0x1D, 0xBD, 0xCF, 0x99))
        fmt = (0x9F4C2855, 0x9F79, 0x4B39, (c_ubyte * 8)(0xA8, 0xD0, 0xE1, 0xD4, 0x2D, 0xE1, 0xD5, 0xF3))
        props = [(2, relaunch), (4, "AWBook"), (3, f"{icon},0"), (5, "AWBook.PhanMemKeToan")]  # ID gắn sau cùng
        n = 0
        for h in hwnds:
            for flag, wp in (("big", 1), ("small", 0)):
                if _ICON_HANDLES.get(flag):
                    user32.SendMessageW(h, 0x0080, wp, _ICON_HANDLES[flag])   # WM_SETICON
            if h in _BRANDED:
                continue
            store = c_void_p()
            if shell32.SHGetPropertyStoreForWindow(wintypes.HWND(h), byref(iid), byref(store)) != 0 or not store:
                continue
            vt = ctypes.cast(store, POINTER(POINTER(c_void_p)))[0]
            set_value = ctypes.WINFUNCTYPE(ctypes.c_long, c_void_p, POINTER(PKEY), POINTER(PROPVARIANT))(vt[6])
            commit = ctypes.WINFUNCTYPE(ctypes.c_long, c_void_p)(vt[7])
            release = ctypes.WINFUNCTYPE(c_ulong, c_void_p)(vt[2])
            keep = []
            for pid, val in props:
                buf = ctypes.create_unicode_buffer(val)
                keep.append(buf)
                pv = PROPVARIANT(vt=31, p=ctypes.cast(buf, c_void_p))   # VT_LPWSTR
                set_value(store, byref(PKEY(GUID(*fmt), pid)), byref(pv))
            commit(store)
            release(store)
            _BRANDED.add(h)
            n += 1
        return n
    except Exception:  # noqa: BLE001 - không gắn được icon thì vẫn chạy bình thường
        return 0


def _brand_soon(icon: Path, relaunch: str) -> None:
    """Gắn icon ngay khi cửa sổ vừa hiện (dò nhanh trong 40 giây đầu)."""
    def run() -> None:
        for _ in range(80):
            _brand_windows(icon, relaunch)
            time.sleep(0.5)
    threading.Thread(target=run, daemon=True).start()


def _icon_path() -> Path:
    base = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))
    return base / "app" / "assets" / "awbook.ico"


def _relaunch_cmd() -> str:
    if getattr(sys, "frozen", False):
        return f'"{sys.executable}"'
    return f'"{sys.executable}" "{Path(__file__).resolve()}"'


def _open_window(url: str, root: Path):
    exe = _find_browser()
    if not exe:
        webbrowser.open(url)
        return None
    prof = root / "data" / "ui-profile"  # hồ sơ riêng cho cửa sổ phần mềm (khác hồ sơ đăng nhập cổng thuế)
    prof.mkdir(parents=True, exist_ok=True)
    _forget_window_placement(prof)
    w, h, x, y = _window_geometry()
    proc = subprocess.Popen([exe, f"--app={url}", f"--user-data-dir={prof}", "--no-first-run",
                             "--no-default-browser-check", f"--window-size={w},{h}", f"--window-position={x},{y}"])
    _brand_soon(_icon_path(), _relaunch_cmd())   # logo AWBook trên thanh taskbar thay cho biểu tượng trình duyệt
    return proc


def main() -> None:
    multiprocessing.freeze_support()  # bắt buộc với bản đóng gói (đọc PDF dùng nhiều tiến trình)
    from app.config import ROOT, settings

    if getattr(sys, "frozen", False):
        _setup_logging(ROOT)

    port = settings.port
    if "--nhac-vat" in sys.argv:   # Windows Task Scheduler gọi đúng ngày giờ đã đặt: hiện bảng nhắc dự báo VAT
        from app import vat_lich
        running = not _port_free(port) and _is_our_app(port)
        if not vat_lich.show_reminder(running):
            return
        if running:
            _open_window(f"http://127.0.0.1:{port}/#vat", ROOT)
            return
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
    url = f"http://127.0.0.1:{port}/" + ("#vat" if "--nhac-vat" in sys.argv else "")   # từ bảng nhắc: vào Dự báo VAT
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
        _brand_windows(_icon_path(), _relaunch_cmd())  # cửa sổ mở thêm / tải lại vẫn giữ logo AWBook
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
