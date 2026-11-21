"""OCR chứng từ scan / ảnh chụp bằng bộ OCR có sẵn của Windows (Windows.Media.Ocr).

* Không cần cài phần mềm OCR riêng; chạy nhanh trên máy cấu hình thấp.
* Dùng tiếng Việt nếu Windows đã cài gói OCR tiếng Việt, nếu không dùng tiếng Anh (vẫn đọc được số, ngày,
  MST, số tiền; chữ có dấu có thể bị mất dấu – bộ đọc trường của phần mềm đã nhận được chữ không dấu).
  Cài gói tiếng Việt (PowerShell quyền quản trị):
      Add-WindowsCapability -Online -Name "Language.OCR~~~vi-VN~0.0.1.0"
* Dùng cho trang chứng từ (hóa đơn, sao kê). Không dùng cho mục đích khác.
"""
from __future__ import annotations

import asyncio
import io
import os
import tempfile
from typing import Any

MAX_PAGES = 6


def _engine():
    from winsdk.windows.globalization import Language
    from winsdk.windows.media.ocr import OcrEngine
    for tag in ("vi-VN", "vi"):
        try:
            if OcrEngine.is_language_supported(Language(tag)):
                return OcrEngine.try_create_from_language(Language(tag)), "vi"
        except Exception:  # noqa: BLE001
            pass
    eng = OcrEngine.try_create_from_language(Language("en-US")) or OcrEngine.try_create_from_user_profile_languages()
    return eng, "en"


def ocr_language() -> str:
    try:
        return _engine()[1]
    except Exception:  # noqa: BLE001
        return "none"


def _prepare(img, ink: bool = False):
    """Ảnh xám, tăng tương phản, phóng to ảnh nhỏ (OCR đọc tốt chữ cao ~30px).
    ink=True (lượt đọc thứ 2): lấy kênh màu TỐI NHẤT của mỗi điểm -> chữ in đỏ/xanh (số hóa đơn, ký hiệu) đậm như
    chữ đen. Không dùng mặc định vì trang có nền / hoa văn màu sẽ bị tối theo và mất chữ khác."""
    from PIL import ImageChops, ImageOps
    img = ImageOps.exif_transpose(img)
    if ink and img.mode not in ("L", "1", "I", "I;16", "F"):
        r, g, b = img.convert("RGB").split()
        img = ImageChops.darker(ImageChops.darker(r, g), b)
    img = img.convert("L")
    img = ImageOps.autocontrast(img, cutoff=1)
    w, h = img.size
    if w < 1800:
        k = 1800 / w
        img = img.resize((int(w * k), int(h * k)))
    if max(img.size) > 9500:  # giới hạn kích thước của Windows OCR
        k = 9500 / max(img.size)
        img = img.resize((int(img.size[0] * k), int(img.size[1] * k)))
    return img


_ENG = None


async def _recognize_file(path: str) -> str:
    from winsdk.windows.graphics.imaging import BitmapDecoder
    from winsdk.windows.storage import FileAccessMode, StorageFile
    f = await StorageFile.get_file_from_path_async(path)
    stream = await f.open_async(FileAccessMode.READ)
    bmp = None
    try:
        dec = await BitmapDecoder.create_async(stream)
        bmp = await dec.get_software_bitmap_async()
        global _ENG
        if _ENG is None:                 # chỉ luồng OCR dùng -> tạo 1 lần, dùng lại
            _ENG = _engine()[0]
        res = await _ENG.recognize_async(bmp)
    finally:
        # ĐÓNG ảnh + luồng đọc file: không đóng thì lần OCR kế tiếp báo "Not enough memory resources"
        for o in (bmp, stream):
            try:
                if o is not None:
                    o.close()
            except Exception:  # noqa: BLE001
                pass
    words = []
    for line in res.lines:
        for w in line.words:
            r = w.bounding_rect
            words.append((r.y + r.height / 2, r.x, r.height, r.width, w.text))
    return _rows(words)


def _rows(words: list[tuple[float, float, float, float, str]]) -> str:
    """Dựng lại các hàng theo tọa độ: chữ có tâm cùng độ cao -> 1 hàng, xếp trái -> phải.
    Khoảng trống lớn giữa 2 chữ (cột bảng) -> chèn 3 dấu cách."""
    if not words:
        return ""
    hs = sorted(w[2] for w in words)
    tol = max(4.0, hs[len(hs) // 2] * 0.55)
    rows: list[list[tuple]] = []
    for w in sorted(words, key=lambda x: x[0]):
        if rows and abs(rows[-1][0][0] - w[0]) <= tol:
            rows[-1].append(w)
        else:
            rows.append([w])
    out = []
    for row in rows:
        row.sort(key=lambda x: x[1])
        s, prev_end = "", None
        for yc, x, h, wd, txt in row:
            if prev_end is not None:
                s += "   " if x - prev_end > h * 1.8 else " "
            s += txt
            prev_end = x + wd
        out.append(s)
    return "\n".join(out)


class _OcrThread:
    """MỌI lần OCR của Windows (WinRT) chạy trên 1 luồng riêng, sống suốt phiên, 1 vòng lặp asyncio.
    Trước đây mỗi luồng đọc hóa đơn tự gọi asyncio.run(...) -> đọc vài trăm bản scan liên tục thì phần mềm bị tắt
    hẳn (lỗi bộ nhớ trong thư viện WinRT, mã thoát 139). Các luồng khác chỉ gửi việc vào hàng đợi rồi chờ kết quả."""

    def __init__(self) -> None:
        import queue
        import threading
        self._q: "queue.Queue" = queue.Queue()
        self._lock = threading.Lock()
        self._thread: "threading.Thread | None" = None

    def _loop(self) -> None:
        loop = asyncio.new_event_loop()
        asyncio.set_event_loop(loop)
        while True:
            path, box, ev = self._q.get()
            try:
                box["r"] = loop.run_until_complete(_recognize_file(path))
            except BaseException as e:  # noqa: BLE001 - trả lỗi về luồng gọi
                box["e"] = e
            ev.set()

    def run(self, path: str, timeout: float = 180) -> str:
        import threading
        with self._lock:
            if not (self._thread and self._thread.is_alive()):
                self._thread = threading.Thread(target=self._loop, daemon=True, name="windows-ocr")
                self._thread.start()
        box: dict[str, Any] = {}
        ev = threading.Event()
        self._q.put((path, box, ev))
        if not ev.wait(timeout):
            raise OSError("OCR của Windows không phản hồi (quá thời gian)")
        if "e" in box:
            raise box["e"]
        return box["r"]


_OCR = _OcrThread()


def ocr_pil(img, ink: bool = False) -> str:
    """Bộ OCR của Windows KHÔNG chạy được nhiều lần nhận dạng cùng lúc (nhiều luồng / tiến trình -> lỗi
    "Not enough memory resources", mã -2147024882): mọi lần OCR xếp hàng trên 1 luồng riêng, gặp lỗi đó thì thử lại."""
    import time
    img = _prepare(img, ink)
    fd, path = tempfile.mkstemp(suffix=".png")
    os.close(fd)
    try:
        img.save(path, "PNG")
        for attempt in range(6):
            try:
                return _OCR.run(path)
            except OSError as e:   # OCR đang bận ở tiến trình khác / tài nguyên tạm thiếu
                if "-2147024882" not in str(e) and "memory" not in str(e).lower() or attempt == 5:
                    raise
                time.sleep(0.4 * (attempt + 1))
        return ""
    finally:
        try:
            os.remove(path)
        except OSError:
            pass


def ocr_image(data: bytes) -> str:
    from PIL import Image
    return ocr_pil(Image.open(io.BytesIO(data)))


_PDFIUM_LOCK = __import__("threading").Lock()


def pdf_images(data: bytes, dpi: int = 300, max_pages: int = MAX_PAGES) -> list[Any]:
    """PDFium KHÔNG an toàn khi nhiều luồng dùng cùng lúc (đọc song song hàng trăm PDF scan có thể làm treo / tắt
    hẳn phần mềm) -> mỗi lúc chỉ 1 luồng mở / vẽ PDF; phần OCR và đọc chữ vẫn chạy song song."""
    import pypdfium2 as pdfium
    with _PDFIUM_LOCK:
        pdf = pdfium.PdfDocument(data)
        try:
            return [pdf[i].render(scale=dpi / 72).to_pil() for i in range(min(len(pdf), max_pages))]
        finally:
            pdf.close()


def ocr_pdf(data: bytes, max_pages: int = MAX_PAGES) -> str:
    return "\n".join(ocr_pil(img) for img in pdf_images(data, max_pages=max_pages))


def page_images(data: bytes, max_pages: int = 3) -> list[Any]:
    """Ảnh từng trang của PDF scan hoặc ảnh chụp – render 1 lần, dùng cho cả 2 lượt OCR."""
    if data[:5] == b"%PDF-":
        return pdf_images(data, max_pages=max_pages)
    from PIL import Image
    return [Image.open(io.BytesIO(data))]


def is_image(name: str, data: bytes) -> bool:
    return name.lower().endswith((".jpg", ".jpeg", ".png", ".bmp", ".tif", ".tiff", ".webp")) or \
        data[:3] == b"\xff\xd8\xff" or data[:8] == b"\x89PNG\r\n\x1a\n"
