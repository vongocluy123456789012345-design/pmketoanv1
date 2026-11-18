"""Bộ giải CAPTCHA tự động dùng ddddocr (hỗ trợ PNG/JPG bytes và SVG string)."""
from __future__ import annotations

import logging
import re

log = logging.getLogger(__name__)
_OCR = None


def _get_ocr():
    global _OCR
    if _OCR is None:
        import ddddocr
        _OCR = ddddocr.DdddOcr(show_ad=False)
    return _OCR


def solve_bytes(img_bytes: bytes) -> str:
    """Giải CAPTCHA từ dữ liệu ảnh (PNG/JPEG/BMP). Trả về chuỗi rỗng nếu lỗi."""
    if not img_bytes:
        return ""
    try:
        ocr = _get_ocr()
        text = ocr.classification(img_bytes) or ""
        return re.sub(r"[^a-zA-Z0-9]", "", text).strip()
    except Exception as e:  # noqa: BLE001
        log.warning("ddddocr solve_bytes error: %s", e)
        return ""


def solve_svg(svg_text: str) -> str:
    """Chuyển SVG captcha sang PNG rồi giải bằng ddddocr (dùng cho cổng thuế GDT)."""
    if not svg_text:
        return ""
    try:
        import fitz  # PyMuPDF (thường đã có sẵn trong project để đọc/in PDF)
        doc = fitz.open(stream=svg_text.encode("utf-8"), filetype="svg")
        pix = doc[0].get_pixmap(dpi=200)
        return solve_bytes(pix.tobytes("png"))
    except Exception as e:  # noqa: BLE001
        log.warning("ddddocr solve_svg error: %s", e)
        return ""