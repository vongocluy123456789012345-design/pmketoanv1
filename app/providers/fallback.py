"""Dựng PDF "bản thể hiện" từ dữ liệu XML khi không lấy được PDF gốc của NCC.

PDF này được đánh dấu pdf_is_original = false và có dòng ghi chú trên trang.
"""
from __future__ import annotations

import io
from pathlib import Path
from typing import Any

from jinja2 import Environment, FileSystemLoader, select_autoescape
from xhtml2pdf import pisa

from ..config import settings

_env = Environment(
    loader=FileSystemLoader(Path(__file__).resolve().parent.parent / "templates"),
    autoescape=select_autoescape(["html"]),
)


def _money(v: Any) -> str:
    if v is None:
        return ""
    try:
        return f"{float(v):,.0f}".replace(",", ".")
    except (TypeError, ValueError):
        return str(v)


_env.filters["money"] = _money


_FONT_NAME: str | None = None


def _register_font() -> str | None:
    """Đăng ký font TTF có tiếng Việt với reportlab (xhtml2pdf không cho đọc file ngoài thư mục)."""
    global _FONT_NAME
    if _FONT_NAME:
        return _FONT_NAME
    regular = Path(settings.fallback_font)
    if not regular.exists():
        return None
    from reportlab.lib.fonts import addMapping
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont

    stem, ext = regular.stem, regular.suffix
    # arial.ttf -> arialbd.ttf / ariali.ttf / arialbi.ttf (quy ước tên font Windows)
    variants = {"": regular, "bd": regular.with_name(stem + "bd" + ext),
                "i": regular.with_name(stem + "i" + ext), "bi": regular.with_name(stem + "bi" + ext)}
    names = {}
    for suf, path in variants.items():
        src = path if path.exists() else regular
        names[suf] = f"VNFont{suf}"
        pdfmetrics.registerFont(TTFont(names[suf], str(src)))
    addMapping("VNFont", 0, 0, names[""])
    addMapping("VNFont", 1, 0, names["bd"])
    addMapping("VNFont", 0, 1, names["i"])
    addMapping("VNFont", 1, 1, names["bi"])
    _FONT_NAME = "VNFont"
    return _FONT_NAME


def render_pdf(parsed: dict[str, Any]) -> bytes:
    font = _register_font()
    html = _env.get_template("invoice_fallback.html").render(inv=parsed, font_name=font)
    if font:
        from xhtml2pdf.default import DEFAULT_FONT
        DEFAULT_FONT["vnfont"] = font
    out = io.BytesIO()
    status = pisa.CreatePDF(html, dest=out, encoding="utf-8")
    if status.err:
        raise RuntimeError("Không dựng được PDF từ XML")
    return out.getvalue()
