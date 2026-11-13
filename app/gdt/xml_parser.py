"""Đọc XML hóa đơn điện tử theo định dạng TT78 / QĐ 1450 (HDon > DLHDon > ...)."""
from __future__ import annotations

import re
import unicodedata
import xml.etree.ElementTree as ET
from typing import Any

URL_RE = re.compile(r"https?://[^\s\"'<>]+", re.I)
SKIP_URL_HOSTS = ("w3.org", "xmlsoap.org", "schemas.", "uri.etsi.org", "gdt.gov.vn/TTHDDT")
SKIP_SUBTREES = {"DSCKS", "Signature", "CKS"}

# Tên trường (đã bỏ dấu, viết thường, bỏ khoảng trắng) thường chứa mã tra cứu.
LOOKUP_CODE_KEYS = (
    "matracuu", "matra cuu", "macuutra", "mtc", "tracuu", "fkey", "searchkey", "keysearch",
    "lookupcode", "transactionid", "sobaomat", "mabaomat", "secretcode", "invoicekey",
)
LOOKUP_URL_KEYS = ("link", "url", "website", "trangtracuu", "portal", "tracuutai", "diachitracuu")


def norm(s: str) -> str:
    s = unicodedata.normalize("NFD", s or "")
    s = "".join(c for c in s if unicodedata.category(c) != "Mn")
    return s.replace("đ", "d").replace("Đ", "D").lower().replace(" ", "").replace("_", "")


def _local(tag: str) -> str:
    return tag.split("}", 1)[-1]


def _child(el: ET.Element | None, name: str) -> ET.Element | None:
    if el is None:
        return None
    for c in el:
        if _local(c.tag) == name:
            return c
    return None


def _text(el: ET.Element | None, *path: str) -> str | None:
    cur = el
    for p in path:
        cur = _child(cur, p)
        if cur is None:
            return None
    t = (cur.text or "").strip()
    return t or None


def _num(v: str | None) -> float | None:
    if v in (None, ""):
        return None
    try:
        return float(v.replace(",", ""))
    except ValueError:
        return None


def _iter_texts(el: ET.Element):
    if _local(el.tag) in SKIP_SUBTREES:
        return
    if el.text and el.text.strip():
        yield el.text.strip()
    for c in el:
        yield from _iter_texts(c)


def _party(el: ET.Element | None) -> dict[str, Any]:
    return {
        "ten": _text(el, "Ten"),
        "mst": _text(el, "MST"),
        "dia_chi": _text(el, "DChi"),
        "sdt": _text(el, "SDThoai"),
        "email": _text(el, "DCTDTu"),
        "stk": _text(el, "STKNHang"),
        "ngan_hang": _text(el, "TNHang"),
        "ho_ten_nguoi_mua": _text(el, "HVTNMHang"),
    }


def parse_invoice_xml(data: bytes) -> dict[str, Any]:
    root = ET.fromstring(data)
    dl = next((e for e in root.iter() if _local(e.tag) == "DLHDon"), root)
    tt = _child(dl, "TTChung")
    nd = _child(dl, "NDHDon")
    toan = _child(nd, "TToan")

    items = []
    ds = _child(nd, "DSHHDVu")
    if ds is not None:
        for h in ds:
            if _local(h.tag) != "HHDVu":
                continue
            items.append({
                "stt": _text(h, "STT"),
                "tinh_chat": _text(h, "TChat"),
                "ma": _text(h, "MHHDVu"),
                "ten": _text(h, "THHDVu"),
                "dvt": _text(h, "DVTinh"),
                "so_luong": _num(_text(h, "SLuong")),
                "don_gia": _num(_text(h, "DGia")),
                "thanh_tien": _num(_text(h, "ThTien")),
                "thue_suat": _text(h, "TSuat"),
                "chiet_khau": _num(_text(h, "STCKhau")),
            })

    # Toàn bộ cặp "Thông tin khác" (TTKhac/TTin) ở mọi cấp.
    extras: dict[str, str] = {}
    for e in root.iter():
        if _local(e.tag) == "TTin":
            k, v = _text(e, "TTruong"), _text(e, "DLieu")
            if k and v:
                extras[k] = v

    mccqt = next((((e.text or "").strip()) for e in root.iter() if _local(e.tag) == "MCCQT"), None)

    # Tổng hợp theo thuế suất (TToan > THTTLTSuat > LTSuat: TSuat, ThTien, TThue)
    rates: dict[str, list[float]] = {}
    lts = _child(toan, "THTTLTSuat")
    if lts is not None:
        for e in lts:
            if _local(e.tag) == "LTSuat":
                r = (_text(e, "TSuat") or "").upper()
                a = rates.setdefault(r, [0.0, 0.0])
                a[0] += _num(_text(e, "ThTien")) or 0
                a[1] += _num(_text(e, "TThue")) or 0
    if not rates:  # XML không có bảng tổng hợp -> cộng từ dòng hàng
        for it in items:
            r = (it.get("thue_suat") or "").upper()
            a = rates.setdefault(r, [0.0, 0.0])
            a[0] += it.get("thanh_tien") or 0
            pct = _num(r.rstrip("%")) if r.endswith("%") else None
            a[1] += round((it.get("thanh_tien") or 0) * pct / 100) if pct is not None else 0

    lookup_code = None
    lookup_urls: list[str] = []
    for k, v in extras.items():
        nk = norm(k)
        if URL_RE.match(v) or any(x in nk for x in LOOKUP_URL_KEYS):
            lookup_urls += URL_RE.findall(v)
        elif lookup_code is None and any(x.replace(" ", "") in nk for x in LOOKUP_CODE_KEYS):
            lookup_code = v
    for t in _iter_texts(root):
        for u in URL_RE.findall(t):
            if not any(h in u for h in SKIP_URL_HOSTS):
                lookup_urls.append(u.rstrip(".,;)"))
    lookup_urls = list(dict.fromkeys(lookup_urls))

    return {
        "khmshdon": _text(tt, "KHMSHDon"),
        "khhdon": _text(tt, "KHHDon"),
        "shdon": _text(tt, "SHDon"),
        "ngay_lap": _text(tt, "NLap"),
        "ten_hd": _text(tt, "THDon"),
        "dvtte": _text(tt, "DVTTe"),
        "tgia": _num(_text(tt, "TGia")),
        "htttoan": _text(tt, "HTTToan"),
        "msttcgp": _text(tt, "MSTTCGP"),
        "mccqt": mccqt,
        "nguoi_ban": _party(_child(nd, "NBan")),
        "nguoi_mua": _party(_child(nd, "NMua")),
        "hang_hoa": items,
        "tong_chua_thue": _num(_text(toan, "TgTCThue")),
        "tong_thue": _num(_text(toan, "TgTThue")),
        "tong_chiet_khau": _num(_text(toan, "TTCKTMai")),
        "tong_thanh_toan": _num(_text(toan, "TgTTTBSo")),
        "bang_chu": _text(toan, "TgTTTBChu"),
        "theo_thue_suat": rates,
        "thong_tin_khac": extras,
        "lookup_code": lookup_code,
        "lookup_urls": lookup_urls,
    }
