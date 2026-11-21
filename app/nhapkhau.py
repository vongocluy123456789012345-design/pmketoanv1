"""Tờ khai hải quan hàng hóa NHẬP KHẨU -> chỉ tiêu [23a]/[24a] (cộng vào [23]/[24]) của tờ khai 01/GTGT.

Đọc bản in PDF / ảnh chụp / Excel / XML (ECUS, VNACCS): số tờ khai, ngày đăng ký, người xuất khẩu (nhà cung cấp),
hàng hóa, trị giá tính thuế GTGT, thuế suất và số thuế GTGT hàng nhập khẩu. AI đọc chính xác nhất (mọi mẫu, ảnh chụp);
không có AI thì đọc theo mẫu chữ của bản in VNACCS (số tờ khai, ngày, tổng thuế GTGT) – người dùng sửa / bổ sung trên bảng.
"""
from __future__ import annotations

import base64
import io
import json
import re
from pathlib import Path
from typing import Any

FIELDS = ("so_tk", "ngay", "ncc", "nuoc", "hang", "tri_gia_tinh_thue", "thue_suat", "thue_gtgt", "ghi_chu", "file")

PROMPT = """Bạn đọc TỜ KHAI HẢI QUAN HÀNG HÓA NHẬP KHẨU của Việt Nam (VNACCS / ECUS: bản in PDF, ảnh chụp, Excel, XML).
Trả về DUY NHẤT 1 JSON:
{"to_khai":[{"so_tk":"số tờ khai (12 chữ số)","ngay":"dd/mm/yyyy – ngày đăng ký tờ khai","ncc":"tên người xuất khẩu (bên
bán ở nước ngoài)","nuoc":"nước xuất khẩu","hang":"tên hàng hóa tóm tắt (nhiều dòng hàng thì liệt kê ngắn, cách nhau ;)",
"tri_gia_tinh_thue":0,"thue_suat":"10%","thue_gtgt":0}]}
Quy tắc: tiền là VND, số thuần (không dấu phân cách nghìn). tri_gia_tinh_thue = TỔNG trị giá tính thuế của sắc thuế GTGT
(mã sắc thuế V) cộng mọi dòng hàng (= trị giá tính thuế nhập khẩu + thuế nhập khẩu + thuế TTĐB + thuế BVMT nếu có);
thue_gtgt = TỔNG số tiền thuế GTGT hàng nhập khẩu. Nhiều tờ khai thì mỗi tờ khai 1 phần tử. Nhiều mức thuế suất GTGT thì
thue_suat ghi các mức cách nhau ;. Không bịa – không thấy thì để rỗng / 0."""


def _num(v: Any) -> float:
    if isinstance(v, (int, float)):
        return float(v)
    s = re.sub(r"[^\d,.\-]", "", str(v or ""))
    if not s:
        return 0.0
    if s.count(".") > 1 or (s.count(".") == 1 and len(s.split(".")[1]) == 3 and "," not in s):
        s = s.replace(".", "")          # 12.345.678 (VN) -> 12345678
    s = s.replace(",", "" if s.count(",") > 1 or (s.count(",") == 1 and len(s.split(",")[1]) == 3) else ".")
    try:
        return float(s)
    except ValueError:
        return 0.0


def clean(row: dict[str, Any]) -> dict[str, Any]:
    out = {k: str(row.get(k) or "").strip() for k in ("so_tk", "ngay", "ncc", "nuoc", "hang", "thue_suat", "ghi_chu", "file")}
    out["tri_gia_tinh_thue"] = round(_num(row.get("tri_gia_tinh_thue")))
    out["thue_gtgt"] = round(_num(row.get("thue_gtgt")))
    if not out["thue_suat"] and out["tri_gia_tinh_thue"] and out["thue_gtgt"]:
        p = round(out["thue_gtgt"] / out["tri_gia_tinh_thue"] * 100)
        if p in (5, 8, 10):
            out["thue_suat"] = f"{p}%"
    if not out["tri_gia_tinh_thue"] and out["thue_gtgt"]:
        m = re.findall(r"(\d+(?:[.,]\d+)?)\s*%", out["thue_suat"])
        if len(m) == 1 and float(m[0].replace(",", ".")) > 0:
            out["tri_gia_tinh_thue"] = round(out["thue_gtgt"] * 100 / float(m[0].replace(",", ".")))
            out["ghi_chu"] = (out["ghi_chu"] + " Trị giá tính thuế suy ra từ thuế GTGT ÷ thuế suất – kiểm tra lại.").strip()
    return out


# ------------------------------------------------------------------ đọc không cần AI (bản in VNACCS)
def parse_text(text: str) -> list[dict[str, Any]]:
    """Đọc theo mẫu chữ bản in tờ khai VNACCS: số tờ khai, ngày đăng ký, người xuất khẩu, tổng thuế GTGT."""
    t = re.sub(r"[ \t]+", " ", text or "")
    so = re.search(r"Số tờ khai\D{0,40}?(\d{12})", t, re.I) or re.search(r"\b(1\d{11})\b", t)
    if not so:
        return []
    ngay = re.search(r"Ngày đăng ký\D{0,30}?(\d{1,2}/\d{1,2}/\d{4})", t, re.I)
    ncc = re.search(r"Người xuất khẩu.{0,80}?Tên\s*:?\s*([^\n]{3,120})", t, re.I | re.S)
    vat = 0.0
    for pat in (r"\bV\b\s*(?:Thuế\s*)?(?:GTGT|giá trị gia tăng)\D{0,20}([\d.,]{4,})",
                r"(?:Thuế\s*GTGT|Thuế giá trị gia tăng)[^\n]{0,60}?(?:Số tiền thuế|Tổng tiền thuế)\D{0,20}([\d.,]{4,})"):
        found = [_num(x) for x in re.findall(pat, t, re.I)]
        if found:
            vat = max(found)
            break
    ts = re.search(r"(?:GTGT|giá trị gia tăng)[^\n]{0,80}?(\d{1,2})\s*%", t, re.I)
    return [clean({"so_tk": so.group(1), "ngay": ngay.group(1) if ngay else "",
                   "ncc": ncc.group(1).strip() if ncc else "", "thue_gtgt": vat,
                   "thue_suat": f"{ts.group(1)}%" if ts else "",
                   "ghi_chu": "Đọc theo mẫu bản in (không AI) – kiểm tra / bổ sung trị giá và người xuất khẩu"})]


# ------------------------------------------------------------------ AI
def _ai(text: str, images: list[bytes]) -> list[dict[str, Any]]:
    from . import assistant
    from .kiemxuat import _ai_provider
    p, models = _ai_provider()
    last: Exception | None = None
    for m in models:
        try:
            if p == "claude":
                import anthropic
                content: list[dict[str, Any]] = [{"type": "image", "source": {"type": "base64", "media_type": "image/jpeg",
                                                                               "data": base64.b64encode(b).decode()}}
                                                 for b in images]
                content.append({"type": "text", "text": text[:30000] or "Đọc tờ khai trong ảnh."})
                r = anthropic.Anthropic(api_key=assistant.key_for(p)).messages.create(
                    model=m, max_tokens=3000, system=PROMPT, messages=[{"role": "user", "content": content}])
                out = "".join(b.text for b in r.content if b.type == "text")
            else:
                user: Any = text[:30000]
                if images:
                    user = [{"type": "text", "text": text[:30000] or "Đọc tờ khai trong ảnh."}] + \
                           [{"type": "image_url", "image_url": {"url": "data:image/jpeg;base64," + base64.b64encode(b).decode()}}
                            for b in images]
                out = assistant._post_plain(p, m, [{"role": "system", "content": PROMPT}, {"role": "user", "content": user}])
            js = re.search(r"\{.*\}", out or "", re.S)
            rows = [clean(x) for x in (json.loads(js.group(0)).get("to_khai") or [])] if js else []
            rows = [r for r in rows if r["so_tk"] or r["thue_gtgt"] or r["tri_gia_tinh_thue"]]
            if rows:
                return rows
        except Exception as e:  # noqa: BLE001 - model không đọc được / hết lượt -> model kế tiếp
            last = e
    if last:
        raise RuntimeError(f"AI không đọc được tờ khai: {last}")
    return []


def _text_of(name: str, data: bytes) -> tuple[str, list[bytes]]:
    """(chữ, ảnh) của 1 file tờ khai."""
    from .kiemxuat import _png
    from .ocr import page_images
    ext = Path(name).suffix.lower()
    if data[:5] == b"%PDF-":
        import pdfplumber
        with pdfplumber.open(io.BytesIO(data)) as pdf:
            txt = "\n".join((pg.extract_text() or "") for pg in pdf.pages[:12])
        if len(txt.strip()) > 80:
            return txt, []
        return "", [_png(im) for im in page_images(data, 4)]
    if ext in (".xlsx", ".xls", ".xlsm", ".csv", ".ods"):
        from .misa_import import load_any_workbook
        wb = load_any_workbook(data, values=True)
        return "\n".join(" | ".join(str(c) for c in r if c not in (None, ""))
                         for ws in wb.worksheets for r in ws.iter_rows(values_only=True)), []
    if ext == ".xml" or data.lstrip()[:1] == b"<":
        t = data.decode("utf-8", "ignore")
        t = re.sub(r"<([A-Za-z0-9_:]+)[^>]*>([^<]+)</\1>", r"\1: \2\n", t)   # <Tag>giá trị</Tag> -> "Tag: giá trị"
        return re.sub(r"<[^>]+>", " ", t), []
    return "", [_png(im) for im in page_images(data, 1)]


def read_files(files: list[tuple[str, bytes]], dung_ai: bool = True) -> tuple[list[dict[str, Any]], list[str]]:
    rows: list[dict[str, Any]] = []
    notes: list[str] = []
    for name, data in files:
        got: list[dict[str, Any]] = []
        how = ""
        try:
            text, images = _text_of(name, data)
            if dung_ai:
                try:
                    got, how = _ai(text, images), "AI đọc"
                except Exception as e:  # noqa: BLE001 - chưa cài khóa AI / hết lượt -> đọc theo mẫu
                    notes.append(f"{name}: {str(e)[:160]}")
            if not got and text:
                got, how = parse_text(text), "đọc theo mẫu bản in"
        except Exception as e:  # noqa: BLE001
            notes.append(f"{name}: không đọc được ({str(e)[:160]})")
        for g in got:
            g["file"] = name
        rows += got
        notes.append(f"{name}: {len(got)} tờ khai ({how or 'không đọc được – nhập tay'})")
    return rows, notes


def totals(rows: list[dict[str, Any]]) -> tuple[int, int]:
    return sum(int(r.get("tri_gia_tinh_thue") or 0) for r in rows), sum(int(r.get("thue_gtgt") or 0) for r in rows)
