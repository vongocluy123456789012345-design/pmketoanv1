"""Kiểm tra đầu vào trước khi xuất hóa đơn bán ra.

Khách gửi yêu cầu xuất hóa đơn (ảnh chụp, PDF, Excel, tin nhắn Zalo...) -> đọc ra từng mặt hàng (tên, ĐVT, SL, đơn giá)
-> dò với file tồn kho chi tiết (nếu có) và với hàng mua vào / đã bán ra trên hóa đơn điện tử đã tải của công ty
-> từng mặt hàng: Đủ / Thiếu / Gần giống cần xác nhận / Chỉ có vật tư liên quan / Không có đầu vào, kèm giá vốn,
   hóa đơn mua vào liên quan, file Excel và lời nhắn trả lời khách.
"""
from __future__ import annotations

import base64
import io
import json
import re
from datetime import date, datetime
from pathlib import Path
from typing import Any

from .config import settings
from .docreader import fold, vn_num
from .inventory import _family, _kw, _kw_match, _norm, _score

_UNITS = (r"cái|cai|chiếc|chiec|bộ|bo|thùng|thung|hộp|hop|kg|kgs|gam|g|m2|m²|m3|md|mét|met|m|cuộn|cuon|tấm|tam|gói|goi|"
          r"chai|lon|đôi|doi|cây|cay|bao|lít|lit|tấn|tan|hũ|hu|túi|tui|vỉ|vi|viên|vien|ram|quyển|quyen|cặp|cap|lốc|loc|"
          r"két|ket|bình|binh|can|ống|ong|sợi|soi|khối|khoi|công|cong|chuyến|chuyen|lần|lan|gói dịch vụ|trọn gói|tron goi")
_HEAD = {"ten": ["ten hang", "ten hang hoa", "hang muc", "ten san pham", "san pham", "mat hang", "ten vat tu", "hang hoa",
                 "noi dung", "dien giai", "description", "ten hhdv", "cong viec", "ten"],
         "mo_ta": ["mo ta", "quy cach", "chi tiet", "thong so", "chat lieu"],
         "dvt": ["dvt", "don vi tinh", "don vi", "dv", "unit"],
         "sl": ["so luong", "sl", "khoi luong", "kl", "quantity", "qty"],
         "tong": ["tong"],
         "dg": ["don gia", "gia ban", "gia", "unit price", "price"],
         "tt": ["thanh tien", "thanh toan", "amount", "tien"],
         "ma": ["ma hang", "ma sp", "ma vat tu", "ma", "code", "sku"]}
_AREA = {"m2", "m²", "md", "m3", "m"}
_STOP = set("cai chiec bo loai hang hoa san pham cac va cho theo mau size voi cua tren duoi trong ngoai chong mot hai ba "
            "so luong don gia thanh tien dich vu thi cong lap dat cung cap ban giao phi tien".split())


# ------------------------------------------------------------------ đọc yêu cầu của khách
def _head_map(row: list[Any]) -> dict[int, str]:
    m: dict[int, str] = {}
    for j, cell in enumerate(row):
        t = _norm(cell)
        if not t:
            continue
        for k, pats in _HEAD.items():
            if k in m.values():
                continue
            if any(t == p or t.startswith(p + " ") or (len(p) > 3 and t.startswith(p)) for p in pats):
                m[j] = k
                break
    return m


def _rows_to_items(rows: list[list[Any]]) -> list[dict[str, Any]]:
    """Bảng có dòng tiêu đề (Tên hàng / Hạng mục, ĐVT, SL, Đơn giá, Thành tiền...), tiêu đề có thể 2 dòng."""
    for i, row in enumerate(rows[:25]):
        m = _head_map(row)
        if "ten" not in m.values() or not ({"sl", "tt", "dg"} & set(m.values())):
            continue
        if i + 1 < len(rows):          # tiêu đề 2 dòng: "Kích thước: Dài | Rộng | Cao | Tổng | SL"
            for j, k in _head_map(rows[i + 1]).items():
                if k in ("sl", "tong", "dvt", "dg", "tt") and k not in m.values():
                    m[j] = k
        col = {k: j for j, k in m.items()}
        out = []
        for r in rows[i + 1:]:
            g = lambda k: (r[col[k]] if k in col and col[k] < len(r) else None)  # noqa: E731
            ten = re.sub(r"\s+", " ", str(g("ten") or "")).strip()
            if not ten or _head_map(r).get(col["ten"]) == "ten":
                continue
            f = fold(ten).lower()
            if re.match(r"(tong cong|cong tien|tong tien|thue gtgt|thue suat|tong thanh toan|cong)\b", f):
                break
            if re.fullmatch(r"[\d\s()=x+\-*/.,]*", ten):
                continue
            sl, tong, dg, tt = _n(g("sl")), _n(g("tong")), _n(g("dg")), _n(g("tt"))
            dvt = str(g("dvt") or "").strip()
            if tong and _norm(dvt) in _AREA:
                sl = round(tong * (sl or 1), 3)
            if dg and tt and (not sl or abs(sl * dg - tt) > max(1000, tt * 0.02)):
                sl = round(tt / dg, 3)      # SL đọc sai dấu thập phân ("43.516" m2) -> theo Thành tiền / Đơn giá
            if not (sl or tt):
                continue
            out.append({"ten": ten, "mo_ta": str(g("mo_ta") or "").strip(), "dvt": dvt, "sl": sl or 0.0,
                        "dg": dg or (round(tt / sl) if tt and sl else 0.0), "tt": tt or round((sl or 0) * (dg or 0)),
                        "ma": str(g("ma") or "").strip()})
        if out:
            return out
    return []


def _n(v: Any) -> float:
    if v is None or v == "":
        return 0.0
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip()
    m = re.search(r"=\s*([\d.,]+)\s*$", s)           # "(44.086 - 0.57) = 43.516"
    if m:
        s = m.group(1)
    return float(vn_num(s) or 0)


def items_from_text(text: str) -> list[dict[str, Any]]:
    """Tin nhắn kiểu "10 thùng sữa Ensure 850g, bàn làm việc x 2 cái, Ghế xoay: 5" -> mặt hàng (không cần AI)."""
    out = []
    parts = [p.strip(" -•*+\t") for p in re.split(r"[\n;]+|,(?!\d)", text or "") if p.strip()]
    u = rf"(?:{_UNITS})"
    for p in parts:
        p = re.sub(r"^\d{1,2}[.)]\s+", "", p)                       # "1. ", "2) "
        p = re.sub(r"^[^\d:]{3,60}:\s*(?=\d[\d.,]*\s+\D)", "", p)                # "Em xuất giúp: 200 đôi gót PC"
        m = re.match(rf"^(?P<sl>\d[\d.,]*)\s*(?P<dvt>{u})?\.?\s+(?P<ten>.{{3,}})$", p, re.I)
        if not m:
            m = re.match(rf"^(?P<ten>.{{3,}}?)\s*(?:[:x×=]|sl|so luong|số lượng)?\s*(?P<sl>\d[\d.,]*)\s*(?P<dvt>{u})?\.?$",
                         p, re.I)
        if not m:
            continue
        ten = m.group("ten").strip(" :-x×")
        if len(_norm(ten)) < 3 or re.search(r"(?i)chị|anh|em|ạ|nhé|giúp|hộ|xuất hóa đơn|xuất hoá đơn", ten) and \
                len(ten.split()) > 8:
            continue
        dg = 0.0
        g = re.search(r"(?i)(?:giá|gia|@|đơn giá)\s*[:=]?\s*([\d.,]+)\s*(k|nghìn|tr|triệu)?", ten)
        if g:
            dg = float(vn_num(g.group(1)) or 0) * {"k": 1e3, "nghìn": 1e3, "tr": 1e6, "triệu": 1e6}.get((g.group(2) or "").lower(), 1)
            ten = ten[:g.start()].strip(" ,-")
        sl = float(vn_num(m.group("sl")) or 0)
        if sl <= 0 or sl > 1e7:
            continue
        out.append({"ten": ten, "mo_ta": "", "dvt": (m.group("dvt") or "").strip(), "sl": sl, "dg": dg, "tt": round(sl * dg),
                    "ma": ""})
    return out


AI_PROMPT = """Bạn đọc yêu cầu xuất hóa đơn / báo giá / hợp đồng / đơn hàng / tin nhắn của khách hàng Việt Nam và liệt kê TỪNG
mặt hàng / hạng mục cần xuất hóa đơn. Trả về DUY NHẤT 1 JSON:
{"items":[{"ten":"tên hàng / hạng mục","mo_ta":"mô tả, chất liệu, quy cách","dvt":"","sl":0,"dg":0,"tt":0}],
 "tong_tien_hang":0,"thue_suat":"","ghi_chu":""}
Quy tắc: số là số thuần (không dấu phân cách nghìn, dấu thập phân là "."); sl = số lượng theo ĐVT – nếu ĐVT là m2/md/m3
và có cột kích thước "Tổng" thì sl = Tổng × số lượng; tt = thành tiền chưa thuế; bỏ dòng tổng cộng, thuế, tổng thanh toán;
giữ đúng tên tiếng Việt có dấu; không bịa – thiếu thì để rỗng / 0."""


def _ai_provider() -> tuple[str, list[str]]:
    from . import assistant
    if assistant.key_for("gemini"):
        return "gemini", ["gemini-flash-latest", "gemini-flash-lite-latest", "gemini-3.5-flash-lite", "gemini-3.1-flash-lite"]
    p = assistant.provider()
    if not assistant.key_for(p) and p != "ollama":
        raise RuntimeError("Chưa cài khóa AI (Trợ lý AI → Cài đặt) – ảnh / PDF scan cần AI để đọc chính xác")
    return p, [assistant.model_for(p)]


def _png(img) -> bytes:
    img = img.convert("RGB")
    if max(img.size) > 2200:
        img.thumbnail((2200, 2200))
    b = io.BytesIO()
    img.save(b, "JPEG", quality=88)
    return b.getvalue()


def _parse_json(out: str) -> list[dict[str, Any]]:
    js = re.search(r"\{.*\}", out or "", re.S)
    if not js:
        return []
    d = json.loads(js.group(0))
    items = []
    for it in d.get("items") or []:
        ten = str(it.get("ten") or "").strip()
        if not ten:
            continue
        sl, dg, tt = _n(it.get("sl")), _n(it.get("dg")), _n(it.get("tt"))
        items.append({"ten": ten, "mo_ta": str(it.get("mo_ta") or "").strip(), "dvt": str(it.get("dvt") or "").strip(),
                      "sl": sl or (round(tt / dg, 3) if dg and tt else 0.0), "dg": dg or (round(tt / sl) if sl and tt else 0.0),
                      "tt": tt or round(sl * dg), "ma": ""})
    return items


def ai_items(text: str = "", images: list[bytes] | None = None) -> list[dict[str, Any]]:
    """AI đọc yêu cầu: ảnh gửi thẳng cho model đọc ảnh (chính xác hơn OCR với chữ tiếng Việt), chữ thì gửi chữ."""
    from . import assistant
    p, models = _ai_provider()
    images = images or []
    last: Exception | None = None
    for m in models:
        try:
            if p == "claude":
                import anthropic
                content: list[dict[str, Any]] = [{"type": "image", "source": {"type": "base64", "media_type": "image/jpeg",
                                                                               "data": base64.b64encode(b).decode()}}
                                                 for b in images]
                content.append({"type": "text", "text": text[:14000] or "Đọc các mặt hàng trong ảnh."})
                r = anthropic.Anthropic(api_key=assistant.key_for(p)).messages.create(
                    model=m, max_tokens=4000, system=AI_PROMPT, messages=[{"role": "user", "content": content}])
                out = "".join(b.text for b in r.content if b.type == "text")
            else:
                user: Any = text[:14000]
                if images:
                    user = [{"type": "text", "text": text[:14000] or "Đọc các mặt hàng trong ảnh."}] + \
                           [{"type": "image_url", "image_url": {"url": "data:image/jpeg;base64," + base64.b64encode(b).decode()}}
                            for b in images]
                out = assistant._post_plain(p, m, [{"role": "system", "content": AI_PROMPT}, {"role": "user", "content": user}])
            items = _parse_json(out)
            if items:
                return items
        except Exception as e:  # noqa: BLE001 - model không đọc ảnh / hết lượt -> model kế tiếp
            last = e
    if last:
        raise RuntimeError(f"AI không đọc được yêu cầu: {last}")
    return []


def read_request(files: list[tuple[str, bytes]], text: str = "", dung_ai: bool = True) -> tuple[list[dict[str, Any]], list[str]]:
    """-> (mặt hàng, ghi chú cách đọc). Excel / CSV / PDF có chữ: đọc bảng trực tiếp; ảnh / PDF scan / tin nhắn khó: AI."""
    from .misa_import import load_any_workbook
    from .ocr import page_images
    items: list[dict[str, Any]] = []
    notes: list[str] = []
    text_used = False
    for name, data in files:
        ext = Path(name).suffix.lower()
        got: list[dict[str, Any]] = []
        how = ""
        if data[:5] == b"%PDF-":
            try:
                import pdfplumber
                with pdfplumber.open(io.BytesIO(data)) as pdf:
                    rows = [r for pg in pdf.pages[:6] for tb in (pg.extract_tables() or []) for r in tb]
                    txt = "\n".join((pg.extract_text() or "") for pg in pdf.pages[:6])
                got, how = _rows_to_items(rows), "bảng trong PDF"
                if not got and len(txt.strip()) > 40:
                    got, how = (ai_items(txt) if dung_ai else items_from_text(txt)), ("AI đọc chữ PDF" if dung_ai else "chữ PDF")
            except Exception as e:  # noqa: BLE001
                notes.append(f"{name}: {e}")
            if not got and dung_ai:
                got, how = ai_items("", [_png(im) for im in page_images(data, 3)]), "AI đọc ảnh PDF scan"
        elif ext in (".xlsx", ".xls", ".xlsm", ".csv", ".ods", ".xltx"):
            wb = load_any_workbook(data, values=True)
            for ws in wb.worksheets:
                got = _rows_to_items([list(r) for r in ws.iter_rows(values_only=True)])
                if got:
                    how = f"bảng Excel (sheet {ws.title})"
                    break
            if not got and dung_ai:
                txt = "\n".join(" | ".join(str(c) for c in r if c not in (None, ""))
                                for ws in wb.worksheets for r in ws.iter_rows(values_only=True))
                got, how = ai_items(txt), "AI đọc file Excel"
        elif ext in (".txt",):
            t = data.decode("utf-8", "ignore")
            got, how = (items_from_text(t), "tin nhắn") if not dung_ai else (ai_items(t), "AI đọc tin nhắn")
        else:                                              # ảnh chụp màn hình / ảnh chụp
            if dung_ai:
                got, how = ai_items(text, [_png(im) for im in page_images(data, 1)]), "AI đọc ảnh (kèm tin nhắn)"
                text_used = bool(text.strip())
            else:
                from .misa_import import items_from_ocr_text
                from .ocr import ocr_image
                t = ocr_image(data)
                got = [{**i, "mo_ta": "", "ma": ""} for i in items_from_ocr_text(t)] or items_from_text(t)
                how = "OCR (không AI – chữ tiếng Việt có thể sai dấu)"
        for g in got:
            g["nguon"] = name
        items += got
        notes.append(f"{name}: {len(got)} mặt hàng ({how or 'không đọc được'})")
    if text.strip() and not text_used:
        got = items_from_text(text)
        how = "tin nhắn"
        if dung_ai and (not got or len(got) < len([x for x in re.split(r"[\n;]+", text) if re.search(r"\d", x)]) // 2):
            try:
                got, how = ai_items(text), "AI đọc tin nhắn"
            except RuntimeError as e:
                notes.append(str(e))
        for g in got:
            g["nguon"] = "tin nhắn"
        items += got
        notes.append(f"Tin nhắn: {len(got)} mặt hàng ({how})")
    return items, notes


# ------------------------------------------------------------------ nguồn đối chiếu: tồn kho + hóa đơn
def kho_dir(mst: str) -> Path:
    return Path(settings.output_dir) / mst / "_HoSo" / "TonKho"


def save_stock(mst: str, name: str, data: bytes) -> Path:
    d = kho_dir(mst)
    d.mkdir(parents=True, exist_ok=True)
    f = d / f"{datetime.now():%Y%m%d_%H%M%S}_{re.sub(r'[^A-Za-z0-9._-]+', '_', name)[-60:]}"
    f.write_bytes(data)
    return f


def last_stock(mst: str) -> Path | None:
    d = kho_dir(mst)
    fs = sorted(d.glob("*")) if d.is_dir() else []
    return fs[-1] if fs else None


def read_stock(data: bytes) -> list[dict[str, Any]]:
    from .inventory import _read_generic, read_opening
    from .misa_import import load_any_workbook
    d = read_opening(data)
    items = d["items"]
    if d["loai"] == "danhmuc":          # bảng có Mã hàng + Tên hàng + Số lượng (không phải danh mục trống)
        wb = load_any_workbook(data, values=True)
        for ws in wb.worksheets:
            g = _read_generic(ws) or []
            if not any(i["sl"] for i in g):
                g = _simple_stock([list(r) for r in ws.iter_rows(values_only=True)])
            if g and any(i["sl"] for i in g):
                items = g
                break
    return [{**i, "ton": float(i.get("sl") or 0)} for i in items]


def _simple_stock(rows: list[list[Any]]) -> list[dict[str, Any]]:
    """Bảng tồn kho 1 dòng tiêu đề: Mã hàng | Tên hàng | ĐVT | Số lượng (tồn) | Giá trị."""
    for i, row in enumerate(rows[:25]):
        heads = {j: _norm(c) for j, c in enumerate(row) if c is not None}
        f = lambda *ps: next((j for j, h in heads.items() if any(h == p or h.startswith(p) for p in ps)), None)  # noqa: E731
        c_ten = f("ten hang", "ten vat tu", "hang hoa", "ten san pham", "ten")
        c_sl = f("so luong ton", "sl ton", "ton cuoi", "ton kho", "so luong", "sl")
        if c_ten is None or c_sl is None:
            continue
        c_ma, c_dvt, c_gt = f("ma hang", "ma vat tu", "ma"), f("dvt", "don vi tinh"), f("gia tri", "thanh tien")
        out = []
        for r in rows[i + 1:]:
            g = lambda c: r[c] if c is not None and c < len(r) else None  # noqa: E731
            ten = str(g(c_ten) or "").strip()
            if not ten or _norm(ten).startswith(("tong cong", "cong")):
                continue
            out.append({"kho": "", "ma_kho": "", "ma": str(g(c_ma) or "").strip(), "ten": ten,
                        "dvt": str(g(c_dvt) or "").strip(), "sl": _n(g(c_sl)), "gt": _n(g(c_gt))})
        return out
    return []


def invoice_items(mst: str, d1: date, d2: date) -> tuple[list[dict[str, Any]], dict[str, int]]:
    """Gộp từng mặt hàng trên HĐ mua vào / bán ra đã tải (đọc XML) theo tên chuẩn hóa."""
    from .misa_import import from_xml
    from .reports import tthai_code
    from .tienich import invoices
    agg: dict[tuple[str, str], dict[str, Any]] = {}
    st = {"mua": 0, "ban": 0, "khong_xml": 0}
    for r in invoices(mst, d1, d2):
        if tthai_code(r) in (4, 6):
            continue
        buy = r.get("huong") == "purchase"
        st["mua" if buy else "ban"] += 1
        f = r.get("file_xml")
        if not (f and Path(str(f)).exists()):
            st["khong_xml"] += 1
            continue
        try:
            lines = from_xml(Path(str(f)).read_bytes()).get("items") or []
        except Exception:  # noqa: BLE001
            st["khong_xml"] += 1
            continue
        for it in lines:
            ten = str(it.get("ten") or "").strip()
            k = _norm(ten)
            if not k or float(it.get("tt") or 0) < 0:
                continue
            a = agg.setdefault((k, _norm(it.get("dvt"))), {"ten": ten, "dvt": str(it.get("dvt") or ""), "nhap_sl": 0.0,
                                                           "nhap_tt": 0.0, "xuat_sl": 0.0, "xuat_tt": 0.0, "hd_mua": []})
            sl, tt = float(it.get("sl") or 0), float(it.get("tt") or 0)
            if buy:
                a["nhap_sl"] += sl
                a["nhap_tt"] += tt
                if len(a["hd_mua"]) < 6:
                    a["hd_mua"].append(f"{r.get('khhdon') or ''}-{str(r.get('shdon') or '').lstrip('0')} "
                                       f"{str(r.get('ngay_lap'))[:10]} {r.get('nbten') or ''}".strip())
            else:
                a["xuat_sl"] += sl
                a["xuat_tt"] += tt
    return list(agg.values()), st


# ------------------------------------------------------------------ dò từng mặt hàng
def _words(s: str) -> set[str]:
    return {w for w in _norm(s).split() if len(w) >= 3 and not w.isdigit() and w not in _STOP}


def _best(name: str, dvt: str, cands: list[dict[str, Any]], kws: list[frozenset[str]]) -> tuple[dict | None, float, str]:
    if not cands:
        return None, 0.0, ""
    k = _norm(name)
    sc, best = max(((_score(k, _norm(c["ten"])), c) for c in cands), key=lambda x: x[0])
    if sc < 0.999:
        toks = set(k.split())
        if len(toks) >= 2 and sum(1 for c in cands if toks <= set(_norm(c["ten"]).split())) >= 2:
            return _group(k, cands)
    if sc >= 0.86:
        return best, sc, "khop"
    it, s2, _ = _kw_match(name, cands, kws, dvt)
    if it is not None and s2 >= 0.9:
        return it, s2, "khop"
    fam = [c for c in cands if _family(name, c["ten"])]
    if fam:
        c = max(fam, key=lambda c: _score(k, _norm(c["ten"])))
        return c, max(0.7, _score(k, _norm(c["ten"]))), "gan"
    if it is not None:
        return it, s2, "gan"
    if sc >= 0.62:
        return best, sc, "gan"
    # tên khách ghi ngắn, nằm gọn trong tên đầy đủ trên HĐ / kho ("Vải giả da PU" ~ "Vải giả da được tráng phủ ... PU
    # BONG NANO, khổ 1m4"): gộp mọi biến thể thành 1 nhóm
    return _group(k, cands)


def _group(k: str, cands: list[dict[str, Any]]) -> tuple[dict | None, float, str]:
    toks = set(k.split())
    grp = [c for c in cands if len(toks) >= 2 and toks <= set(_norm(c["ten"]).split())]
    if len(grp) == 1:
        return grp[0], 0.75, "gan"
    if grp:
        dv: dict[str, int] = {}
        for c in grp:
            dv[c.get("dvt") or ""] = dv.get(c.get("dvt") or "", 0) + 1
        g = {"ten": f"Nhóm {len(grp)} mặt hàng: " + " | ".join(c["ten"][:60] for c in grp[:4]),
             "dvt": max(dv, key=dv.get), "ma": "", "_group": True}
        for f in ("ton", "gt", "sl", "nhap_sl", "nhap_tt", "xuat_sl", "xuat_tt"):
            if any(f in c for c in grp):
                g[f] = sum(float(c.get(f) or 0) for c in grp)
        g["hd_mua"] = [h for c in grp for h in c.get("hd_mua", [])][:6]
        return g, 0.72, "gan"
    return None, 0.0, ""


def _dvt_ok(a: str, b: str) -> bool:
    a, b = _norm(a), _norm(b)
    same = [{"cai", "chiec"}, {"m2", "m 2"}, {"md", "m", "met"}, {"bo", "set"}, {"kg", "kgs"}]
    return not a or not b or a == b or any(a in g and b in g for g in same)


def check(mst: str, req: list[dict[str, Any]], stock: list[dict[str, Any]] | None, d1: date, d2: date) -> dict[str, Any]:
    inv, st = invoice_items(mst, d1, d2)
    bought = [x for x in inv if x["nhap_sl"] > 0 or x["nhap_tt"] > 0]
    s_kw = [_kw(x["ten"]) for x in stock or []]
    b_kw = [_kw(x["ten"]) for x in bought]
    left_stock = {id(x): x["ton"] for x in stock or []}
    left_inv: dict[int, float] = {}
    rows = []
    for q in req:
        sl = float(q.get("sl") or 0)
        r: dict[str, Any] = {**q, "kho_ten": "", "kho_ton": None, "hd_ten": "", "hd_nhap": None, "hd_ban": None,
                             "hd_con": None, "gia_von": None, "lien_quan": [], "ket_qua": "", "muc": "", "ghi_chu": []}
        sm, ssc, skind = _best(q["ten"], q.get("dvt", ""), stock or [], s_kw)
        hm, hsc, hkind = _best(q["ten"], q.get("dvt", ""), bought, b_kw)
        avail, basis, kind = None, "", ""
        if sm is not None:
            left_stock.setdefault(id(sm), sm.get("ton", 0.0))
            r["kho_ten"], r["kho_ton"] = f"{sm.get('ma') or ''} {sm['ten']}".strip(), left_stock[id(sm)]
            avail, basis, kind = left_stock[id(sm)], "tồn kho", skind
            if sm.get("ton") and sm.get("gt"):
                r["gia_von"] = sm["gt"] / sm["ton"]
            if not _dvt_ok(q.get("dvt", ""), sm.get("dvt", "")):
                r["ghi_chu"].append(f"ĐVT khác: yêu cầu '{q.get('dvt')}', kho '{sm.get('dvt')}' – kiểm tra quy đổi")
                kind = "gan"
        if hm is not None:
            sold = float(hm.get("xuat_sl") or 0)
            con = left_inv.get(id(hm), hm["nhap_sl"] - sold)
            r.update({"hd_ten": hm["ten"], "hd_nhap": hm["nhap_sl"], "hd_ban": sold, "hd_con": con,
                      "hd_mua_so": hm["hd_mua"]})
            if r["gia_von"] is None and hm["nhap_sl"]:
                r["gia_von"] = hm["nhap_tt"] / hm["nhap_sl"]
            if avail is None:
                avail, basis, kind = con, "HĐ mua vào − đã bán ra", hkind
                if not _dvt_ok(q.get("dvt", ""), hm.get("dvt", "")):
                    r["ghi_chu"].append(f"ĐVT khác: yêu cầu '{q.get('dvt')}', HĐ mua '{hm.get('dvt')}' – kiểm tra quy đổi")
                    kind = "gan"
        if basis == "tồn kho" and r["hd_con"] is not None and abs(r["hd_con"] - (r["kho_ton"] or 0)) > max(1, sl * 0.05):
            r["ghi_chu"].append(f"Theo HĐ mua vào − bán ra còn {r['hd_con']:g} (tồn kho ghi {r['kho_ton']:g}) – đối chiếu lại")
        if avail is None:
            # hàng dịch vụ / gia công / thi công: tìm vật tư liên quan theo từ khóa của tên + mô tả
            w = _words(f"{q['ten']} {q.get('mo_ta', '')}")
            rel = []
            for x in bought:
                common = w & _words(x["ten"])
                if len(common) >= 2 or any(len(c) >= 6 for c in common):
                    rel.append((len(common), x))
            rel.sort(key=lambda t: (-t[0], -t[1]["nhap_tt"]))
            r["lien_quan"] = [f"{x['ten']} ({x['nhap_sl']:g} {x['dvt']}, {x['nhap_tt']:,.0f} đ)".replace(",", ".")
                              for _, x in rel[:5]]
            if rel:
                r["muc"], r["ket_qua"] = "LIENQUAN", "Chỉ có vật tư liên quan (hàng gia công / thi công – kiểm tra định mức)"
            else:
                r["muc"], r["ket_qua"] = "KHONG", "Không có đầu vào"
        else:
            need = sl
            if sm is not None and basis == "tồn kho":
                left_stock[id(sm)] = max(0.0, left_stock[id(sm)] - need)
            elif hm is not None:
                left_inv[id(hm)] = (left_inv.get(id(hm), avail)) - need
            if avail >= need - 1e-9 and need > 0:
                r["muc"], r["ket_qua"] = "DU", f"Đủ ({basis}: còn {avail:g})"
            elif avail <= 0:
                r["muc"], r["ket_qua"] = "THIEU", f"Hết hàng ({basis}: còn {avail:g})"
            else:
                r["muc"], r["ket_qua"] = "THIEU", f"Thiếu {need - avail:g} ({basis}: còn {avail:g}, cần {need:g})"
            if kind == "gan":
                r["ket_qua"] = "Gần giống – cần xác nhận: " + r["ket_qua"]
                r["muc"] = "GAN" if r["muc"] == "DU" else r["muc"]
        if r["gia_von"] and q.get("dg") and q["dg"] < r["gia_von"] * 0.98:
            r["ghi_chu"].append(f"Giá bán {q['dg']:,.0f} thấp hơn giá vốn {r['gia_von']:,.0f}".replace(",", "."))
        r["ghi_chu"] = "; ".join(r["ghi_chu"])
        rows.append(r)
    tong = {k: sum(1 for r in rows if r["muc"] == k) for k in ("DU", "GAN", "THIEU", "LIENQUAN", "KHONG")}
    return {"mst": mst, "tu": d1.isoformat(), "den": d2.isoformat(), "rows": rows, "tong": tong, "hoa_don": st,
            "co_ton_kho": bool(stock), "so_hang_ton": len(stock or []), "so_mat_hang_mua": len(bought)}


def reply_text(res: dict[str, Any], ten_kh: str = "") -> str:
    rows = res["rows"]
    t = res["tong"]
    ok = t["DU"]
    lines = [f"Dạ {ten_kh or 'anh/chị'} ơi, em đã kiểm tra đầu vào cho {len(rows)} mặt hàng anh/chị gửi:"]
    lines.append(f"- Đủ hàng / đủ đầu vào: {ok} mặt hàng")
    bad = [r for r in rows if r["muc"] in ("THIEU", "KHONG")]
    gan = [r for r in rows if r["muc"] in ("GAN", "LIENQUAN")]
    if bad:
        lines.append("- Chưa đủ đầu vào (cần bổ sung hóa đơn mua vào trước khi xuất):")
        lines += [f"  + {r['ten']} – {r['sl']:g} {r.get('dvt') or ''}: {r['ket_qua']}" for r in bad]
    if gan:
        lines.append("- Cần anh/chị xác nhận lại tên hàng / quy cách:")
        lines += [f"  + {r['ten']}: {r['ket_qua']}" + (f" (đang có: {(r['kho_ten'] or r['hd_ten'])[:70]})"
                                                      if (r['kho_ten'] or r['hd_ten']) else "") for r in gan]
    lines.append("Anh/chị kiểm tra giúp em, đủ thì em xuất hóa đơn ngay ạ.")
    return "\n".join(lines)


def write_excel(res: dict[str, Any]) -> bytes:
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font, PatternFill
    wb = Workbook()
    ws = wb.active
    ws.title = "Kiem tra dau vao"
    ws.append([f"KIỂM TRA ĐẦU VÀO TRƯỚC KHI XUẤT HÓA ĐƠN – MST {res['mst']} – HĐ từ {res['tu']} đến {res['den']}"])
    ws["A1"].font = Font(bold=True, size=13)
    t = res["tong"]
    ws.append([f"Đủ {t['DU']} · Gần giống cần xác nhận {t['GAN']} · Thiếu {t['THIEU']} · Chỉ vật tư liên quan {t['LIENQUAN']}"
               f" · Không có đầu vào {t['KHONG']}"])
    head = ["Mặt hàng khách yêu cầu", "Mô tả", "ĐVT", "SL", "Đơn giá", "Thành tiền", "KẾT QUẢ", "Hàng trong kho khớp",
            "Tồn kho", "Hàng trên HĐ mua vào khớp", "SL mua vào", "SL đã bán ra", "Còn theo HĐ", "Giá vốn bình quân",
            "HĐ mua vào", "Vật tư liên quan", "Ghi chú", "Nguồn"]
    ws.append(head)
    for c in ws[3]:
        c.font = Font(bold=True)
        c.fill = PatternFill("solid", fgColor="DDEBF7")
        c.alignment = Alignment(wrap_text=True, vertical="center")
    fill = {"DU": "E2EFDA", "GAN": "FFF2CC", "THIEU": "F8CBAD", "LIENQUAN": "FCE4D6", "KHONG": "F8CBAD"}
    for r in res["rows"]:
        ws.append([r["ten"], r.get("mo_ta"), r.get("dvt"), r.get("sl"), r.get("dg"), r.get("tt"), r["ket_qua"], r["kho_ten"],
                   r["kho_ton"], r["hd_ten"], r["hd_nhap"], r["hd_ban"], r["hd_con"],
                   round(r["gia_von"]) if r["gia_von"] else None, "\n".join(r.get("hd_mua_so") or []),
                   "\n".join(r["lien_quan"]), r["ghi_chu"], r.get("nguon")])
        ws.cell(ws.max_row, 7).fill = PatternFill("solid", fgColor=fill.get(r["muc"], "FFFFFF"))
        for c in (5, 6, 14):
            ws.cell(ws.max_row, c).number_format = "#,##0"
        for c in ws[ws.max_row]:
            c.alignment = Alignment(wrap_text=True, vertical="top")
    for col, w in zip("ABCDEFGHIJKLMNOPQR", (34, 24, 7, 9, 13, 15, 40, 30, 9, 30, 10, 10, 10, 13, 40, 45, 34, 16)):
        ws.column_dimensions[col].width = w
    ws.freeze_panes = "A4"
    ws2 = wb.create_sheet("Loi nhan khach")
    for ln in reply_text(res).splitlines():
        ws2.append([ln])
    ws2.column_dimensions["A"].width = 120
    b = io.BytesIO()
    wb.save(b)
    return b.getvalue()
