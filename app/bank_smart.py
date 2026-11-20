"""Đối chiếu sao kê THÔNG MINH: dò từng dòng sao kê xem thu / chi của AI, cho việc gì, mã đối tượng nào.

Nguồn dữ liệu (càng nhiều càng chính xác):
* Sao kê ngân hàng (bắt buộc).
* Hóa đơn trên HĐĐT từ tháng sao kê lùi về các tháng trước – đủ 5 loại: mua vào có mã, không mã, máy tính tiền,
  bán ra, bán ra máy tính tiền (dữ liệu đã tải / đọc nhanh trên máy + lấy thêm từ cổng nếu đang đăng nhập).
* Danh mục / tổng hợp công nợ PHẢI THU (khách hàng) và PHẢI TRẢ (nhà cung cấp) xuất từ MISA: lấy MÃ ĐỐI TƯỢNG, tên, MST,
  số tài khoản ngân hàng, số dư công nợ -> nhận ra đối tượng dù nội dung chuyển khoản ghi tắt / ghi tên người chuyển.
* Sổ chi tiết bán hàng: ghép tiền thu về với chứng từ bán hàng (kể cả hàng bán lẻ không có trên HĐĐT).
* AI (nếu có khóa): đoán đối tượng + mục đích cho các dòng còn chưa rõ – luôn đánh dấu CẦN KIỂM TRA.
"""
from __future__ import annotations

import json
import re
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Any

from .bank_post import core_name, fold, name_sim

# ------------------------------------------------------------------ đọc file danh mục / sổ
_SUMMARY = re.compile(r"^(so dong|tong cong|cong cong)\b|^(cong|tong)$")   # "Công ty..." KHÔNG phải dòng cộng


def _k(v: Any) -> str:
    return re.sub(r"\s+", " ", re.sub(r"[^a-z0-9 ]", " ", fold(v))).strip()


def _num(v: Any) -> float:
    if v is None or v == "":
        return 0.0
    if isinstance(v, (int, float)):
        return float(v)
    s = re.sub(r"[^\d,.\-]", "", str(v))
    if not s:
        return 0.0
    if s.count(".") > 1 or ("." in s and "," in s and s.rfind(",") > s.rfind(".")):
        s = s.replace(".", "").replace(",", ".")
    elif s.count(",") >= 1 and "." not in s:
        s = s.replace(",", "") if len(s.split(",")[-1]) == 3 else s.replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return 0.0


def _date(v: Any) -> date | None:
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    m = re.match(r"\s*(\d{1,2})[/\-.](\d{1,2})[/\-.](\d{4})", str(v or ""))
    if m:
        try:
            return date(int(m.group(3)), int(m.group(2)), int(m.group(1)))
        except ValueError:
            return None
    m = re.match(r"\s*(\d{4})-(\d{2})-(\d{2})", str(v or ""))
    return date(int(m.group(1)), int(m.group(2)), int(m.group(3))) if m else None


def _rows(name: str, data: bytes) -> list[list[Any]]:
    from .misa_import import load_any_workbook
    wb = load_any_workbook(data, values=True)
    best: list[list[Any]] = []
    for ws in wb.worksheets:
        rows = [list(r) for r in ws.iter_rows(values_only=True)]
        if len(rows) > len(best):
            best = rows
    return best


def _header(rows: list[list[Any]], need: tuple[str, ...]) -> tuple[int, list[str]] | None:
    """Dòng tiêu đề (trong 30 dòng đầu) có đủ các cột cần; gộp dòng con (Nợ / Có) ngay dưới nếu có."""
    for i, r in enumerate(rows[:30]):
        heads = [_k(c) for c in r]
        if all(any(h.startswith(n) for h in heads) for n in need):
            sub = [_k(c) for c in rows[i + 1]] if i + 1 < len(rows) else []
            if sub and sum(1 for s in sub if s in ("no", "co")) >= 2:
                cur, out = "", []
                for j in range(max(len(heads), len(sub))):
                    h = heads[j] if j < len(heads) else ""
                    cur = h or cur
                    s = sub[j] if j < len(sub) else ""
                    out.append(f"{cur} {s}".strip() if s in ("no", "co") else h)
                return i + 1, out
            return i, heads
    return None


def _col(heads: list[str], *names: str, avoid: tuple[str, ...] = ()) -> int | None:
    for n in names:
        for j, h in enumerate(heads):
            if h.startswith(n) and not any(a in h for a in avoid):
                return j
    return None


def norm_mst(s: Any) -> str:
    m = re.search(r"(?<!\d)\d{12}(?!\d)|\d{10}(?:\s*-\s*\d{3})?", str(s or ""))
    return re.sub(r"\s", "", m.group(0)) if m else ""


def read_partners(name: str, data: bytes, loai: str) -> list[dict[str, Any]]:
    """Danh mục / tổng hợp / công nợ đầu kỳ khách hàng (loai="KH") hoặc nhà cung cấp (loai="NCC").
    Trả [{ma, ten, mst, tk: [số TK ngân hàng], du: số dư công nợ (dương = còn phải thu / phải trả), loai}]."""
    rows = _rows(name, data)
    h = _header(rows, ("ma",)) or _header(rows, ("ten",))
    if not h:
        raise ValueError(f"{name}: không thấy dòng tiêu đề (cần cột Mã / Tên khách hàng hoặc nhà cung cấp)")
    hr, heads = h
    c_ma = _col(heads, "ma khach hang", "ma nha cung cap", "ma doi tuong", "ma ncc", "ma kh", "ma dt",
                avoid=("so thue",))
    if c_ma is None:
        c_ma = next((j for j, x in enumerate(heads) if x == "ma"), None)
    c_ten = _col(heads, "ten khach hang", "ten nha cung cap", "ten doi tuong", "ten ncc", "ten kh", "ten")
    c_mst = _col(heads, "ma so thue", "mst")
    c_tk = _col(heads, "so tai khoan ngan hang", "tai khoan ngan hang", "so tk ngan hang", "stk", "so tai khoan")
    c_dia = _col(heads, "dia chi")
    # số dư: tổng hợp công nợ (Số dư cuối kỳ Nợ / Có), công nợ đầu kỳ (Dư Nợ / Dư Có), danh sách (Công nợ)
    c_no = _col(heads, "so du cuoi ky no", "du no cuoi", "so du no", "du no")
    c_co = _col(heads, "so du cuoi ky co", "du co cuoi", "so du co", "du co")
    # tổng hợp công nợ: số dư = đầu kỳ + phát sinh (Nợ - Có) – chắc chắn hơn cột cuối kỳ
    c_dkn, c_dkc = _col(heads, "so du dau ky no"), _col(heads, "so du dau ky co")
    c_psn, c_psc = _col(heads, "phat sinh no", "so phat sinh no"), _col(heads, "phat sinh co", "so phat sinh co")
    c_cn = _col(heads, "cong no", "so con no", "con phai thu", "con phai tra")
    detail = _col(heads, "so chung tu") is not None     # sổ CHI TIẾT công nợ: số dư lũy kế từng dòng
    out: dict[str, dict[str, Any]] = {}
    for r in rows[hr + 1:]:
        g = lambda c: (r[c] if c is not None and c < len(r) else None)  # noqa: E731
        ma, ten = str(g(c_ma) or "").strip(), str(g(c_ten) or "").strip()
        if not ma and not ten or _SUMMARY.match(_k(ma or ten)) or _k(ten).startswith("ten khach hang"):
            continue
        mst = norm_mst(g(c_mst)) or (norm_mst(ma) if re.fullmatch(r"[\d\- ]{10,16}", ma) else "")
        tk = [x for x in re.findall(r"\d{6,20}", str(g(c_tk) or "")) if len(x) >= 6]
        if None not in (c_dkn, c_dkc, c_psn, c_psc):
            du = _num(g(c_dkn)) - _num(g(c_dkc)) + _num(g(c_psn)) - _num(g(c_psc))
            du = du if loai == "KH" else -du
        elif c_no is not None or c_co is not None:
            du = _num(g(c_no)) - _num(g(c_co))
            du = du if loai == "KH" else -du
        else:
            du = _num(g(c_cn))
        key = ma or _k(ten)
        p = out.setdefault(key, {"ma": ma, "ten": ten, "mst": mst, "tk": [], "du": 0.0, "dia_chi": "", "loai": loai})
        p["ten"] = p["ten"] or ten
        p["mst"] = p["mst"] or mst
        p["tk"] = sorted(set(p["tk"]) | set(tk))
        if detail:
            if g(c_no) is not None or g(c_co) is not None:
                p["du"] = du                         # lấy số dư của dòng cuối cùng
        else:
            p["du"] += du
        p["dia_chi"] = p["dia_chi"] or str(g(c_dia) or "")
    return list(out.values())


def read_sales(name: str, data: bytes) -> list[dict[str, Any]]:
    """Sổ chi tiết bán hàng (MISA) -> chứng từ bán hàng gộp theo Số chứng từ:
    [{so_ct, so_hd, ngay, ma_kh, ten, mst, tong, dien_giai}]."""
    rows = _rows(name, data)
    h = _header(rows, ("so chung tu",)) or _header(rows, ("so hoa don",))
    if not h:
        raise ValueError(f"{name}: không thấy tiêu đề sổ chi tiết bán hàng (cần cột Số chứng từ / Số hóa đơn)")
    hr, heads = h
    c = {k: _col(heads, *v) for k, v in {
        "ngay": ("ngay chung tu", "ngay hach toan", "ngay hoa don", "ngay"),
        "so_ct": ("so chung tu",), "so_hd": ("so hoa don",), "ma_kh": ("ma khach hang", "ma doi tuong"),
        "ten": ("ten khach hang", "ten doi tuong", "khach hang"), "mst": ("ma so thue",),
        "tong": ("tong thanh toan", "tong tien thanh toan", "tong cong tien thanh toan"),
        "ds": ("doanh so ban", "doanh so", "thanh tien"), "ck": ("chiet khau",), "thue": ("thue gtgt", "tien thue"),
        "tra": ("gia tri tra lai",), "giam": ("gia tri giam gia",), "dg": ("dien giai chung", "dien giai"),
    }.items()}
    vouchers: dict[str, dict[str, Any]] = {}
    for r in rows[hr + 1:]:
        g = lambda k: (r[c[k]] if c[k] is not None and c[k] < len(r) else None)  # noqa: E731
        so_ct, so_hd = str(g("so_ct") or "").strip(), str(g("so_hd") or "").strip()
        if not (so_ct or so_hd) or _SUMMARY.match(_k(r[0] if r else "")):
            continue
        if c["tong"] is not None:
            amt = _num(g("tong"))
        else:
            amt = _num(g("ds")) - _num(g("ck")) - _num(g("tra")) - _num(g("giam")) + _num(g("thue"))
        v = vouchers.setdefault(so_ct or so_hd, {"so_ct": so_ct, "so_hd": so_hd, "ngay": _date(g("ngay")),
                                                  "ma_kh": str(g("ma_kh") or "").strip(),
                                                  "ten": str(g("ten") or "").strip(), "mst": norm_mst(g("mst")),
                                                  "tong": 0.0, "dien_giai": str(g("dg") or "")[:200]})
        v["tong"] += amt
    return [v for v in vouchers.values() if v["tong"] > 1]


# ------------------------------------------------------------------ danh bạ đối tượng
_VOWEL_GROUPS = re.compile(r"[aeiouy]+")


def _syllable(w: str) -> bool:
    """Chữ là 1 âm tiết tiếng Việt (1 cụm nguyên âm: "ngan", "khang", "phuong") – dễ trùng; chữ viết tắt ("hdt") hoặc
    tên nước ngoài nhiều âm tiết ("alpha", "nikita") thì không."""
    return len(_VOWEL_GROUPS.findall(w)) == 1 and w.isalpha()


_WEAK = set("cong ty tnhh mtv co phan cp dntn chi nhanh cn tai thanh pho tp ho chi minh hcm viet nam vn "
            "thuong mai dich vu tm dv san xuat sx xay dung xd dau tu dt ky thuat cong nghe va ha noi "
            "ho kinh doanh hkd xuat nhap khau xnk phat trien tong hop thiet bi vat tu giai phap quoc te "
            # chữ hay gặp trong nội dung chuyển khoản
            "chuyen tien khoan ck tt thanh toan mua ban hang hoa don hd so ngay ma giao dich cho cua tu den "
            "nhan gui tra no cu cong qua lan dot thang nam phi".split())


class Directory:
    """Danh bạ khách hàng + nhà cung cấp (từ danh mục công nợ) để nhận diện đối tượng trên sao kê."""

    def __init__(self, partners: list[dict[str, Any]]):
        self.items = partners
        self.by_mst: dict[str, dict[str, Any]] = {}
        self.by_tk: dict[str, dict[str, Any]] = {}
        self.by_tok: dict[str, list[dict[str, Any]]] = {}
        from collections import Counter
        df = Counter(w for p in partners for w in set(core_name(p["ten"]).split()))
        lim = max(3, len(partners) // 100)            # chữ có ở > 1% đối tượng: chữ chung, không phân biệt
        self.common = {w for w, n in df.items() if n > lim} | _WEAK
        for p in partners:
            p["_core"] = core_name(p["ten"])
            # phần tên PHÂN BIỆT (bỏ chữ loại hình / chữ chung), giữ thứ tự: "Cty TNHH Kỹ Thuật HDT" -> "hdt"
            p["_key"] = " ".join(w for w in p["_core"].split() if w not in self.common)
            if p["mst"]:
                self.by_mst.setdefault(p["mst"], p)
                self.by_mst.setdefault(p["mst"][:10], p)
            for t in p["tk"]:
                self.by_tk.setdefault(t, p)
            for w in set(p["_core"].split()) - _WEAK:
                if len(w) >= 3:
                    self.by_tok.setdefault(w, []).append(p)

    def __bool__(self) -> bool:
        return bool(self.items)

    def by_partner(self, mst: str, ten: str = "", loai: str = "") -> dict[str, Any] | None:
        """Đối tượng trong danh mục ứng với MST / tên trên hóa đơn."""
        mst = norm_mst(mst)
        p = self.by_mst.get(mst) or (self.by_mst.get(mst[:10]) if mst else None)
        if p:
            return p
        if ten:
            cands = self._cands(core_name(ten))
            best = max(cands, key=lambda q: name_sim(q["ten"], ten), default=None)
            if best and name_sim(best["ten"], ten) >= 0.9 and (not loai or best["loai"] == loai):
                return best
        return None

    def _cands(self, text: str) -> list[dict[str, Any]]:
        seen, out = set(), []
        for w in set(text.split()) - _WEAK:
            for p in self.by_tok.get(w, [])[:60]:
                if id(p) not in seen:
                    seen.add(id(p))
                    out.append(p)
        return out

    def identify(self, t: dict[str, Any], own: str = "") -> dict[str, Any] | None:
        """Đối tượng của 1 giao dịch: số TK ngân hàng, MST, mã đối tượng, tên trong nội dung / tên đối ứng.
        own: tên chủ tài khoản (bỏ khỏi nội dung, vd "Khang Phat tt cty Nikita" -> chỉ còn "Nikita")."""
        raw = " ".join(str(t.get(k) or "") for k in ("mo_ta", "doi_ung", "tham_chieu", "tk_doi_ung"))
        text = _k(raw)
        own_core = re.split(r"\b(?:tai khoan|account|so tk|ngay thuc hien|date)\b", core_name(own))[0]
        own_key = [w for w in own_core.split() if not w.isdigit()]
        for k in range(len(own_key) - 1):              # bỏ từng cặp chữ liền nhau của tên mình ("khang phat")
            text = re.sub(rf"\b{re.escape(own_key[k])} {re.escape(own_key[k + 1])}\b", " ", text)
        text = re.sub(r"\s+", " ", text).strip()
        padded = f" {text} "
        digits = set(re.findall(r"\d{6,20}", raw.replace(" ", " "))) | {str(t.get("tk_doi_ung") or "").strip()}
        inflow = bool(t.get("co"))
        found: list[tuple[int, dict[str, Any], str]] = []
        for d in digits:
            if d in self.by_tk:
                found.append((100, self.by_tk[d], f"số TK ngân hàng {d}"))
        for m in re.findall(r"\d{10}(?:\s*-\s*\d{3})?", raw):
            p = self.by_mst.get(re.sub(r"\s", "", m)) or self.by_mst.get(m[:10])
            if p:
                found.append((95, p, f"MST {p['mst']} trong nội dung"))
        words = set(text.split())
        cands = self._cands(text + " " + _k(t.get("doi_ung")))
        for p in cands:
            ma = _k(p["ma"])
            if ma and len(ma) >= 4 and not ma.isdigit() and ma in words:
                found.append((85, p, f"mã {p['ma']} trong nội dung"))
            if t.get("doi_ung") and name_sim(p["ten"], t["doi_ung"]) >= 0.85:
                found.append((90, p, "tên người chuyển / người nhận"))
            elif len(p["_core"]) >= 6 and f" {p['_core']} " in padded:
                # tên ngắn 2 âm tiết tiếng Việt ("Hải Dương", "Kiều An") dễ trùng tên người chuyển -> cần kiểm tra
                short = len(p["_core"].split()) <= 2 and all(_syllable(w) for w in p["_core"].split())
                found.append((74 if short else 84, p, "tên đối tượng trong nội dung"))
            elif " " in p["_key"] and f" {p['_key']} " in padded:
                # tên riêng >= 2 chữ nằm LIỀN NHAU trong nội dung ("tt cty Tuan Duc mua hang" -> "tuan duc")
                found.append((80, p, f"tên riêng \"{p['_key']}\" trong nội dung"))
            elif p["_key"] and " " not in p["_key"] and len(p["_key"]) >= 3 and not _syllable(p["_key"]) and re.search(
                    r"(?:^| )(?:cty|cong ty|ct|dntn|hkd|cua hang|ch|tiem|shop) " + re.escape(p["_key"]) + r"(?: |$)", text):
                # tên riêng 1 chữ chỉ tính khi đứng ngay sau "cty / công ty / DNTN / HKD / cửa hàng" ("tt cty HDT")
                found.append((78, p, f"\"cty {p['_key']}\" trong nội dung"))
        if not found:
            return None
        scored = []
        for s, p, why in found:
            s += 5 if (p["loai"] == "KH") == inflow else -10   # tiền vào thường từ khách hàng, tiền ra cho NCC
            scored.append((s, p, why))
        scored.sort(key=lambda x: -x[0])
        s, p, why = scored[0]
        rivals = {q["ma"] or q["ten"] for s2, q, _ in scored if s2 >= s - 2 and q is not p}
        if s < 70:
            return None
        return {"ma": p["ma"], "ten": p["ten"], "mst": p["mst"], "loai": p["loai"], "why": why, "diem": s,
                "du": p.get("du", 0.0), "nghi_ngo": sorted(rivals)[:3]}


# ------------------------------------------------------------------ ghép sổ chi tiết bán hàng
def apply_sales(txns: list[dict[str, Any]], sales: list[dict[str, Any]], dirx: Directory | None) -> int:
    """Tiền thu về chưa ghép được hóa đơn: ghép với chứng từ trên sổ chi tiết bán hàng (1 CT đúng số tiền, hoặc
    nhiều CT của cùng khách hàng cộng đúng số tiền). Trả số giao dịch ghép được."""
    from .bank_post import _subset
    for v in sales:
        v["_con_lai"] = v["tong"]
        v["ngay_lap"] = v["ngay"]
    n = 0
    for t in sorted(txns, key=lambda t: t["ngay"]):
        if not t.get("co") or t.get("alloc"):
            continue
        amt, d = t["co"], t["ngay"]
        pool = [v for v in sales if v["_con_lai"] > 1 and (v["ngay"] is None or
                                                           d - timedelta(days=200) <= v["ngay"] <= d + timedelta(days=5))]
        dt = t.get("dt")
        mine = [v for v in pool if dt and ((dt["mst"] and v["mst"] and v["mst"][:10] == dt["mst"][:10]) or
                                            (v["ma_kh"] and v["ma_kh"] == dt["ma"]) or name_sim(v["ten"], dt["ten"]) >= 0.9)]
        pick, how = None, ""
        exact = [v for v in (mine or pool) if abs(v["_con_lai"] - amt) <= 1]
        if exact:
            exact.sort(key=lambda v: abs(((v["ngay"] or d) - d).days))
            if mine or len(exact) == 1 or abs(((exact[0]["ngay"] or d) - d).days) <= 3:
                pick, how = [exact[0]], ("Khớp 1 chứng từ sổ bán hàng của đúng khách hàng" if mine else
                                         "Khớp số tiền 1 chứng từ sổ bán hàng (cần kiểm tra khách hàng)")
        if not pick and len(mine) >= 2:
            combo = _subset(mine, amt, ref=d)
            if combo:
                pick, how = combo, f"Trả gộp {len(combo)} chứng từ sổ bán hàng cùng khách hàng"
        if not pick:
            continue
        for v in pick:
            inv = {"huong": "sold", "nmmst": v["mst"], "nmten": v["ten"], "khhdon": v["so_ct"], "shdon": v["so_hd"],
                   "ngay_lap": v["ngay"].isoformat() if v["ngay"] else "", "tgtttbso": v["tong"], "_con_lai": 0.0,
                   "_tt": [t], "_so_ban": v}
            t.setdefault("alloc", []).append({"inv": inv, "so_tien": v["_con_lai"]})
            v["_con_lai"] = 0.0
        t["cach_khop"] = (t.get("cach_khop", "") + "; " + how).strip("; ")
        if not dt:
            v = pick[0]
            p = dirx.by_partner(v["mst"], v["ten"], "KH") if dirx else None
            t["dt"] = {"ma": (p or {}).get("ma") or v["ma_kh"], "ten": (p or {}).get("ten") or v["ten"],
                       "mst": v["mst"], "loai": "KH", "why": "theo sổ chi tiết bán hàng", "diem": 75, "nghi_ngo": []}
        n += 1
    return n


# ------------------------------------------------------------------ AI cho các dòng còn chưa rõ
AI_KEYS = {"thanh_toan_ncc": "phai_tra", "thu_khach_hang": "phai_thu", "luong": "luong", "bao_hiem": "bhxh",
           "thue": "thue_khac", "phi_ngan_hang": "phi_nh", "lai_tien_gui": "lai", "rut_tien_mat": "tien_mat",
           "nop_tien_mat": "tien_mat", "tam_ung": "tam_ung", "vay": "vay", "noi_bo": "noi_bo", "khac": "khac"}

AI_PROMPT = """Bạn là kế toán Việt Nam. Với mỗi giao dịch sao kê ngân hàng dưới đây (chưa xác định được đối tượng),
hãy xác định: tiền thu từ ai / chi cho ai và cho việc gì. Chỉ được chọn đối tượng trong danh sách "ung_vien" của chính
giao dịch đó (theo "ma"), không chắc thì để null. "loai" là 1 trong: thanh_toan_ncc, thu_khach_hang, luong, bao_hiem, thue,
phi_ngan_hang, lai_tien_gui, rut_tien_mat, nop_tien_mat, tam_ung, vay, noi_bo, khac.
Trả về DUY NHẤT 1 mảng JSON: [{"id": <id>, "ma": <mã hoặc null>, "loai": "...", "muc_dich": "<diễn giải ngắn tiếng Việt>",
"tin_cay": "cao" | "thap"}]"""


def ai_classify(txns: list[dict[str, Any]], dirx: Directory | None, say=print) -> int:
    """Hỏi AI cho các giao dịch chưa có đối tượng / hóa đơn / quy tắc. Kết quả lưu t['_ai'] (luôn cần kiểm tra)."""
    from . import assistant
    p = "gemini" if assistant.key_for("gemini") else assistant.provider()
    if not assistant.key_for(p):
        say("Chưa có khóa AI – bỏ qua bước AI phân loại")
        return 0
    todo = [t for t in txns if not t.get("alloc") and not t.get("dt") and not t.get("_rule")]
    n = 0
    for k in range(0, len(todo), 40):
        batch = todo[k:k + 40]
        items = []
        for i, t in enumerate(batch):
            cands = dirx._cands(_k(" ".join([t.get("mo_ta", ""), t.get("doi_ung", "")]))) if dirx else []
            cands.sort(key=lambda q: -name_sim(q["ten"], t.get("doi_ung") or t.get("mo_ta")))
            items.append({"id": i, "ngay": t["ngay"].isoformat(), "thu_chi": "thu" if t.get("co") else "chi",
                          "so_tien": t.get("co") or t.get("no"), "noi_dung": t.get("mo_ta", "")[:300],
                          "doi_ung": t.get("doi_ung", ""),
                          "ung_vien": [{"ma": q["ma"], "ten": q["ten"], "loai": q["loai"]} for q in cands[:6]]})
        models = (["gemini-flash-lite-latest", "gemini-flash-latest"] if p == "gemini" else [assistant.model_for(p)])
        res = None
        for m in models:
            try:
                out = assistant._post_plain(p, m, [{"role": "system", "content": AI_PROMPT},
                                                   {"role": "user", "content": json.dumps(items, ensure_ascii=False)}])
                js = re.search(r"\[.*\]", out, re.S)
                if js:
                    res = json.loads(js.group(0))
                    break
            except Exception as e:  # noqa: BLE001
                say(f"AI lỗi ({m}): {str(e)[:120]}")
        for r in res or []:
            try:
                t = batch[int(r["id"])]
            except (KeyError, ValueError, IndexError, TypeError):
                continue
            ma = r.get("ma")
            cand = next((q for q in (dirx.items if dirx else []) if ma and q["ma"] == ma), None)
            t["_ai"] = {"key": AI_KEYS.get(r.get("loai"), "khac"), "muc_dich": str(r.get("muc_dich") or "")[:200],
                        "tin_cay": r.get("tin_cay", "thap")}
            if cand:
                t["dt"] = {"ma": cand["ma"], "ten": cand["ten"], "mst": cand["mst"], "loai": cand["loai"],
                           "why": "AI gợi ý", "diem": 60, "nghi_ngo": [], "ai": True}
            n += 1
    return n


# ------------------------------------------------------------------ gán mã đối tượng vào chứng từ
def fill_codes(vouchers: list[dict[str, Any]], dirx: Directory | None) -> None:
    for v in vouchers:
        t = v["tx"]
        dt = t.get("dt")
        for ln in v["lines"]:
            p = None
            if dirx and ln.get("mst"):
                p = dirx.by_partner(ln["mst"], ln.get("ten", ""))
            if p:
                ln["ma"], ln["ten"] = p["ma"] or p["mst"], p["ten"] or ln.get("ten", "")
            elif dt and (not ln.get("mst") or norm_mst(ln["mst"])[:10] == (dt["mst"] or "")[:10]) \
                    and (ln["no"] in ("131", "331", "1388", "3388", "141") or ln["co"] in ("131", "331", "1388", "3388", "141")):
                ln["ma"], ln["ten"], ln["mst"] = dt["ma"] or dt["mst"], dt["ten"], ln.get("mst") or dt["mst"]
        first = v["lines"][0]
        if first.get("ma"):
            v["ma"] = first["ma"]
            v["ten"] = first.get("ten") or v["ten"]


def status_of(t: dict[str, Any], v: dict[str, Any]) -> str:
    if t.get("alloc") and not v.get("kiem_tra"):
        return "Chắc chắn"
    if t.get("_rule") and not v.get("kiem_tra"):
        return "Chắc chắn"
    if t.get("alloc") or (t.get("dt") and not t["dt"].get("ai")) or t.get("_rule") or t.get("doi_ung"):
        return "Cần kiểm tra"
    return "CHƯA RÕ thu / chi cho ai"


# ------------------------------------------------------------------ file kết quả ghép thu chi
def write_result(dest: Path, vouchers: list[dict[str, Any]], info: dict[str, str], notes: list[str],
                 sources: list[str]) -> dict[str, int]:
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter
    fills = {"Chắc chắn": "E2EFDA", "Cần kiểm tra": "FFF2CC", "CHƯA RÕ thu / chi cho ai": "F8CBAD"}
    thin = Side(style="thin", color="BFBFBF")
    box = Border(left=thin, right=thin, top=thin, bottom=thin)
    heads = [("STT", 5), ("Ngày", 11), ("Thu / Chi", 8), ("Số tiền", 15), ("Nội dung trên sao kê", 50),
             ("Tên / TK đối ứng (sao kê)", 28), ("Mã đối tượng", 16), ("Tên đối tượng", 36), ("KH / NCC", 7),
             ("Căn cứ nhận diện đối tượng", 30), ("Thu / chi cho việc gì", 46), ("Hóa đơn / chứng từ ghép", 46),
             ("TK Nợ", 7), ("TK Có", 7), ("Số chứng từ", 12), ("Kết quả", 22), ("Cần kiểm tra", 36)]
    rows, stat = [], {"Chắc chắn": 0, "Cần kiểm tra": 0, "CHƯA RÕ thu / chi cho ai": 0}
    for v in vouchers:
        t = v["tx"]
        dt = t.get("dt") or {}
        st = status_of(t, v)
        stat[st] += 1
        docs = "; ".join(
            (f"CT bán hàng {a['inv']['_so_ban']['so_ct']}" + (f" (HĐ {a['inv']['_so_ban']['so_hd']})" if a['inv']['_so_ban']['so_hd'] else "")
             if a["inv"].get("_so_ban") else f"HĐ {a['inv'].get('khhdon') or ''}-{a['inv'].get('shdon')}")
            + f": {a['so_tien']:,.0f}".replace(",", ".") for a in t.get("alloc", []))
        why = dt.get("why", "") or ("chưa có trong danh mục – theo tên trên sao kê" if t.get("doi_ung") else "")
        if dt.get("nghi_ngo"):
            why += " (cũng giống: " + ", ".join(dt["nghi_ngo"]) + ")"
        purpose = v.get("ly_do") or ""
        if t.get("_ai"):
            purpose = f"[AI gợi ý – {t['_ai']['tin_cay']}] {t['_ai']['muc_dich']}"
        ln = v["lines"][0]
        rows.append(([len(rows) + 1, t["ngay"], "Thu" if t.get("co") else "Chi", t.get("co") or t.get("no"),
                      t.get("mo_ta", ""), " ".join(x for x in (t.get("doi_ung"), t.get("tk_doi_ung")) if x),
                      v.get("ma") or ln.get("ma") or "", v.get("ten") or dt.get("ten", ""),
                      dt.get("loai", ""), why, purpose, docs, ln["no"], ln["co"], v.get("so_ct", ""), st,
                      v.get("kiem_tra", "")], fills[st]))
    wb = Workbook()
    ws = wb.active
    ws.title = "Tong hop"
    ws["A1"] = "KẾT QUẢ GHÉP THU CHI THEO SAO KÊ"
    ws["A1"].font = Font(bold=True, size=14)
    lines = [f"Tài khoản: {info.get('so_tk', '')} {info.get('ngan_hang', '')} – {info.get('chu_tk', '')}",
             f"Tổng giao dịch: {len(vouchers)}",
             f"• Chắc chắn: {stat['Chắc chắn']}", f"• Cần kiểm tra: {stat['Cần kiểm tra']}",
             f"• CHƯA RÕ thu / chi cho ai: {stat['CHƯA RÕ thu / chi cho ai']}", "", "Nguồn dữ liệu đã dò:"] + \
        [f"• {s}" for s in sources] + ([""] + ["Lưu ý:"] + [f"• {n}" for n in notes] if notes else [])
    for i, s in enumerate(lines, 3):
        ws.cell(i, 1, s)
    ws.column_dimensions["A"].width = 120

    def sheet(title: str, data: list[tuple[list[Any], str]]) -> None:
        sh = wb.create_sheet(title)
        for c, (h, w) in enumerate(heads, 1):
            cell = sh.cell(1, c, h)
            cell.font, cell.border = Font(bold=True), box
            cell.fill = PatternFill("solid", fgColor="DDEBF7")
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            sh.column_dimensions[get_column_letter(c)].width = w
        for r, (vals, fill) in enumerate(data, 2):
            for c, val in enumerate(vals, 1):
                cell = sh.cell(r, c, val)
                cell.border = box
                cell.fill = PatternFill("solid", fgColor=fill)
                cell.alignment = Alignment(vertical="top", wrap_text=isinstance(val, str) and len(val) > 40)
                if isinstance(val, date):
                    cell.number_format = "dd/mm/yyyy"
                elif isinstance(val, float) and c == 4:
                    cell.number_format = "#,##0"
        sh.freeze_panes = "A2"
        sh.auto_filter.ref = f"A1:{get_column_letter(len(heads))}{max(2, len(data) + 1)}"

    sheet("Ghep thu chi", rows)
    sheet("Chua ro", [x for x in rows if x[1] == fills["CHƯA RÕ thu / chi cho ai"]])
    sheet("Can kiem tra", [x for x in rows if x[1] == fills["Cần kiểm tra"]])
    dest.parent.mkdir(parents=True, exist_ok=True)
    wb.save(dest)
    return {"chac": stat["Chắc chắn"], "kiem_tra": stat["Cần kiểm tra"], "chua_ro": stat["CHƯA RÕ thu / chi cho ai"]}
