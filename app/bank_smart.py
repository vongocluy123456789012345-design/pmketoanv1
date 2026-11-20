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
            main = re.split(r"\b(?:chi nhanh|cn)\b", p["_core"])[0].strip() or p["_core"]
            p["_main2"] = " ".join(main.split()[:2]) if len(main.split()) >= 2 else ""
            # phần tên PHÂN BIỆT (bỏ chữ loại hình / chữ chung), giữ thứ tự: "Cty TNHH Kỹ Thuật HDT" -> "hdt"
            p["_key"] = " ".join(w for w in main.split() if w not in self.common)
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
            elif p["_main2"] and len(p["_main2"]) >= 6 and re.search(r"(?:^| )" + re.escape(p["_main2"]) + r"(?: |$)", text) \
                    and text.find(p["_main2"]) <= 3:
                # nội dung bắt đầu bằng tên chính của đối tượng ("HOA SEN-040726..." -> Hoa Sen Home)
                found.append((76, p, f"nội dung bắt đầu bằng \"{p['_main2']}\""))
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
    if t.get("booked"):
        return "Đã có chứng từ trên sổ"
    if t.get("hoc") and t["hoc"]["share"] >= 0.8 and t["hoc"]["n"] >= 2:
        return "Chắc chắn"
    if t.get("alloc") and not v.get("kiem_tra"):
        return "Chắc chắn"
    if t.get("_rule") and not v.get("kiem_tra"):
        return "Chắc chắn"
    if t.get("alloc") or (t.get("dt") and not t["dt"].get("ai")) or t.get("_rule") or t.get("doi_ung") or t.get("hoc"):
        return "Cần kiểm tra"
    return "CHƯA RÕ thu / chi cho ai"


# ------------------------------------------------------------------ file kết quả ghép thu chi
def write_result(dest: Path, vouchers: list[dict[str, Any]], info: dict[str, str], notes: list[str],
                 sources: list[str], invoices: list[dict[str, Any]] | None = None,
                 period: tuple[date, date] | None = None) -> dict[str, int]:
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter
    fills = {"Chắc chắn": "E2EFDA", "Cần kiểm tra": "FFF2CC", "CHƯA RÕ thu / chi cho ai": "F8CBAD",
             "Đã có chứng từ trên sổ": "EDEDED"}
    thin = Side(style="thin", color="BFBFBF")
    box = Border(left=thin, right=thin, top=thin, bottom=thin)
    heads = [("STT", 5), ("Ngày", 11), ("Thu / Chi", 8), ("Số tiền", 15), ("Nội dung trên sao kê", 50),
             ("Tên / TK đối ứng (sao kê)", 28), ("Mã đối tượng", 16), ("Tên đối tượng", 36), ("KH / NCC", 7),
             ("Căn cứ nhận diện đối tượng", 30), ("Thu / chi cho việc gì", 46), ("Hóa đơn / chứng từ ghép", 46),
             ("TK Nợ", 7), ("TK Có", 7), ("Số chứng từ", 12), ("Kết quả", 22), ("Cần kiểm tra", 36),
             ("Gợi ý từ lịch sử (đã hạch toán trước đây)", 50), ("Độ tin cậy (%)", 9),
             ("Lý do khớp (Match Reason)", 60), ("Cảnh báo thuế / công nợ", 50)]
    rows, stat = [], {"Chắc chắn": 0, "Cần kiểm tra": 0, "CHƯA RÕ thu / chi cho ai": 0, "Đã có chứng từ trên sổ": 0}
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
        purpose = v.get("_dien_giai") or v.get("ly_do") or ""
        if v.get("_ly_do_misa"):
            purpose = f"{v['_ly_do_misa']}: {purpose}"
        b, hoc = t.get("booked"), t.get("hoc")
        if b:
            purpose = f"{b['ly_do'] + ': ' if b['ly_do'] else ''}{b['dien_giai']}"
            dt = {"ten": b["doi_tuong"], "why": f"đã hạch toán {b['so_ct']} ({b['nguon']})", "loai": dt.get("loai", "")}
        hoc_txt = (f"{hoc['doi_tuong']} | {hoc['ly_do']} | {hoc['dien_giai'][:60]} – {hoc['how']}, {hoc['n']} lần "
                   f"({int(hoc['share'] * 100)}%)") if hoc else ""
        if t.get("_ai"):
            purpose = f"[AI gợi ý – {t['_ai']['tin_cay']}] {t['_ai']['muc_dich']}"
        ln = v["lines"][0]
        rows.append(([len(rows) + 1, t["ngay"], "Thu" if t.get("co") else "Chi", t.get("co") or t.get("no"),
                      t.get("mo_ta", ""), " ".join(x for x in (t.get("doi_ung"), t.get("tk_doi_ung")) if x),
                      v.get("ma") or ln.get("ma") or "", (b or {}).get("doi_tuong") or v.get("_dt_out") or v.get("ten") or dt.get("ten", ""),
                      dt.get("loai", ""), dt.get("why", why) if b else why, purpose, docs, ln["no"], ln["co"],
                      b["so_ct"] if b else v.get("so_ct", ""), st, "" if b else v.get("kiem_tra", ""), hoc_txt,
                      confidence(t, v), match_reason(t), t.get("canh_bao", "")], fills[st]))
    wb = Workbook()
    ws = wb.active
    ws.title = "Tong hop"
    ws["A1"] = "KẾT QUẢ GHÉP THU CHI THEO SAO KÊ"
    ws["A1"].font = Font(bold=True, size=14)
    lines = [f"Tài khoản: {info.get('so_tk', '')} {info.get('ngan_hang', '')} – {info.get('chu_tk', '')}",
             f"Tổng giao dịch: {len(vouchers)}",
             f"• Chắc chắn: {stat['Chắc chắn']}", f"• Cần kiểm tra: {stat['Cần kiểm tra']}",
             f"• CHƯA RÕ thu / chi cho ai: {stat['CHƯA RÕ thu / chi cho ai']}",
             f"• Đã có chứng từ trên sổ (không xuất lại vào file nhập): {stat['Đã có chứng từ trên sổ']}",
             "", "Nguồn dữ liệu đã dò:"] + \
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
    # ---- truy vết 2 chiều xuống từng dòng hàng của hóa đơn đã ghép
    from .bank_post import NON_CASH_LIMIT
    sh = wb.create_sheet("Chi tiet dong hang")
    h2 = [("Ngày GD", 11), ("Số tiền GD", 14), ("Chứng từ NH", 22), ("Hóa đơn", 22), ("Đối tác", 34), ("Tổng HĐ", 14),
          ("Tỷ lệ đã TT", 9), ("Trạng thái TT", 22), ("Tên hàng / dịch vụ", 44), ("ĐVT", 8), ("SL", 9),
          ("Thành tiền dòng (gồm thuế)", 15), ("Đã TT qua NH", 15), ("Còn phải thu / trả", 15)]
    for c, (h, w) in enumerate(h2, 1):
        cell = sh.cell(1, c, h)
        cell.font, cell.fill, cell.border = Font(bold=True), PatternFill("solid", fgColor="DDEBF7"), box
        sh.column_dimensions[get_column_letter(c)].width = w
    r = 2
    seen_inv: set[int] = set()
    for v in vouchers:
        t = v["tx"]
        for it in item_lines(t):
            inv = it["inv"]
            key = (id(inv), it["ten"], it["line"])
            if hash(key) in seen_inv:
                continue
            seen_inv.add(hash(key))
            pm = inv.get("nbten") if inv.get("huong") == "purchase" else inv.get("nmten")
            vals = [t["ngay"], t.get("co") or t.get("no"), f"{t['ngay']:%d/%m/%Y} {t.get('tham_chieu') or ''}".strip(),
                    f"{inv.get('khhdon') or ''}-{inv.get('shdon') or ''}", pm or "", float(inv.get("tgtttbso") or 0),
                    f"{it['k'] * 100:.0f}%", it["status"], it["ten"], it["dvt"], it["sl"], it["line"], it["paid"], it["con"]]
            for c, val in enumerate(vals, 1):
                cell = sh.cell(r, c, val)
                cell.border = box
                if isinstance(val, date):
                    cell.number_format = "dd/mm/yyyy"
                elif isinstance(val, float) or (isinstance(val, int) and c in (2, 6, 12, 13, 14)):
                    cell.number_format = "#,##0"
            r += 1
    sh.freeze_panes = "A2"
    # ---- cảnh báo thuế: HĐ mua vào >= ngưỡng chưa thanh toán qua ngân hàng
    risks = tax_risks(invoices or [], *(period or (date.min, date.max)), NON_CASH_LIMIT) if invoices else []
    sh3 = wb.create_sheet("Canh bao thue")
    sh3.cell(1, 1, f"HĐ MUA VÀO TỪ {NON_CASH_LIMIT:,.0f} Đ CHƯA THANH TOÁN QUA NGÂN HÀNG TRONG KỲ SAO KÊ – rủi ro không được "
                   "khấu trừ thuế GTGT đầu vào (Luật thuế GTGT 2024: thanh toán không dùng tiền mặt từ 5 triệu)".replace(",", ".")
             ).font = Font(bold=True)
    h3 = [("Ngày HĐ", 11), ("Ký hiệu - Số HĐ", 22), ("MST người bán", 15), ("Tên người bán", 40), ("Tổng thanh toán", 15),
          ("Đã TT qua NH", 15), ("Còn chưa TT", 15), ("Cảnh báo", 50)]
    for c, (h, w) in enumerate(h3, 1):
        cell = sh3.cell(2, c, h)
        cell.font, cell.fill, cell.border = Font(bold=True), PatternFill("solid", fgColor="F8CBAD"), box
        sh3.column_dimensions[get_column_letter(c)].width = w
    for r3, x in enumerate(risks, 3):
        i = x["inv"]
        for c, val in enumerate([str(i.get("ngay_lap"))[:10], f"{i.get('khhdon') or ''}-{i.get('shdon') or ''}",
                                 i.get("nbmst"), i.get("nbten"), float(i.get("tgtttbso") or 0), x["paid"], x["con"],
                                 x["why"]], 1):
            cell = sh3.cell(r3, c, val)
            cell.border = box
            if isinstance(val, float):
                cell.number_format = "#,##0"
    dest.parent.mkdir(parents=True, exist_ok=True)
    wb.save(dest)
    return {"chac": stat["Chắc chắn"], "kiem_tra": stat["Cần kiểm tra"], "chua_ro": stat["CHƯA RÕ thu / chi cho ai"],
            "da_co": stat["Đã có chứng từ trên sổ"], "canh_bao_thue": len(risks)}


# ================================================================== sổ chi tiết công nợ (MISA, nhóm theo đối tượng)
_GROUP = re.compile(r"^\s*ten (khach hang|nha cung cap|doi tuong)\s*:?\s*(.+)$")


def read_ledger(name: str, data: bytes, loai: str) -> dict[str, Any] | None:
    """Chi tiết công nợ phải thu / phải trả kiểu MISA: dòng "Tên khách hàng: X" rồi các dòng chứng từ
    (Ngày, Số CT, Diễn giải, Phát sinh [phải thu | đã thu] / [đã trả | phải trả], Số dư).
    Trả {"partners": [...], "entries": [{ngay, so_ct, dien_giai, ten, phat_sinh, da, loai}]} hoặc None nếu không phải
    dạng này."""
    rows = _rows(name, data)
    hr = next((i for i, r in enumerate(rows[:30]) if any(_k(c).startswith("so chung tu") for c in r)), None)
    if hr is None or not any(_GROUP.match(_k(r[0] if r else "")) for r in rows[hr:hr + 40]):
        return None
    heads = [_k(c) for c in rows[hr]]
    sub = [_k(c) for c in rows[hr + 1]] if hr + 1 < len(rows) else []

    def col(*ns):
        return next((j for j, h in enumerate(heads) if any(h.startswith(n) for n in ns)), None)

    def sc(*ns):
        return next((j for j, h in enumerate(sub) if any(n in h for n in ns)), None)

    c_ngay, c_ct, c_dg = col("ngay chung tu", "ngay hach toan"), col("so chung tu"), col("dien giai")
    c_ps = sc("so phai thu", "so phai tra")
    c_da = sc("so da thu", "so da tra")
    c_con = sc("so con phai thu", "so con phai tra")
    partners: dict[str, dict[str, Any]] = {}
    entries: list[dict[str, Any]] = []
    cur = None
    for r in rows[hr + 1:]:
        def g(c):
            return r[c] if c is not None and c < len(r) else None
        if _GROUP.match(_k(r[0] if r else "")):
            ten = re.split(r":", str(r[0]), maxsplit=1)[-1].strip()
            cur = partners.setdefault(_k(ten), {"ma": "", "ten": ten, "mst": norm_mst(ten), "tk": [], "du": 0.0,
                                                "dia_chi": "", "loai": loai})
            continue
        if cur is None or not g(c_ct):
            continue
        if c_con is not None and g(c_con) not in (None, ""):
            cur["du"] = _num(g(c_con))
        entries.append({"ngay": _date(g(c_ngay)), "so_ct": str(g(c_ct)).strip(), "dien_giai": str(g(c_dg) or "")[:200],
                        "ten": cur["ten"], "phat_sinh": _num(g(c_ps)), "da": _num(g(c_da)), "loai": loai})
    return {"partners": list(partners.values()), "entries": entries}


# ================================================================== sổ thu / chi tiền gửi đã nhập (lịch sử)
def read_bank_book(name: str, data: bytes) -> dict[str, Any] | None:
    """Danh sách Thu tiền gửi / Chi tiền gửi xuất từ MISA (STT, Ngày hạch toán, Số chứng từ, Diễn giải, Số tiền,
    Đối tượng, Lý do thu/chi, Loại chứng từ) -> {"huong": "thu"|"chi", "entries": [...], "header", "hr"}."""
    rows = _rows(name, data)
    h = _header(rows, ("so chung tu", "so tien"))
    if not h:
        return None
    hr, heads = h
    title = _k(rows[0][0] if rows and rows[0] else "")
    if any(x.startswith("ly do chi") for x in heads):
        huong = "chi"
    elif any(x.startswith("ly do thu") for x in heads):
        huong = "thu"
    else:
        huong = "chi" if "chi" in title else "thu"
    names = {"ngay": ("ngay hach toan", "ngay chung tu", "ngay"), "so_ct": ("so chung tu",), "dg": ("dien giai",),
             "tien": ("so tien",), "dt": ("doi tuong", "ten doi tuong"), "ma": ("ma doi tuong",),
             "ly_do": ("ly do",), "loai_ct": ("loai chung tu",)}
    c = {k: _col(heads, *v) for k, v in names.items()}
    out = []
    for r in rows[hr + 1:]:
        def g(k):
            return r[c[k]] if c[k] is not None and c[k] < len(r) else None
        if not g("so_ct") or _num(g("tien")) <= 0:
            continue
        out.append({"ngay": _date(g("ngay")), "so_ct": str(g("so_ct")).strip(), "dien_giai": str(g("dg") or ""),
                    "so_tien": _num(g("tien")), "doi_tuong": str(g("dt") or "").strip(), "ma": str(g("ma") or "").strip(),
                    "ly_do": str(g("ly_do") or "").strip(), "loai_ct": str(g("loai_ct") or "").strip(), "huong": huong})
    return {"huong": huong, "entries": out, "header": [str(x) if x is not None else "" for x in rows[hr]], "hr": hr + 1}


# ================================================================== dòng sao kê ĐÃ hạch toán + học theo lịch sử
def mark_booked(txns: list[dict[str, Any]], book: list[dict[str, Any]], ledger: list[dict[str, Any]]) -> int:
    """Dòng sao kê đã có chứng từ thu / chi (sổ thu chi tiền gửi đã nhập, hoặc khoản đã thu / đã trả trên chi tiết công
    nợ): cùng chiều, đúng số tiền, lệch ngày <= 5. Ghi t["booked"] để không xuất trùng vào file nhập."""
    cands = [{"ngay": e["ngay"], "so_ct": e["so_ct"], "tien": e["so_tien"], "thu": e["huong"] == "thu",
              "doi_tuong": e["doi_tuong"], "ly_do": e["ly_do"], "dien_giai": e["dien_giai"], "nguon": "sổ thu chi tiền gửi"}
             for e in book]
    cands += [{"ngay": e["ngay"], "so_ct": e["so_ct"], "tien": e["da"], "thu": e["loai"] == "KH", "doi_tuong": e["ten"],
               "ly_do": "", "dien_giai": e["dien_giai"], "nguon": "chi tiết công nợ"} for e in ledger if e["da"] > 0]
    pairs = []
    for t in txns:
        amt, thu = t.get("co") or t.get("no"), bool(t.get("co"))
        for e in cands:
            if e["thu"] == thu and abs(e["tien"] - amt) <= 1 and e["ngay"] and abs((e["ngay"] - t["ngay"]).days) <= 5:
                pairs.append((abs((e["ngay"] - t["ngay"]).days), e["nguon"] != "sổ thu chi tiền gửi", id(t), t, e))
    pairs.sort(key=lambda x: (x[0], x[1]))
    used: set[tuple[str, float]] = set()
    n = 0
    for _g, _s, _i, t, e in pairs:
        key = (e["so_ct"], e["tien"])
        if t.get("booked") or key in used:
            continue
        used.add(key)
        t["booked"] = e
        n += 1
    return n


def _sig(text: str) -> str:
    """Dấu hiệu nội dung ngân hàng: bỏ số, mã, ngày giờ -> "tiktok shop", "ho kinh doanh genme dieu khac voc dang ck"."""
    ws = []
    for w in _k(text).split():
        if re.fullmatch(r"\d+x+\d+", w):
            ws.append("the")                  # số thẻ che (4229XX6023): giao dịch thẻ
        elif re.fullmatch(r"[a-z]{2,5}\d{5,}", w):
            ws.append(re.match(r"[a-z]+", w).group(0) + "#")   # mã đơn có tiền tố chữ (DHN13082694026 -> dhn#)
        elif not any(ch.isdigit() for ch in w) and len(w) > 1:
            ws.append(w)
    return " ".join(ws[:8])


_FIRST_WEAK = {"thanh", "toan", "ho", "kinh", "doanh", "chuyen", "tien", "tt", "ck", "cong", "ty", "cty", "nhan", "gui",
               "the", "ib", "ibft", "mbvcb", "vietcombank", "ft", "dh"}


def learn(txns: list[dict[str, Any]]) -> int:
    """Học từ các dòng đã hạch toán (t["booked"]): cùng dấu hiệu nội dung ngân hàng + cùng chiều -> đối tượng, lý do,
    diễn giải kế toán đã dùng. Gán t["hoc"] cho dòng chưa hạch toán. Trả số dòng gợi ý được."""
    from collections import Counter, defaultdict
    memo: dict[tuple[bool, str], Counter] = defaultdict(Counter)
    first: dict[tuple[bool, str], Counter] = defaultdict(Counter)
    money: dict[tuple[bool, int], Counter] = defaultdict(Counter)
    for t in txns:
        b = t.get("booked")
        if not b:
            continue
        thu, sg = bool(t.get("co")), _sig(t.get("mo_ta", ""))
        if sg:
            memo[(thu, sg)][(b["doi_tuong"], b["ly_do"], b["dien_giai"])] += 1
            w = next((x for x in sg.split() if x not in _FIRST_WEAK and len(x) >= 3), "")
            if w:
                first[(thu, w)][(b["doi_tuong"], b["ly_do"])] += 1
        amt = round(t.get("co") or t.get("no"))
        if amt >= 1_000_000:
            money[(thu, amt)][(b["doi_tuong"], b["ly_do"], b["dien_giai"])] += 1
    keys = list(memo)
    n = 0
    for t in txns:
        if t.get("booked"):
            continue
        s, thu = _sig(t.get("mo_ta", "")), bool(t.get("co"))
        if not s:
            continue
        cnt = memo.get((thu, s))
        how = "cùng nội dung ngân hàng"
        if not cnt:
            ws = set(s.split())
            best = (0.0, None)
            for k in keys:
                if k[0] != thu:
                    continue
                kw = set(k[1].split())
                sim = len(ws & kw) / len(ws | kw)
                if sim > best[0]:
                    best = (sim, k)
            if best[1] is not None and best[0] >= 0.6:
                cnt, how = memo[best[1]], f"nội dung gần giống \"{best[1][1][:40]}\""
        if not cnt:
            w = next((x for x in s.split() if x not in _FIRST_WEAK and len(x) >= 3), "")
            c1 = first.get((thu, w))
            if c1 and sum(c1.values()) >= 2 and c1.most_common(1)[0][1] / sum(c1.values()) >= 0.7:
                (dt, ly), k = c1.most_common(1)[0]
                cnt, how = Counter({(dt, ly, ""): k, ("", "", "~"): sum(c1.values()) - k}), f"cùng từ khóa \"{w}\""
        if not cnt:
            c2 = money.get((thu, round(t.get("co") or t.get("no"))))
            if c2 and sum(c2.values()) >= 2:
                cnt, how = c2, "cùng số tiền định kỳ"
        if not cnt:
            continue
        total = sum(cnt.values())
        (dt, ly, dg), k = cnt.most_common(1)[0]
        share = k / total
        if share < 0.5 or (how == "cùng số tiền định kỳ" and share < 0.7):
            continue
        t["hoc"] = {"doi_tuong": dt, "ly_do": ly, "dien_giai": dg or "", "n": total, "share": round(share, 2), "how": how}
        n += 1
    return n


# ================================================================== xuất theo mẫu sổ thu / chi tiền gửi của khách
_LYDO = {"thu": {"phai_thu": "Thu tiền khách hàng (không theo hóa đơn)", "lai": "Thu lãi đầu tư tài chính",
                 "vay": "Vay nợ"},
         "chi": {"phai_tra": "Trả tiền nhà cung cấp (không theo hóa đơn)", "tam_ung": "Tạm ứng cho nhân viên",
                 "luong": "Trả lương nhân viên", "thue_gtgt": "Nộp thuế", "thue_tndn": "Nộp thuế",
                 "thue_tncn": "Nộp thuế", "thue_khac": "Nộp thuế", "bhxh": "Nộp bảo hiểm", "vay": "Trả nợ vay"}}


def next_numbers(book: list[dict[str, Any]], huong: str, default: str) -> list[Any]:
    """Tiếp số chứng từ theo sổ đã nhập: PTB2736 -> PTB2737 ([prefix, số tiếp theo, độ dài số])."""
    best = None
    for e in book:
        if e["huong"] != huong:
            continue
        m = re.fullmatch(r"([A-Za-z]+)(\d+)", e["so_ct"])
        if m and (best is None or int(m.group(2)) > best[1]) and m.group(1) in (default, "PTB", "PCB", "NTTK", "UNC"):
            best = (m.group(1), int(m.group(2)), len(m.group(2)))
    return [best[0], best[1] + 1, best[2]] if best else [default, 1, 5]


def prepare_export(vouchers: list[dict[str, Any]], book: list[dict[str, Any]]) -> list[dict[str, Any]]:
    """Chứng từ CHƯA hạch toán: số chứng từ tiếp theo sổ, lý do thu / chi theo danh mục MISA, đối tượng (tên), diễn
    giải theo cách kế toán đã ghi (học từ lịch sử). Trả danh sách chứng từ cần nhập (bỏ dòng đã có trên sổ)."""
    from .bank_post import accounts
    acc_key: dict[str, str] = {}
    for k, v in accounts("TT200").items():
        acc_key.setdefault(v, k)
    for k, v in accounts("TT133").items():
        acc_key.setdefault(v, k)
    nums = {"thu": next_numbers(book, "thu", "PTB"), "chi": next_numbers(book, "chi", "PCB")}
    seen = {h: sorted({e["ly_do"] for e in book if e["huong"] == h and e["ly_do"]}) for h in ("thu", "chi")}
    out = []
    for v in sorted(vouchers, key=lambda v: v["ngay"]):
        t = v["tx"]
        if t.get("booked"):
            continue
        h = v["loai"]
        pre, no, width = nums[h]
        v["so_ct"] = f"{pre}{no:0{width}d}"
        nums[h][1] += 1
        ln = v["lines"][0]
        other = ln["co"] if h == "thu" else ln["no"]
        ly = _LYDO[h].get(acc_key.get(other, ""), "Thu khác" if h == "thu" else "Chi khác")
        hoc = t.get("hoc")
        weak_hoc = hoc and not hoc["how"].startswith("cùng nội dung")
        if hoc and hoc["ly_do"] and not (weak_hoc and t.get("_rule") == "Tạm ứng cho nhân viên"):
            ly = hoc["ly_do"]
        elif seen[h] and ly not in seen[h] and acc_key.get(other, "") in ("luong", "bhxh", "thue_gtgt", "thue_tndn",
                                                                          "thue_tncn", "thue_khac"):
            other_ly = "Thu khác" if h == "thu" else "Chi khác"       # công ty chưa từng dùng lý do này -> Chi khác
            ly = other_ly if other_ly in seen[h] else ly
        v["_ly_do_misa"] = ly
        if hoc and hoc["share"] >= 0.7:
            v["_dt_out"] = hoc["doi_tuong"]       # kế toán đã ghi thế nào thì theo vậy (kể cả để trống đối tượng)
        else:
            v["_dt_out"] = (hoc or {}).get("doi_tuong") or (t.get("dt") or {}).get("ten") or v.get("ten") or ""
        v["_dien_giai"] = (hoc or {}).get("dien_giai") or v.get("ly_do") or ln["dien_giai"]
        v["_loai_ct"] = "Thu tiền gửi" if h == "thu" else "Ủy nhiệm chi"
        out.append(v)
    for h in ("thu", "chi"):
        for i, v in enumerate([x for x in out if x["loai"] == h], 1):
            v["_stt"] = i
    return out


# ================================================================== sổ chi tiết mua hàng + ghép khoản chi
def read_purchases(name: str, data: bytes) -> list[dict[str, Any]]:
    """Sổ chi tiết mua hàng (MISA) -> chứng từ mua gộp theo Số chứng từ: [{so_ct, so_hd, ngay, ten, mst, tong}].
    Sổ không có cột nhà cung cấp thì để trống (nối qua chi tiết công nợ phải trả theo số chứng từ)."""
    rows = _rows(name, data)
    h = _header(rows, ("so chung tu",))
    if not h:
        raise ValueError(f"{name}: không thấy tiêu đề sổ chi tiết mua hàng (cần cột Số chứng từ)")
    hr, heads = h
    names = {"ngay": ("ngay chung tu", "ngay hach toan", "ngay"), "so_ct": ("so chung tu",), "so_hd": ("so hoa don",),
             "ten": ("ten nha cung cap", "nha cung cap", "ten doi tuong"), "mst": ("ma so thue",),
             "tong": ("tong thanh toan", "tong tien thanh toan"), "gt": ("gia tri mua", "thanh tien"),
             "ck": ("chiet khau",), "tra": ("gia tri tra lai",), "giam": ("gia tri giam gia",), "thue": ("thue gtgt", "tien thue")}
    c = {k: _col(heads, *v, avoid=("theo dong",)) for k, v in names.items()}
    docs: dict[str, dict[str, Any]] = {}
    for r in rows[hr + 1:]:
        def g(k):
            return r[c[k]] if c[k] is not None and c[k] < len(r) else None
        so_ct = str(g("so_ct") or "").strip()
        if not so_ct or _SUMMARY.match(_k(r[0] if r else "")):
            continue
        amt = _num(g("tong")) if c["tong"] is not None else (
            _num(g("gt")) - _num(g("ck")) - _num(g("tra")) - _num(g("giam")) + _num(g("thue")))
        d = docs.setdefault(so_ct, {"so_ct": so_ct, "so_hd": str(g("so_hd") or "").strip(), "ngay": _date(g("ngay")),
                                    "ten": str(g("ten") or "").strip(), "mst": norm_mst(g("mst")), "tong": 0.0})
        d["tong"] += amt
    return [d for d in docs.values() if d["tong"] > 1]


def payable_docs(ledger: list[dict[str, Any]], purchases: list[dict[str, Any]]) -> list[dict[str, Any]]:
    """Chứng từ phải trả: gộp phát sinh phải trả trên chi tiết công nợ theo số chứng từ (có tên NCC) + chứng từ trên sổ
    mua hàng (lấy tên NCC theo số chứng từ trên công nợ nếu sổ mua hàng không có)."""
    by: dict[str, dict[str, Any]] = {}
    for e in ledger:
        if e["loai"] == "NCC" and e["phat_sinh"] > 0:
            d = by.setdefault(e["so_ct"], {"so_ct": e["so_ct"], "so_hd": "", "ngay": e["ngay"], "ten": e["ten"],
                                           "mst": "", "tong": 0.0})
            d["tong"] += e["phat_sinh"]
    for p in purchases:
        if p["so_ct"] in by:
            by[p["so_ct"]]["so_hd"] = by[p["so_ct"]]["so_hd"] or p["so_hd"]
            continue
        by[p["so_ct"]] = dict(p)
    return list(by.values())


def apply_payables(txns: list[dict[str, Any]], docs: list[dict[str, Any]], dirx: Directory | None) -> int:
    """Khoản chi chưa ghép hóa đơn / chưa hạch toán: đúng số tiền 1 chứng từ mua hàng (hoặc gộp nhiều chứng từ cùng NCC)
    trong 120 ngày trước -> nhà cung cấp + chứng từ."""
    from .bank_post import _subset
    for d in docs:
        d["_con_lai"] = d["tong"]
        d["ngay_lap"] = d["ngay"]
    n = 0
    for t in sorted(txns, key=lambda t: t["ngay"]):
        if not t.get("no") or t.get("alloc") or t.get("booked") or t.get("hoc"):
            continue
        amt, day = t["no"], t["ngay"]
        pool = [d for d in docs if d["_con_lai"] > 1 and (d["ngay"] is None or
                                                          day - timedelta(days=120) <= d["ngay"] <= day + timedelta(days=3))]
        dt = t.get("dt")
        mine = [d for d in pool if dt and d["ten"] and name_sim(d["ten"], dt["ten"]) >= 0.9]
        pick, how = None, ""
        exact = sorted([d for d in (mine or pool) if abs(d["_con_lai"] - amt) <= 1],
                       key=lambda d: abs(((d["ngay"] or day) - day).days))
        if exact and mine:
            pick = [exact[0]]
            how = "Khớp 1 chứng từ mua hàng của đúng NCC"
        elif exact and len(exact) == 1 and abs(((exact[0]["ngay"] or day) - day).days) <= 10:
            pick = [exact[0]]
            how = "Khớp số tiền 1 chứng từ mua hàng gần ngày (kiểm tra NCC)"
        elif exact and len({_k(d["ten"]) for d in exact if d["ten"]}) == 1 and all(d["ten"] for d in exact)                 and abs(((exact[0]["ngay"] or day) - day).days) <= 15:
            pick = [exact[0]]                     # nhiều chứng từ cùng số tiền nhưng cùng 1 NCC (vd quảng cáo TikTok)
            how = f"Khớp số tiền chứng từ mua hàng – {len(exact)} chứng từ cùng NCC"
        elif len(mine) >= 2:
            combo = _subset(mine, amt, ref=day)
            if combo:
                pick, how = combo, f"Trả gộp {len(combo)} chứng từ mua hàng cùng NCC"
        if not pick:
            continue
        for d in pick:
            inv = {"huong": "purchase", "nbmst": d["mst"], "nbten": d["ten"], "khhdon": d["so_ct"], "shdon": d["so_hd"],
                   "ngay_lap": d["ngay"].isoformat() if d["ngay"] else "", "tgtttbso": d["tong"], "_con_lai": 0.0,
                   "_tt": [t], "_so_mua": d}
            t.setdefault("alloc", []).append({"inv": inv, "so_tien": d["_con_lai"]})
            d["_con_lai"] = 0.0
        t["cach_khop"] = (t.get("cach_khop", "") + "; " + how).strip("; ")
        if not dt and pick[0]["ten"]:
            p = dirx.by_partner(pick[0]["mst"], pick[0]["ten"], "NCC") if dirx else None
            t["dt"] = {"ma": (p or {}).get("ma", ""), "ten": (p or {}).get("ten") or pick[0]["ten"], "mst": pick[0]["mst"],
                       "loai": "NCC", "why": "theo chứng từ mua hàng / công nợ phải trả", "diem": 75, "nghi_ngo": []}
        n += 1
    return n


# ================================================================== điền đúng file mẫu nhập khẩu Phiếu thu / chi tiền gửi
def _tpl_value(head: str, v: dict[str, Any], ln: dict[str, Any], info: dict[str, str], addr: str) -> Any:
    h = _k(head).replace("(*)", "").strip()
    t, thu = v["tx"], v["loai"] == "thu"
    code = ln.get("ma") or (t.get("dt") or {}).get("ma") or ""
    dg = v.get("_dien_giai") or v.get("ly_do") or ln["dien_giai"]
    rules = [
        (("ngay hach toan", "ngay chung tu"), v["ngay"]),
        (("so chung tu",), v["so_ct"]),
        (("ma doi tuong hach toan",), code),
        (("ma doi tuong",), code),
        (("ten doi tuong",), v.get("_dt_out") or v.get("ten") or ""),
        (("dia chi",), addr),
        (("nop vao tk", "so tai khoan chi"), info.get("so_tk") or None),
        (("mo tai ngan hang", "ten ngan hang chi"), info.get("ngan_hang") or None),
        (("so tai khoan nhan",), t.get("tk_doi_ung") or None),
        (("ten ngan hang nhan",), t.get("nh_doi_ung") or None),
        (("ly do thu", "ly do chi"), v.get("_ly_do_misa")),
        (("dien giai ly do thu", "noi dung thanh toan"), dg),
        (("dien giai hach toan", "dien giai"), ln.get("_dg_line") or dg),
        (("so tien",), ln["so_tien"]),
    ]
    for keys, val in rules:
        if any(h.startswith(k) for k in keys):
            if h.startswith("ly do") and not thu and h.startswith("ly do thu"):
                return None
            return val
    return None


def fill_misa_template(tpl: bytes, tpl_name: str, vouchers: list[dict[str, Any]], info: dict[str, str],
                       addr_of: dict[str, str] | None = None) -> tuple[bytes, str]:
    """Điền chứng từ vào ĐÚNG file mẫu nhập khẩu MISA (Phiếu thu / Phiếu chi tiền gửi) – giữ nguyên định dạng mẫu, xóa
    dòng ví dụ; Lý do thu / chi chọn theo danh sách chọn của mẫu. Trả (dữ liệu file, đuôi file)."""
    from . import fill_template as ft
    from . import misa_import as mi
    kind, ext = ft.kind_of(tpl), ft.out_ext(tpl, tpl_name)
    use_excel = kind in ("xls", "ods", "html", "xml2003") and ft.excel_available()
    probe = ft.probe_with_excel(tpl, ext) if use_excel else tpl
    wb, ws, hr, _cols, _unk = mi.read_template(probe)
    heads = {c: str(ws.cell(hr, c).value or "") for c in range(1, ws.max_column + 1) if ws.cell(hr, c).value}
    lists = {}
    try:
        lists = mi.template_lists(ws, hr, {c: f"c{c}" for c in heads})
    except Exception:  # noqa: BLE001
        pass
    text_cols = {c for c, h in heads.items() if any(_k(h).startswith(x) for x in
                                                   ("so chung tu", "ma doi tuong", "nop vao tk", "so tai khoan"))}
    keys = {c: (f"so_chung_tu_{c}" if c in text_cols else f"c{c}") for c in heads}
    ft.TEXT_KEYS.update(k for c, k in keys.items() if c in text_cols)
    rows = []
    for v in vouchers:
        for ln in v["lines"]:
            addr = (addr_of or {}).get(ln.get("ma") or "", "")
            row = {}
            for c, h in heads.items():
                val = _tpl_value(h, v, ln, info, addr)
                if val is not None and c in lists and isinstance(val, str):
                    val = mi._fit_option(val, lists[c])
                row[keys[c]] = val
            rows.append(row)
    sheet_index = wb.worksheets.index(ws) + 1
    if kind == "ooxml":
        return ft.fill_ooxml(tpl, ws.title, hr, keys, rows), ext
    if use_excel:
        return ft.fill_with_excel(tpl, ext, sheet_index, hr, keys, rows), ext
    import io
    from openpyxl import load_workbook
    wb2 = load_workbook(io.BytesIO(probe))
    ws2 = wb2.worksheets[sheet_index - 1]
    if ws2.max_row > hr:
        ws2.delete_rows(hr + 1, ws2.max_row - hr)
    for i, row in enumerate(rows, hr + 1):
        for c, k in keys.items():
            if row.get(k) is not None:
                ws2.cell(i, c, row[k])
    out = io.BytesIO()
    wb2.save(out)
    return out.getvalue(), ".xlsx"


# ================================================================== độ tin cậy, lý do khớp, cảnh báo thuế / công nợ
def confidence(t: dict[str, Any], v: dict[str, Any]) -> int:
    """Độ tin cậy (%) của kết quả ghép 1 dòng sao kê."""
    if t.get("booked"):
        return 100
    if t.get("alloc"):
        return int(t.get("tin_cay") or 80)
    hoc = t.get("hoc")
    if hoc:
        base = 95 if hoc["how"].startswith("cùng nội dung") else 85 if hoc["how"].startswith("bộ nhớ") else 80
        return int(base * hoc["share"])
    if t.get("_rule"):
        return 75 if v.get("kiem_tra") else 95
    dt = t.get("dt")
    if dt:
        return 50 if dt.get("ai") else int(min(dt.get("diem", 70), 85))
    if t.get("_ai"):
        return 40
    return 0


def match_reason(t: dict[str, Any]) -> str:
    if t.get("booked"):
        b = t["booked"]
        return f"Đã có chứng từ {b['so_ct']} trên {b['nguon']} (đúng số tiền, lệch ngày ≤ 5)"
    parts = []
    if t.get("cach_khop"):
        parts.append(t["cach_khop"])
    if t.get("hoc"):
        h = t["hoc"]
        parts.append(f"Theo lịch sử hạch toán: {h['how']} ({h['n']} lần, {int(h['share'] * 100)}%)")
    if t.get("_rule"):
        parts.append(f"Theo từ khóa: {t['_rule']}")
    if t.get("dt") and t["dt"].get("why"):
        parts.append(f"Đối tượng: {t['dt']['why']}")
    if t.get("_ai"):
        parts.append("AI gợi ý")
    return "; ".join(parts) or "Chưa xác định – không đoán (cần kế toán xem)"


def tax_risks(invoices: list[dict[str, Any]], d1: date, d2: date, limit: float) -> list[dict[str, Any]]:
    """HĐ mua vào từ ngưỡng (5 triệu) lập trong kỳ sao kê mà chưa thanh toán qua ngân hàng (hoặc cá nhân trả thay):
    rủi ro KHÔNG được khấu trừ thuế GTGT đầu vào."""
    out = []
    for i in invoices:
        if i.get("huong") != "purchase" or float(i.get("tgtttbso") or 0) < limit:
            continue
        try:
            d = date.fromisoformat(str(i.get("ngay_lap"))[:10])
        except ValueError:
            continue
        if not (d1 <= d <= d2):
            continue
        paid = float(i.get("tgtttbso") or 0) - float(i.get("_con_lai") or 0)
        by_person = any("Cá nhân trả thay" in (t.get("canh_bao") or "") for t in i.get("_tt", []))
        if i.get("_con_lai", 0) > 1 or by_person:
            out.append({"inv": i, "paid": paid, "con": float(i.get("_con_lai") or 0),
                        "why": ("Cá nhân trả thay – kiểm tra điều kiện khấu trừ" if by_person else
                                "Chưa thấy thanh toán qua ngân hàng trong sao kê" if paid <= 1 else
                                "Mới thanh toán 1 phần qua ngân hàng")})
    return out


def item_lines(t: dict[str, Any]) -> list[dict[str, Any]]:
    """Truy vết xuống từng dòng hàng của HĐ đã ghép: tỷ lệ k = đã trả / tổng HĐ, tiền đã trả từng dòng = thành tiền x k."""
    from .misa_import import from_xml
    rows = []
    for a in t.get("alloc", []):
        inv = a["inv"]
        total = float(inv.get("tgtttbso") or 0)
        paid_all = total - float(inv.get("_con_lai") or 0)
        k = min(1.0, paid_all / total) if total else 0.0
        items = []
        xml = inv.get("file_xml")
        if xml and Path(str(xml)).exists():
            try:
                items = from_xml(Path(str(xml)).read_bytes()).get("items") or []
            except Exception:  # noqa: BLE001
                items = []
        if not items:
            items = [{"ten": "(tổng hóa đơn / chứng từ – chưa có chi tiết dòng hàng)", "dvt": "", "sl": None,
                      "tt": total, "thue": 0, "ck": 0}]
        for it in items:
            line = float(it.get("tt") or 0) - float(it.get("ck") or 0) + float(it.get("thue") or 0)
            rows.append({"t": t, "inv": inv, "k": k, "ten": it.get("ten") or "", "dvt": it.get("dvt") or "",
                         "sl": it.get("sl"), "line": line, "paid": round(line * k), "con": round(line - line * k),
                         "status": "Đã thanh toán đủ" if k >= 0.999 else f"Thanh toán một phần ({k * 100:.0f}%)"})
    return rows


# ================================================================== bộ nhớ ánh xạ (học cho các kỳ sau)


def _mem_file() -> Path:
    from .config import ROOT
    return ROOT / "data" / "bank_memory.json"


def load_memory(acct: str) -> list[dict[str, Any]]:
    try:
        return json.loads(_mem_file().read_text(encoding="utf-8")).get(acct, [])
    except Exception:  # noqa: BLE001
        return []


def save_memory(acct: str, txns: list[dict[str, Any]]) -> int:
    """Lưu cặp (nội dung ngân hàng / STK đối ứng -> đối tượng, lý do, diễn giải) từ các dòng ĐÃ hạch toán của tài khoản
    này để kỳ sau tự gợi ý (kể cả khi không tải lại sổ thu chi cũ)."""
    if not acct:
        return 0
    f = _mem_file()
    try:
        allm = json.loads(f.read_text(encoding="utf-8"))
    except Exception:  # noqa: BLE001
        allm = {}
    cur = {(m["thu"], m["sig"], m["doi_tuong"], m["ly_do"], m["dien_giai"]): m for m in allm.get(acct, [])}
    n = 0
    for t in txns:
        b = t.get("booked")
        if not b:
            continue
        key = (bool(t.get("co")), _sig(t.get("mo_ta", "")), b["doi_tuong"], b["ly_do"], b["dien_giai"])
        if not key[1]:
            continue
        m = cur.setdefault(key, {"thu": key[0], "sig": key[1], "doi_tuong": key[2], "ly_do": key[3],
                                 "dien_giai": key[4], "n": 0, "tk": ""})
        m["n"] += 1
        m["tk"] = m["tk"] or str(t.get("tk_doi_ung") or "")
        n += 1
    allm[acct] = list(cur.values())[-3000:]
    f.parent.mkdir(parents=True, exist_ok=True)
    f.write_text(json.dumps(allm, ensure_ascii=False), encoding="utf-8")
    return n


def apply_memory(txns: list[dict[str, Any]], memory: list[dict[str, Any]]) -> int:
    """Dòng chưa hạch toán, chưa có gợi ý: dùng bộ nhớ các kỳ trước (cùng STK đối ứng, hoặc cùng nội dung ngân hàng)."""
    from collections import Counter
    if not memory:
        return 0
    by_sig: dict[tuple[bool, str], Counter] = {}
    by_tk: dict[tuple[bool, str], Counter] = {}
    for m in memory:
        by_sig.setdefault((m["thu"], m["sig"]), Counter())[(m["doi_tuong"], m["ly_do"], m["dien_giai"])] += m["n"]
        if m.get("tk"):
            by_tk.setdefault((m["thu"], m["tk"]), Counter())[(m["doi_tuong"], m["ly_do"], m["dien_giai"])] += m["n"]
    n = 0
    for t in txns:
        if t.get("booked") or t.get("hoc"):
            continue
        thu = bool(t.get("co"))
        cnt = by_tk.get((thu, str(t.get("tk_doi_ung") or ""))) if t.get("tk_doi_ung") else None
        how = "bộ nhớ: cùng STK đối ứng"
        if not cnt:
            cnt, how = by_sig.get((thu, _sig(t.get("mo_ta", "")))), "bộ nhớ: cùng nội dung ngân hàng các kỳ trước"
        if not cnt:
            continue
        (dt, ly, dg), k = cnt.most_common(1)[0]
        total = sum(cnt.values())
        if k / total < 0.6:
            continue
        t["hoc"] = {"doi_tuong": dt, "ly_do": ly, "dien_giai": dg, "n": total, "share": round(k / total, 2), "how": how}
        n += 1
    return n
