"""Đối chiếu dòng tiền sao kê với hóa đơn (kể cả hóa đơn các tháng trước lấy từ cổng HĐĐT),
hạch toán Nợ/Có và xuất file nhập khẩu chứng từ ngân hàng cho MISA SME và EasyBooks.

Phần đọc sao kê nằm ở bank.py (giữ nguyên). File này:
  * statement_info     – số tài khoản / ngân hàng / tên chủ tài khoản trên sao kê
  * portal_invoices    – danh sách hóa đơn mua vào / bán ra các tháng trước từ tab cổng đã đăng nhập
  * allocate           – gán mỗi giao dịch cho 1 hoặc nhiều hóa đơn (trả gộp, trả 1 phần, theo số HĐ trong nội dung)
  * post               – định khoản Nợ/Có theo chế độ TT200 / TT133
  * export_misa / export_easybooks – điền theo file mẫu của phần mềm (nếu có) hoặc bộ cột mặc định
  * write_report       – file Excel đối chiếu: theo giao dịch, theo hóa đơn, hóa đơn còn nợ, chứng từ hạch toán
"""
from __future__ import annotations

import json
import re
import time
import unicodedata
from datetime import date, timedelta
from difflib import SequenceMatcher
from pathlib import Path
from typing import Any

from .bank import _score, _tokens

TOL = 1000  # sai lệch cho phép khi cộng nhiều hóa đơn (đồng)


# ------------------------------------------------------------------ tiện ích chữ
def fold(s: Any) -> str:
    s = unicodedata.normalize("NFD", str(s or ""))
    s = "".join(c for c in s if unicodedata.category(c) != "Mn").replace("đ", "d").replace("Đ", "D").lower()
    return re.sub(r"\s+", " ", re.sub(r"[^a-z0-9 ]", " ", s)).strip()


_LEGAL = (r"\b(cong ty|cty|ct|tnhh|mtv|mot thanh vien|co phan|ctcp|cp|trach nhiem huu han|dntn|doanh nghiep tu nhan|"
          r"chi nhanh|ho kinh doanh|hkd|tm|dv|tmdv|thuong mai dich vu)\b")


def core_name(s: Any) -> str:
    return re.sub(r"\s+", " ", re.sub(_LEGAL, " ", fold(s))).strip()


def name_sim(a: Any, b: Any) -> float:
    a, b = core_name(a), core_name(b)
    if len(a) < 4 or len(b) < 4:
        return 0.0
    if a in b or b in a:
        return 1.0
    return SequenceMatcher(None, a, b).ratio()


def _money(v: float) -> str:
    return f"{v:,.0f}".replace(",", ".")


def _vn(d: Any) -> str:
    try:
        return date.fromisoformat(str(d)[:10]).strftime("%d/%m/%Y")
    except ValueError:
        return str(d or "")


def _inv_date(inv: dict[str, Any]) -> date | None:
    try:
        return date.fromisoformat(str(inv.get("ngay_lap"))[:10])
    except ValueError:
        return None


def inv_label(inv: dict[str, Any]) -> str:
    return f"{inv.get('khhdon') or ''}-{inv.get('shdon')} ngày {_vn(inv.get('ngay_lap'))}"


def partner(inv: dict[str, Any]) -> tuple[str, str]:
    """(MST, tên) của đối tác trên hóa đơn: mua vào -> người bán, bán ra -> người mua."""
    if inv.get("huong") == "purchase":
        return str(inv.get("nbmst") or ""), str(inv.get("nbten") or "")
    return str(inv.get("nmmst") or ""), str(inv.get("nmten") or "")


# ------------------------------------------------------------------ thông tin sao kê
_BANKS = [("techcombank", "Techcombank"), ("ky thuong", "Techcombank"), ("vietcombank", "Vietcombank"),
          ("ngoai thuong", "Vietcombank"), ("vietinbank", "VietinBank"), ("cong thuong", "VietinBank"),
          ("bidv", "BIDV"), ("dau tu va phat trien", "BIDV"), ("agribank", "Agribank"),
          ("mbbank", "MB"), ("quan doi", "MB"), ("acb", "ACB"), ("a chau", "ACB"), ("vpbank", "VPBank"),
          ("viet nam thinh vuong", "VPBank"), ("sacombank", "Sacombank"), ("sai gon thuong tin", "Sacombank"),
          ("tpbank", "TPBank"), ("tien phong", "TPBank"), ("hdbank", "HDBank"), ("vib", "VIB"),
          ("shb", "SHB"), ("ocb", "OCB"), ("msb", "MSB"), ("seabank", "SeABank")]


def statement_info(name: str, data: bytes) -> dict[str, str]:
    """Số tài khoản, ngân hàng, tên chủ TK in trên đầu sao kê (PDF / Excel)."""
    text = ""
    try:
        if name.lower().endswith(".pdf") or data[:5] == b"%PDF-":
            import io

            import pdfplumber
            with pdfplumber.open(io.BytesIO(data)) as pdf:
                text = pdf.pages[0].extract_text() or ""
        else:
            from .bank import _excel_rows
            for rows in _excel_rows(data, name):
                text += "\n".join(" ".join(str(c) for c in r if c not in (None, "")) for r in rows[:25])
                break
    except Exception:  # noqa: BLE001
        pass
    f = fold(text)
    info = {"so_tk": "", "ngan_hang": "", "chu_tk": ""}
    m = re.search(r"(?:so tai khoan|account number|account no|so tk)\D{0,25}?(\d[\d ]{3,20}\d)", f)
    if m:
        info["so_tk"] = m.group(1).replace(" ", "")
    for k, v in _BANKS:
        if re.search(r"\b" + k + r"\b", f[:600]):
            info["ngan_hang"] = v
            break
    m = re.search(r"(?:chu tai khoan|ten tai khoan|account name|ten khach hang|customer name)\s*:?\s*(?:customer name|account name)?\s*([a-z0-9 ]{5,80})", f)
    if m:
        info["chu_tk"] = m.group(1).split(" so ")[0].strip().upper()
    return info


# ------------------------------------------------------------------ hóa đơn các tháng trước từ cổng HĐĐT
_PORTAL_KINDS = ("mua-vao", "mua-vao-khong-ma", "ban-ra")


def _months(d1: date, d2: date) -> list[tuple[int, int]]:
    out, cur = [], d1.replace(day=1)
    while cur <= d2:
        out.append((cur.year, cur.month))
        cur = date(cur.year + (cur.month == 12), cur.month % 12 + 1, 1)
    return out


def portal_invoices(root: Path, d1: date, d2: date, say=print) -> tuple[list[dict[str, Any]], str | None]:
    """Danh sách hóa đơn (mua vào có mã / không mã, bán ra) từ cổng HĐĐT cho các tháng d1..d2, gọi qua
    tab cổng thuế đang đăng nhập (không mở trang mới). Lưu bộ đệm theo tháng trong DocChungTu/.
    Tháng đã qua > 40 ngày dùng lại bộ đệm; tháng gần đây luôn lấy mới. Trả (hóa đơn, lỗi)."""
    from .auto_download import KINDS, PAGE_SIZE, summary_row
    from .config import settings
    from .gdt.browser import BROWSER

    cache_dir = root / "DocChungTu" / ".hoadon_cong"
    cache_dir.mkdir(parents=True, exist_ok=True)
    st = BROWSER.status()
    logged = bool(st.get("logged_in"))
    out: list[dict[str, Any]] = []
    err = None
    for y, m in _months(d1, d2):
        f = cache_dir / f"{y:04d}-{m:02d}.json"
        first = date(y, m, 1)
        last = date(y + (m == 12), m % 12 + 1, 1) - timedelta(days=1)
        fresh_ok = f.exists() and (date.today() - last).days > 40
        if fresh_ok or (f.exists() and not logged):
            out += json.loads(f.read_text(encoding="utf-8"))
            continue
        if not logged:
            err = "Chưa đăng nhập cổng HĐĐT trên Chrome – chỉ dùng hóa đơn đã tải sẵn trên máy"
            continue
        rows: list[dict[str, Any]] = []
        try:
            for key in _PORTAL_KINDS:
                kind = KINDS[key]
                search = f"tdlap=ge={first:%d/%m/%Y}T00:00:00;tdlap=le={last:%d/%m/%Y}T23:59:59"
                if kind["filter"]:
                    search += ";" + kind["filter"]
                state = None
                for _page in range(200):
                    q = f"sort=tdlap:desc&size={PAGE_SIZE}" + (f"&state={state}" if state else "") + f"&search={search}"
                    r = BROWSER.fetch(f"/{kind['ep']}/invoices/{kind['direction']}", q, f"Tìm kiếm ({kind['tab']})")
                    time.sleep(settings.gdt_delay)
                    if not r.ok:
                        raise RuntimeError(f"cổng trả HTTP {r.status}")
                    data = r.json()
                    page = data.get("datas") or []
                    for inv in page:
                        row = summary_row(kind, first, inv)
                        row["nguon"] = "Cổng HĐĐT"
                        rows.append(row)
                    state = data.get("state")
                    if not state or not page:
                        break
            f.write_text(json.dumps(rows, ensure_ascii=False, default=str), encoding="utf-8")
            say(f"Cổng HĐĐT {m:02d}/{y}: {len(rows)} hóa đơn")
            out += rows
        except Exception as e:  # noqa: BLE001
            err = f"Lấy hóa đơn {m:02d}/{y} từ cổng lỗi: {e}"
            if f.exists():
                out += json.loads(f.read_text(encoding="utf-8"))
    return out, err


def merge_invoices(*lists: list[dict[str, Any]]) -> list[dict[str, Any]]:
    """Gộp, bỏ trùng (ưu tiên dòng đã tải đủ trên máy), bỏ HĐ bị thay thế / bị hủy và HĐ 0 đồng."""
    from .reports import tthai_code
    seen: dict[tuple, dict[str, Any]] = {}
    for lst in lists:
        for r in lst:
            k = (r.get("huong"), str(r.get("nbmst") or ""), str(r.get("khhdon") or ""),
                 str(r.get("shdon") or "").lstrip("0"))
            if k not in seen:
                seen[k] = r
    return [r for r in seen.values() if tthai_code(r) not in (4, 6) and float(r.get("tgtttbso") or 0) > 0]


# ------------------------------------------------------------------ gán giao dịch <-> hóa đơn
def _desc(t: dict[str, Any]) -> str:
    return " ".join([t.get("mo_ta", ""), t.get("doi_ung", ""), t.get("tham_chieu", "")])


def _is_partner(inv: dict[str, Any], t: dict[str, Any]) -> bool:
    mst, name = partner(inv)
    dt = t.get("dt")          # đối tượng đã nhận diện từ danh mục công nợ (số TK, MST, mã, tên)
    if dt and not dt.get("ai") and dt.get("diem", 0) >= 85:   # chỉ dùng đối tượng nhận diện chắc chắn
        if mst and dt.get("mst") and mst[:10] == dt["mst"][:10]:
            return True
        if dt.get("ten") and name_sim(name, dt["ten"]) >= 0.9:
            return True
    d = _desc(t)
    if mst and mst in re.sub(r"\D", " ", d).split():
        return True
    if name_sim(name, t.get("doi_ung")) >= 0.85:
        return True
    cn = core_name(name)
    return len(cn) >= 8 and cn in fold(t.get("mo_ta"))


def _subset(cands: list[dict[str, Any]], amount: float, limit: int = 8,
            ref: date | None = None) -> list[dict[str, Any]] | None:
    """Tập hóa đơn (còn nợ) cộng lại đúng số tiền. Ưu tiên ít hóa đơn, rồi HĐ gần ngày thanh toán nhất
    (nhiều tổ hợp cùng tổng – vd HĐ tháng 8 và tháng 11 cùng số tiền – thì chọn HĐ gần ngày trả)."""
    from itertools import combinations
    target = round(amount)
    ref = ref or date.today()

    def gap(i):
        d = _inv_date(i)
        return abs((ref - d).days) if d else 9999

    small = sorted(cands, key=gap)[:30]
    for size in range(1, min(4, limit) + 1):
        best = None
        for combo in combinations(small, size):
            if abs(sum(round(i["_con_lai"]) for i in combo) - target) <= TOL:
                g = sum(gap(i) for i in combo)
                if best is None or g < best[0]:
                    best = (g, combo)
        if best:
            return list(best[1])
    states: dict[int, tuple[int, ...]] = {0: ()}
    for i, inv in enumerate(cands):
        v = round(inv["_con_lai"])
        new = {}
        for s, combo in states.items():
            ns = s + v
            if ns <= target + TOL and len(combo) < limit and ns not in states and ns not in new:
                new[ns] = combo + (i,)
        states.update(new)
        if len(states) > 50000:
            break
    best = None
    for s, combo in states.items():
        if combo and abs(s - target) <= TOL and (best is None or len(combo) < len(best)):
            best = combo
    return [cands[i] for i in best] if best else None


def allocate(invoices: list[dict[str, Any]], txns: list[dict[str, Any]]) -> None:
    """Ghi vào mỗi giao dịch: t['alloc'] = [{'inv', 'so_tien'}], t['cach_khop'] (giải thích).
    Ghi vào mỗi hóa đơn: inv['_con_lai'] (còn phải thu / phải trả), inv['_tt'] (các giao dịch)."""
    for inv in invoices:
        inv["_con_lai"] = float(inv.get("tgtttbso") or 0)
        inv["_tt"] = []
    for t in txns:
        t["alloc"], t["cach_khop"] = [], ""

    def pool(t):
        h = "sold" if t["co"] else "purchase"
        dt = t["ngay"]
        return [i for i in invoices if i.get("huong") == h and i["_con_lai"] > 1
                and (_inv_date(i) is None or _inv_date(i) <= dt + timedelta(days=5))]

    def amt(t):
        return t["co"] or t["no"]

    def left(t):
        return amt(t) - sum(a["so_tien"] for a in t["alloc"])

    def give(t, inv, v, how):
        v = min(v, inv["_con_lai"])
        if v <= 0:
            return
        t["alloc"].append({"inv": inv, "so_tien": v})
        inv["_con_lai"] -= v
        inv["_tt"].append(t)
        if how and how not in t["cach_khop"]:
            t["cach_khop"] = (t["cach_khop"] + "; " + how).strip("; ")

    def round_off(t):
        """Lệch vài đồng do làm tròn khi cộng nhiều HĐ: gộp vào HĐ cuối, không tạo dòng lẻ."""
        r = left(t)
        if t["alloc"] and 0 < r <= TOL:
            t["alloc"][-1]["so_tien"] += r
            t["alloc"][-1]["inv"]["_con_lai"] -= r

    order = sorted(txns, key=lambda t: t["ngay"])

    # 1) Đúng số tiền + có bằng chứng (số HĐ / MST / tên đối tác trong nội dung) -> 1 hóa đơn
    cands = []
    for j, t in enumerate(order):
        for inv in pool(t):
            s, why = _score(inv, t)
            if s >= 70 and "đúng số tiền" in why:
                cands.append((s, -abs(((_inv_date(inv) or t["ngay"]) - t["ngay"]).days), j, id(inv), inv, why))
    cands.sort(key=lambda x: (-x[0], -x[1]))
    for s, _g, j, _k, inv, why in cands:
        t = order[j]
        if t["alloc"] or inv["_con_lai"] < amt(t) - 1:
            continue
        give(t, inv, amt(t), "Khớp 1 HĐ: " + ", ".join(why))

    # 2) Nội dung ghi số hóa đơn (1 hoặc nhiều HĐ) của đúng đối tác
    for t in order:
        if left(t) <= 1:
            continue
        digits = {d.lstrip("0") for d in re.findall(r"\d{2,8}", _desc(t))}
        hit = [i for i in pool(t) if str(i.get("shdon") or "").lstrip("0") in digits and _is_partner(i, t)]
        if not hit:
            hit = [i for i in pool(t) if str(i.get("shdon") or "").lstrip("0") in digits
                   and len(str(i.get("shdon") or "").lstrip("0")) >= 3]
            if sum(i["_con_lai"] for i in hit) > left(t) + TOL * 50:
                hit = []  # số trùng ngẫu nhiên, không cùng đối tác -> bỏ
        hit.sort(key=lambda i: _inv_date(i) or date.min)
        if hit:
            how = (f"Nội dung ghi số HĐ ({len(hit)} HĐ)" if len(hit) > 1 else "Nội dung ghi số HĐ")
            for inv in hit:
                if left(t) <= 1:
                    break
                give(t, inv, left(t), how)

    def recent_first(t, days):
        """HĐ còn nợ của đúng đối tác, ưu tiên HĐ trong `days` ngày trước giao dịch (HĐ quá cũ thường đã trả
        từ kỳ sao kê trước), sắp xếp cũ -> mới."""
        mine = [i for i in pool(t) if _is_partner(i, t)]
        near = [i for i in mine if (_inv_date(i) or t["ngay"]) >= t["ngay"] - timedelta(days=days)]
        return sorted(near, key=lambda i: _inv_date(i) or date.min), sorted(mine, key=lambda i: _inv_date(i) or date.min)

    # 3) Trả gộp nhiều hóa đơn cùng đối tác, cộng lại đúng số tiền (ưu tiên HĐ gần ngày)
    for t in order:
        if t["alloc"]:
            continue
        near, mine = recent_first(t, 120)
        for group in (near, mine):
            combo = _subset(group, amt(t), ref=t["ngay"]) if len(group) >= 2 else None
            if combo and len(combo) >= 2:
                for inv in sorted(combo, key=lambda i: _inv_date(i) or date.min):
                    give(t, inv, inv["_con_lai"], f"Trả gộp {len(combo)} HĐ cùng đối tác, cộng đúng số tiền")
                round_off(t)
                break

    # 3b) Nhiều lần chuyển gần nhau (≤ 7 ngày) cho cùng đối tác, tổng các lần = tổng 1 nhóm HĐ
    #     (vd 2 lần 554.999.999 + 600.000.000 trả cho HĐ 444.000.000 + 710.999.999)
    free = [t for t in order if not t["alloc"]]
    done: set[int] = set()
    for k, t in enumerate(free):
        if id(t) in done:
            continue
        near, mine = recent_first(t, 120)
        if not mine:
            continue
        grp = [t] + [u for u in free[k + 1:] if id(u) not in done and bool(u["co"]) == bool(t["co"])
                     and (u["ngay"] - t["ngay"]).days <= 7 and _is_partner(mine[0], u)]
        if len(grp) < 2:
            continue
        total = sum(amt(u) for u in grp)
        for group in (near, mine):
            combo = _subset(group, total, limit=10, ref=t["ngay"]) if group else None
            if combo:
                invs = sorted(combo, key=lambda i: _inv_date(i) or date.min)
                how = f"{len(grp)} lần chuyển gần nhau trả {len(invs)} HĐ cùng đối tác, tổng khớp {_money(total)}"
                for u in sorted(grp, key=lambda u: u["ngay"]):  # rải lần lượt vào các HĐ, cũ trước
                    for inv in invs:
                        if left(u) <= 1:
                            break
                        give(u, inv, left(u), how)
                    done.add(id(u))
                round_off(grp[-1])
                break

    # 4) Chỉ khớp số tiền (không có bằng chứng khác): 1–1, ưu tiên ngày gần nhất
    cands, n_same = [], {}
    for j, t in enumerate(order):
        if t["alloc"]:
            continue
        for inv in pool(t):
            if abs(inv["_con_lai"] - amt(t)) <= 1:
                gap = abs(((_inv_date(inv) or t["ngay"]) - t["ngay"]).days)
                if gap <= 120:
                    cands.append((gap, j, id(inv), inv))
                    n_same[j] = n_same.get(j, 0) + 1
    cands.sort(key=lambda x: (x[0], x[1]))
    for gap, j, _k, inv in cands:
        t = order[j]
        if t["alloc"] or inv["_con_lai"] < amt(t) - 1:
            continue
        sure = gap <= 3 and n_same.get(j) == 1  # duy nhất 1 HĐ đúng số tiền, lập trong 3 ngày
        give(t, inv, amt(t), "Khớp số tiền, HĐ duy nhất cùng ngày" if sure
             else "Chỉ khớp số tiền (cần kiểm tra lại đối tác)")

    # 5) Cùng đối tác nhưng không khớp tổng: trừ dần vào HĐ gần đây, cũ trước (thanh toán 1 phần / trả trước).
    #    Chỉ xét HĐ trong 90 ngày trước giao dịch – HĐ cũ hơn thường đã thanh toán ở kỳ sao kê trước.
    for t in order:
        if left(t) <= 1:
            continue
        mine, _all = recent_first(t, 90)
        for inv in mine:
            if left(t) <= 1:
                break
            part = left(t) < inv["_con_lai"] - 1
            give(t, inv, left(t), "Thanh toán 1 phần HĐ" if part else "Trừ dần công nợ HĐ cũ nhất của đối tác")
        if t["alloc"] and left(t) > 1:
            t["cach_khop"] += f"; thừa {_money(left(t))} so với HĐ còn nợ (trả trước / ứng trước)"


def describe(t: dict[str, Any]) -> str:
    """Câu mô tả rõ: nhận tiền từ HĐ nào / thanh toán cho HĐ nào."""
    if not t.get("alloc"):
        return ""
    verb = "Thu tiền HĐ bán ra" if t["co"] else "Thanh toán HĐ mua vào"
    parts = []
    for a in t["alloc"]:
        inv = a["inv"]
        full = abs(a["so_tien"] - float(inv.get("tgtttbso") or 0)) <= 1
        parts.append(f"{inv_label(inv)}: {_money(a['so_tien'])}" + ("" if full else
                     f" (HĐ {_money(float(inv.get('tgtttbso') or 0))})"))
    return f"{verb} " + "; ".join(parts)


# ------------------------------------------------------------------ hạch toán Nợ / Có
def accounts(che_do: str) -> dict[str, str]:
    tt133 = (che_do or "").upper() == "TT133"
    return {"ngan_hang": "1121", "tien_mat": "1111", "phai_thu": "131", "phai_tra": "331", "luong": "334",
            "phi_nh": "6422" if tt133 else "6427", "lai": "515", "bhxh": "3383", "tam_ung": "141", "vay": "341",
            "thue_gtgt": "33311" if not tt133 else "33311", "thue_tndn": "3334", "thue_tncn": "3335",
            "thue_khac": "3339", "phai_tra_khac": "3388", "phai_thu_khac": "1388", "von": "4111"}


_RULES_OUT = [  # (từ khóa trong nội dung đã bỏ dấu, khóa tài khoản, diễn giải, cần kiểm tra?)
    (r"\bung (luong|tien)\b|\bung\b.{0,25}\bluong\b|tam ung", "tam_ung", "Tạm ứng cho nhân viên", True),
    (r"\bluong\b|\bsalary\b|thuong tet", "luong", "Trả lương nhân viên", False),
    (r"\bbhxh\b|bao hiem xa hoi|\bbhyt\b|bao hiem y te", "bhxh", "Nộp bảo hiểm", False),
    (r"thue gtgt|thue gia tri gia tang|\bvat\b", "thue_gtgt", "Nộp thuế GTGT", True),
    (r"thue tndn|thu nhap doanh nghiep", "thue_tndn", "Nộp thuế TNDN", True),
    (r"thue tncn|thu nhap ca nhan", "thue_tncn", "Nộp thuế TNCN", True),
    (r"\bnsnn\b|ngan sach|nop thue|le phi|mon bai", "thue_khac", "Nộp thuế / lệ phí NSNN", True),
    (r"\bphi\b.*\b(chuyen|dich vu|sms|quan ly|duy tri|thuong nien)\b|\bfee\b|phi ngan hang", "phi_nh",
     "Phí ngân hàng", False),
    (r"rut tien mat|\brut tien\b|\batm\b", "tien_mat", "Rút tiền gửi về quỹ tiền mặt", False),
    (r"tam ung", "tam_ung", "Chi tạm ứng", True),
    (r"tra no vay|tra goc|tra vay|\bvay\b", "vay", "Trả nợ vay", True),
    (r"tra lai ck nham|chuyen nham|ck nham|hoan tra|tra lai tien", "phai_tra_khac",
     "Trả lại tiền chuyển nhầm", True),
]
_RULES_IN = [
    (r"lai tien gui|tra lai tien gui|\binterest\b|lai nhap goc", "lai", "Lãi tiền gửi ngân hàng", False),
    (r"nop tien mat|nop tm\b|\bnop tien\b(?!.*\b(chuyen|ck|mbvcb|ibft)\b)", "tien_mat",
     "Nộp tiền mặt vào tài khoản", True),
    (r"gop von|von gop", "von", "Nhận vốn góp", True),
    (r"hoan tien|tra lai ck nham|chuyen nham|ck nham", "phai_thu_khac", "Nhận lại tiền chuyển nhầm / hoàn tiền", True),
    (r"giai ngan|nhan no vay|khoan vay", "vay", "Nhận tiền vay", True),
]


def post(t: dict[str, Any], acc: dict[str, str], own_name: str = "") -> dict[str, Any]:
    """Định khoản 1 giao dịch -> chứng từ có 1..n dòng (Nợ, Có, số tiền, đối tượng, diễn giải)."""
    inflow = bool(t["co"])
    amount = t["co"] or t["no"]
    bank = acc["ngan_hang"]
    lines, review = [], []
    for a in t.get("alloc", []):
        inv = a["inv"]
        mst, name = partner(inv)
        if inflow:
            lines.append({"no": bank, "co": acc["phai_thu"], "so_tien": a["so_tien"], "mst": mst, "ten": name,
                          "dien_giai": f"Thu tiền HĐ {inv_label(inv)} – {name}"[:250]})
        else:
            lines.append({"no": acc["phai_tra"], "co": bank, "so_tien": a["so_tien"], "mst": mst, "ten": name,
                          "dien_giai": f"Thanh toán HĐ {inv_label(inv)} – {name}"[:250]})
    rest = amount - sum(x["so_tien"] for x in lines)
    if rest > 1:
        f = fold(" ".join([t.get("mo_ta", ""), t.get("doi_ung", "")]))
        same_owner = own_name and name_sim(own_name, t.get("doi_ung")) >= 0.85
        key, text, chk = None, "", True
        company_peer = re.search(r"\b(cty|cong ty|tnhh|co phan|ctcp|dntn)\b", fold(t.get("doi_ung")))
        for pat, k, txt, c in (_RULES_IN if inflow else _RULES_OUT):
            if k == "phi_nh" and company_peer:
                continue        # "tt tiền phí dịch vụ" chuyển cho 1 công ty: trả nhà cung cấp, không phải phí ngân hàng
            if re.search(pat, f):
                key, text, chk = k, txt, c
                t["_rule"] = txt
                break
        ai = t.get("_ai")
        if key is None and ai and ai["key"] not in ("phai_thu", "phai_tra", "khac", "noi_bo"):
            key, text, chk = ai["key"], f"{ai['muc_dich']} (AI gợi ý)", True
        elif key is None and ai and ai["key"] == "noi_bo":
            key, text, chk = ("phai_thu_khac" if inflow else "phai_tra_khac"), f"{ai['muc_dich']} (AI gợi ý)", True
        if key is None and same_owner:
            key, text, chk = ("phai_thu_khac" if inflow else "phai_tra_khac"), "Chuyển tiền nội bộ giữa các TK", True
        first = lines[0] if lines else None
        if key is None:
            # không rõ: tiền vào -> khách hàng ứng trước / chưa có HĐ; tiền ra -> trả trước người bán
            key = "phai_thu" if inflow else "phai_tra"
            text = ("Thu tiền khách hàng (chưa xác định HĐ)" if inflow else "Trả tiền người bán (chưa xác định HĐ)")
            if t.get("alloc"):
                text = "Khách hàng trả trước / trả thừa" if inflow else "Trả trước cho người bán"
            chk = True
        mst = first["mst"] if first and key in ("phai_thu", "phai_tra") else ""
        ten = (first["ten"] if first and key in ("phai_thu", "phai_tra") else "") or t.get("doi_ung", "")
        dt = t.get("dt")
        if not first and dt and key in ("phai_thu", "phai_tra", "phai_thu_khac", "phai_tra_khac"):
            mst, ten = dt.get("mst") or "", dt.get("ten") or ten     # đối tượng nhận diện từ danh mục công nợ
            if not t.get("_rule") and not ai:
                text = ("Thu tiền khách hàng" if inflow else "Trả tiền nhà cung cấp") + f" {dt['ten']}"
            if ai and ai["key"] in ("phai_thu", "phai_tra"):
                text = f"{ai['muc_dich']} (AI gợi ý)"
        lines.append({"no": bank if inflow else acc[key], "co": acc[key] if inflow else bank, "so_tien": rest,
                      "mst": mst, "ten": ten,
                      "dien_giai": f"{text} – {t.get('doi_ung') or ''} – {t.get('mo_ta') or ''}".strip(" –")[:250]})
        if chk:
            review.append(text)
    if "Chỉ khớp số tiền" in (t.get("cach_khop") or ""):
        review.append("Chỉ khớp số tiền – kiểm tra lại đối tác")
    main = lines[0]
    return {"loai": "thu" if inflow else "chi", "ngay": t["ngay"], "mst": main["mst"], "ten": main["ten"] or t.get("doi_ung", ""),
            "ly_do": (describe(t) or main["dien_giai"])[:250], "lines": lines, "kiem_tra": "; ".join(review),
            "tx": t}


def build_vouchers(txns: list[dict[str, Any]], che_do: str, own_name: str, start_no: int = 1) -> list[dict[str, Any]]:
    acc = accounts(che_do)
    out, n_thu, n_chi = [], start_no, start_no
    for t in sorted(txns, key=lambda t: (t["ngay"], t.get("tham_chieu") or "")):
        v = post(t, acc, own_name)
        if v["loai"] == "thu":
            v["so_ct"] = f"NTTK{n_thu:05d}"
            n_thu += 1
        else:
            v["so_ct"] = f"UNC{n_chi:05d}"
            n_chi += 1
        out.append(v)
    return out


# ------------------------------------------------------------------ mã đối tượng
def load_codes(folder: Path) -> dict[str, str]:
    """Bảng mã đối tượng MST -> Mã trong phần mềm kế toán (file MauNhapKhau/MaDoiTuong.xlsx, 2 cột MST | Mã).
    Không có bảng thì dùng MST làm mã đối tượng."""
    f = folder / "MaDoiTuong.xlsx"
    out: dict[str, str] = {}
    if not f.exists():
        return out
    from openpyxl import load_workbook
    ws = load_workbook(f, read_only=True, data_only=True).active
    for r in ws.iter_rows(values_only=True):
        if r and len(r) >= 2 and r[0] and r[1] and re.fullmatch(r"[\d-]{9,14}", str(r[0]).strip()):
            out[str(r[0]).strip()] = str(r[1]).strip()
    return out


# ------------------------------------------------------------------ xuất file nhập khẩu
# Bộ cột mặc định theo tên cột trên file mẫu nhập khẩu của MISA SME (Tệp > Nhập khẩu từ Excel > Ngân hàng).
# Nếu đặt file mẫu tải từ MISA vào thư mục MauNhapKhau, phần mềm đọc dòng tiêu đề của file mẫu và điền theo đó.
MISA_THU = ["Hiển thị trên sổ", "Ngày hạch toán (*)", "Ngày chứng từ (*)", "Số chứng từ (*)", "Mã đối tượng",
            "Tên đối tượng", "Địa chỉ", "Nộp vào TK", "Mở tại ngân hàng", "Lý do thu", "Diễn giải lý do thu",
            "Loại tiền", "Tỷ giá", "Diễn giải", "TK Nợ (*)", "TK Có (*)", "Số tiền", "Số tiền quy đổi", "Đối tượng"]
MISA_CHI = ["Hiển thị trên sổ", "Phương thức thanh toán", "Ngày hạch toán (*)", "Ngày chứng từ (*)",
            "Số chứng từ (*)", "Mã đối tượng", "Tên đối tượng", "Địa chỉ", "Tài khoản chi", "Ngân hàng chi",
            "Tài khoản nhận", "Ngân hàng nhận", "Lý do chi", "Nội dung thanh toán", "Loại tiền", "Tỷ giá",
            "Diễn giải", "TK Nợ (*)", "TK Có (*)", "Số tiền", "Số tiền quy đổi", "Đối tượng"]
EASY_COLS = ["Loại chứng từ", "Ngày chứng từ", "Ngày hạch toán", "Số chứng từ", "Mã đối tượng", "Tên đối tượng",
             "Tài khoản ngân hàng", "Tên ngân hàng", "Diễn giải chung", "Diễn giải", "TK Nợ", "TK Có", "Số tiền",
             "Đối tượng Nợ", "Đối tượng Có", "Số TK đối ứng", "Ngân hàng đối ứng", "Ghi chú kiểm tra"]


def _key(h: Any) -> str:
    return re.sub(r"[^a-z0-9]", "", fold(h))


def _value(h: str, v: dict[str, Any], ln: dict[str, Any], info: dict[str, str], code: str) -> Any:
    k = _key(h)
    t = v["tx"]
    thu = v["loai"] == "thu"
    table = [
        ("hienthitrenso", None),
        ("phuongthucthanhtoan", None if thu else "Ủy nhiệm chi"),
        ("loaichungtu", "Báo có" if thu else "Báo nợ"),
        ("ngayhachtoan", v["ngay"]), ("ngaychungtu", v["ngay"]), ("sochungtu", v["so_ct"]),
        ("doituongno", code if ln["no"] in ("131", "331", "1388", "3388", "141") else None),
        ("doituongco", code if ln["co"] in ("131", "331", "1388", "3388", "141") else None),
        ("madoituong", code), ("tendoituong", v["ten"]), ("diachi", None),
        ("noptk", info.get("so_tk")), ("nopvaotk", info.get("so_tk")), ("taikhoanchi", info.get("so_tk")), ("taikhoannganhang", info.get("so_tk")),
        ("sotaikhoan", info.get("so_tk")),
        ("motainganhang", info.get("ngan_hang")), ("nganhangchi", info.get("ngan_hang")),
        ("tennganhang", info.get("ngan_hang")),
        ("taikhoannhan", t.get("tk_doi_ung")), ("sotkdoiung", t.get("tk_doi_ung")),
        ("nganhangnhan", t.get("nh_doi_ung")), ("nganhangdoiung", t.get("nh_doi_ung")),
        ("lydothu", v["ly_do"]), ("lydochi", v["ly_do"]), ("diengiailydo", v["ly_do"]),
        ("noidungthanhtoan", v["ly_do"]), ("diengiaichung", v["ly_do"]),
        ("loaitien", "VND"), ("tygia", 1),
        ("diengiai", ln["dien_giai"]),
        ("tkno", ln["no"]), ("tkco", ln["co"]),
        ("sotienquydoi", ln["so_tien"]), ("sotien", ln["so_tien"]), ("thanhtien", ln["so_tien"]),
        ("doituong", code),
        ("ghichukiemtra", v.get("kiem_tra")),
    ]
    for prefix, val in table:
        if k.startswith(prefix):
            return val
    return None


def _template_header(path: Path) -> tuple[list[str], int] | None:
    """Dòng tiêu đề trong file mẫu của phần mềm kế toán (dòng có 'Số chứng từ' và 'Số tiền')."""
    from openpyxl import load_workbook
    ws = load_workbook(path, read_only=True).active
    for i, r in enumerate(ws.iter_rows(max_row=15, values_only=True), 1):
        keys = [_key(c) for c in r if c]
        if any(k.startswith("sochungtu") for k in keys) and any(k.startswith("sotien") for k in keys):
            return [str(c) if c is not None else "" for c in r], i
    return None


def _write(dest: Path, header: list[str], vouchers: list[dict[str, Any]], info: dict[str, str],
           codes: dict[str, str], template: Path | None, head_row: int = 1) -> Path:
    from openpyxl import Workbook, load_workbook
    from openpyxl.styles import Font, PatternFill

    if template:
        wb = load_workbook(template)
        ws = wb.active
        r = head_row + 1
        if ws.max_row > head_row:  # xóa dòng ví dụ trong file mẫu
            ws.delete_rows(head_row + 1, ws.max_row - head_row)
    else:
        wb = Workbook()
        ws = wb.active
        ws.title = "Chung tu"
        for c, h in enumerate(header, 1):
            cell = ws.cell(1, c, h)
            cell.font = Font(bold=True)
            cell.fill = PatternFill("solid", fgColor="DDEBF7")
            ws.column_dimensions[cell.column_letter].width = max(12, min(45, len(h) + 4))
        r = 2
    warn = PatternFill("solid", fgColor="FFF2CC")
    for v in vouchers:
        for ln in v["lines"]:
            code = ln.get("ma") or codes.get(ln["mst"]) or ln["mst"] or ""
            for c, h in enumerate(header, 1):
                val = _value(h, v, ln, info, code) if h else None
                if val is None:
                    continue
                cell = ws.cell(r, c, val)
                if isinstance(val, date):
                    cell.number_format = "dd/mm/yyyy"
                elif isinstance(val, float):
                    cell.number_format = "#,##0"
                if v.get("kiem_tra"):
                    cell.fill = warn
            r += 1
    dest.parent.mkdir(parents=True, exist_ok=True)
    wb.save(dest)
    return dest


def _find_template(folder: Path, *words: str) -> Path | None:
    if not folder.is_dir():
        return None
    for f in sorted(folder.glob("*.xlsx")):
        n = _key(f.stem)
        if all(w in n for w in words):
            return f
    return None


def export_misa(folder: Path, out_dir: Path, tag: str, vouchers, info, codes) -> list[Path]:
    out = []
    for loai, words, default, name in (("thu", ("misa", "thu"), MISA_THU, "MISA_ThuTienGui"),
                                       ("chi", ("misa", "chi"), MISA_CHI, "MISA_ChiTienGui")):
        vs = [v for v in vouchers if v["loai"] == loai]
        if not vs:
            continue
        tpl = _find_template(folder, *words)
        head = _template_header(tpl) if tpl else None
        header, row = (head if head else (default, 1))
        out.append(_write(out_dir / f"{name}_{tag}.xlsx", header, vs, info, codes, tpl if head else None, row))
    return out


def export_easybooks(folder: Path, out_dir: Path, tag: str, vouchers, info, codes) -> list[Path]:
    out = []
    for loai, words, name in (("thu", ("easybook", "co"), "EasyBooks_BaoCo"),
                              ("chi", ("easybook", "no"), "EasyBooks_BaoNo")):
        vs = [v for v in vouchers if v["loai"] == loai]
        if not vs:
            continue
        tpl = _find_template(folder, *words)
        head = _template_header(tpl) if tpl else None
        header, row = (head if head else (EASY_COLS, 1))
        out.append(_write(out_dir / f"{name}_{tag}.xlsx", header, vs, info, codes, tpl if head else None, row))
    return out


# ------------------------------------------------------------------ file đối chiếu
def write_report(dest: Path, txns, invoices, vouchers, info, errors: list[str], note: str = "") -> Path:
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font, PatternFill

    from .reports import _header
    bold = Font(bold=True)
    warn = PatternFill("solid", fgColor="FFF2CC")
    ok = PatternFill("solid", fgColor="E2EFDA")
    wb = Workbook()
    ws = wb.active
    ws.title = "Tom tat"
    ws.column_dimensions["A"].width = 55
    ws.column_dimensions["B"].width = 20
    matched = [t for t in txns if t.get("alloc")]
    used = [i for i in invoices if i["_tt"]]
    lines = [
        ("ĐỐI CHIẾU SAO KÊ – HÓA ĐƠN – HẠCH TOÁN", None),
        ("Tài khoản", f"{info.get('so_tk') or ''} {info.get('ngan_hang') or ''}".strip()),
        ("Chủ tài khoản", info.get("chu_tk")),
        ("Số giao dịch", len(txns)),
        ("Tổng tiền vào (ghi có)", sum(t["co"] for t in txns)),
        ("Tổng tiền ra (ghi nợ)", sum(t["no"] for t in txns)),
        ("Giao dịch đã gán được hóa đơn", len(matched)),
        ("Giao dịch chưa có hóa đơn", len(txns) - len(matched)),
        ("Hóa đơn được thanh toán trong sao kê", len(used)),
        ("Hóa đơn dùng để đối chiếu (máy + cổng HĐĐT)", len(invoices)),
        ("Chứng từ cần kiểm tra lại (tô vàng)", sum(1 for v in vouchers if v.get("kiem_tra"))),
    ]
    for i, (a, b) in enumerate(lines, 1):
        ws.cell(i, 1, a).font = bold if i == 1 else Font()
        if b is not None:
            c = ws.cell(i, 2, b)
            if isinstance(b, float):
                c.number_format = "#,##0"
    k = len(lines) + 2
    for e in ([note] if note else []) + errors:
        ws.cell(k, 1, e)
        k += 1

    # Theo giao dịch
    w2 = wb.create_sheet("Theo giao dich")
    _header(w2, [("Ngày GD", 11), ("Tiền vào", 15), ("Tiền ra", 15), ("Diễn giải ngân hàng", 40),
                 ("Tên đối ứng", 30), ("Số TK đối ứng", 17), ("Ngân hàng đối ứng", 22), ("Số bút toán", 18),
                 ("NHẬN TỪ HĐ / THANH TOÁN CHO HĐ", 60), ("Đối tác trên HĐ", 34), ("MST đối tác", 13),
                 ("Cách khớp", 42), ("Chưa gán HĐ", 14), ("Số CT", 11), ("Hạch toán", 30), ("Cần kiểm tra", 30)])
    vmap = {id(v["tx"]): v for v in vouchers}
    for r, t in enumerate(sorted(txns, key=lambda t: t["ngay"]), 2):
        v = vmap.get(id(t), {})
        partners = {partner(a["inv"]) for a in t.get("alloc", [])}
        rest = (t["co"] or t["no"]) - sum(a["so_tien"] for a in t.get("alloc", []))
        hach = "; ".join(f"Nợ {ln['no']}/Có {ln['co']}: {_money(ln['so_tien'])}" for ln in v.get("lines", []))
        vals = [t["ngay"], t["co"] or None, t["no"] or None, t.get("mo_ta"), t.get("doi_ung"), t.get("tk_doi_ung"),
                t.get("nh_doi_ung"), t.get("tham_chieu"), describe(t) or "— chưa xác định được hóa đơn —",
                "; ".join(p[1] for p in partners), "; ".join(p[0] for p in partners if p[0]), t.get("cach_khop"),
                rest if rest > 1 else None, v.get("so_ct"), hach, v.get("kiem_tra")]
        for c, val in enumerate(vals, 1):
            cell = w2.cell(r, c, val)
            if c in (2, 3, 13):
                cell.number_format = "#,##0"
            if c == 1:
                cell.number_format = "dd/mm/yyyy"
            if c == 9:
                cell.alignment = Alignment(wrap_text=True, vertical="top")
                cell.fill = ok if t.get("alloc") else warn
    w2.freeze_panes = "B2"

    # Theo hóa đơn
    w3 = wb.create_sheet("Theo hoa don")
    _header(w3, [("Loại", 24), ("Ngày HĐ", 11), ("Ký hiệu", 10), ("Số HĐ", 9), ("MST đối tác", 13), ("Đối tác", 36),
                 ("Tổng thanh toán", 15), ("Đã thu / đã trả", 15), ("Còn lại", 15), ("Các lần thanh toán", 55),
                 ("Nguồn", 12)])
    rows = sorted(used, key=lambda i: (i.get("huong") or "", str(i.get("ngay_lap"))))
    for r, inv in enumerate(rows, 2):
        mst, name = partner(inv)
        paid = float(inv.get("tgtttbso") or 0) - inv["_con_lai"]
        pays = []
        for t in inv["_tt"]:
            a = next(x["so_tien"] for x in t["alloc"] if x["inv"] is inv)
            pays.append(f"{t['ngay']:%d/%m/%Y} {_money(a)} ({t.get('tham_chieu') or ''})")
        vals = [inv.get("loai"), inv.get("ngay_lap"), inv.get("khhdon"), inv.get("shdon"), mst, name,
                float(inv.get("tgtttbso") or 0), paid, inv["_con_lai"] if inv["_con_lai"] > 1 else 0, "; ".join(pays),
                inv.get("nguon") or "Đã tải"]
        for c, val in enumerate(vals, 1):
            cell = w3.cell(r, c, val)
            if c in (7, 8, 9):
                cell.number_format = "#,##0"

    # Hóa đơn còn nợ (trong kỳ sao kê) – để theo dõi công nợ
    w4 = wb.create_sheet("HD con no")
    _header(w4, [("Loại", 24), ("Ngày HĐ", 11), ("Ký hiệu", 10), ("Số HĐ", 9), ("MST đối tác", 13), ("Đối tác", 36),
                 ("Tổng thanh toán", 15), ("Còn phải thu / trả", 16), ("Nguồn", 12)])
    d1 = min((t["ngay"] for t in txns), default=date.today())
    d2 = max((t["ngay"] for t in txns), default=date.today())
    open_inv = [i for i in invoices if i["_con_lai"] > 1 and (_inv_date(i) or d1) <= d2
                and (_inv_date(i) or d2) >= d1 - timedelta(days=62)]
    for r, inv in enumerate(sorted(open_inv, key=lambda i: str(i.get("ngay_lap"))), 2):
        mst, name = partner(inv)
        for c, val in enumerate([inv.get("loai"), inv.get("ngay_lap"), inv.get("khhdon"), inv.get("shdon"), mst, name,
                                 float(inv.get("tgtttbso") or 0), inv["_con_lai"], inv.get("nguon") or "Đã tải"], 1):
            cell = w4.cell(r, c, val)
            if c in (7, 8):
                cell.number_format = "#,##0"

    # Chứng từ hạch toán (xem trước khi nhập phần mềm)
    w5 = wb.create_sheet("Hach toan")
    _header(w5, [("Số CT", 11), ("Ngày", 11), ("Loại", 9), ("Mã/MST đối tượng", 14), ("Tên đối tượng", 34),
                 ("Diễn giải", 60), ("TK Nợ", 8), ("TK Có", 8), ("Số tiền", 15), ("Cần kiểm tra", 36)])
    r = 2
    for v in vouchers:
        for ln in v["lines"]:
            for c, val in enumerate([v["so_ct"], v["ngay"], "Báo có" if v["loai"] == "thu" else "Báo nợ",
                                     ln["mst"], ln["ten"] or v["ten"], ln["dien_giai"], ln["no"], ln["co"],
                                     ln["so_tien"], v.get("kiem_tra")], 1):
                cell = w5.cell(r, c, val)
                if c == 2:
                    cell.number_format = "dd/mm/yyyy"
                if c == 9:
                    cell.number_format = "#,##0"
                if v.get("kiem_tra"):
                    cell.fill = warn
            r += 1
    dest.parent.mkdir(parents=True, exist_ok=True)
    wb.save(dest)
    return dest


_ = _tokens  # giữ import cho tương thích
