"""Đối chiếu SAO KÊ GỐC của ngân hàng (PDF / Excel) với SỔ TIỀN GỬI đã nhập trên phần mềm kế toán
(MISA, EasyBooks, Fast, Bravo... xuất ra Excel / PDF: Sổ tiền gửi ngân hàng, Sổ chi tiết TK 112, sổ nhật ký thu/chi,
hoặc chính file nhập khẩu Thu/Chi tiền gửi).

Cách đối chiếu (làm riêng cho tiền VÀO và tiền RA; trên sổ: Thu / Gửi vào / Phát sinh Nợ 112 = tiền vào):
  1. Khớp 1-1: cùng số tiền, ưu tiên có số chứng từ / số tham chiếu trùng, rồi ngày gần nhất, rồi nội dung giống.
     Khác ngày vẫn khớp nhưng ghi rõ "lệch N ngày".
  2. Khớp gộp: 1 dòng sao kê = tổng 2-4 dòng sổ (hoặc ngược lại) trong ±3 ngày (vd tách tiền hàng + phí).
  3. Sai chiều: sao kê là tiền vào nhưng sổ ghi tiền ra (hoặc ngược lại) cùng số tiền.
  4. Lệch số tiền: cùng ngày (±3), nội dung / số tham chiếu giống nhưng số tiền khác -> ghi chênh lệch.
  5. Còn lại: chỉ có trên sao kê (chưa hạch toán) / chỉ có trên sổ (không có trên sao kê).
Dòng sổ nằm ngoài kỳ của sao kê được để riêng, không tính là lệch. So thêm tổng phát sinh và số dư đầu/cuối kỳ.
"""
from __future__ import annotations

import re
from datetime import date, timedelta
from difflib import SequenceMatcher
from itertools import combinations
from pathlib import Path
from typing import Any

from .bank import _parse_date
from .docreader import fold, vn_num
from .gdt.xml_parser import norm

TOL = 0.5  # đồng

BOOK_COLS: list[tuple[str, tuple[str, ...]]] = [
    ("ngay", ("ngayhachtoan", "ngayghiso", "ngaythangghiso", "ngaychungtu", "ngayct", "ngaygiaodich", "ngaythang",
              "ngay", "date")),
    ("so_ct", ("sochungtu", "sohieuchungtu", "soct", "sohieu", "chungtu", "voucherno", "sophieu")),
    ("mo_ta", ("diengiai", "noidung", "mota", "lydo", "description")),
    ("vao", ("sotienthu", "thutien", "guivao", "tienvao", "phatsinhno", "psno", "sotienno", "no", "thu", "tang", "debit")),
    ("ra", ("sotienchi", "chitien", "rutra", "tienra", "phatsinhco", "psco", "sotienco", "co", "chi", "giam", "credit")),
    ("so_du", ("sotonquy", "tonquy", "conlai", "sodu", "ton", "balance")),
    ("doi_tuong", ("tendoituong", "doituong", "tenkhachhang", "tennhacungcap", "nguoinop", "nguoinhan")),
    ("tk_du", ("tkdoiung", "taikhoandoiung", "tkdu")),
    ("so_tien", ("sotien", "thanhtien", "sotienquydoi", "amount")),
    ("loai", ("loaichungtu", "loaict", "loai")),
]
SHORT = {"no", "co", "thu", "chi", "tang", "giam", "ton", "ngay", "loai", "soct", "date", "psno", "psco"}
# dòng tổng / số dư (chỉ xét phần ĐẦU ô diễn giải – tránh nhầm "CONG TY" thành "Cộng")
SUMMARY = re.compile(r"^\s*(cong|tong cong|tong|so du|du dau|du cuoi|dau ky|cuoi ky|luy ke|so ton|ton dau|ton cuoi)\b")
IN_WORDS = re.compile(r"\b(thu|bao co|nop tien|gui vao|tien vao)\b")
OUT_WORDS = re.compile(r"\b(chi|bao no|uy nhiem chi|rut|tien ra)\b")
IN_CODE = re.compile(r"^(nttk|ntk|bc|pt|ct?bc|gbc)[\W_]*\d", re.I)   # số chứng từ thu tiền gửi
OUT_CODE = re.compile(r"^(unc|ctk|bn|pc|ct?bn|gbn)[\W_]*\d", re.I)  # số chứng từ chi tiền gửi


# ------------------------------------------------------------------ đọc sổ
def _map(names: list[str]) -> dict[str, int]:
    m: dict[str, int] = {}
    for field, keys in BOOK_COLS:
        for k in keys:
            for i, h in enumerate(names):
                if h and i not in m.values() and (h == k or (k not in SHORT and k in h)):
                    m[field] = i
                    break
            if field in m:
                break
    return m


def _header(rows: list[list[Any]]) -> tuple[int, dict[str, int]] | None:
    """Tìm dòng tiêu đề (kể cả tiêu đề 2 tầng: 'Số phát sinh' gộp ô trên 'Nợ | Có')."""
    best = None
    for i in range(min(len(rows), 40)):
        one = [norm(str(c or "")) for c in rows[i]]
        cands = [(i, _map(one))]
        if i + 1 < len(rows):
            two = [norm(str(c or "")) for c in rows[i + 1]]
            parent, filled = "", []
            for c in range(max(len(one), len(two))):
                p = one[c] if c < len(one) else ""
                parent = p or parent
                sub = two[c] if c < len(two) else ""
                filled.append((parent + sub) if sub else p)
            # tầng dưới là "no/co/thu/chi" -> ghép với tầng trên ("phatsinh"+"no" = "phatsinhno")
            mixed = [s if (c < len(two) and two[c]) else one[c] if c < len(one) else "" for c, s in enumerate(filled)]
            m2 = _map(mixed)
            m3 = _map(two)
            for k, v in m3.items():
                if k in ("vao", "ra") and k not in m2:
                    m2[k] = v
            cands.append((i + 1, m2))
        for hr, m in cands:
            ok = "ngay" in m and (("vao" in m and "ra" in m) or "so_tien" in m)
            sc = len(m) + (5 if "vao" in m and "ra" in m else 0)  # có đủ cột Thu/Chi tốt hơn 1 cột Số tiền
            if ok and (best is None or sc > best[2]):
                best = (hr, m, sc)
    return (best[0], best[1]) if best else None


def _direction(text: str) -> str | None:
    t = fold(text or "").strip()
    if IN_CODE.search(t):
        return "vao"
    if OUT_CODE.search(t):
        return "ra"
    if IN_WORDS.search(t):
        return "vao"
    if OUT_WORDS.search(t):
        return "ra"
    return None


def _entries_from_rows(rows: list[list[Any]], src: str) -> tuple[list[dict[str, Any]], dict[str, Any]]:
    h = _header(rows)
    if not h:
        return [], {}
    hr, mp = h
    info: dict[str, Any] = {}
    out: list[dict[str, Any]] = []
    sheet_dir = _direction(src)
    for r_i, r in enumerate(rows[hr + 1:], hr + 2):
        get = lambda f: r[mp[f]] if f in mp and mp[f] < len(r) else None  # noqa: E731
        txt_all = fold(" ".join(str(x) for x in r if x not in (None, "")))
        lead = fold(str(get("mo_ta") or "")) or fold(next((str(x) for x in r if isinstance(x, str) and x.strip()), ""))
        d = _parse_date(get("ngay"))
        if not d or SUMMARY.search(lead):
            if "dau ky" in txt_all or "so du dau" in txt_all:
                v = vn_num(get("so_du"))
                if v is not None:
                    info["dau_ky"] = v
            if "cuoi ky" in txt_all or "so du cuoi" in txt_all:
                v = vn_num(get("so_du"))
                if v is not None:
                    info["cuoi_ky"] = v
            continue
        vao, ra = vn_num(get("vao")) or 0.0, vn_num(get("ra")) or 0.0
        if "so_tien" in mp and not ("vao" in mp and "ra" in mp):
            amt = vn_num(get("so_tien")) or 0.0
            dr = _direction(str(get("loai") or "")) or _direction(str(get("so_ct") or "")) or sheet_dir
            if amt < 0:
                dr = "ra" if dr != "ra" else "vao"
            vao, ra = (abs(amt), 0.0) if dr != "ra" else (0.0, abs(amt))
        if not vao and not ra:
            continue
        out.append({"ngay": d, "so_ct": str(get("so_ct") or "").strip(), "mo_ta": str(get("mo_ta") or "").strip(),
                    "vao": abs(vao), "ra": abs(ra), "so_du": vn_num(get("so_du")),
                    "doi_tuong": str(get("doi_tuong") or "").strip(), "tk_du": str(get("tk_du") or "").strip(),
                    "nguon": f"{src} – dòng {r_i}"})
    if out and out[-1].get("so_du") is not None:
        info.setdefault("cuoi_ky_cot", out[-1]["so_du"])
    return out, info


def read_book(name: str, data: bytes) -> tuple[list[dict[str, Any]], dict[str, Any], str | None]:
    """-> (dòng sổ, {dau_ky, cuoi_ky}, lỗi)"""
    tables: list[tuple[str, list[list[Any]]]] = []
    try:
        if data[:5] == b"%PDF-":
            from .bank import _pdf_grid_rows, _pdf_rows
            rows = _pdf_grid_rows(data)
            if not _header(rows):
                rows = _pdf_rows(data)
            tables.append((name, rows))
        else:
            from .misa_import import load_any_workbook
            wb = load_any_workbook(data)
            for ws in wb.worksheets:
                tables.append((f"{name} [{ws.title}]", [list(r) for r in ws.iter_rows(values_only=True)]))
    except Exception as e:  # noqa: BLE001
        return [], {}, f"không đọc được file: {e}"
    entries: list[dict[str, Any]] = []
    info: dict[str, Any] = {}
    for src, rows in tables:
        e, i = _entries_from_rows(rows, src)
        entries += e
        for k, v in i.items():
            info.setdefault(k, v)
    if not entries:
        return [], info, ("không thấy bảng sổ (cần các cột Ngày, Diễn giải, Thu/Chi hoặc Phát sinh Nợ/Có "
                          "hoặc Số tiền)")
    return entries, info, None


# ------------------------------------------------------------------ đối chiếu
def _sim(a: str, b: str) -> float:
    a, b = fold(a or "")[:200], fold(b or "")[:200]
    return SequenceMatcher(None, a, b).ratio() if a and b else 0.0


def _ref_hit(s: dict[str, Any], b: dict[str, Any]) -> bool:
    sct = re.sub(r"\s", "", b.get("so_ct") or "")
    ref = re.sub(r"\s", "", s.get("tham_chieu") or "")
    st = re.sub(r"\s", "", fold(s.get("mo_ta") or ""))
    bt = re.sub(r"\s", "", fold(b.get("mo_ta") or ""))
    return bool((len(sct) >= 4 and fold(sct) in st) or (len(ref) >= 6 and (fold(ref) in bt or fold(ref) == fold(sct))))


def _amt(x: dict[str, Any], side: str, d: str) -> float:
    if side == "s":
        return x["co"] if d == "vao" else x["no"]
    return x["vao"] if d == "vao" else x["ra"]


def reconcile(stmt: list[dict[str, Any]], book: list[dict[str, Any]]) -> dict[str, Any]:
    stmt = sorted(stmt, key=lambda t: t["ngay"])
    d1, d2 = stmt[0]["ngay"], stmt[-1]["ngay"]
    inside = [b for b in book if d1 - timedelta(days=3) <= b["ngay"] <= d2 + timedelta(days=3)]
    outside = [b for b in book if b not in inside]
    rows: list[dict[str, Any]] = []
    used_s: set[int] = set()
    used_b: set[int] = set()
    S = {id(s): s for s in stmt}
    B = {id(b): b for b in inside}

    def add(status: str, s_list: list, b_list: list, note: str = "", diff: float = 0.0) -> None:
        for s in s_list:
            used_s.add(id(s))
        for b in b_list:
            used_b.add(id(b))
        rows.append({"trang_thai": status, "sk": s_list, "so": b_list, "ghi_chu": note, "chenh_lech": diff})

    for d in ("vao", "ra"):
        ss = [s for s in stmt if _amt(s, "s", d) > 0]
        bb = [b for b in inside if _amt(b, "b", d) > 0]
        # 1) khớp 1-1 cùng số tiền
        pairs = []
        for s in ss:
            for b in bb:
                if abs(_amt(s, "s", d) - _amt(b, "b", d)) <= TOL:
                    dd = abs((s["ngay"] - b["ngay"]).days)
                    if dd <= 45:
                        pairs.append((not _ref_hit(s, b), dd, -_sim(s["mo_ta"], b["mo_ta"]), id(s), id(b)))
        for _r, dd, _ms, si, bi in sorted(pairs):
            if si in used_s or bi in used_b:
                continue
            add("Khớp" if dd == 0 else "Khớp – lệch ngày", [S[si]], [B[bi]],
                "" if dd == 0 else f"Ngày sổ lệch {dd} ngày so với sao kê")
        # 2) khớp gộp (1 dòng = tổng nhiều dòng phía bên kia, ±3 ngày)
        for one_side, many_side in (("s", "b"), ("b", "s")):
            ones = [x for x in (ss if one_side == "s" else bb) if id(x) not in (used_s if one_side == "s" else used_b)]
            for x in ones:
                amt = _amt(x, one_side, d)
                pool = [y for y in (bb if many_side == "b" else ss)
                        if id(y) not in (used_b if many_side == "b" else used_s)
                        and abs((y["ngay"] - x["ngay"]).days) <= 3 and _amt(y, many_side, d) < amt + TOL][:25]
                hit = None
                for k in (2, 3, 4):
                    for combo in combinations(pool, k):
                        if abs(sum(_amt(y, many_side, d) for y in combo) - amt) <= TOL:
                            hit = combo
                            break
                    if hit:
                        break
                if hit:
                    if one_side == "s":
                        add("Khớp gộp", [x], list(hit), f"1 dòng sao kê = {len(hit)} dòng sổ")
                    else:
                        add("Khớp gộp", list(hit), [x], f"{len(hit)} dòng sao kê = 1 dòng sổ")
    # 3) sai chiều vào/ra
    for s in stmt:
        if id(s) in used_s:
            continue
        sd, amt = ("vao", s["co"]) if s["co"] else ("ra", s["no"])
        other = "ra" if sd == "vao" else "vao"
        cand = [b for b in inside if id(b) not in used_b and abs(_amt(b, "b", other) - amt) <= TOL
                and abs((b["ngay"] - s["ngay"]).days) <= 3]
        if cand:
            b = min(cand, key=lambda b: abs((b["ngay"] - s["ngay"]).days))
            add("Sai chiều thu/chi", [s], [b],
                f"Sao kê là tiền {'vào' if sd == 'vao' else 'ra'} nhưng sổ ghi tiền {'vào' if other == 'vao' else 'ra'}")
    # 4) lệch số tiền
    pairs = []
    for s in stmt:
        if id(s) in used_s:
            continue
        sd = "vao" if s["co"] else "ra"
        for b in inside:
            if id(b) in used_b or _amt(b, "b", sd) <= 0:
                continue
            dd = abs((b["ngay"] - s["ngay"]).days)
            sim = _sim(s["mo_ta"], b["mo_ta"])
            if dd <= 3 and (_ref_hit(s, b) or sim >= 0.45):
                pairs.append((not _ref_hit(s, b), -sim, dd, id(s), id(b), sd))
    for _r, _ms, _dd, si, bi, sd in sorted(pairs):
        if si in used_s or bi in used_b:
            continue
        diff = _amt(S[si], "s", sd) - _amt(B[bi], "b", sd)
        add("Lệch số tiền", [S[si]], [B[bi]], f"Sao kê nhiều hơn sổ {diff:,.0f}" if diff > 0 else
            f"Sổ nhiều hơn sao kê {-diff:,.0f}", diff)
    # 5) còn lại
    for s in stmt:
        if id(s) not in used_s:
            add("Chỉ có trên sao kê", [s], [], "Chưa hạch toán trên sổ", s["co"] - s["no"])
    for b in inside:
        if id(b) not in used_b:
            add("Chỉ có trên sổ", [], [b], "Không thấy trên sao kê", -(b["vao"] - b["ra"]))
    order = {"Lệch số tiền": 0, "Sai chiều thu/chi": 1, "Chỉ có trên sao kê": 2, "Chỉ có trên sổ": 3,
             "Khớp – lệch ngày": 4, "Khớp gộp": 5, "Khớp": 6}
    rows.sort(key=lambda r: (order[r["trang_thai"]], min([x["ngay"] for x in r["sk"] + r["so"]])))
    return {"rows": rows, "outside": outside, "d1": d1, "d2": d2}


def summary(res: dict[str, Any], stmt: list[dict[str, Any]], book_info: dict[str, Any]) -> dict[str, Any]:
    inside = [b for r in res["rows"] for b in r["so"]]
    cnt: dict[str, int] = {}
    for r in res["rows"]:
        cnt[r["trang_thai"]] = cnt.get(r["trang_thai"], 0) + 1
    st = sorted(stmt, key=lambda t: t["ngay"])
    sk_dau = sk_cuoi = None
    if st and st[0].get("so_du") is not None and st[-1].get("so_du") is not None:
        sk_dau = st[0]["so_du"] - st[0]["co"] + st[0]["no"]
        sk_cuoi = st[-1]["so_du"]
    so_cuoi = book_info.get("cuoi_ky", book_info.get("cuoi_ky_cot"))
    s = {"tu": res["d1"].isoformat(), "den": res["d2"].isoformat(), "dem": cnt,
         "sk_vao": sum(t["co"] for t in stmt), "sk_ra": sum(t["no"] for t in stmt),
         "so_vao": sum(b["vao"] for b in inside), "so_ra": sum(b["ra"] for b in inside),
         "sk_dau_ky": sk_dau, "sk_cuoi_ky": sk_cuoi, "so_dau_ky": book_info.get("dau_ky"), "so_cuoi_ky": so_cuoi,
         "ngoai_ky": len(res["outside"]),
         "lech": sum(v for k, v in cnt.items() if not k.startswith("Khớp"))}
    s["chenh_vao"] = s["sk_vao"] - s["so_vao"]
    s["chenh_ra"] = s["sk_ra"] - s["so_ra"]
    return s


# ------------------------------------------------------------------ báo cáo Excel
def write_report(dest: Path, res: dict[str, Any], summ: dict[str, Any], errors: list[str]) -> Path:
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font, PatternFill

    color = {"Khớp": "E2EFDA", "Khớp gộp": "E2EFDA", "Khớp – lệch ngày": "FFF2CC", "Lệch số tiền": "F8CBAD",
             "Sai chiều thu/chi": "F8CBAD", "Chỉ có trên sao kê": "FCE4D6", "Chỉ có trên sổ": "DDEBF7"}
    bold, money = Font(bold=True), "#,##0"
    wb = Workbook()
    ws = wb.active
    ws.title = "Tong hop"
    ws.append(["ĐỐI CHIẾU SAO KÊ NGÂN HÀNG VỚI SỔ TIỀN GỬI TRÊN PHẦN MỀM KẾ TOÁN"])
    ws["A1"].font = Font(bold=True, size=13)
    ws.append([f"Kỳ sao kê: {summ['tu']} → {summ['den']}"])
    ws.append([])
    ws.append(["Chỉ tiêu", "Sao kê ngân hàng", "Sổ kế toán", "Chênh lệch (SK - Sổ)"])
    for c in ws[4]:
        c.font = bold
    lines = [("Tổng tiền vào", summ["sk_vao"], summ["so_vao"]), ("Tổng tiền ra", summ["sk_ra"], summ["so_ra"]),
             ("Số dư đầu kỳ", summ["sk_dau_ky"], summ["so_dau_ky"]), ("Số dư cuối kỳ", summ["sk_cuoi_ky"], summ["so_cuoi_ky"])]
    for name, a, b in lines:
        ws.append([name, a, b, (a - b) if a is not None and b is not None else None])
        for c in ws[ws.max_row][1:]:
            c.number_format = money
    ws.append([])
    ws.append(["Kết quả", "Số dòng"])
    ws[ws.max_row][0].font = bold
    for k in ("Khớp", "Khớp gộp", "Khớp – lệch ngày", "Lệch số tiền", "Sai chiều thu/chi", "Chỉ có trên sao kê", "Chỉ có trên sổ"):
        ws.append([k, summ["dem"].get(k, 0)])
        ws[ws.max_row][0].fill = PatternFill("solid", fgColor=color[k])
    ws.append(["Dòng sổ ngoài kỳ sao kê (không đối chiếu)", summ["ngoai_ky"]])
    for e in errors:
        ws.append(["Lưu ý: " + e])
    ws.column_dimensions["A"].width = 44
    for c in "BCD":
        ws.column_dimensions[c].width = 22

    heads = ["Trạng thái", "Ghi chú", "Chênh lệch", "Ngày SK", "Tiền vào SK", "Tiền ra SK", "Nội dung sao kê", "Số tham chiếu",
             "Ngày sổ", "Số CT", "Thu (vào) sổ", "Chi (ra) sổ", "Diễn giải sổ", "Đối tượng", "Nguồn sổ"]

    def sheet(title: str, keep) -> None:
        w = wb.create_sheet(title)
        w.append(heads)
        for c in w[1]:
            c.font = bold
        for r in res["rows"]:
            if not keep(r["trang_thai"]):
                continue
            n = max(len(r["sk"]), len(r["so"]))
            for k in range(n):
                s = r["sk"][k] if k < len(r["sk"]) else None
                b = r["so"][k] if k < len(r["so"]) else None
                w.append([r["trang_thai"] if k == 0 else "", r["ghi_chu"] if k == 0 else "",
                          (r["chenh_lech"] or None) if k == 0 else None,
                          s["ngay"] if s else None, (s["co"] or None) if s else None, (s["no"] or None) if s else None,
                          s["mo_ta"] if s else "", s.get("tham_chieu", "") if s else "",
                          b["ngay"] if b else None, b["so_ct"] if b else "", (b["vao"] or None) if b else None,
                          (b["ra"] or None) if b else None, b["mo_ta"] if b else "", b["doi_tuong"] if b else "",
                          b["nguon"] if b else ""])
                row = w[w.max_row]
                row[0].fill = PatternFill("solid", fgColor=color[r["trang_thai"]])
                for i in (2, 4, 5, 10, 11):
                    row[i].number_format = money
                for i in (3, 8):
                    row[i].number_format = "dd/mm/yyyy"
                for i in (6, 12):
                    row[i].alignment = Alignment(wrap_text=True, vertical="top")
        for i, wd in enumerate([20, 30, 14, 11, 15, 15, 50, 16, 11, 14, 15, 15, 45, 28, 30], 1):
            w.column_dimensions[w.cell(1, i).column_letter].width = wd
        w.freeze_panes = "A2"
        w.auto_filter.ref = w.dimensions

    sheet("Can xu ly", lambda s: not s.startswith("Khớp") or s == "Khớp – lệch ngày")
    sheet("Tat ca", lambda s: True)
    if res["outside"]:
        w = wb.create_sheet("So ngoai ky")
        w.append(["Ngày", "Số CT", "Thu (vào)", "Chi (ra)", "Diễn giải", "Nguồn"])
        for b in res["outside"]:
            w.append([b["ngay"], b["so_ct"], b["vao"] or None, b["ra"] or None, b["mo_ta"], b["nguon"]])
    dest.parent.mkdir(parents=True, exist_ok=True)
    wb.save(dest)
    return dest
