"""Đối chiếu sao kê thông minh: python -m pytest tests/test_bank_smart.py"""
import io
import sys
from datetime import date
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from openpyxl import Workbook  # noqa: E402

from app import bank_post, bank_smart as bs  # noqa: E402


def _xlsx(rows):
    wb = Workbook()
    ws = wb.active
    for r in rows:
        ws.append(r)
    b = io.BytesIO()
    wb.save(b)
    return b.getvalue()


def test_read_partners_and_sales():
    ncc = _xlsx([["TỔNG HỢP CÔNG NỢ PHẢI TRẢ"], ["Tài khoản: 331"],
                 ["Mã nhà cung cấp", "Tên nhà cung cấp", "TK công nợ", "Số dư đầu kỳ", None, "Phát sinh", None, "Số dư cuối kỳ", None],
                 [None, None, None, "Nợ", "Có", "Nợ", "Có", "Nợ", "Có"],
                 ["NCC_HDT", "CTy TNHH TM DV Kỹ thuật HDT", "331", 0, 0, 52578720, 58756320, 0, 6177600],
                 ["Số dòng = 1", None, None, 0, 0, 52578720, 58756320, 0, 6177600]])
    ps = bs.read_partners("ncc.xlsx", ncc, "NCC")
    assert len(ps) == 1 and ps[0]["ma"] == "NCC_HDT" and round(ps[0]["du"]) == 6177600
    kh = _xlsx([["Danh sách khách hàng"], ["STT", "Mã khách hàng", "Tên khách hàng", "Địa chỉ", "Công nợ", "Mã số thuế"],
                [1, "KH_GN", "Công ty TNHH Kỹ Thuật Cơ Điện Gia Nghĩa", "HCM", 281264667, "0313940386"],
                [2, "001168008628", "Hộ kinh doanh Gật Gù", "HCM", 0, "001168008628"]])
    ps2 = bs.read_partners("kh.xlsx", kh, "KH")
    assert [p["mst"] for p in ps2] == ["0313940386", "001168008628"]          # MST 12 số (CCCD) giữ đủ
    so = _xlsx([["SỔ CHI TIẾT BÁN HÀNG"],
                ["Ngày hạch toán", "Ngày chứng từ", "Số chứng từ", "Ngày hóa đơn", "Số hóa đơn", "Diễn giải chung",
                 "Tên khách hàng", "Mã số thuế", "Doanh số bán", "Chiết khấu", "Thuế GTGT", "Tổng thanh toán"],
                [date(2026, 1, 3), date(2026, 1, 3), "BDV001", date(2026, 1, 3), "217", "Bán hàng",
                 "Công ty TNHH Kỹ Thuật Cơ Điện Gia Nghĩa", "0313940386", 200000000, 0, 16000000, 216000000],
                [date(2026, 1, 3), date(2026, 1, 3), "BDV001", date(2026, 1, 3), "217", "Bán hàng",
                 "Công ty TNHH Kỹ Thuật Cơ Điện Gia Nghĩa", "0313940386", 60430247, 0, 4834420, 65264667],
                ["Tổng cộng", None, None, None, None, None, None, None, 0, 0, 0, 281264667]])
    vs = bs.read_sales("so.xlsx", so)
    assert len(vs) == 1 and round(vs[0]["tong"]) == 281264667 and vs[0]["so_hd"] == "217"


def test_identify_and_pipeline(tmp_path):
    partners = [
        {"ma": "NCC_HDT", "ten": "CTy TNHH TM DV Kỹ thuật HDT", "mst": "3702693431", "tk": [], "du": 0, "loai": "NCC"},
        {"ma": "NCC_ALPHA", "ten": "CTy TNHH TM Thiết Bị Điện Alpha", "mst": "0300000001", "tk": ["19033123456789"],
         "du": 0, "loai": "NCC"},
        {"ma": "KH_GN", "ten": "Công ty TNHH Kỹ Thuật Cơ Điện Gia Nghĩa", "mst": "0313940386", "tk": [], "du": 0, "loai": "KH"},
        {"ma": "KH_KK", "ten": "Cty TNHH Thực Phẩm Khang Ký", "mst": "0300000002", "tk": [], "du": 0, "loai": "KH"},
        {"ma": "KH_HD", "ten": "Công Ty TNHH MTV Hải Dương", "mst": "0300000003", "tk": [], "du": 0, "loai": "KH"},
    ]
    dx = bs.Directory(partners)
    own = "CONG TY TNHH THUONG MAI DICH VU DIEN KHANG PHAT"
    mk = lambda mo, co=0.0, no=0.0, dd="", tk="": {"ngay": date(2026, 1, 6), "mo_ta": mo, "doi_ung": dd,  # noqa: E731
                                                  "tham_chieu": "", "tk_doi_ung": tk, "co": co, "no": no}
    t1 = mk("Khang Phat tt cong no cu cty HDT hd so 755 756", no=342984583)
    t2 = mk("Cty Gia Nghia thanh toan Cty Khang Phat HD 217", co=281264667)
    t3 = mk("chuyen tien mua hang", no=2336796, tk="19033123456789")
    t4 = mk("LE THI HAI DUONG chuyen tien", co=286000)
    t5 = mk("Phi quan ly tai khoan", no=11000)
    t6 = mk("Thanh toan tien hang dot 3", co=5000000)
    assert dx.identify(t1, own)["ma"] == "NCC_HDT"                   # "cty HDT" (chữ viết tắt)
    assert dx.identify(t2, own) is None or dx.identify(t2, own)["ma"] != "KH_KK"   # tên mình (Khang Phát) bị bỏ
    assert dx.identify(t3, own)["ma"] == "NCC_ALPHA"                 # số TK ngân hàng người nhận
    d4 = dx.identify(t4, own)
    assert d4 is None or d4["diem"] < 85                             # tên người trùng tên công ty: chỉ cần kiểm tra
    txns = [t1, t2, t3, t4, t5, t6]
    for t in txns:
        t["dt"] = dx.identify(t, own)
    sales = [{"so_ct": "BDV001", "so_hd": "217", "ngay": date(2026, 1, 3), "ma_kh": "", "mst": "0313940386",
              "ten": "Công ty TNHH Kỹ Thuật Cơ Điện Gia Nghĩa", "tong": 281264667.0, "dien_giai": ""}]
    bank_post.allocate([], txns)
    assert bs.apply_sales(txns, sales, dx) == 1
    assert t2["dt"]["ma"] == "KH_GN" and t2["alloc"][0]["inv"]["_so_ban"]["so_ct"] == "BDV001"
    vouchers = bank_post.build_vouchers(txns, "TT200", own)
    bs.fill_codes(vouchers, dx)
    by = {id(v["tx"]): v for v in vouchers}
    assert by[id(t1)]["lines"][0]["ma"] == "NCC_HDT" and by[id(t1)]["lines"][0]["no"] == "331"
    assert by[id(t2)]["lines"][0]["ma"] == "KH_GN" and by[id(t2)]["lines"][0]["co"] == "131"
    assert by[id(t5)]["lines"][0]["no"] in ("6427", "6422")          # phí ngân hàng theo quy tắc
    st = bs.write_result(tmp_path / "kq.xlsx", vouchers, {"so_tk": "1"}, [], ["test"])
    assert st["chua_ro"] >= 1 and (tmp_path / "kq.xlsx").exists()   # "tiền hàng đợt 3" không rõ của ai
    assert bs.status_of(t6, by[id(t6)]).startswith("CHƯA RÕ")


def test_ledger_book_booked_learn_export():
    led = _xlsx([["CHI TIẾT CÔNG NỢ PHẢI TRẢ NHÀ CUNG CẤP"], ["Từ ngày 01/01/2026"], [],
                 ["Ngày hạch toán", "Ngày chứng từ", "Số chứng từ", "Diễn giải", "Phát sinh", None, "Số dư"],
                 [None, None, None, None, "Số đã trả", "Số phải trả", "Số trả trước", "Số còn phải trả"],
                 ["Tên nhà cung cấp: TIKTOK PTE. LTD."],
                 [date(2026, 6, 1), date(2026, 6, 1), "MDV1", "Mua dịch vụ quảng cáo", 0, 33000000, 0, 33000000],
                 [date(2026, 6, 2), date(2026, 6, 2), "PCB2600", "Thanh toán Mua dịch vụ quảng cáo", 33000000, 0, 0, 0],
                 [None, None, None, "Cộng phát sinh trong kỳ", 33000000, 33000000, 0, 0]])
    lg = bs.read_ledger("ncc.xlsx", led, "NCC")
    assert lg and lg["partners"][0]["ten"] == "TIKTOK PTE. LTD." and len(lg["entries"]) == 2
    assert lg["entries"][1]["da"] == 33000000 and lg["entries"][0]["phat_sinh"] == 33000000
    chi = _xlsx([["CHI TIỀN GỬI"], [], ["STT", "Ngày hạch toán", "Số chứng từ", "Diễn giải", "Số tiền", "Đối tượng",
                                         "Lý do chi", "Loại chứng từ"],
                 [1, date(2026, 6, 10), "PCB2601", "Lương T5 Phúc", 9000000, "", "Chi khác", "Ủy nhiệm chi"],
                 [2, date(2026, 6, 11), "PCB2602", "Lương T5 Tuyền", 8000000, "", "Chi khác", "Ủy nhiệm chi"]])
    bb = bs.read_bank_book("Chi_tien_gui.xlsx", chi)
    assert bb["huong"] == "chi" and [e["so_ct"] for e in bb["entries"]] == ["PCB2601", "PCB2602"]
    mk = lambda d, mo, no: {"ngay": d, "mo_ta": mo, "doi_ung": "", "tham_chieu": "", "tk_doi_ung": "", "co": 0.0, "no": no}  # noqa: E731
    t1 = mk(date(2026, 6, 2), "4229XX6023 HO KINH DOANH GENME", 33000000)
    t2 = mk(date(2026, 6, 10), "LUONG T5 PHUC-100626", 9000000)
    t3 = mk(date(2026, 6, 11), "LUONG T5 TUYEN-110626", 8000000)
    t4 = mk(date(2026, 7, 10), "LUONG T6 HOANG-100726", 7692308)          # tháng mới, tên người khác
    txns = [t1, t2, t3, t4]
    assert bs.mark_booked(txns, bb["entries"], lg["entries"]) == 3
    assert t1["booked"]["so_ct"] == "PCB2600" and t2["booked"]["so_ct"] == "PCB2601" and not t4.get("booked")
    assert bs.learn(txns) == 1 and t4["hoc"]["ly_do"] == "Chi khác"      # học theo từ khóa "luong"
    vouchers = bank_post.build_vouchers(txns, "TT133", "HO KINH DOANH GENME")
    out = bs.prepare_export(vouchers, bb["entries"])
    assert len(out) == 1 and out[0]["so_ct"] == "PCB2603" and out[0]["_ly_do_misa"] == "Chi khác"
