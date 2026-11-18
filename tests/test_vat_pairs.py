"""HĐ thay thế / điều chỉnh khi lên tờ khai 01/GTGT: python -m pytest tests/test_vat_pairs.py"""
import json
import sys
import tempfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from app import vat  # noqa: E402
from app.config import settings  # noqa: E402

MST = "0312345678"


def row(so, val, code, tthai, goc=None, kh="C26TAA"):
    return {"huong": "sold", "loai": "Bán ra", "thang": "10/2026", "ngay_lap": f"2026-10-{10 + so % 15:02d}",
            "khmshdon": "1", "khhdon": kh, "shdon": str(so), "nbmst": MST, "nbten": "Cty A", "nmmst": "0100000002",
            "nmten": "Khách B", "tgtcthue": val, "tgtthue": round(val * 0.1), "tgtttbso": val + round(val * 0.1),
            "tthai": tthai, "tthai_code": code, "khhdgoc": kh if goc else "", "shdgoc": str(goc or ""),
            "tdlhdgoc": "2026-10-01" if goc else "", "file_xml": "", "file_pdf": ""}


ROWS = [
    row(32, 10_000_000, 4, "Đã bị thay thế"),                 # bị thay thế bởi 33 -> loại
    row(33, 11_000_000, 2, "Thay thế", goc=32),                # thay thế -> kê khai 11tr
    row(34, 5_000_000, 5, "Đã bị điều chỉnh"),                 # bị điều chỉnh giảm hết bởi 35
    row(35, -5_000_000, 3, "Điều chỉnh", goc=34),              # -> 34 + 35 triệt tiêu, loại cả 2
    row(36, 10_000_000, 5, "Đã bị điều chỉnh"),                # bị điều chỉnh giảm 1tr bởi 37 -> kê khai 9tr
    row(37, -1_000_000, 3, "Điều chỉnh", goc=36),
    row(38, 2_000_000, 3, "Điều chỉnh", goc=7),                # HĐ gốc ở kỳ khác -> kê khai +2tr
    row(40, 4_000_000, "Hóa đơn đã bị thay thế", "Hóa đơn đã bị thay thế"),  # MTT, trạng thái dạng chữ -> loại
    row(41, 3_000_000, 1, "Hóa đơn mới"),
]


def test_pairs_on_return():
    old = settings.output_dir
    with tempfile.TemporaryDirectory() as tmp:
        object.__setattr__(settings, "output_dir", tmp)
        try:
            d = Path(tmp) / MST / "2026-10"
            d.mkdir(parents=True)
            (d / ".dulieu_2026-10.json").write_text(json.dumps({"saved": "2026-11-05", "rows": ROWS}, ensure_ascii=False),
                                                    encoding="utf-8")
            r = vat.compute(MST, "M", 10, 2026, {})
        finally:
            object.__setattr__(settings, "output_dir", old)
    # 33 (11tr) + 36 sau điều chỉnh (9tr) + 38 (+2tr) + 41 (3tr) = 25tr
    assert r["ct"]["32"] == 25_000_000 and r["ct"]["33"] == 2_500_000, r["ct"]
    declared = sorted(int(b["so"]) for b in r["bk"]["sold"])
    assert declared == [33, 36, 38, 41], declared
    line36 = next(b for b in r["bk"]["sold"] if b["so"] == "36")
    assert line36["gt"] == 9_000_000 and "37" in line36["ghi_chu"]
    out = {int(x["so"]): x["why"] for x in r["excluded"]}
    assert set(out) == {32, 34, 35, 37, 40}, out
    assert "33" in out[32] and "triệt tiêu" in out[34] and "triệt tiêu" in out[35]
    text = " ".join(r["warnings"])
    assert "bị thay thế bởi HĐ C26TAA-33" in text and "triệt tiêu về 0" in text and "điều chỉnh giảm" in text
