"""Test dự báo thuế GTGT (không cần mạng): python -m tests.test_vat"""
import json
import sys
import tempfile
from datetime import date
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from app import vat  # noqa: E402
from app.config import settings  # noqa: E402


def test_rate_and_period():
    assert vat.rate_class("8%") == "8" and vat.rate_class("10%") == "10" and vat.rate_class("0.08") == "8"
    assert vat.rate_class("KCT") == "KCT" and vat.rate_class("kkknt") == "KKKNT" and vat.rate_class("") == "?"
    assert vat.rate_class("KHAC:5.26%") == "K5.26"
    assert vat.last_completed(date(2026, 12, 1), "M") == ("M", 11, 2026)
    assert vat.last_completed(date(2026, 1, 1), "M") == ("M", 12, 2025)
    assert vat.last_completed(date(2026, 10, 1), "Q") == ("Q", 3, 2026)
    assert vat.last_completed(date(2026, 2, 15), "Q") == ("Q", 4, 2025)
    assert vat.deadline("M", 11, 2026) == date(2026, 12, 20)
    assert vat.deadline("Q", 4, 2026) == date(2027, 1, 31)
    assert "204/2025" in vat.reduction_basis(date(2026, 11, 30))


def _row(huong, so, rates, tthai=1, mau="1", tong=0):
    return {"huong": huong, "loai": "x", "thang": "11/2026", "ngay_lap": "2026-11-10", "khmshdon": mau, "khhdon": "C26TAA",
            "shdon": str(so), "nbmst": "0100000001" if huong == "purchase" else "0312345678", "nbten": "A",
            "nmmst": "0312345678" if huong == "purchase" else "0100000002", "nmten": "B",
            "tgtcthue": sum(v[0] for v in rates.values()), "tgtthue": sum(v[1] for v in rates.values()),
            "tgtttbso": tong, "tthai_code": tthai, "tthai": "", "file_xml": "", "_rates": rates}


def test_compute():
    old = settings.output_dir
    with tempfile.TemporaryDirectory() as tmp:
        object.__setattr__(settings, "output_dir", tmp)
        try:
            folder = Path(tmp) / "0312345678" / "2026-11"
            folder.mkdir(parents=True)
            rows = [
                _row("sold", 1, {"8%": [100_000_000, 8_000_000]}),
                _row("sold", 2, {"10%": [50_000_000, 5_000_000], "KCT": [10_000_000, 0]}),
                _row("sold", 3, {"5%": [20_000_000, 1_000_000]}),
                _row("sold", 4, {"10%": [999_000_000, 99_900_000]}, tthai=6),        # bị hủy -> loại
                _row("purchase", 5, {"10%": [80_000_000, 8_000_000]}, tong=88_000_000),
                _row("purchase", 6, {"KCT": [3_000_000, 0]}),                         # KCT: vào [23], không thuế
                _row("purchase", 7, {"": [4_000_000, 0]}, mau="2"),                   # HĐ bán hàng
            ]
            (folder / ".dulieu_2026-11.json").write_text(json.dumps({"saved": "2026-12-01", "rows": rows}), encoding="utf-8")
            r = vat.compute("0312345678", "M", 11, 2026, {"ct22": 1_000_000, "khong_kt": 500_000})
            ct = r["ct"]
            assert ct["23"] == 83_000_000 and ct["24"] == 8_000_000 and ct["25"] == 7_500_000
            assert ct["26"] == 10_000_000 and ct["30"] == 20_000_000 and ct["31"] == 1_000_000
            assert ct["32"] == 150_000_000 and ct["33"] == 13_000_000
            assert ct["27"] == 170_000_000 and ct["28"] == 14_000_000 and ct["34"] == 180_000_000 and ct["35"] == 14_000_000
            assert ct["36"] == 6_500_000 and ct["40a"] == 5_500_000 and ct["40"] == 5_500_000 and ct["43"] == 0
            assert r["stat"]["loai_tru"] == 1 and r["pl8"]["gt"] == 100_000_000 and r["pl8"]["giam"] == 2_000_000
            assert not r["missing"] and not r["partial"]
            # còn được khấu trừ chuyển kỳ sau
            r2 = vat.compute("0312345678", "M", 11, 2026, {"ct22": 10_000_000, "ct42": 1_000_000})
            assert r2["ct"]["40"] == 0 and r2["ct"]["41"] == 4_000_000 and r2["ct"]["42"] == 1_000_000 and r2["ct"]["43"] == 3_000_000
            html = vat.render_html({**r, "created": None})
            assert "TỜ KHAI THUẾ GIÁ TRỊ GIA TĂNG" in html and "[40a]" in html and "Phụ lục III" in html
            # kỳ 11/2026 dùng mẫu TT 89/2026: [27] trừ [32b], [34] cộng [34a]
            assert r["mau"] == "TT89" and "89/2026" in html and "[32b]" in html and "[34a]" in html
            r3 = vat.compute("0312345678", "M", 11, 2026, {"ct32b": 5_000_000, "ct34a": 7_000_000})
            assert r3["ct"]["27"] == 165_000_000 and r3["ct"]["34"] == 182_000_000
            assert not vat.uses_tt89("M", 6, 2026) and vat.uses_tt89("Q", 3, 2026) and not vat.uses_tt89("Q", 2, 2026)
        finally:
            object.__setattr__(settings, "output_dir", old)


if __name__ == "__main__":
    for name, fn in list(globals().items()):
        if name.startswith("test_") and callable(fn):
            fn()
            print("OK", name)
