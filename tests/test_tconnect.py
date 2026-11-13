"""Test offline cho connector TConnect (không gọi mạng)."""
import os
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from app.providers.fallback import render_pdf  # noqa: E402
from app.sources.tconnect import (  # noqa: E402
    decrypt_payload, encrypt_payload, split_symbol, to_parsed, to_row,
)

DETAIL = {
    "id": "7b1e0000-0000-0000-0000-0000000000c2", "direction": "purchase", "isSco": False,
    "invoiceSymbol": "C26TAB", "invoiceNumber": "205", "issueDate": "2026-07-18",
    "sellerName": "CÔNG TY TNHH VẬT TƯ ABC", "sellerTaxCode": "0101243150", "sellerAddress": "Hà Nội",
    "buyerName": "CÔNG TY MUA", "buyerTaxCode": "0312345678", "paymentMethod": "TM/CK",
    "totalBeforeTax": 1472727, "totalTax": 147273, "totalAmount": 1620000,
    "totalInWords": "Một triệu sáu trăm hai mươi nghìn đồng", "currency": "VND",
    "mccqt": "00ABC", "lookupCode": "XYZ123",
    "taxBreakdown": [{"taxRate": "10%", "amountBeforeTax": 1472727, "taxAmount": 147273}],
    "items": [{"lineNumber": 1, "description": "Mực in HP 12A", "unit": "Hộp", "quantity": 3,
               "unitPrice": 490909, "amount": 1472727, "taxRate": "10%", "taxAmount": 147273,
               "totalAmount": 1620000}],
}


def test_encrypt_roundtrip():
    key = os.urandom(32).hex()
    obj = {"username": "a@b.c", "password": "p@ss#1", "client_id": "x", "client_secret": "y"}
    enc = encrypt_payload(obj, key)
    assert all(c in "0123456789abcdef" for c in enc) and len(enc) % 32 == 0
    assert decrypt_payload(enc, key) == obj


def test_split_symbol():
    assert split_symbol("1C26TAB") == ("1", "C26TAB")
    assert split_symbol("C26TAB") == ("", "C26TAB")
    assert split_symbol("K25TAA") == ("", "K25TAA")


def test_mapping_and_render():
    row = to_row("0312345678", DETAIL)
    assert row["nbmst"] == "0101243150" and row["khhdon"] == "C26TAB" and row["shdon"] == "205"
    assert row["tdlap"] == "2026-07-18T00:00:00+07:00" and row["tgtttbso"] == 1620000
    p = to_parsed(DETAIL)
    assert p["lookup_code"] == "XYZ123" and p["hang_hoa"][0]["ten"] == "Mực in HP 12A"
    assert render_pdf(p)[:5] == b"%PDF-"


if __name__ == "__main__":
    for name, fn in list(globals().items()):
        if name.startswith("test_"):
            fn()
            print("OK", name)
