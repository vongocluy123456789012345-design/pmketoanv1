"""Client cho TConnect Open API – lấy hộ hóa đơn đầu vào/đầu ra từ cổng thuế.

Tài liệu: https://docs.tconnect.vn/docs/open-api/invoice-purchase/overview (v1.7.0, 11/2026)

Luồng: login (Keycloak, body mã hóa AES-256-CBC) -> POST /tax-accounts -> POST /tax-accounts/{id}/sync
       -> poll GET /sync-jobs/{jobId} -> GET /invoices (phân trang) -> GET /invoices/{id}
Phí: sync / fetch-detail tính theo hóa đơn; GET /invoices và GET /invoices/{id} mỗi lần gọi = 10% phí 1 lượt sync.
"""
from __future__ import annotations

import json
import os
import time
from dataclasses import dataclass
from typing import Any, Iterator

import httpx
from cryptography.hazmat.primitives import padding
from cryptography.hazmat.primitives.ciphers import Cipher, algorithms, modes

API_PATH = "/openapi/v1/api-invoice"


class TConnectError(Exception):
    pass


class TConnectAuthError(TConnectError):
    pass


def encrypt_payload(obj: dict[str, Any], aes_key_hex: str) -> str:
    """AES-256-CBC, PKCS7, IV 16 byte nối đầu ciphertext, xuất hex (theo đặc tả TConnect)."""
    key = bytes.fromhex(aes_key_hex)
    if len(key) != 32:
        raise TConnectError("TCONNECT_AES_KEY phải là 64 ký tự hex (32 byte)")
    iv = os.urandom(16)
    padder = padding.PKCS7(128).padder()
    data = padder.update(json.dumps(obj, ensure_ascii=False).encode("utf-8")) + padder.finalize()
    enc = Cipher(algorithms.AES(key), modes.CBC(iv)).encryptor()
    return (iv + enc.update(data) + enc.finalize()).hex()


def decrypt_payload(hex_data: str, aes_key_hex: str) -> dict[str, Any]:
    """Giải mã ngược (dùng cho test)."""
    raw = bytes.fromhex(hex_data)
    iv, ct = raw[:16], raw[16:]
    dec = Cipher(algorithms.AES(bytes.fromhex(aes_key_hex)), modes.CBC(iv)).decryptor()
    padded = dec.update(ct) + dec.finalize()
    unpadder = padding.PKCS7(128).unpadder()
    return json.loads(unpadder.update(padded) + unpadder.finalize())


@dataclass
class TConnectConfig:
    auth_base: str          # vd https://sme-open-api-sandbox.tconnect.vn
    partner_code: str
    client_id: str
    client_secret: str
    aes_key: str
    username: str
    password: str

    @property
    def configured(self) -> bool:
        return all([self.auth_base, self.partner_code, self.client_id, self.client_secret,
                    self.aes_key, self.username, self.password])


class TConnectClient:
    def __init__(self, cfg: TConnectConfig, timeout: float = 60.0):
        if not cfg.configured:
            raise TConnectError("Chưa cấu hình đủ biến TCONNECT_* trong file .env")
        self.cfg = cfg
        base = cfg.auth_base.rstrip("/")
        self.http = httpx.Client(timeout=timeout, headers={"Partner-Code": cfg.partner_code})
        self.auth_base = base
        self.api_base = base + API_PATH
        self._token: str | None = None
        self._refresh: str | None = None
        self._expires_at = 0.0

    # ---------- auth ----------
    def _store_token(self, data: dict[str, Any]) -> None:
        data = data.get("data", data) if isinstance(data, dict) else {}
        token = data.get("access_token")
        if not token:
            raise TConnectAuthError(f"Đăng nhập TConnect không trả về access_token: {str(data)[:200]}")
        self._token = token
        self._refresh = data.get("refresh_token") or self._refresh
        self._expires_at = time.time() + float(data.get("expires_in") or 300) - 30

    def login(self) -> None:
        payload = encrypt_payload({
            "username": self.cfg.username, "password": self.cfg.password,
            "client_id": self.cfg.client_id, "client_secret": self.cfg.client_secret,
        }, self.cfg.aes_key)
        r = self.http.post(self.auth_base + "/openapi/v1/auth/login", json={"data": payload})
        if r.status_code >= 400:
            raise TConnectAuthError(f"Đăng nhập TConnect thất bại ({r.status_code}): {r.text[:300]}")
        self._store_token(r.json())

    def _refresh_token(self) -> None:
        if not self._refresh:
            return self.login()
        r = self.http.post(self.auth_base + "/openapi/v1/auth/refresh-token",
                           json={"refresh_token": self._refresh})
        if r.status_code >= 400:
            return self.login()
        self._store_token(r.json())

    def _headers(self) -> dict[str, str]:
        if not self._token:
            self.login()
        elif time.time() >= self._expires_at:
            self._refresh_token()
        return {"Authorization": f"Bearer {self._token}"}

    def _call(self, method: str, path: str, **kw) -> Any:
        for attempt in range(2):
            r = self.http.request(method, self.api_base + path, headers=self._headers(), **kw)
            if r.status_code == 401 and attempt == 0:
                self._token = None  # token hết hạn sớm -> đăng nhập lại 1 lần
                continue
            if r.status_code >= 400:
                raise TConnectError(f"TConnect {method} {path} lỗi {r.status_code}: {r.text[:300]}")
            return r.json() if r.content else None
        raise TConnectAuthError("TConnect từ chối token")

    # ---------- tài khoản cổng thuế ----------
    def list_tax_accounts(self) -> list[dict[str, Any]]:
        data = self._call("GET", "/tax-accounts")
        return data.get("data", data) if isinstance(data, dict) else (data or [])

    def register_tax_account(self, mst: str, password: str, initial_sync: bool = False) -> dict[str, Any]:
        return self._call("POST", "/tax-accounts", json={
            "mst": mst, "password": password, "triggerInitialSync": initial_sync,
        })

    def update_tax_password(self, account_id: str, password: str) -> dict[str, Any]:
        return self._call("PATCH", f"/tax-accounts/{account_id}", json={"password": password})

    # ---------- đồng bộ ----------
    def start_sync(self, account_id: str, date_from: str, date_to: str,
                   directions: list[str], fetch_details: bool = True) -> dict[str, Any]:
        return self._call("POST", f"/tax-accounts/{account_id}/sync", json={
            "dateFrom": date_from, "dateTo": date_to,
            "directions": directions, "fetchDetails": fetch_details,
        })

    def get_job(self, job_id: str) -> dict[str, Any]:
        return self._call("GET", f"/sync-jobs/{job_id}")

    def wait_job(self, job_id: str, on_progress=None, timeout: float = 3600) -> dict[str, Any]:
        deadline, delay = time.time() + timeout, 3.0
        while True:
            job = self.get_job(job_id)
            if on_progress:
                on_progress(job)
            if job.get("status") == "succeeded":
                return job
            if job.get("status") == "failed":
                code = job.get("errorCode") or "SYNC_FAILED"
                if code == "INVALID_CREDENTIALS":
                    raise TConnectAuthError("TConnect báo sai mật khẩu cổng thuế (INVALID_CREDENTIALS) – "
                                            "hãy cập nhật mật khẩu.")
                raise TConnectError(f"Đồng bộ TConnect thất bại: {code} {job.get('errorMessage') or ''}")
            if time.time() > deadline:
                raise TConnectError("Quá thời gian chờ TConnect đồng bộ")
            time.sleep(delay)
            delay = min(delay * 1.3, 15)

    # ---------- hóa đơn ----------
    def iter_invoices(self, account_id: str, direction: str, date_from: str, date_to: str,
                      limit: int = 100) -> Iterator[dict[str, Any]]:
        page = 1
        while True:
            data = self._call("GET", "/invoices", params={
                "accountId": account_id, "direction": direction, "dateFrom": date_from,
                "dateTo": date_to, "page": page, "limit": limit, "sort": "issueDate:desc",
            })
            rows = data.get("data") or []
            yield from rows
            if page >= int(data.get("totalPage") or 1) or not rows:
                break
            page += 1

    def get_invoice(self, invoice_id: str) -> dict[str, Any]:
        return self._call("GET", f"/invoices/{invoice_id}")

    def close(self) -> None:
        self.http.close()


# ---------- chuyển đổi sang định dạng nội bộ ----------
def split_symbol(symbol: str) -> tuple[str, str]:
    """'1C26TAB' -> ('1', 'C26TAB'); 'C26TAB' -> ('', 'C26TAB')."""
    s = (symbol or "").strip()
    if len(s) == 7 and s[0].isdigit():
        return s[0], s[1:]
    return "", s


def to_parsed(d: dict[str, Any]) -> dict[str, Any]:
    """Chi tiết TConnect -> cấu trúc giống parse_invoice_xml (để dùng chung renderer + connector PDF)."""
    khms, khhd = split_symbol(d.get("invoiceSymbol", ""))
    return {
        "khmshdon": khms, "khhdon": khhd, "shdon": str(d.get("invoiceNumber") or ""),
        "ngay_lap": d.get("issueDate"), "ten_hd": None, "dvtte": d.get("currency"),
        "tgia": None, "htttoan": d.get("paymentMethod"), "msttcgp": None,
        "mccqt": d.get("mccqt"),
        "nguoi_ban": {"ten": d.get("sellerName"), "mst": d.get("sellerTaxCode"),
                      "dia_chi": d.get("sellerAddress")},
        "nguoi_mua": {"ten": d.get("buyerName"), "mst": d.get("buyerTaxCode"),
                      "dia_chi": d.get("buyerAddress")},
        "hang_hoa": [{
            "stt": it.get("lineNumber"), "ten": it.get("description"), "dvt": it.get("unit"),
            "so_luong": it.get("quantity"), "don_gia": it.get("unitPrice"),
            "thanh_tien": it.get("amount"), "thue_suat": it.get("taxRate"), "chiet_khau": None,
        } for it in (d.get("items") or [])],
        "tong_chua_thue": d.get("totalBeforeTax"), "tong_thue": d.get("totalTax"),
        "tong_chiet_khau": None, "tong_thanh_toan": d.get("totalAmount"),
        "bang_chu": d.get("totalInWords"),
        "thong_tin_khac": {}, "lookup_code": d.get("lookupCode"), "lookup_urls": [],
        "tax_breakdown": d.get("taxBreakdown") or [],
        "is_sco": d.get("isSco"),
        "source": "tconnect",
    }


def to_row(owner_mst: str, d: dict[str, Any]) -> dict[str, Any]:
    khms, khhd = split_symbol(d.get("invoiceSymbol", ""))
    return {
        "owner_mst": owner_mst,
        "direction": d.get("direction") or "purchase",
        "nbmst": str(d.get("sellerTaxCode") or ""),
        "khmshdon": khms, "khhdon": khhd, "shdon": str(d.get("invoiceNumber") or ""),
        "tdlap": f"{d['issueDate']}T00:00:00+07:00" if d.get("issueDate") else None,
        "nbten": d.get("sellerName"), "nmmst": d.get("buyerTaxCode"), "nmten": d.get("buyerName"),
        "tgtcthue": d.get("totalBeforeTax"), "tgtthue": d.get("totalTax"),
        "tgtttbso": d.get("totalAmount"),
        "mhdon": d.get("mccqt"),
        "source_query": "tconnect-sco" if d.get("isSco") else "tconnect",
        "raw": d,
    }
