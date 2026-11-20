"""Danh sách tài khoản cổng hoadondientu của các công ty để tải hóa đơn theo lịch (tự đăng nhập lần lượt từng công ty).

Mật khẩu được mã hóa bằng Windows DPAPI (khóa của tài khoản Windows đang dùng): file chép sang máy / tài khoản
Windows khác không giải mã được. Phần mềm KHÔNG bao giờ trả mật khẩu ra giao diện / API.
Captcha đăng nhập luôn do người dùng tự nhập.
"""
from __future__ import annotations

import base64
import ctypes
import json
import re
from ctypes import wintypes
from typing import Any

from ..config import ROOT

FILE = ROOT / "data" / "cong_thue_tk.json"


class _Blob(ctypes.Structure):
    _fields_ = [("cbData", wintypes.DWORD), ("pbData", ctypes.POINTER(ctypes.c_char))]


def _blob(data: bytes) -> _Blob:
    buf = ctypes.create_string_buffer(data, len(data))
    return _Blob(len(data), ctypes.cast(buf, ctypes.POINTER(ctypes.c_char)))


def _out(b: _Blob) -> bytes:
    try:
        return ctypes.string_at(b.pbData, b.cbData)
    finally:
        ctypes.windll.kernel32.LocalFree(b.pbData)


def protect(data: bytes) -> bytes:
    out = _Blob()
    if not ctypes.windll.crypt32.CryptProtectData(ctypes.byref(_blob(data)), "AWBook", None, None, None, 0x1,
                                                   ctypes.byref(out)):
        raise OSError("Không mã hóa được mật khẩu (DPAPI)")
    return _out(out)


def unprotect(data: bytes) -> bytes:
    out = _Blob()
    if not ctypes.windll.crypt32.CryptUnprotectData(ctypes.byref(_blob(data)), None, None, None, None, 0x1,
                                                     ctypes.byref(out)):
        raise OSError("Không giải mã được mật khẩu đã lưu (đổi máy / tài khoản Windows?) – hãy lưu lại tài khoản")
    return _out(out)


def norm_user(u: str) -> str:
    return re.sub(r"\s+", "", str(u or ""))


def _read() -> list[dict[str, Any]]:
    try:
        d = json.loads(FILE.read_text(encoding="utf-8"))
    except (FileNotFoundError, ValueError):
        return []
    if isinstance(d, dict) and "accounts" not in d and d.get("username"):      # dữ liệu 1 tài khoản kiểu cũ
        return [{"username": d["username"], "ten": "", "password": d["password"], "auto": True}]
    return list(d.get("accounts") or [])


def _write(accs: list[dict[str, Any]]) -> None:
    FILE.parent.mkdir(parents=True, exist_ok=True)
    FILE.write_text(json.dumps({"accounts": accs}, ensure_ascii=False, indent=1), encoding="utf-8")


def _pw_ok(a: dict[str, Any]) -> bool:
    try:
        unprotect(base64.b64decode(a["password"]))
        return True
    except Exception:  # noqa: BLE001
        return False


def accounts() -> list[dict[str, Any]]:
    """Danh sách công ty – chỉ tên đăng nhập / tên công ty / bật tự động, KHÔNG có mật khẩu."""
    return [{"username": a["username"], "ten": a.get("ten", ""), "auto": bool(a.get("auto", True)), "ok": _pw_ok(a)}
            for a in _read()]


def save(username: str, password: str | None, ten: str = "", auto: bool = True) -> list[dict[str, Any]]:
    """Thêm hoặc sửa 1 công ty. Sửa mà để trống mật khẩu = giữ mật khẩu cũ."""
    username = norm_user(username)
    if not username:
        raise ValueError("Nhập tên đăng nhập cổng thuế (MST) của công ty")
    accs = _read()
    cur = next((a for a in accs if a["username"] == username), None)
    if not password and not cur:
        raise ValueError("Nhập mật khẩu cổng thuế của công ty này")
    enc = base64.b64encode(protect(password.encode("utf-8"))).decode() if password else cur["password"]
    item = {"username": username, "ten": (ten or "").strip() or (cur or {}).get("ten", ""), "password": enc,
            "auto": bool(auto)}
    if cur:
        accs[accs.index(cur)] = item
    else:
        accs.append(item)
    _write(accs)
    return accounts()


def set_auto(username: str, auto: bool) -> list[dict[str, Any]]:
    accs = _read()
    for a in accs:
        if a["username"] == norm_user(username):
            a["auto"] = bool(auto)
    _write(accs)
    return accounts()


def delete(username: str) -> list[dict[str, Any]]:
    _write([a for a in _read() if a["username"] != norm_user(username)])
    return accounts()


def load(username: str) -> tuple[str, str] | None:
    """(tên đăng nhập, mật khẩu) để phần mềm tự đăng nhập – chỉ dùng nội bộ."""
    for a in _read():
        if a["username"] == norm_user(username):
            try:
                return a["username"], unprotect(base64.b64decode(a["password"])).decode("utf-8")
            except Exception:  # noqa: BLE001
                return None
    return None


def auto_users() -> list[str]:
    return [a["username"] for a in accounts() if a["auto"] and a["ok"]]


def name_of(username: str) -> str:
    return next((a["ten"] for a in accounts() if a["username"] == norm_user(username)), "")
