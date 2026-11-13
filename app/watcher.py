"""Theo dõi thư mục Downloads: file XML/ZIP/Excel tải từ cổng thuế được tự động nhập.

* Lần bật đầu tiên chỉ ghi nhận các file đang có (không nhập file cũ).
* File mới: chờ tải xong (dung lượng không đổi qua 2 lần quét), bỏ qua .crdownload/.part/.tmp.
* Mỗi hóa đơn được gán vào MST khách hàng đã khai báo (bảng companies) dựa trên nội dung file.
"""
from __future__ import annotations

import json
import threading
import time
from collections import deque
from datetime import datetime
from pathlib import Path
from typing import Any

from . import db
from .config import ROOT, settings

EXTS = {".xml", ".zip", ".xlsx", ".xls"}
TEMP_EXTS = {".crdownload", ".part", ".tmp", ".download"}
STATE_FILE = ROOT / "data" / "watcher_state.json"


class Watcher:
    def __init__(self) -> None:
        self.folder = Path(settings.watch_dir).expanduser()
        self.enabled = False
        self.fetch_pdf = True
        self.events: deque[dict[str, Any]] = deque(maxlen=200)
        self._seen: dict[str, str] = {}      # path -> "size:mtime" đã xử lý
        self._pending: dict[str, str] = {}   # path -> "size:mtime" lần quét trước (chờ ổn định)
        self._lock = threading.Lock()
        self._thread: threading.Thread | None = None
        self._load()

    # ---------- trạng thái ----------
    def _load(self) -> None:
        try:
            s = json.loads(STATE_FILE.read_text(encoding="utf-8"))
            self.enabled = bool(s.get("enabled"))
            self.fetch_pdf = bool(s.get("fetch_pdf", True))
            self._seen = s.get("seen", {})
            if s.get("folder"):
                self.folder = Path(s["folder"])
        except (FileNotFoundError, ValueError):
            pass

    def _save(self) -> None:
        STATE_FILE.parent.mkdir(parents=True, exist_ok=True)
        STATE_FILE.write_text(json.dumps({
            "enabled": self.enabled, "fetch_pdf": self.fetch_pdf,
            "folder": str(self.folder), "seen": self._seen,
        }, ensure_ascii=False), encoding="utf-8")

    def log(self, level: str, msg: str, **extra: Any) -> None:
        self.events.appendleft({"time": datetime.now().strftime("%d/%m %H:%M:%S"),
                                "level": level, "msg": msg, **extra})

    def status(self) -> dict[str, Any]:
        return {"enabled": self.enabled, "folder": str(self.folder), "folder_exists": self.folder.is_dir(),
                "fetch_pdf": self.fetch_pdf, "events": list(self.events)[:50]}

    # ---------- điều khiển ----------
    def configure(self, enabled: bool, folder: str | None = None, fetch_pdf: bool | None = None) -> None:
        with self._lock:
            if folder:
                newf = Path(folder).expanduser()
                if newf != self.folder:
                    self.folder, self._seen = newf, {}
            if fetch_pdf is not None:
                self.fetch_pdf = fetch_pdf
            if enabled and not self.enabled:
                self._baseline()
                self.log("info", f"Bắt đầu theo dõi {self.folder}")
            elif not enabled and self.enabled:
                self.log("info", "Đã tắt theo dõi")
            self.enabled = enabled
            self._save()
        self.ensure_running()

    def _baseline(self) -> None:
        """Đánh dấu file đang có là đã thấy -> chỉ nhập file tải về SAU khi bật."""
        for p in self._candidates():
            self._seen[str(p)] = self._sig(p)

    def ensure_running(self) -> None:
        if self._thread and self._thread.is_alive():
            return
        self._thread = threading.Thread(target=self._loop, daemon=True, name="downloads-watcher")
        self._thread.start()

    # ---------- quét ----------
    def _candidates(self):
        if not self.folder.is_dir():
            return []
        out = []
        for p in self.folder.iterdir():
            if p.is_file() and p.suffix.lower() in EXTS and p.suffix.lower() not in TEMP_EXTS:
                out.append(p)
        return out

    @staticmethod
    def _sig(p: Path) -> str:
        st = p.stat()
        return f"{st.st_size}:{int(st.st_mtime)}"

    def _loop(self) -> None:
        while True:
            try:
                if self.enabled:
                    self.scan_once()
            except Exception as e:  # noqa: BLE001
                self.log("error", f"Lỗi quét thư mục: {e}")
            time.sleep(settings.watch_interval)

    def scan_once(self) -> list[Path]:
        ready: list[Path] = []
        with self._lock:
            for p in self._candidates():
                key, sig = str(p), self._sig(p)
                if self._seen.get(key) == sig:
                    continue
                if self._pending.get(key) == sig:  # dung lượng ổn định 2 lần quét -> tải xong
                    ready.append(p)
                    self._pending.pop(key, None)
                    self._seen[key] = sig
                else:
                    self._pending[key] = sig
            if ready:
                self._save()
        if ready:
            self.process(ready)
        return ready

    # ---------- xử lý ----------
    def process(self, paths: list[Path]) -> None:
        from .sources.file_import import detect_owner, iter_items
        from .sync import start_import_job

        try:
            managed = {c["mst"] for c in db.list_companies()}
        except Exception as e:  # noqa: BLE001
            self.log("error", f"Không đọc được danh sách MST: {e}")
            return
        groups: dict[str, list[tuple[str, bytes]]] = {}
        for p in paths:
            try:
                data = p.read_bytes()
            except OSError as e:
                self.log("error", f"{p.name}: không đọc được ({e})")
                continue
            n_ok = 0
            for kind, item_name, payload in iter_items(p.name, data):
                if kind == "skip":
                    continue
                try:
                    owner, why = detect_owner(kind, payload, item_name, managed)
                except Exception as e:  # noqa: BLE001
                    owner, why = None, f"lỗi đọc file: {e}"
                if not owner:
                    self.log("warn", f"{item_name}: bỏ qua – {why}", file=p.name)
                    continue
                groups.setdefault(owner, []).append((item_name, payload))
                n_ok += 1
            if n_ok == 0:
                self.log("warn", f"{p.name}: không có hóa đơn nào nhận được", file=p.name)
        for owner, files in groups.items():
            job = start_import_job(owner, "auto", files, self.fetch_pdf)
            self.log("ok", f"Đang nhập {len(files)} mục cho MST {owner}", job_id=job.id, mst=owner)


WATCHER = Watcher()
