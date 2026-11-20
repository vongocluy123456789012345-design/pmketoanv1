"""Chạy tự động theo lịch: đúng ngày + giờ:phút đã đặt mỗi tháng, tự tải hóa đơn (tháng trước / từ đầu tháng đến hôm nay).

* Lịch được đăng ký vào Windows Task Scheduler ("AWBook\\TaiHoaDonTuDong"): đến giờ Windows chạy
  "AWBook.exe --tai-tu-dong" kể cả khi CHƯA mở phần mềm -> phần mềm chạy nền, Chrome cổng thuế chạy ẩn (dùng phiên
  đăng nhập đã lưu), hiện 1 cửa sổ nhỏ: thanh tiến độ, ô nhập captcha (CHỈ khi cổng bắt nhập), kết quả tải
  (đầy đủ / thiếu gì) + nút vào phần mềm kiểm tra. Máy tắt đúng giờ đó thì chạy ngay khi bật máy lại.
* Phiên đăng nhập cổng thuế hết hạn -> mở cửa sổ Chrome cho bạn tự đăng nhập, đăng nhập xong tự tải tiếp.
* Phần mềm đang mở lúc đến giờ: chạy luôn trong phần mềm (cùng cửa sổ tiến trình).
"""
from __future__ import annotations

import json
import subprocess
import threading
import time
from calendar import monthrange
from datetime import date, datetime, timedelta
from typing import Any

from .config import ROOT, settings

STATE = ROOT / "data" / "schedule.json"
TASK_NAME = r"AWBook\TaiHoaDonTuDong"
LOGIN_POLL = 3             # giây – kiểm tra lại đăng nhập khi đang chờ người dùng đăng nhập
GIVE_UP_HOURS = 24
GIVE_UP_HOURS_BG = 3       # chạy nền (chưa mở phần mềm): chờ đăng nhập tối đa 3 giờ
POPUP_PROFILE = ROOT / "data" / "popup-profile"


def log(msg: str) -> None:
    """Ghi nhật ký lịch tải tự động (bản exe: data/app.log) để kiểm tra lịch có chạy không."""
    try:
        print(f"[{datetime.now():%Y-%m-%d %H:%M:%S}] LỊCH TẢI: {msg}", flush=True)
    except Exception:  # noqa: BLE001
        pass


def toast(title: str, body: str) -> None:
    """Thông báo Windows (không cần cài thêm gì)."""
    esc = lambda s: s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").replace("'", "''")  # noqa: E731
    ps = f"""
[Windows.UI.Notifications.ToastNotificationManager, Windows.UI.Notifications, ContentType = WindowsRuntime] > $null
[Windows.Data.Xml.Dom.XmlDocument, Windows.Data.Xml.Dom.XmlDocument, ContentType = WindowsRuntime] > $null
$x = New-Object Windows.Data.Xml.Dom.XmlDocument
$x.LoadXml('<toast><visual><binding template="ToastGeneric"><text>{esc(title)}</text><text>{esc(body)}</text></binding></visual><audio src="ms-winsoundevent:Notification.IM"/></toast>')
$app = '{{1AC14E77-02E7-4E5D-B744-2EB1AE5198B7}}\\WindowsPowerShell\\v1.0\\powershell.exe'
[Windows.UI.Notifications.ToastNotificationManager]::CreateToastNotifier($app).Show([Windows.UI.Notifications.ToastNotification]::new($x))
"""
    try:
        subprocess.Popen(["powershell", "-NoProfile", "-NonInteractive", "-Command", ps],
                         creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
    except Exception:  # noqa: BLE001
        pass


def prev_month(today: date) -> tuple[date, date]:
    last = today.replace(day=1) - timedelta(days=1)
    return last.replace(day=1), date(last.year, last.month, monthrange(last.year, last.month)[1])


def period_for(today: date, ky: str) -> tuple[date, date]:
    """ky = "truoc": cả tháng trước; "nay": từ đầu tháng này đến hôm nay."""
    if ky == "nay" and today.day > 1:
        return today.replace(day=1), today
    return prev_month(today)


def judge(job: Any) -> dict[str, Any]:
    """Kết quả 1 lượt tải: ĐẦY ĐỦ hay thiếu gì (để báo người dùng)."""
    miss: list[str] = []
    if job.status == "auth_error":
        miss.append("Phiên đăng nhập cổng thuế hết hạn giữa chừng – đăng nhập lại rồi bấm Chạy ngay")
    elif job.status == "cancelled":
        miss.append("Lượt tải đã bị dừng giữa chừng")
    elif job.status == "error":
        miss.append("Lượt tải bị lỗi – xem nhật ký trong mục Tải hóa đơn")
    # cổng thuế KHÔNG có XML gốc (hóa đơn không mã...): không phải lỗi tải, chỉ lưu ý
    no_src = [e for e in job.errors if "không tồn tại hồ sơ gốc" in e.lower()]
    notes = [f"{len(no_src)} hóa đơn cổng thuế không có XML gốc (thường là hóa đơn không mã) – đã có dữ liệu bảng kê"
             " và PDF; chi tiết trong file tổng hợp"] if no_src else []
    xml_short = job.total - job.xml_ok - len(no_src)
    if job.get_xml and xml_short > 0:
        miss.append(f"{xml_short} hóa đơn chưa tải được XML")
    if job.get_pdf and job.pdf_missing:
        miss.append(f"{job.pdf_missing} hóa đơn chưa có PDF (gốc hoặc bản in từ cổng)")
    skipped = [e for e in job.errors if "bỏ qua" in e or "captcha" in e.lower()]
    other = [e for e in job.errors if e not in skipped and e not in no_src and "LỖI" not in e[:4]]
    miss += [e.split("\n")[0][:160] for e in skipped[:6]]
    if not skipped:
        miss += [e.split("\n")[0][:160] for e in other[:4]]
    ok = job.status == "done" and not miss
    head = (f"Đã tải ĐẦY ĐỦ {job.total} hóa đơn" if ok else
            f"Đã tải {job.done}/{job.total} hóa đơn – CÒN THIẾU, cần kiểm tra")
    detail = f"XML {job.xml_ok} · PDF gốc {job.pdf_ok} · PDF bản in từ cổng {job.pdf_hddt} · bảng kê Excel {job.excel_ok}"
    return {"ok": ok, "title": head, "detail": detail, "missing": miss, "notes": notes, "out_dir": job.out_dir,
            "summary_file": job.summary_file, "mst": job.mst}


def _browser_exe() -> str | None:
    try:
        from .gdt.browser import find_browser
        return find_browser()
    except Exception:  # noqa: BLE001
        return None


def open_window(path: str, popup: bool = False) -> subprocess.Popen | None:
    """Mở cửa sổ phần mềm (popup=True: cửa sổ nhỏ tiến trình ở góc phải màn hình, hồ sơ riêng)."""
    exe = _browser_exe()
    if not exe:
        return None
    url = f"http://127.0.0.1:{settings.port}{path}"
    if popup:
        POPUP_PROFILE.mkdir(parents=True, exist_ok=True)
        w, h, x, y = 520, 470, 100, 100
        try:
            import ctypes
            sw, sh = ctypes.windll.user32.GetSystemMetrics(0), ctypes.windll.user32.GetSystemMetrics(1)
            x, y = max(0, sw - w - 24), max(0, sh - h - 70)
        except Exception:  # noqa: BLE001
            pass
        args = [f"--user-data-dir={POPUP_PROFILE}", f"--window-size={w},{h}", f"--window-position={x},{y}",
                "--disable-features=Translate"]
    else:
        prof = ROOT / "data" / "ui-profile"
        prof.mkdir(parents=True, exist_ok=True)
        args = [f"--user-data-dir={prof}", "--window-size=1280,860"]
    return subprocess.Popen([exe, f"--app={url}", *args, "--no-first-run", "--no-default-browser-check"])


class Scheduler:
    def __init__(self) -> None:
        self.enabled = False
        self.day = 5
        self.hour = 8
        self.minute = 0
        self.ky = "truoc"
        self.kinds = ["mua-vao", "mua-vao-khong-ma", "ban-ra"]
        self.last_run = ""        # ngày của lượt theo lịch đã chạy gần nhất (YYYY-MM-DD)
        self.pending_since = ""   # đang chờ đăng nhập từ lúc nào
        self.next_try = 0.0
        self.message = ""
        self.job_id: str | None = None
        self.vat_error = ""
        self.windows = ""
        self.background = False   # tiến trình chạy nền do Windows gọi (chưa mở phần mềm)
        self.phase = "idle"       # idle | login | running | done
        self.result: dict[str, Any] | None = None
        self.login_cap: dict[str, str] | None = None   # (không dùng nữa – giữ cho giao diện cũ)
        self._prefilled = False
        self.login_err = ""
        self.queue: list[str] = []      # các công ty (tên đăng nhập) còn chờ tải trong lượt này
        self.cur = ""                   # công ty đang tải ("" = dùng phiên đang đăng nhập sẵn)
        self.results: list[dict[str, Any]] = []
        self._force = False
        self._scheduled = False   # lượt đang chạy là lượt theo lịch (hiện cửa sổ tiến trình, captcha nhập tay)
        self._popup: subprocess.Popen | None = None
        self._vat_t = 0.0
        self._thread: threading.Thread | None = None
        self._load()

    def _load(self) -> None:
        try:
            s = json.loads(STATE.read_text(encoding="utf-8"))
            for k in ("enabled", "day", "hour", "minute", "ky", "kinds", "last_run", "windows"):
                if k in s:
                    setattr(self, k, s[k])
        except (FileNotFoundError, ValueError):
            pass

    def _save(self) -> None:
        STATE.parent.mkdir(parents=True, exist_ok=True)
        STATE.write_text(json.dumps({"enabled": self.enabled, "day": self.day, "hour": self.hour,
                                     "minute": self.minute, "ky": self.ky, "kinds": self.kinds,
                                     "last_run": self.last_run, "windows": self.windows}, ensure_ascii=False),
                         encoding="utf-8")

    def describe(self) -> str:
        if not self.enabled:
            return "Đang tắt tải tự động"
        what = "từ đầu tháng đến hôm đó" if self.ky == "nay" else "của tháng trước"
        return f"Ngày {self.day} hằng tháng lúc {self.hour:02d}:{self.minute:02d} tự tải hóa đơn {what}"

    def status(self) -> dict[str, Any]:
        from .auto_download import JOBS
        d1, d2 = period_for(date.today(), self.ky)
        job = JOBS.get(self.job_id) if self.job_id else None
        jp = None
        if job:
            last = next((x for x in reversed(job.log) if not x[9:].startswith("OCR")), "")
            jp = {"id": job.id, "status": job.status, "total": job.total, "done": job.done, "xml_ok": job.xml_ok,
                  "pdf_ok": job.pdf_ok, "pdf_hddt": job.pdf_hddt, "captcha": job.captcha, "mst": job.mst,
                  "months": [f"{a:%m/%Y}" for a, _ in job.months], "last": last[9:] if last else ""}
        return {"enabled": self.enabled, "day": self.day, "hour": self.hour, "minute": self.minute, "ky": self.ky,
                "kinds": self.kinds, "last_run": self.last_run, "message": self.message, "job_id": self.job_id,
                "waiting_login": bool(self.pending_since), "next_period": f"{d1:%d/%m/%Y} – {d2:%d/%m/%Y}",
                "mo_ta": self.describe(), "windows": self.windows, "phase": self.phase, "job": jp,
                "result": self.result, "background": self.background, "cur": self.cur,
                "cur_ten": self._ten(self.cur), "queue": list(self.queue), "results": self.results,
                "login_cap": ({k: self.login_cap[k] for k in ("image", "user", "error")} if self.login_cap else None)}

    def configure(self, enabled: bool, day: int, hour: int, kinds: list[str], minute: int = 0, ky: str = "truoc") -> None:
        from . import vat_lich
        self.enabled, self.day, self.hour, self.minute, self.kinds = enabled, day, hour, minute, kinds
        self.ky = ky if ky in ("truoc", "nay") else "truoc"
        self.windows = vat_lich.register_task(
            {"bat": enabled, "ngay": [day], "gio": f"{hour:02d}:{minute:02d}"}, TASK_NAME, "--tai-tu-dong",
            "AWBook: tự tải hóa đơn điện tử theo lịch",
            "Đã đặt lịch trên Windows: đến giờ tự tải kể cả khi chưa mở phần mềm")
        self.message = self.describe()
        self._save()
        self.ensure_running()

    def trigger(self, scheduled: bool = False, mark: bool = True) -> None:
        """Chạy ngay. scheduled=True: lượt theo lịch (Windows gọi) – hiện cửa sổ tiến trình, captcha nhập tay.
        Có lưu tài khoản công ty (bật tải tự động): tải LẦN LƯỢT từng công ty, mỗi công ty tự đăng nhập."""
        from .gdt import credentials
        self._force = True
        self._scheduled = scheduled
        self.next_try = 0
        self.result, self.results = None, []
        self.queue = credentials.auto_users() if scheduled else []
        self.cur = self.queue.pop(0) if self.queue else ""
        self.login_cap, self.login_err, self.pending_since = None, "", ""
        log(f"bắt đầu lượt {'theo lịch' if mark else 'chạy ngay'} – nền={self.background}, công ty: "
            f"{', '.join([self.cur] + self.queue) or '(phiên đang đăng nhập)'}")
        if scheduled:
            self.phase = "running"
        if scheduled and mark:                      # lượt đúng lịch (không tính khi bấm Chạy ngay để thử)
            self.last_run = self._occurrence(datetime.now()).date().isoformat()
            self._save()
        self.ensure_running()

    @staticmethod
    def _ten(user: str) -> str:
        if not user:
            return ""
        from .gdt import credentials
        return credentials.name_of(user)

    def _finish_company(self, res: dict[str, Any]) -> None:
        """Xong 1 công ty (tải xong / bỏ qua / lỗi đăng nhập): ghi kết quả, chuyển sang công ty kế tiếp."""
        res = {**res, "user": self.cur or res.get("mst", ""), "ten": self._ten(self.cur)}
        self.results.append(res)
        log(f"xong {res['ten'] or res['user']}: {res['title']} {'; '.join(res.get('missing') or [])}")
        self.login_cap, self.login_err, self.pending_since, self.job_id = None, "", "", None
        if self.queue:
            self.cur = self.queue.pop(0)
            self._force, self.next_try, self.phase = True, 0, "running"
            return
        self._force, self.phase = False, "done"
        if len(self.results) == 1:
            self.result = self.results[0]
        else:
            bad = [r for r in self.results if not r["ok"]]
            self.result = {"ok": not bad, "notes": [], "missing": [], "summary_file": None, "mst": "",
                           "out_dir": str(settings.output_dir), "companies": self.results, "detail": "",
                           "title": (f"Đã tải ĐẦY ĐỦ hóa đơn {len(self.results)} công ty" if not bad else
                                     f"{len(bad)}/{len(self.results)} công ty CÒN THIẾU – cần kiểm tra")}
        self.message = self.result["title"]
        toast("Tải hóa đơn tự động: " + ("ĐẦY ĐỦ" if self.result["ok"] else "CÒN THIẾU"),
              "; ".join(f"{r['ten'] or r['user']}: {r['title']}" for r in self.results)[:240])

    @staticmethod
    def _fail(title: str, why: str) -> dict[str, Any]:
        return {"ok": False, "title": title, "detail": "", "notes": [], "missing": [why], "out_dir": "",
                "summary_file": None, "mst": ""}

    def skip_company(self) -> dict[str, Any]:
        """Bỏ qua công ty đang chờ đăng nhập -> sang công ty kế tiếp."""
        if self.phase != "login":
            return {"ok": False, "message": "Không có công ty nào đang chờ đăng nhập"}
        self._finish_company(self._fail("Chưa tải: đã bỏ qua đăng nhập",
                                        "Mở phần mềm, đăng nhập cổng thuế rồi bấm Chạy ngay ở mục Lịch tự động"))
        return {"ok": True}

    def refill(self) -> dict[str, Any]:
        """Điền lại tài khoản + mật khẩu vào form đăng nhập trên Chrome (nút trên cửa sổ tiến trình)."""
        from .gdt import credentials
        from .gdt.browser import BROWSER
        creds = credentials.load(self.cur) if self.cur else None
        if not creds:
            return {"ok": False, "message": "Chưa lưu mật khẩu công ty này – hãy tự nhập trên Chrome"}
        try:
            BROWSER.prefill_login(*creds)
            return {"ok": True}
        except Exception as e:  # noqa: BLE001
            return {"ok": False, "message": f"Không điền được: {e}"}

    def busy(self) -> bool:
        return bool(self._force or self.job_id or self.pending_since)

    def ensure_running(self) -> None:
        if self._thread and self._thread.is_alive():
            return
        self._thread = threading.Thread(target=self._loop, daemon=True, name="scheduler")
        self._thread.start()

    # ------------------------------------------------------------------
    def _occurrence(self, now: datetime) -> datetime:
        d = min(self.day, monthrange(now.year, now.month)[1])
        return datetime(now.year, now.month, d, self.hour, self.minute)

    def _due(self, now: datetime) -> bool:
        if not self.enabled:
            return False
        occ = self._occurrence(now)
        if now < occ or self.last_run == occ.date().isoformat():
            return False
        if self.ky == "truoc" and self.last_run == f"{prev_month(now.date())[0]:%Y-%m}":   # dữ liệu kiểu cũ
            return False
        return True

    def _loop(self) -> None:
        from . import vat
        while True:
            try:
                self._tick()
            except Exception as e:  # noqa: BLE001
                self.message = f"Lỗi bộ chạy tự động: {e}"
            if time.time() - self._vat_t >= 30:
                self._vat_t = time.time()
                try:  # dự báo thuế GTGT ngay từ 00:00 ngày đầu kỳ sau
                    vat.auto_tick()
                    from . import vat_lich      # lịch dự báo / nhắc theo ngày tự chọn trong tháng
                    vat_lich.tick()
                    self.vat_error = ""
                except Exception as e:  # noqa: BLE001
                    self.vat_error = f"Lỗi dự báo VAT tự động: {e}"
            time.sleep(2 if self.busy() else 10)

    def _show_popup(self) -> None:
        if self._scheduled and not (self._popup and self._popup.poll() is None):
            self._popup = open_window("/tien-trinh", popup=True)

    def popup_open(self) -> bool:
        return bool(self._popup and self._popup.poll() is None)

    def _tick(self) -> None:
        from .auto_download import JOBS, running_job, start
        from .gdt.browser import BROWSER

        now = datetime.now()
        if self.job_id:  # đang theo dõi lượt tự động
            job = JOBS.get(self.job_id)
            if job and job.status == "running":
                if job.captcha and self._scheduled:
                    self._show_popup()          # cổng bắt nhập captcha -> hiện cửa sổ cho bạn nhập
                return
            self._finish_company(judge(job) if job else self._fail("Lượt tải bị mất", "Hãy bấm Chạy ngay"))
            if self.phase == "done":
                if self._scheduled:
                    self._show_popup()
                if self.background and BROWSER.hidden:
                    BROWSER.close()             # tắt Chrome chạy ẩn
            return
        if not (self._force or self._due(now)) or time.time() < self.next_try:
            return
        if not self._force:                     # đến lịch khi phần mềm đang mở
            self.trigger(scheduled=True)
        if running_job():
            self.next_try = time.time() + 30
            return
        if self._scheduled:
            self._show_popup()
        st = BROWSER.status()
        if not st.get("open"):
            try:
                st = BROWSER.open()   # Chrome thật (cổng chặn Chrome chạy ẩn)
            except Exception as e:  # noqa: BLE001
                self.message = f"Không mở được Chrome: {e}"
        from .gdt import credentials
        other = bool(self.cur) and st.get("logged_in") and credentials.norm_user(st.get("mst")) != self.cur
        if not st.get("logged_in") or other:
            # Đăng nhập trên CỬA SỔ CHROME THẬT của cổng: phần mềm điền sẵn tài khoản + mật khẩu đã lưu,
            # người dùng tự nhập captcha và tự bấm "Đăng nhập" (cổng chặn đăng nhập tự động).
            first = not self.pending_since
            if first:
                self.pending_since = now.isoformat(timespec="minutes")
                self._prefilled = False
                self.phase = "login"
                who = f"{self._ten(self.cur)} ({self.cur})" if self.cur else "công ty cần tải"
                toast("Đến lịch tải hóa đơn tự động",
                      f"Đăng nhập {who} trên cửa sổ Chrome cổng thuế: nhập captcha rồi bấm Đăng nhập.")
                try:
                    if other:
                        BROWSER.logout()                  # đang đăng nhập công ty khác -> đăng xuất
                    if BROWSER.hidden or not st.get("open"):
                        BROWSER.show_for_login()
                    else:
                        BROWSER.open()                    # đưa cửa sổ cổng thuế lên trước
                except Exception as e:  # noqa: BLE001
                    self.message = f"Không mở được Chrome: {e}"
            if self.cur and not self._prefilled:
                creds = credentials.load(self.cur)
                if creds:
                    try:
                        self._prefilled = BROWSER.prefill_login(*creds)
                    except Exception as e:  # noqa: BLE001
                        self.message = f"Chưa điền sẵn được tài khoản: {e}"
                else:
                    self._prefilled = True                # không đọc được mật khẩu: để người dùng tự nhập
            if self._scheduled:
                self._show_popup()
            limit = GIVE_UP_HOURS_BG if self.background else GIVE_UP_HOURS
            if now - datetime.fromisoformat(self.pending_since) > timedelta(hours=limit):
                self._finish_company(self._fail("Chưa tải được: chưa đăng nhập cổng thuế",
                                                "Hãy mở phần mềm, đăng nhập cổng thuế rồi bấm Chạy ngay"))
                return
            who = f"{self._ten(self.cur)} ({self.cur})" if self.cur else "cổng thuế"
            self.message = f"Chờ đăng nhập {who} trên cửa sổ Chrome – đăng nhập xong phần mềm tự tải"
            self.next_try = time.time() + LOGIN_POLL
            return
        d1, d2 = period_for(now.date(), self.ky)
        job = start(d1, d2, self.kinds or ["mua-vao", "ban-ra"], True, True, True, manual_captcha=self._scheduled)
        self.job_id, self.pending_since, self._force, self.phase = job.id, "", False, "running"
        log(f"bắt đầu tải {d1:%d/%m/%Y}–{d2:%d/%m/%Y} MST {st.get('mst')}")
        self.message = f"Đang tự động tải hóa đơn {d1:%d/%m/%Y} – {d2:%d/%m/%Y}"
        if not self._scheduled:
            toast("Bắt đầu tải hóa đơn tự động", f"{d1:%d/%m/%Y} – {d2:%d/%m/%Y} · MST {st.get('mst')}")


SCHEDULER = Scheduler()
