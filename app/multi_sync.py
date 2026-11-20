"""Tải hóa đơn NHIỀU CÔNG TY CÙNG LÚC.

Cổng hoadondientu chỉ xác thực bằng header "Authorization: Bearer <token>" (không dựa vào cookie), nên:
  1. Phần mềm cho đăng nhập LẦN LƯỢT từng công ty trên cửa sổ Chrome cổng thuế (điền sẵn tài khoản + mật khẩu đã lưu,
     người dùng tự nhập captcha và bấm Đăng nhập – phần mềm không tự vượt captcha đăng nhập).
  2. Đăng nhập xong công ty nào thì giữ token của công ty đó và BẮT ĐẦU TẢI NGAY, rồi chuyển sang đăng nhập công ty
     kế tiếp – các công ty đã đăng nhập vẫn tiếp tục tải song song bằng token riêng của mình.
  3. Bản in PDF từ trang cổng (cần tab cổng đang hiện trang Tra cứu) tạm chờ tới khi đăng nhập xong hết.
Token chỉ giữ trong bộ nhớ (tắt phần mềm là mất), không ghi ra đĩa.
"""
from __future__ import annotations

import threading
import time
import uuid
from datetime import date, datetime
from typing import Any

from . import auto_download as ad
from .gdt import credentials
from .gdt.browser import BROWSER

MAX_PARALLEL = 4          # số công ty tải cùng lúc (còn lại xếp hàng, đăng nhập vẫn làm trước)
LOGIN_WAIT = 15 * 60      # chờ người dùng đăng nhập 1 công ty tối đa 15 phút rồi bỏ qua
TOKENS: dict[str, str] = {}

STATE_TEXT = {"cho": "Chờ đăng nhập", "login": "Đang chờ bạn đăng nhập", "cho_tai": "Đã đăng nhập – xếp hàng tải",
              "dang_tai": "Đang tải", "xong": "Xong", "thieu": "Xong – còn thiếu", "loi": "Lỗi", "bo_qua": "Bỏ qua",
              "dung": "Đã dừng"}


def _norm(u: Any) -> str:
    return credentials.norm_user(str(u or ""))


def _valid(tok: str | None, user: str, margin: int = 20 * 60) -> bool:
    if not tok:
        return False
    info = BROWSER.token_info(tok)
    return _norm(info["mst"]) == user and (not info["exp"] or info["exp"] > time.time() + margin)


class Run:
    def __init__(self, users: list[str], d1: date, d2: date, kinds: list[str], excel: bool, xml: bool, pdf: bool,
                 easy: bool) -> None:
        self.id = uuid.uuid4().hex
        self.d1, self.d2, self.kinds = d1, d2, kinds
        self.excel, self.xml, self.pdf, self.easy = excel, xml, pdf, easy
        self.items: list[dict[str, Any]] = [
            {"user": u, "ten": credentials.name_of(u), "state": "cho", "msg": "", "job_id": None, "token": "",
             "co_mk": credentials.load(u) is not None, "ket_qua": None} for u in users]
        self.status = "running"         # running | done | cancelled
        self.cur: int | None = None     # công ty đang chờ đăng nhập
        self.cancel = False
        self.skip = False
        self.log: list[str] = []
        self.started = datetime.now()

    def say(self, msg: str) -> None:
        self.log.append(f"{datetime.now():%H:%M:%S} {msg}")
        self.log = self.log[-200:]

    def _job(self, it: dict[str, Any]) -> ad.AutoJob | None:
        return ad.JOBS.get(it["job_id"]) if it["job_id"] else None

    def running(self) -> int:
        return sum(1 for it in self.items if (j := self._job(it)) and j.status == "running")

    def public(self) -> dict[str, Any]:
        rows = []
        for i, it in enumerate(self.items):
            j = self._job(it)
            jp = None
            if j:
                last = next((x for x in reversed(j.log) if not x[9:].startswith("OCR")), "")
                jp = {"id": j.id, "status": j.status, "mst": j.mst, "total": j.total, "done": j.done,
                      "xml_ok": j.xml_ok, "pdf_ok": j.pdf_ok, "pdf_hddt": j.pdf_hddt, "pdf_missing": j.pdf_missing,
                      "errors": len(j.errors), "last_error": j.errors[-1][:300] if j.errors else "",
                      "captcha": j.captcha, "last": last[9:], "out_dir": j.out_dir, "summary_file": j.summary_file}
            rows.append({k: it[k] for k in ("user", "ten", "state", "msg", "co_mk", "ket_qua")}
                        | {"state_text": STATE_TEXT.get(it["state"], it["state"]), "job": jp, "dang_login": i == self.cur})
        return {"id": self.id, "status": self.status, "items": rows, "running": self.running(), "max": MAX_PARALLEL,
                "tu": f"{self.d1:%d/%m/%Y}", "den": f"{self.d2:%d/%m/%Y}", "kinds": self.kinds, "log": self.log[-60:],
                "cur": self.cur}


RUN: Run | None = None
_lock = threading.Lock()


def current() -> dict[str, Any]:
    return RUN.public() if RUN else {}


def start(users: list[str], d1: date, d2: date, kinds: list[str], excel: bool = True, xml: bool = True,
          pdf: bool = True, easy: bool = True) -> Run:
    global RUN
    with _lock:
        if ad.MULTI["active"]:
            raise ValueError("Đang tải nhiều công ty – chờ xong hoặc bấm Dừng tất cả")
        if ad.running_job():
            raise ValueError("Đang có lượt tải 1 công ty chạy – chờ xong hoặc bấm Dừng ở trên")
        seen: list[str] = []
        for u in users:
            n = _norm(u)
            if n and n not in seen:
                seen.append(n)
        if not seen:
            raise ValueError("Chọn ít nhất một công ty")
        if d1 > d2:
            raise ValueError("Từ ngày phải trước đến ngày")
        RUN = Run(seen, d1, d2, kinds, excel, xml, pdf, easy)
        ad.MULTI["active"] = True
    threading.Thread(target=_main, args=(RUN,), daemon=True, name="multi-sync").start()
    return RUN


def stop() -> None:
    run = RUN
    if not run:
        return
    run.cancel = True
    for it in run.items:
        j = run._job(it)
        if j and j.status == "running":
            j.cancel = True
            j.submit_captcha(None)


def skip_login() -> bool:
    if RUN and RUN.cur is not None:
        RUN.skip = True
        return True
    return False


def refill() -> dict[str, Any]:
    run = RUN
    if not run or run.cur is None:
        return {"ok": False, "message": "Không có công ty nào đang chờ đăng nhập"}
    creds = credentials.load(run.items[run.cur]["user"])
    if not creds:
        return {"ok": False, "message": "Chưa lưu mật khẩu công ty này – hãy tự nhập trên Chrome"}
    try:
        BROWSER.prefill_login(*creds)
        return {"ok": True}
    except Exception as e:  # noqa: BLE001
        return {"ok": False, "message": f"Không điền được: {e}"}


# ---------------------------------------------------------------- chạy
def _prefill(run: Run, it: dict[str, Any]) -> bool:
    creds = credentials.load(it["user"])
    if not creds:
        it["msg"] = "Chưa lưu mật khẩu – tự nhập tài khoản, mật khẩu, captcha trên Chrome rồi bấm Đăng nhập"
        return True
    for i in range(3):
        try:
            BROWSER.prefill_login(*creds)
            it["msg"] = "Đã điền sẵn tài khoản + mật khẩu trên Chrome – nhập captcha rồi bấm Đăng nhập"
            return True
        except Exception as e:  # noqa: BLE001 - trang cổng đang tải
            if i == 2:
                it["msg"] = f"Chưa điền sẵn được ({e}) – tự nhập trên Chrome"
            time.sleep(3)
    return False


def _login(run: Run, it: dict[str, Any]) -> str | None:
    user = it["user"]
    tok = TOKENS.get(user)
    if _valid(tok, user):
        run.say(f"{it['ten'] or user}: dùng phiên đã đăng nhập trước đó")
        return tok
    st = BROWSER.status()
    if not st.get("open"):
        st = BROWSER.open()
    if st.get("logged_in") and _norm(st.get("mst")) == user:
        tok = BROWSER.current_token()
        if _valid(tok, user):
            run.say(f"{it['ten'] or user}: đang đăng nhập sẵn trên Chrome")
            return tok
    # cần đăng nhập công ty này trên cửa sổ Chrome cổng thuế
    BROWSER.login_free.clear()
    it["state"] = "login"
    run.say(f"Chờ đăng nhập {it['ten'] or user} ({user}) trên Chrome")
    if st.get("logged_in"):
        BROWSER.logout()
    try:
        BROWSER.open()                       # đưa cửa sổ cổng thuế lên trước
    except Exception:  # noqa: BLE001
        pass
    _prefill(run, it)
    end = time.time() + LOGIN_WAIT
    while not run.cancel:
        if run.skip:
            run.skip = False
            it["msg"] = "Bạn đã bỏ qua công ty này"
            return None
        if time.time() > end:
            it["msg"] = "Quá 15 phút chưa đăng nhập – đã bỏ qua"
            return None
        st = BROWSER.status()
        if st.get("logged_in"):
            who = _norm(st.get("mst"))
            tok = BROWSER.current_token()
            if who == user and tok:
                return tok
            if tok and who:
                # đăng nhập nhầm công ty khác: giữ lại token đó (nếu công ty đó cũng trong danh sách sẽ khỏi đăng nhập lại)
                TOKENS[who] = tok
                run.say(f"Chrome vừa đăng nhập {who}, không phải {user} – đã giữ phiên {who}, đăng xuất để đăng nhập {user}")
                BROWSER.logout()
                _prefill(run, it)
                it["msg"] = f"Vừa đăng nhập nhầm {who}. " + it["msg"]
        time.sleep(2)
    return None


def _start_ready(run: Run) -> None:
    free = MAX_PARALLEL - run.running()
    for it in run.items:
        if free <= 0 or run.cancel:
            return
        if it["state"] != "cho_tai":
            continue
        if not _valid(it["token"], it["user"], margin=60):
            it["state"], it["msg"] = "loi", "Phiên đăng nhập đã hết hạn trước khi tới lượt tải – chạy lại công ty này"
            continue
        job = ad.start(run.d1, run.d2, run.kinds, run.excel, run.xml, run.pdf, run.easy,
                       token=it["token"], user=it["user"])
        it["job_id"], it["state"], it["msg"] = job.id, "dang_tai", ""
        run.say(f"Bắt đầu tải {it['ten'] or it['user']}")
        free -= 1


def _update(run: Run) -> None:
    from .scheduler import judge
    for it in run.items:
        j = run._job(it)
        if not j or it["state"] != "dang_tai" or j.status == "running":
            continue
        res = judge(j)
        it["ket_qua"] = {k: res.get(k) for k in ("ok", "title", "missing", "notes", "summary_file", "out_dir")}
        it["state"] = {"done": "xong" if res.get("ok") else "thieu", "cancelled": "dung"}.get(j.status, "loi")
        it["msg"] = res.get("title") or ""
        if j.status in ("error", "auth_error") and j.errors:
            it["msg"] = j.errors[-1].splitlines()[0][:300]
        run.say(f"{it['ten'] or it['user']}: {it['msg']}")


def _main(run: Run) -> None:
    try:
        for i, it in enumerate(run.items):
            if run.cancel:
                break
            run.cur = i
            try:
                tok = _login(run, it)
            except Exception as e:  # noqa: BLE001 - không mở được Chrome...
                tok = None
                it["msg"] = f"Không đăng nhập được: {e}"
            if run.cancel:
                break
            if not tok:
                it["state"] = "bo_qua"
                run.say(f"{it['ten'] or it['user']}: {it['msg'] or 'bỏ qua'}")
                continue
            TOKENS[it["user"]] = tok
            it["token"], it["state"], it["msg"] = tok, "cho_tai", ""
            _update(run)
            _start_ready(run)
        run.cur = None
        BROWSER.login_free.set()
        if not run.cancel and len(run.items) > 1:
            run.say("Đã đăng nhập xong các công ty – phần mềm tự tải tiếp, bạn có thể làm việc khác")
        while True:
            _update(run)
            if run.cancel:
                stop()
            else:
                _start_ready(run)
            if not any(it["state"] in ("dang_tai", "cho_tai") for it in run.items) or \
                    (run.cancel and not run.running()):
                break
            time.sleep(2)
        _update(run)
        for it in run.items:
            if it["state"] in ("cho", "login", "cho_tai"):
                it["state"], it["msg"] = ("dung", "Đã dừng") if run.cancel else ("bo_qua", it["msg"] or "Chưa tải")
    except Exception as e:  # noqa: BLE001
        run.say(f"LỖI {e}")
    finally:
        run.cur = None
        BROWSER.login_free.set()
        run.status = "cancelled" if run.cancel else "done"
        ad.MULTI["active"] = False
        ok = [it for it in run.items if it["state"] == "xong"]
        try:
            from .scheduler import toast
            toast(f"Tải nhiều công ty: xong {len(ok)}/{len(run.items)}",
                  "; ".join(f"{it['ten'] or it['user']}: {STATE_TEXT.get(it['state'], '')}" for it in run.items)[:240])
        except Exception:  # noqa: BLE001
            pass
        run.say(f"Kết thúc: đủ {len(ok)}/{len(run.items)} công ty")
