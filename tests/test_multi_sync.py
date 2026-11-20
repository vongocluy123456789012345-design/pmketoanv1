import base64
import json
import time
import types

from app import auto_download as ad
from app import multi_sync as ms
from app.gdt import credentials


def _tok(mst, exp_in=3600):
    enc = lambda d: base64.urlsafe_b64encode(json.dumps(d).encode()).decode().rstrip("=")  # noqa: E731
    return f"{enc({'alg': 'none'})}.{enc({'sub': mst, 'exp': int(time.time()) + exp_in})}.x"


class FakeBrowser:
    """Chrome giả: đang đăng nhập A; khi được điền sẵn tài khoản B thì vài nhịp sau 'người dùng' đăng nhập xong B."""

    def __init__(self):
        self.cookie = _tok("0300000001")
        self.logouts, self.prefills, self.ticks = 0, [], 0
        self.login_free = types.SimpleNamespace(set=lambda: None, clear=lambda: None)
        self.token_info = ms.BROWSER.token_info

    def status(self):
        self.ticks += 1
        if self.cookie is None and self.prefills and self.ticks > 3:
            self.cookie = _tok(self.prefills[-1])
        if not self.cookie:
            return {"open": True, "logged_in": False}
        return {"open": True, "logged_in": True, "mst": self.token_info(self.cookie)["mst"]}

    def open(self):
        return self.status()

    def current_token(self):
        return self.cookie

    def logout(self):
        self.logouts += 1
        self.cookie = None

    def prefill_login(self, user, pw):
        self.prefills.append(user)
        self.ticks = 0
        return True


def test_multi_logs_in_each_company_and_downloads_in_parallel(monkeypatch):
    fb = FakeBrowser()
    monkeypatch.setattr(ms, "BROWSER", fb)
    monkeypatch.setattr(ms, "TOKENS", {})
    monkeypatch.setattr(ms.time, "sleep", lambda s: None)
    monkeypatch.setattr(credentials, "load", lambda u: (u, "pw"))
    monkeypatch.setattr(credentials, "name_of", lambda u: "Cty " + u[-1])
    started = []

    def fake_start(d1, d2, kinds, ex, xml, pdf, easy, token="", user=""):
        j = ad.AutoJob(f"job{len(started)}", [(d1, d2)], kinds, ex, xml, pdf, token=token, user=user)
        j.status, j.total, j.done, j.xml_ok = "done", 2, 2, 2
        ad.JOBS[j.id] = j
        started.append((user, ms.BROWSER.token_info(token)["mst"]))
        return j

    monkeypatch.setattr(ad, "start", fake_start)
    from datetime import date
    run = ms.Run(["0300000001", "0300000002"], date(2026, 10, 1), date(2026, 10, 31), ["mua-vao"], True, True, False, False)
    ad.MULTI["active"] = True
    ms._main(run)
    # công ty 1 đang đăng nhập sẵn -> không cần đăng nhập; công ty 2 đăng xuất rồi điền sẵn tài khoản
    assert started == [("0300000001", "0300000001"), ("0300000002", "0300000002")]
    assert fb.logouts == 1 and fb.prefills == ["0300000002"]
    assert [it["state"] for it in run.items] == ["xong", "xong"]
    assert run.status == "done" and not ad.MULTI["active"]
    assert set(ms.TOKENS) == {"0300000001", "0300000002"}


def test_expired_token_not_reused():
    assert not ms._valid(_tok("0300000001", exp_in=-10), "0300000001")
    assert not ms._valid(_tok("0300000001"), "0300000002")
    assert ms._valid(_tok("0300000001"), "0300000001")


def test_running_job_ignores_multi_jobs():
    j = ad.AutoJob("tokjob", [], [], True, True, True, token="abc")
    ad.JOBS[j.id] = j
    try:
        assert ad.running_job() is None or ad.running_job().id != "tokjob"
        assert j in ad.running_any()
    finally:
        ad.JOBS.pop(j.id)
