"""Test vòng lặp trợ lý AI bằng client giả lập (không gọi API thật): python -m tests.test_assistant"""
import sys
from pathlib import Path
from types import SimpleNamespace as NS

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

import anthropic  # noqa: E402

from app import assistant  # noqa: E402

CALLS = []


class FakeMessages:
    def create(self, **kw):
        CALLS.append(kw)
        if len(CALLS) == 1:  # lượt 1: mô hình gọi công cụ của app + tra cứu web
            return NS(stop_reason="tool_use", content=[
                NS(type="server_tool_use", name="web_search", input={"query": "khấu trừ thuế GTGT 2026"}),
                NS(type="web_search_tool_result", content=[NS(url="https://thuvienphapluat.vn/x", title="Bài X")]),
                NS(type="tool_use", id="tu_1", name="get_status", input={}),
            ])
        last = kw["messages"][-1]
        assert last["role"] == "user" and last["content"][0]["tool_use_id"] == "tu_1"
        assert "noi_luu" in last["content"][0]["content"]
        return NS(stop_reason="end_turn", content=[NS(type="text", text="Đã kiểm tra: chưa đăng nhập cổng thuế.")])


class FakeClient:
    def __init__(self, **kw):
        self.beta = NS(messages=FakeMessages())


def test_loop():
    old_cls, old_key, old_p = anthropic.Anthropic, assistant.api_key, assistant.provider
    anthropic.Anthropic = FakeClient
    assistant.api_key = lambda: "sk-ant-test"
    assistant.provider = lambda: "claude"
    try:
        r = assistant.chat(None, "kiểm tra trạng thái giúp tôi")
    finally:
        anthropic.Anthropic, assistant.api_key, assistant.provider = old_cls, old_key, old_p
    assert r["reply"].startswith("Đã kiểm tra") and len(CALLS) == 2
    assert any("get_status" in a for a in r["actions"]) and r["sources"][0]["url"].startswith("https://thuvienphapluat")
    kw = CALLS[0]
    assert kw["model"] == "claude-opus-5-5" and kw["fallbacks"] == "default"
    web = [t for t in kw["tools"] if t.get("type", "").startswith("web_")]
    assert web and all("thuvienphapluat.vn" in t["allowed_domains"] for t in web)
    assert all(t.get("strict") for t in kw["tools"] if "input_schema" in t)


def test_free_loop():
    """AI miễn phí (Gemini qua chuẩn OpenAI): gọi tra_cuu_luat + get_status, hết lượt thì đổi sang model dự phòng."""
    seen = []

    def fake_post(p, model, msgs):
        seen.append(model)
        if len(seen) == 1:
            raise assistant._RateLimited("429")
        if len(seen) == 2:
            return {"choices": [{"message": {"role": "assistant", "content": None, "tool_calls": [
                {"id": "c1", "type": "function", "function": {"name": "tra_cuu_luat", "arguments": '{"cau_hoi": "ngưỡng 5 triệu"}'}},
                {"id": "c2", "type": "function", "function": {"name": "get_status", "arguments": "{}"}}]}}]}
        tool_msgs = [m for m in msgs if m["role"] == "tool"]
        assert [m["tool_call_id"] for m in tool_msgs] == ["c1", "c2"] and "Điều 14" in tool_msgs[0]["content"]
        assert msgs[0]["role"] == "system" and "tra_cuu_luat" in msgs[0]["content"]
        return {"choices": [{"message": {"role": "assistant", "content": "Từ 5 triệu phải chuyển khoản (Điều 14)."}}]}

    saved = (assistant._post_oa, assistant.law_lookup, assistant.provider, assistant.key_for)
    assistant._post_oa = fake_post
    assistant.law_lookup = lambda q, k="": {"noi_dung": "Luật Thuế GTGT 2024 Điều 14...", "nguon": [{"title": "thuvienphapluat.vn", "url": "https://x/1"}]}
    assistant.provider = lambda: "gemini"
    assistant.key_for = lambda p: "AIza-test"
    try:
        r = assistant.chat(None, "mua 6 triệu trả tiền mặt có được khấu trừ không")
    finally:
        assistant._post_oa, assistant.law_lookup, assistant.provider, assistant.key_for = saved
    assert seen[0] == "gemini-flash-latest" and seen[1] == "gemini-flash-lite-latest"
    assert r["reply"].startswith("Từ 5 triệu") and r["sources"][0]["title"] == "thuvienphapluat.vn"
    assert any("Tra cứu luật" in a for a in r["actions"]) and any("get_status" in a for a in r["actions"])


def test_ui_and_builtin():
    import json as _j
    import tempfile
    # thao tác màn hình được gom vào "ui" trả cho giao diện
    assistant._UI.items = []
    out = assistant.run_tool("mo_trang_app", {"trang": "vat", "mst": "0312345678", "ky_key": "2026-10"})
    assert out == {"ok": True} and assistant._UI.items == [{"type": "page", "page": "vat", "mst": "0312345678", "key": "2026-10"}]
    assert "loi" in assistant.run_tool("mo_file", {"duong_dan": "C:/Windows/notepad.exe"})
    # khóa dựng sẵn: lưu làm rối, đọc lại đúng, dùng khi khách chưa có khóa riêng
    old = assistant.BUILTIN_FILE
    with tempfile.TemporaryDirectory() as tmp:
        assistant.BUILTIN_FILE = Path(tmp) / "b.json"
        try:
            assistant.save_builtin("gemini", "AIzaSyTEST123")
            raw = assistant.BUILTIN_FILE.read_text()
            assert "AIzaSyTEST123" not in raw and _j.loads(raw)["provider"] == "gemini"
            assert assistant.builtin()["keys"]["gemini"] == "AIzaSyTEST123"
        finally:
            assistant.BUILTIN_FILE = old
    assert "HƯỚNG DẪN PHẦN MỀM AWBOOK" in assistant.SYSTEM and "tra_cuu_luat" in assistant.SYSTEM_FREE


def test_group():
    """Nhóm chat: trưởng nhóm trả lời; thành viên soát dựa trên nguồn đã tra; thành viên lỗi thì bỏ qua."""
    saved = (assistant.chat, assistant._post_plain, assistant.key_for)
    seen = []
    assistant.chat = lambda cid, t, ctx="": {"chat_id": "x", "reply": "Không được khấu trừ.", "evidence": ["[tra_cuu_luat] Điều 14"],
                                             "actions": [], "sources": [], "ui": []}

    def fake_plain(p, model, msgs):
        seen.append((model, msgs[1]["content"]))
        if "Kiểm soát rủi ro" in msgs[0]["content"]:  # thành viên này bận ở mọi model dự phòng -> bị bỏ qua
            raise RuntimeError("503")
        return "✅ Đồng ý"
    assistant_time_sleep = __import__("time").sleep
    __import__("time").sleep = lambda s: None
    assistant._post_plain = fake_plain
    assistant.key_for = lambda p: "k"
    try:
        r = assistant.group_chat(None, "mua 6 triệu tiền mặt")
    finally:
        assistant.chat, assistant._post_plain, assistant.key_for = saved
        __import__("time").sleep = assistant_time_sleep
    assert r["reply"] == "Không được khấu trừ." and [g["name"] for g in r["group"]] == ["Chuyên viên thuế"]
    assert sum(1 for m, _ in seen if m) == 4  # 1 lượt chuyên viên thuế + 3 lượt thử của thành viên bận
    assert "Điều 14" in seen[0][1] and "Không được khấu trừ." in seen[0][1]


if __name__ == "__main__":
    test_group()
    print("OK test_group")
    test_ui_and_builtin()
    print("OK test_ui_and_builtin")
    test_loop()
    print("OK test_loop")
    test_free_loop()
    print("OK test_free_loop")
