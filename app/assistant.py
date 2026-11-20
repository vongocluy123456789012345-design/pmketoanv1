"""Trợ lý AI trong app (Claude API): hiểu câu lệnh tiếng Việt để làm việc trên app và tư vấn thuế – kế toán.

* Làm việc trên app bằng "tool use": tải hóa đơn ("tải hóa đơn tháng 10 về cho tôi"), dự báo VAT, lấy bù PDF,
  tra PDF theo mã, xem trạng thái / thông báo...
* Tư vấn luật: chỉ tra cứu trên nguồn chính thống (thuvienphapluat.vn, vbpl.vn, chinhphu.vn, gdt.gov.vn, mof.gov.vn…)
  qua web_search / web_fetch phía Anthropic, trả lời có dẫn điều khoản và link.
* Cần khóa API Anthropic: biến ANTHROPIC_API_KEY trong .env, hoặc nhập trong app (lưu tại data/ai.json trên máy).
"""
from __future__ import annotations

import json
import os
import threading
import uuid
from datetime import date, datetime
from pathlib import Path
from typing import Any

from .config import ROOT, settings

MODEL = "claude-opus-5-5"
AI_FILE = ROOT / "data" / "ai.json"
MAX_STEPS = 12
LAW_DOMAINS = ["thuvienphapluat.vn", "vbpl.vn", "chinhphu.vn", "vanban.chinhphu.vn", "gdt.gov.vn", "mof.gov.vn",
               "baohiemxahoi.gov.vn", "quochoi.vn", "luatvietnam.vn", "molisa.gov.vn", "customs.gov.vn"]

SYSTEM = """Bạn là trợ lý của phần mềm kế toán AWBook (hóa đơn điện tử, sao kê, thuế GTGT) và đóng vai một kế toán dịch vụ \
giàu kinh nghiệm tại Việt Nam. Xưng "em", gọi người dùng là "anh/chị". Trả lời bằng tiếng Việt, ngắn gọn, rõ ràng.

1. LÀM VIỆC TRÊN APP: khi người dùng nhờ làm việc (tải hóa đơn, dự báo VAT, lấy PDF, xem tiến độ…) thì dùng công cụ, \
không chỉ hướng dẫn. Tự suy ra ngày tháng từ câu nói ("tháng 10" = tháng 10 năm nay nếu tháng đó đã qua, ngược lại năm \
trước; "quý 3" = 01/07–30/09). Trước khi tải, gọi get_status: nếu chưa đăng nhập cổng thuế thì gọi open_portal rồi nhờ \
người dùng tự đăng nhập trên cửa sổ Chrome vừa mở (em không nhập mật khẩu hay captcha). Hóa đơn máy tính tiền cần người \
dùng nhập captcha của cổng: phần mềm hỏi tất cả captcha một lượt ngay đầu, sau đó tự chạy – nói rõ điều này. Báo lại kết \
quả bằng số liệu cụ thể (số hóa đơn, file lưu ở đâu, số thuế phải nộp…).

2. TƯ VẤN THUẾ / LUẬT: không trả lời luật từ trí nhớ. Luôn web_search trước (từ khóa tiếng Việt kèm năm hiện tại, ưu tiên \
thuvienphapluat.vn, vbpl.vn, chinhphu.vn, gdt.gov.vn, mof.gov.vn) và đọc trang khi cần. Mẫu trả lời: Kết luận ngắn → Căn \
cứ (tên văn bản số/năm, Điều/khoản, hiệu lực, có bị sửa đổi/thay thế không) → Áp dụng cho trường hợp của khách (tính ra số \
nếu có dữ liệu, có thể gọi công cụ của app để lấy số liệu thật) → Lưu ý (phân biệt quy định trực tiếp và suy luận; điểm \
chưa chắc thì khuyên xác nhận với cơ quan thuế) → Nguồn (link). Không bịa số hiệu văn bản, điều khoản, mức tiền. Lưu ý mốc \
mới: Luật Thuế GTGT 48/2024/QH15 (từ 01/7/2025), Luật Quản lý thuế 2025 và Thông tư 89/2026/TT-BTC (mẫu 01/GTGT mới từ kỳ \
7/2026, quý III/2026), giảm thuế GTGT 8% theo NQ 204/2025/QH15 – NĐ 174/2025 đến 31/12/2026 – vẫn phải tra lại.

3. AN TOÀN: không vượt captcha / kiểm tra chống robot của bất kỳ trang nào; không nhận hay lưu mật khẩu; nói rõ em là trợ \
lý AI, không thay thế chuyên viên thuế / luật sư cho quyết định quan trọng.

4. MÀN HÌNH: mỗi tin nhắn có kèm [Màn hình hiện tại: …] cho biết khách đang ở trang nào và trạng thái app – dùng nó để hiểu \
câu hỏi ("cái này", "nút này"). Khi khách cần tự thao tác, gọi mo_trang_app để mở đúng trang rồi chỉ từng bước; muốn cho \
khách xem file kết quả thì gọi mo_file. Hỏi số liệu hóa đơn thì dùng tim_hoa_don, không đoán.

5. TOÀN QUYỀN DỮ LIỆU: em đọc được dữ liệu của TẤT CẢ công ty trên máy (tong_quan_khach_hang, thong_tin_cong_ty, han_nop, \
kết quả sao kê, rủi ro HĐ, lãi lỗ, định khoản…) và được sửa cài đặt / số liệu (cai_dat_vat, nhap_so_lieu_vat, \
cai_lich_tai_tu_dong, cai_lich_nhac_vat, danh_dau_da_nop, hoc_dinh_khoan, them_mst_rui_ro, xuất XML). Khi khách hỏi chung \
("hôm nay cần làm gì", "khách nào sắp đến hạn", "công ty X tháng này nộp bao nhiêu") thì tự gọi công cụ lấy số thật rồi \
trả lời kèm việc nên làm tiếp. Trước khi THAY ĐỔI cài đặt/số liệu mà khách chưa nói rõ giá trị thì hỏi lại một câu; sau khi \
sửa thì báo đã đổi gì. Không xóa dữ liệu. Hướng dẫn chi tiết từng bước khi khách cần tự làm (mở đúng trang bằng mo_trang_app)."""

try:  # "dạy" AI cách dùng app: hướng dẫn nội bộ đưa vào lời dặn hệ thống
    SYSTEM += "\n\n" + (Path(__file__).resolve().parent / "ai_guide.md").read_text(encoding="utf-8")
except OSError:
    pass


# ------------------------------------------------------------------ nhà cung cấp AI + khóa API
# Miễn phí: Gemini (Google AI Studio – có gói free, không cần thẻ), Groq, OpenRouter (model ":free"), Ollama (chạy trên máy).
# Trả phí: Claude (Anthropic). Các nhà cung cấp miễn phí đều nói chuyện qua chuẩn "OpenAI-compatible chat completions".
PROVIDERS: dict[str, dict[str, Any]] = {
    "gemini": {"name": "Google Gemini (miễn phí)", "base": "https://generativelanguage.googleapis.com/v1beta/openai",
               "model": "gemini-flash-latest", "fallback": "gemini-flash-lite-latest", "env": "GEMINI_API_KEY",
               # gói free: Flash ~20 lượt/ngày, Flash-Lite nhiều lượt hơn -> hết lượt model nào thì chuyển model kế
               "chain": ["gemini-flash-lite-latest", "gemini-3.5-flash-lite", "gemini-3.1-flash-lite"],
               "key_url": "https://aistudio.google.com/apikey", "free": True},
    "groq": {"name": "Groq (miễn phí, giới hạn lượt)", "base": "https://api.groq.com/openai/v1",
             "model": "llama-3.3-70b-versatile", "env": "GROQ_API_KEY", "key_url": "https://console.groq.com/keys",
             "free": True},
    "openrouter": {"name": "OpenRouter (model :free)", "base": "https://openrouter.ai/api/v1",
                   "model": "deepseek/deepseek-chat-v3.1:free", "env": "OPENROUTER_API_KEY",
                   "key_url": "https://openrouter.ai/settings/keys", "free": True},
    "ollama": {"name": "Ollama – chạy trên máy (miễn phí, cần máy mạnh)", "base": "http://127.0.0.1:11434/v1",
               "model": "qwen3:8b", "env": "", "key_url": "https://ollama.com/download", "free": True, "no_key": True},
    "claude": {"name": "Claude – Anthropic (trả phí, mạnh nhất)", "model": MODEL, "env": "ANTHROPIC_API_KEY",
               "key_url": "https://console.anthropic.com/settings/keys", "free": False},
}


def _load() -> dict[str, Any]:
    try:
        return json.loads(AI_FILE.read_text(encoding="utf-8"))
    except (FileNotFoundError, ValueError):
        return {}


def _save(d: dict[str, Any]) -> None:
    AI_FILE.parent.mkdir(parents=True, exist_ok=True)
    AI_FILE.write_text(json.dumps(d, ensure_ascii=False), encoding="utf-8")


def provider() -> str:
    p = _load().get("provider") or builtin()["provider"]
    return p if p in PROVIDERS else "gemini"


# Khóa dựng sẵn của nhà phát hành: file app/ai_builtin.json (KHÔNG đưa lên git), đóng gói kèm bản .exe để khách
# dùng được ngay mà không phải tự lấy khóa. Khóa chỉ được làm rối (không phải mã hóa): ai mở file exe vẫn lấy được.
BUILTIN_FILE = Path(__file__).resolve().parent / "ai_builtin.json"
_MASK = b"AWBook-ai-builtin"


def _xor(b: bytes) -> bytes:
    return bytes(x ^ _MASK[i % len(_MASK)] for i, x in enumerate(b))


def builtin() -> dict[str, Any]:
    import base64
    try:
        d = json.loads(BUILTIN_FILE.read_text(encoding="utf-8"))
        return {"provider": d.get("provider", "gemini"),
                "keys": {k: _xor(base64.b64decode(v)).decode() for k, v in (d.get("keys") or {}).items()}}
    except (FileNotFoundError, ValueError, OSError):
        return {"provider": None, "keys": {}}


def save_builtin(p: str, key: str) -> None:
    """Chỉ dùng ở bản chạy từ mã nguồn (máy nhà phát hành) trước khi chạy build_exe.bat."""
    import base64
    d = {"provider": p, "keys": {}}
    try:
        d = json.loads(BUILTIN_FILE.read_text(encoding="utf-8"))
    except (FileNotFoundError, ValueError):
        pass
    d["provider"] = p
    d.setdefault("keys", {})
    if key:
        d["keys"][p] = base64.b64encode(_xor(key.encode())).decode()
    else:
        d["keys"].pop(p, None)
    BUILTIN_FILE.write_text(json.dumps(d), encoding="utf-8")


def key_for(p: str) -> str | None:
    env = PROVIDERS[p].get("env")
    d = _load()
    old = d.get("api_key") if p == "claude" else None  # khóa Claude đã lưu ở bản trước
    return ((os.getenv(env) if env else None) or (d.get("keys") or {}).get(p) or old
            or builtin()["keys"].get(p) or None)


def api_key() -> str | None:
    return key_for("claude")


def model_for(p: str) -> str:
    return (_load().get("models") or {}).get(p) or PROVIDERS[p]["model"]


def configure(p: str, key: str | None = None, model: str | None = None) -> None:
    if p not in PROVIDERS:
        raise ValueError("Nhà cung cấp AI không hợp lệ")
    d = _load()
    d["provider"] = p
    if key is not None:
        d.setdefault("keys", {})
        if key.strip():
            d["keys"][p] = key.strip()
        else:
            d["keys"].pop(p, None)
            if p == "claude":
                d.pop("api_key", None)
    if model is not None:
        d.setdefault("models", {})[p] = model.strip() or PROVIDERS[p]["model"]
    _save(d)


def set_api_key(key: str | None) -> None:  # tương thích bản trước (khóa Claude)
    configure("claude", key or "")


def status() -> dict[str, Any]:
    from .config import FROZEN
    p = provider()
    k = key_for(p)
    bk = builtin()["keys"].get(p)
    uses_builtin = bool(k) and k == bk
    return {"provider": p, "configured": bool(k) or bool(PROVIDERS[p].get("no_key")),
            "from_env": bool(PROVIDERS[p].get("env") and os.getenv(PROVIDERS[p]["env"])),
            "builtin": uses_builtin, "can_set_builtin": not FROZEN,
            "hint": "khóa có sẵn của phần mềm" if uses_builtin else ((k[:8] + "…" + k[-4:]) if k else ""),
            "model": model_for(p),
            "law_search": bool(key_for("gemini")) or p == "claude",
            "providers": {c: {"name": v["name"], "key_url": v["key_url"], "free": v["free"], "model": model_for(c),
                              "has_key": bool(key_for(c)) or bool(v.get("no_key"))} for c, v in PROVIDERS.items()}}


def list_models(p: str) -> list[str]:
    """Danh sách model của nhà cung cấp (để người dùng chọn)."""
    import httpx
    if p == "claude":
        return [MODEL, "claude-sonnet-5-5", "claude-haiku-5-5"]
    cfg = PROVIDERS[p]
    k = key_for(p)
    r = httpx.get(cfg["base"] + "/models", headers={"Authorization": f"Bearer {k}"} if k else {}, timeout=20)
    r.raise_for_status()
    ids = [m.get("id", "").replace("models/", "") for m in r.json().get("data", [])]
    if p == "openrouter":
        ids = [i for i in ids if i.endswith(":free")]
    return sorted(i for i in ids if i)


# ------------------------------------------------------------------ công cụ của app
def _tool(name: str, desc: str, props: dict[str, Any], required: list[str] | None = None) -> dict[str, Any]:
    return {"name": name, "description": desc, "strict": True,
            "input_schema": {"type": "object", "properties": props, "required": required or list(props),
                             "additionalProperties": False}}


S = {"type": "string"}
I = {"type": "integer"}
KINDS = ["mua-vao", "mua-vao-khong-ma", "ban-ra", "mua-vao-mtt", "ban-ra-mtt"]

APP_TOOLS = [
    _tool("get_status", "Trạng thái app: Chrome/cổng thuế đã đăng nhập chưa (MST nào), lượt tải đang chạy, nơi lưu, "
                        "các MST và tháng đã có dữ liệu hóa đơn.", {}),
    _tool("open_portal", "Mở cửa sổ Chrome của phần mềm tới hoadondientu.gdt.gov.vn để người dùng TỰ đăng nhập.", {}),
    _tool("start_download",
          "Tải ĐẦY ĐỦ hóa đơn từ cổng thuế (bảng kê Excel, XML, PDF) – chạy lâu. CHỈ dùng khi khách yêu cầu tải hóa đơn "
          "về máy; KHÔNG dùng để dự báo VAT (dùng doc_nhanh_hoa_don). Cần đã đăng nhập cổng. "
          "kinds: mua-vao (có mã), mua-vao-khong-ma, ban-ra, mua-vao-mtt, ban-ra-mtt (máy tính tiền – người dùng nhập "
          "captcha một lượt lúc đầu). Mặc định nên chọn mua-vao, mua-vao-khong-ma, ban-ra.",
          {"date_from": {**S, "description": "YYYY-MM-DD"}, "date_to": {**S, "description": "YYYY-MM-DD"},
           "kinds": {"type": "array", "items": {"type": "string", "enum": KINDS}}}),
    _tool("download_progress", "Tiến độ lượt tải hóa đơn đang chạy hoặc vừa xong (số hóa đơn, XML, PDF, lỗi, log).", {}),
    _tool("stop_download", "Dừng lượt tải hóa đơn đang chạy.", {}),
    _tool("vat_forecast",
          "Dự báo thuế GTGT theo tờ khai 01/GTGT cho 1 MST và 1 kỳ, từ hóa đơn đã tải. Trả về các chỉ tiêu, số phải nộp "
          "/ chuyển kỳ sau, cảnh báo, đường dẫn file PDF/Excel. ct22 = thuế còn được khấu trừ kỳ trước (-1 = tự lấy "
          "theo [43] kỳ trước hoặc số đã nhập).",
          {"mst": S, "kieu": {"type": "string", "enum": ["M", "Q"]}, "ky": {**I, "description": "tháng 1-12 hoặc quý 1-4"},
           "nam": I, "ct22": I}),
    _tool("doc_nhanh_hoa_don",
          "ĐỌC NHANH hóa đơn mua vào/bán ra của 1 kỳ trực tiếp trên cổng thuế (Chrome phải đang đăng nhập đúng MST) – chỉ "
          "đọc danh sách nên xong trong vài giây, rồi trả luôn số liệu tờ khai 01/GTGT. Dùng khi khách hỏi VAT/thuế của "
          "kỳ CHƯA có dữ liệu trên máy. Chi tiết từng mặt hàng + bảng kê đầy đủ lập tiếp ở chế độ nền.",
          {"mst": S, "kieu": {"type": "string", "enum": ["M", "Q"]}, "ky": I, "nam": I, "ct22": I}),
    _tool("tien_do_bang_ke", "Tiến độ lấy chi tiết hóa đơn và đường dẫn file tờ khai + bảng kê đầy đủ đã lập của 1 kỳ.",
          {"mst": S, "kieu": {"type": "string", "enum": ["M", "Q"]}, "ky": I, "nam": I}),
    _tool("vat_history", "Các kỳ đã dự báo VAT của 1 MST (số phải nộp / chuyển kỳ sau).", {"mst": S}),
    _tool("fill_missing_pdf", "Lấy bù PDF gốc còn thiếu (đọc mã tra cứu trong XML, Chrome chạy ẩn) cho các tháng đã tải.",
          {"mst": S, "month_from": {**S, "description": "YYYY-MM"}, "month_to": {**S, "description": "YYYY-MM"}}),
    _tool("lookup_pdf", "Lấy PDF hóa đơn từ trang tra cứu của nhà cung cấp theo link + mã tra cứu (không cần đăng nhập).",
          {"url": S, "code": S}),
    _tool("notifications", "Các thông báo gần đây của app (dự báo VAT tự động, nhắc đăng nhập…).", {}),
    _tool("tim_hoa_don",
          "Tìm và cộng hóa đơn ĐÃ TẢI về máy theo MST, khoảng tháng, hướng (purchase = mua vào, sold = bán ra, all) và từ "
          "khóa (tên/MST đối tác, ký hiệu, số HĐ; để trống = tất cả). Trả về tổng tiền và tối đa 40 hóa đơn.",
          {"mst": S, "thang_tu": {**S, "description": "YYYY-MM"}, "thang_den": {**S, "description": "YYYY-MM"},
           "huong": {"type": "string", "enum": ["purchase", "sold", "all"]}, "tu_khoa": S}),
    _tool("mo_trang_app", "Mở một màn hình của app cho khách xem/thao tác: tai (tải hóa đơn), vat (dự báo VAT), misa (chuyển "
                          "hóa đơn vào mẫu MISA), tonkho (tổng hợp tồn kho mẫu MISA), chungtu (sao kê & chứng từ), doichieu, lich (lịch tự động "
                          "tải HĐ + nhắc VAT + tài khoản tải tự động), nhapfile, dulieu; menu Tiện ích: tongquan (bảng điều khiển "
                          "tất cả khách hàng), hannop (lịch hạn nộp), ruiro (rủi ro HĐ đầu vào), dinhkhoan (định khoản mua vào), "
                          "bbcn (biên bản đối chiếu công nợ), laulo (lãi lỗ & thuế tạm tính), xmltk (xuất XML tờ khai). Với vat có "
                          "thể kèm kỳ (ky_key dạng YYYY-MM hoặc YYYY-Qn, để trống nếu không cần) và mst.",
          {"trang": {"type": "string", "enum": ["tai", "vat", "misa", "tonkho", "chungtu", "doichieu", "lich", "nhapfile", "dulieu", "tongquan", "hannop", "ruiro", "dinhkhoan", "bbcn", "laulo", "xmltk"]},
           "mst": S, "ky_key": S}),
    _tool("mo_file", "Mở file kết quả (PDF/Excel) hoặc thư mục trong nơi lưu hóa đơn cho khách xem.", {"duong_dan": S}),
    # ---- đọc toàn bộ dữ liệu các công ty
    _tool("tong_quan_khach_hang", "Bảng điều khiển TẤT CẢ công ty (MST) trên máy: tên, tháng HĐ mới nhất, đã tải đủ kỳ vừa "
                                  "xong chưa, VAT kỳ vừa xong / kỳ này, hạn nộp gần nhất chưa nộp, số hạn quá hạn, sao kê "
                                  "gần nhất, rủi ro HĐ đầu vào. Dùng khi hỏi chung về các khách hàng / việc cần làm.", {}),
    _tool("thong_tin_cong_ty", "Mọi thông tin của 1 MST: tên, địa chỉ, cơ quan thuế (HTKK), kê khai tháng/quý, các tháng đã tải "
                               "HĐ, các kỳ đã dự báo VAT, số liệu đã nhập tay cho VAT, kết quả đối chiếu sao kê gần đây.", {"mst": S}),
    _tool("han_nop", "Lịch hạn nộp (tờ khai GTGT/TNCN, tạm nộp TNDN, BHXH, môn bài, quyết toán) của 1 MST, hoặc mst rỗng = tất "
                     "cả khách hàng (chỉ hạn chưa nộp). ngay_toi = số ngày xem trước (vd 30).", {"mst": S, "ngay_toi": I}),
    _tool("danh_dau_da_nop", "Đánh dấu 1 hạn nộp là đã nộp / chưa nộp (key lấy từ han_nop).",
          {"mst": S, "key": S, "da_nop": {"type": "boolean"}}),
    _tool("kiem_tra_rui_ro_hd", "Kiểm tra rủi ro HĐ mua vào đã tải trong khoảng ngày (YYYY-MM-DD): hủy/thay thế/điều chỉnh, "
                                "trùng, sai MST/tên người mua, MST bán sai, thuế lệch thuế suất, 8% hết hạn, MST rủi ro, ≥5tr "
                                "cần chứng từ không tiền mặt.", {"mst": S, "tu": S, "den": S}),
    _tool("them_mst_rui_ro", "Thêm các MST doanh nghiệp rủi ro / bỏ địa chỉ kinh doanh vào danh sách kiểm tra (chuỗi chứa MST).",
          {"danh_sach": S}),
    _tool("uoc_tinh_lai_lo", "Ước tính lãi lỗ + thuế TNDN tạm tính (doanh nghiệp) hoặc thuế khoán % doanh thu (hộ kinh doanh) từ "
                             "HĐ đã tải. luong, chi_phi_khac: số tiền (0 nếu không có); thue_suat 0 = tự xác định; nhom_hkd "
                             "cho hộ kinh doanh.",
          {"mst": S, "tu": S, "den": S, "luong": {"type": "number"}, "chi_phi_khac": {"type": "number"},
           "thue_suat": {"type": "number"},
           "nhom_hkd": {"type": "string", "enum": ["phan_phoi", "dich_vu", "san_xuat", "khac", "cho_thue_ts"]}}),
    _tool("goi_y_dinh_khoan", "Gợi ý TK Nợ cho từng dòng hàng HĐ mua vào trong khoảng ngày (đã học / từ khóa / mặc định).",
          {"mst": S, "tu": S, "den": S, "che_do": {"type": "string", "enum": ["TT200", "TT133"]}}),
    _tool("hoc_dinh_khoan", "Ghi nhớ cách hạch toán: danh sách {ten_hang, tk} – lần sau tự gợi ý và tự điền TK khi chuyển MISA.",
          {"cap": {"type": "array", "items": {"type": "object", "properties": {"ten_hang": S, "tk": S},
                                               "required": ["ten_hang", "tk"], "additionalProperties": False}}}),
    _tool("xuat_xml_to_khai", "Xuất XML tờ khai 01/GTGT (cấu trúc HTKK) từ kết quả dự báo VAT đã lập. ky_key: YYYY-MM hoặc "
                              "YYYY-Qn.", {"mst": S, "ky_key": S, "nguoi_ky": S}),
    _tool("ket_qua_sao_ke", "Các lần đối chiếu sao kê ngân hàng gần đây của 1 MST: kỳ, số giao dịch, số khớp hóa đơn, thống kê, "
                            "đường dẫn file kết quả.", {"mst": S}),
    # ---- chỉnh sửa / cài đặt (quyền cao)
    _tool("cai_dat_vat", "Đổi cài đặt kê khai VAT của 1 MST: kieu M (tháng) / Q (quý), auto = tự dự báo đầu kỳ sau.",
          {"mst": S, "kieu": {"type": "string", "enum": ["M", "Q"]}, "auto": {"type": "boolean"}}),
    _tool("nhap_so_lieu_vat", "Nhập / sửa số liệu tay cho tờ khai VAT 1 kỳ rồi dự báo lại (vd 22 = thuế còn khấu trừ kỳ trước, "
                              "23a/24a nhập khẩu, 37/38 điều chỉnh, 40b…). Giữ các số đã nhập khác.",
          {"mst": S, "kieu": {"type": "string", "enum": ["M", "Q"]}, "ky": I, "nam": I,
           "chi_tieu": {"type": "array", "items": {"type": "object", "properties": {"ma": S, "so": {"type": "number"}},
                                                   "required": ["ma", "so"], "additionalProperties": False}}}),
    _tool("xem_lich_tu_dong", "Xem cài đặt lịch tự động tải HĐ (ngày giờ, kỳ, loại HĐ, công ty đã lưu tài khoản) và lịch nhắc "
                              "dự báo VAT.", {}),
    _tool("cai_lich_tai_tu_dong", "Cài lịch tự động tải HĐ: bat, ngay (1-31), gio, phut, ky (truoc = tháng trước, nay, quy_truoc, "
                                  "quy_nay, thang/quy/khoang dùng tu-den YYYY-MM-DD), loai HĐ. Hỏi lại khách trước khi đổi.",
          {"bat": {"type": "boolean"}, "ngay": I, "gio": I, "phut": I,
           "ky": {"type": "string", "enum": ["truoc", "nay", "quy_truoc", "quy_nay", "thang", "quy", "khoang"]},
           "tu": S, "den": S, "loai": {"type": "array", "items": {"type": "string", "enum": KINDS}}}),
    _tool("cai_lich_nhac_vat", "Cài lịch nhắc / tự dự báo VAT: bat, ngay = các ngày trong tháng (vd '1,15'), gio 'HH:MM'.",
          {"bat": {"type": "boolean"}, "ngay": S, "gio": S}),
]
WEB_TOOLS = [
    {"type": "web_search_20260209", "name": "web_search", "max_uses": 6, "allowed_domains": LAW_DOMAINS},
    {"type": "web_fetch_20260209", "name": "web_fetch", "max_uses": 6, "allowed_domains": LAW_DOMAINS},
]


def _months_with_data() -> dict[str, list[str]]:
    base = Path(settings.output_dir)
    out: dict[str, list[str]] = {}
    from .vat import _msts
    for m in _msts():
        out[m] = sorted(p.name for p in (base / m).iterdir() if p.is_dir() and len(p.name) == 7 and p.name[4] == "-")
    return out


def find_invoices(mst: str, m1: str, m2: str, huong: str = "all", kw: str = "") -> dict[str, Any]:
    from .auto_download import month_list
    from .vat import _key, _num
    base = Path(settings.output_dir) / mst
    rows: dict[tuple, dict[str, Any]] = {}
    missing = []
    for d1, _ in month_list(m1, m2):
        f = base / f"{d1:%Y-%m}" / f".dulieu_{d1:%Y-%m}.json"
        if not f.exists():
            missing.append(f"{d1:%m/%Y}")
            continue
        for r in json.loads(f.read_text(encoding="utf-8")).get("rows", []):
            rows.setdefault(_key(r), r)
    kw = (kw or "").strip().lower()
    out = []
    for r in rows.values():
        if huong in ("purchase", "sold") and r.get("huong") != huong:
            continue
        hay = " ".join(str(r.get(k) or "") for k in ("nbten", "nbmst", "nmten", "nmmst", "khhdon", "shdon")).lower()
        if kw and kw not in hay:
            continue
        out.append(r)
    out.sort(key=lambda r: r.get("ngay_lap") or "")
    tot = {h: {"so_hd": 0, "chua_thue": 0.0, "thue": 0.0, "thanh_toan": 0.0} for h in ("purchase", "sold")}
    for r in out:
        t = tot["purchase" if r.get("huong") == "purchase" else "sold"]
        t["so_hd"] += 1
        t["chua_thue"] += _num(r.get("tgtcthue"))
        t["thue"] += _num(r.get("tgtthue"))
        t["thanh_toan"] += _num(r.get("tgtttbso"))
    return {"tong_mua_vao": tot["purchase"], "tong_ban_ra": tot["sold"], "thang_chua_co_du_lieu": missing,
            "hoa_don": [{"ngay": r.get("ngay_lap"), "huong": "mua vào" if r.get("huong") == "purchase" else "bán ra",
                         "ky_hieu": r.get("khhdon"), "so": r.get("shdon"),
                         "doi_tac": r.get("nbten") if r.get("huong") == "purchase" else r.get("nmten"),
                         "mst_doi_tac": r.get("nbmst") if r.get("huong") == "purchase" else r.get("nmmst"),
                         "chua_thue": r.get("tgtcthue"), "thue": r.get("tgtthue"), "tong": r.get("tgtttbso"),
                         "trang_thai": r.get("tthai"), "pdf": r.get("file_pdf")} for r in out[:40]],
            "so_hd_tim_thay": len(out)}


def _vat_quick(mst: str, kieu: str, ky: int, nam: int, ct22: int = -1, export: bool = True) -> dict[str, Any]:
    """Tính tờ khai ngay (1–2 giây, không in file) để trả lời liền; in PDF/Excel bảng kê chạy nền rồi báo chuông."""
    from . import vat
    key = vat.period_key(kieu, ky, nam)
    inputs = dict(vat.get_cfg(mst)["inputs"].get(key, {}))
    if ct22 is not None and ct22 >= 0:
        inputs["ct22"] = ct22
        vat.set_cfg(mst, inputs={key: inputs})
    elif "ct22" not in inputs:
        inputs["ct22"] = vat.suggest_ct22(mst, kieu, ky, nam)["value"]
    res = vat.compute(mst, kieu, ky, nam, inputs)
    if res["missing"] and not res["rows"]:
        return {"chua_co_du_lieu": res["missing"],
                "huong_dan": "Chưa có hóa đơn của kỳ này trên máy. Gọi get_status: nếu Chrome đã đăng nhập đúng MST thì "
                             "gọi doc_nhanh_hoa_don; nếu chưa thì gọi open_portal, nhờ khách đăng nhập MST này rồi nhắn lại."}
    if export:
        threading.Thread(target=lambda: vat.forecast(mst, kieu, ky, nam, inputs, source="ai"), daemon=True).start()
    ct = res["ct"]
    return {"ky": res["label"], "mau_to_khai": res["mau"], "cong_ty": res["company"]["ten"],
            "chi_tieu": {k: ct[k] for k in ("22", "23", "24", "25", "26", "29", "30", "31", "32", "33", "32a", "34", "35",
                                            "36", "40", "41", "43")},
            "phai_nop_40": ct["40"], "chuyen_ky_sau_43": ct["43"], "han_nop": res["deadline"],
            "so_hd": res["stat"], "ban_ra_8pt_giam_thue": res["pl8"]["giam"], "canh_bao": res["warnings"][:6],
            "file": "PDF/Excel tờ khai + bảng kê đang lập ở chế độ nền (vài giây), có thông báo ở chuông khi xong; "
                    "khách có thể mở trang vat để xem."}


def run_tool(name: str, a: dict[str, Any]) -> Any:
    import time
    t0 = time.time()
    out = _run_tool(name, a)
    _log(f"TOOL {name} {time.time() - t0:.1f}s")
    if isinstance(out, dict) and out.get("ui"):
        getattr(_UI, "items", []).append(out.pop("ui"))
    return out


def _run_tool(name: str, a: dict[str, Any]) -> Any:
    from . import auto_download as ad
    from . import pdf_tools, vat
    from .gdt.browser import BROWSER

    if name == "get_status":
        st = BROWSER.status()
        job = ad.running_job()
        return {"chrome_mo": st.get("open"), "da_dang_nhap": st.get("logged_in"), "mst_dang_nhap": st.get("mst"),
                "luot_tai_dang_chay": job.public() | {"log": job.log[-5:]} if job else None,
                "noi_luu": settings.output_dir, "du_lieu_theo_mst": _months_with_data(), "hom_nay": date.today().isoformat()}
    if name == "open_portal":
        return BROWSER.open()
    if name == "start_download":
        if ad.running_job():
            return {"loi": "Đang có lượt tải chạy – dùng download_progress để xem."}
        st = BROWSER.status()
        if not st.get("logged_in"):
            return {"loi": "Chưa đăng nhập cổng thuế. Gọi open_portal và nhờ người dùng tự đăng nhập."}
        d1, d2 = date.fromisoformat(a["date_from"]), date.fromisoformat(a["date_to"])
        if d1 > d2 or d2 > date.today():
            return {"loi": "Khoảng ngày không hợp lệ (đến ngày không được sau hôm nay)."}
        kinds = [k for k in ad.KINDS if k in (a.get("kinds") or ["mua-vao", "mua-vao-khong-ma", "ban-ra"])]
        job = ad.start(d1, d2, kinds, True, True, True)
        return {"da_bat_dau": True, "job_id": job.id, "mst": st.get("mst"), "ky": [ad.period(x, y) for x, y in job.months],
                "loai": kinds, "ghi_chu": "Có loại máy tính tiền thì app sẽ hiện ô captcha ở mục Tải hóa đơn."
                if any("mtt" in k for k in kinds) else ""}
    if name == "download_progress":
        job = ad.running_job() or (list(ad.JOBS.values())[-1] if ad.JOBS else None)
        if not job:
            return {"thong_bao": "Chưa có lượt tải nào trong phiên này."}
        p = job.public()
        return {k: p[k] for k in ("status", "mst", "months", "total", "done", "xml_ok", "pdf_ok", "pdf_hddt", "pdf_missing",
                                  "excel_ok", "summary_file", "out_dir")} | {"loi": job.errors[-5:], "log": job.log[-8:],
                                                                             "dang_cho_captcha": bool(job.captcha)}
    if name == "stop_download":
        job = ad.running_job()
        if job:
            job.cancel = True
            job.submit_captcha(None)
        return {"ok": bool(job)}
    if name == "vat_forecast":
        return _vat_quick(a["mst"], a["kieu"], a["ky"], a["nam"], a.get("ct22", -1))
    if name == "doc_nhanh_hoa_don":
        from . import quick_read
        st = quick_read.read_period(a["mst"], a["kieu"], a["ky"], a["nam"])
        if st.get("loi"):
            return st
        return {"da_doc": st} | _vat_quick(a["mst"], a["kieu"], a["ky"], a["nam"], a.get("ct22", -1), export=False) | {
            "ghi_chu": f"Đã đọc nhanh {st['mua_vao']} HĐ mua vào, {st['ban_ra']} HĐ bán ra trong {st['giay']} giây. Đang lấy "
                       f"chi tiết từng mặt hàng ({st['can_lay_chi_tiet']} HĐ) và lập tờ khai + bảng kê đầy đủ ở chế độ nền, "
                       "xong sẽ báo ở chuông thông báo. Hóa đơn máy tính tiền (cần captcha) chưa gồm trong lượt đọc nhanh."}
    if name == "tien_do_bang_ke":
        from . import quick_read
        key = vat.period_key(a["kieu"], a["ky"], a["nam"])
        s = vat.saved_result(a["mst"], key) or {}
        return {"lay_chi_tiet": quick_read.STATE.get(a["mst"]), "files": s.get("files"), "lap_luc": s.get("created")}
    if name == "vat_history":
        return vat.history(a["mst"])
    if name == "fill_missing_pdf":
        from .auto_download import month_list
        months = [f"{x:%Y-%m}" for x, _ in month_list(a["month_from"], a["month_to"])]
        job = pdf_tools.start_fill(a["mst"], months)
        return {"da_bat_dau": True, "job_id": job.id, "ghi_chu": "Xem tiến độ ở mục Tải hóa đơn > Lấy bù PDF."}
    if name == "lookup_pdf":
        return pdf_tools.lookup_one(a["url"], a["code"])
    if name == "notifications":
        return [{k: n.get(k) for k in ("time", "title", "body", "warnings")} for n in vat.notifications()[:8]]
    if name == "tim_hoa_don":
        return find_invoices(a["mst"], a["thang_tu"], a["thang_den"], a.get("huong", "all"), a.get("tu_khoa", ""))
    if name == "mo_trang_app":
        return {"ok": True, "ui": {"type": "page", "page": a["trang"], "mst": a.get("mst", ""), "key": a.get("ky_key", "")}}
    if name in TI_TOOLS:
        return _run_ti(name, a)
    if name == "mo_file":
        p = Path(a["duong_dan"])
        base = Path(settings.output_dir).resolve()
        if not p.resolve().is_relative_to(base) or not p.exists():
            return {"loi": "Chỉ mở được file/thư mục có thật trong nơi lưu hóa đơn."}
        return {"ok": True, "ui": {"type": "file", "path": str(p)}}
    return {"loi": f"Không có công cụ {name}"}


TI_TOOLS = {"tong_quan_khach_hang", "thong_tin_cong_ty", "han_nop", "danh_dau_da_nop", "kiem_tra_rui_ro_hd",
            "them_mst_rui_ro", "uoc_tinh_lai_lo", "goi_y_dinh_khoan", "hoc_dinh_khoan", "xuat_xml_to_khai", "ket_qua_sao_ke",
            "cai_dat_vat", "nhap_so_lieu_vat", "xem_lich_tu_dong", "cai_lich_tai_tu_dong", "cai_lich_nhac_vat"}


def _range(a: dict[str, Any]) -> tuple[Any, Any]:
    from datetime import date
    t = date.today()
    try:
        d1 = date.fromisoformat(str(a.get("tu") or "")[:10])
    except ValueError:
        d1 = date(t.year, t.month, 1)
    try:
        d2 = date.fromisoformat(str(a.get("den") or "")[:10])
    except ValueError:
        d2 = t
    return d1, max(d1, d2)


def _run_ti(name: str, a: dict[str, Any]) -> Any:
    import re
    import json as _json

    from . import tienich, vat
    mst = re.sub(r"[^\d-]", "", str(a.get("mst") or ""))
    if name == "tong_quan_khach_hang":
        return tienich.dashboard()
    if name == "thong_tin_cong_ty":
        cfg = vat.get_cfg(mst)
        return {**tienich.company(mst), "kieu_ke_khai": cfg["kieu"], "tu_du_bao": cfg["auto"],
                "so_lieu_nhap_tay": cfg["inputs"], "thang_da_tai": tienich.months_of(mst), "du_bao_vat": vat.history(mst)[:12],
                "sao_ke": _run_ti("ket_qua_sao_ke", a)}
    if name == "han_nop":
        n = int(a.get("ngay_toi") or 45)
        rows = tienich.deadlines(mst, ahead=n) if mst else tienich.all_deadlines(ahead=n)
        return rows[:60]
    if name == "danh_dau_da_nop":
        return tienich.mark_done(mst, a["key"], bool(a.get("da_nop", True)))
    if name == "kiem_tra_rui_ro_hd":
        r = tienich.risk_check(mst, *_range(a))
        return {**{k: v for k, v in r.items() if k != "issues"}, "issues": r["issues"][:40],
                "ghi_chu": "Xem đầy đủ / tải Excel ở menu Tiện ích → Kiểm tra rủi ro hóa đơn đầu vào"}
    if name == "them_mst_rui_ro":
        return {"danh_sach": tienich.set_risky_list("\n".join(tienich.risky_list()) + "\n" + a.get("danh_sach", ""))}
    if name == "uoc_tinh_lai_lo":
        return tienich.profit(mst, *_range(a), float(a.get("chi_phi_khac") or 0), float(a.get("luong") or 0),
                              a.get("nhom_hkd") or "phan_phoi", float(a.get("thue_suat") or 0) or None)
    if name == "goi_y_dinh_khoan":
        rows = tienich.purchase_lines(mst, *_range(a), a.get("che_do") or "TT200")
        return {"so_dong": len(rows), "dong": rows[:60]}
    if name == "hoc_dinh_khoan":
        return {"da_hoc": tienich.learn_tk([(x["ten_hang"], x["tk"]) for x in a.get("cap", [])])}
    if name == "xuat_xml_to_khai":
        try:
            f, note = tienich.vat_xml(mst, a["ky_key"], a.get("nguoi_ky", ""))
        except ValueError as e:
            return {"loi": str(e)}
        return {"file": str(f), "ghi_chu": note}
    if name == "ket_qua_sao_ke":
        d = Path(settings.output_dir) / mst / "DocChungTu"
        out = []
        for f in sorted(d.glob("KetQua_GhepThuChi_*.json"))[-5:] if d.is_dir() else []:
            try:
                out.append(_json.loads(f.read_text(encoding="utf-8")))
            except ValueError:
                pass
        if not out and d.is_dir():
            out = [{"file": str(f)} for f in sorted(d.glob("KetQua_GhepThuChi_*.xlsx"))[-5:]]
        return out or {"ghi_chu": "Chưa đối chiếu sao kê cho MST này"}
    if name == "cai_dat_vat":
        return vat.set_cfg(mst, kieu=a.get("kieu", "M"), auto=bool(a.get("auto", True)))
    if name == "nhap_so_lieu_vat":
        kieu, ky, nam = a.get("kieu", "M"), int(a["ky"]), int(a["nam"])
        key = vat.period_key(kieu, ky, nam)
        inputs = dict(vat.get_cfg(mst)["inputs"].get(key, {}))
        for x in a.get("chi_tieu", []):
            inputs[str(x["ma"]).strip().lstrip("[").rstrip("]")] = float(x["so"])
        r = vat.forecast(mst, kieu, ky, nam, inputs, source="manual")
        return {k: r.get(k) for k in ("label", "phai_nop", "chuyen_ky_sau", "warnings", "files")} | {"so_lieu_nhap": inputs}
    if name == "xem_lich_tu_dong":
        from . import vat_lich
        from .scheduler import SCHEDULER
        d = vat_lich.get()
        return {"tai_hoa_don": SCHEDULER.status(), "nhac_vat": {**d, "mo_ta": vat_lich.describe(d)}}
    if name == "cai_lich_tai_tu_dong":
        from .scheduler import SCHEDULER
        SCHEDULER.configure(bool(a.get("bat")), int(a.get("ngay") or 5), int(a.get("gio") or 8),
                            a.get("loai") or ["mua-vao", "mua-vao-khong-ma", "ban-ra"], int(a.get("phut") or 0),
                            a.get("ky") or "truoc", a.get("tu", ""), a.get("den", ""))
        return SCHEDULER.status()
    if name == "cai_lich_nhac_vat":
        from . import vat_lich
        try:
            d = vat_lich.save(bool(a.get("bat")), a.get("ngay", ""), a.get("gio") or "08:00")
        except ValueError as e:
            return {"loi": str(e)}
        return {**d, "mo_ta": vat_lich.describe(d)}
    return {"loi": f"Không có công cụ {name}"}


# ------------------------------------------------------------------ hội thoại
CHATS: dict[str, dict[str, Any]] = {}
_lock = threading.Lock()


def _chat_claude(chat_id: str | None, text: str) -> dict[str, Any]:
    import anthropic

    key = api_key()
    if not key:
        raise RuntimeError("Chưa có khóa API Anthropic – nhập ở ô cài đặt của Trợ lý AI.")
    with _lock:
        c = CHATS.get(chat_id or "") or {"id": uuid.uuid4().hex, "messages": [], "busy": False}
        if c["busy"]:
            raise RuntimeError("Trợ lý đang xử lý câu trước, chờ chút nhé.")
        c["busy"] = True
        CHATS[c["id"]] = c
    client = anthropic.Anthropic(api_key=key, max_retries=3, timeout=300)
    msgs = c["messages"]
    start = len(msgs)  # lỗi giữa chừng -> bỏ cả lượt này để lịch sử hội thoại luôn hợp lệ
    msgs.append({"role": "user", "content": f"[Hôm nay {datetime.now():%d/%m/%Y %H:%M}] {text}"})
    actions: list[str] = []
    sources: dict[str, str] = {}
    reply = ""
    try:
        for _ in range(MAX_STEPS):
            resp = client.beta.messages.create(
                model=MODEL, max_tokens=16000, system=SYSTEM, messages=msgs, tools=APP_TOOLS + WEB_TOOLS,
                output_config={"effort": "medium"}, cache_control={"type": "ephemeral"},
                betas=["server-side-fallback-2026-07-01"], fallbacks="default",
            )
            msgs.append({"role": "assistant", "content": resp.content})
            for b in resp.content:
                if b.type == "web_search_tool_result" and isinstance(b.content, list):
                    for r in b.content:
                        if getattr(r, "url", None):
                            sources.setdefault(r.url, getattr(r, "title", "") or r.url)
                elif b.type == "server_tool_use":
                    q = (b.input or {}).get("query") or (b.input or {}).get("url") or ""
                    actions.append(("Tra cứu: " if b.name == "web_search" else "Đọc trang: ") + str(q))
            if resp.stop_reason == "refusal":
                reply = "Xin lỗi, em không xử lý được yêu cầu này."
                break
            if resp.stop_reason == "pause_turn":
                continue
            uses = [b for b in resp.content if b.type == "tool_use"]
            if resp.stop_reason != "tool_use" or not uses:
                reply = "".join(b.text for b in resp.content if b.type == "text").strip()
                break
            results = []
            for u in uses:
                actions.append(f"Thực hiện: {u.name} {json.dumps(u.input, ensure_ascii=False)}")
                try:
                    out = run_tool(u.name, dict(u.input))
                    results.append({"type": "tool_result", "tool_use_id": u.id,
                                    "content": json.dumps(out, ensure_ascii=False, default=str)[:20000]})
                except Exception as e:  # noqa: BLE001 - báo lỗi cho mô hình tự xử lý
                    results.append({"type": "tool_result", "tool_use_id": u.id, "content": f"Lỗi: {e}",
                                    "is_error": True})
            msgs.append({"role": "user", "content": results})
        else:
            reply = reply or "Em đã làm nhiều bước nhưng chưa xong – anh/chị hỏi tiếp để em tiếp tục nhé."
    except anthropic.AuthenticationError:
        del msgs[start:]
        raise RuntimeError("Khóa API Anthropic không hợp lệ.") from None
    except anthropic.RateLimitError:
        del msgs[start:]
        raise RuntimeError("Đang vượt giới hạn gọi API, thử lại sau ít phút.") from None
    except anthropic.APIConnectionError:
        del msgs[start:]
        raise RuntimeError("Không kết nối được tới Claude API – kiểm tra mạng.") from None
    except anthropic.APIStatusError as e:
        del msgs[start:]
        raise RuntimeError(f"Claude API báo lỗi {e.status_code}: {e.message}") from None
    finally:
        c["busy"] = False
    return {"chat_id": c["id"], "reply": reply, "actions": actions,
            "sources": [{"url": u, "title": t} for u, t in list(sources.items())[:12]]}


# ------------------------------------------------------------------ AI miễn phí (chuẩn OpenAI-compatible)
SYSTEM_FREE = SYSTEM.replace(
    "Luôn web_search trước (từ khóa tiếng Việt kèm năm hiện tại, ưu tiên thuvienphapluat.vn",
    "Luôn gọi công cụ tra_cuu_luat trước (câu hỏi rõ ràng, tiếng Việt, kèm năm hiện tại; nguồn: thuvienphapluat.vn"
).replace("và đọc trang khi cần.", "Chỉ dùng nội dung và nguồn mà tra_cuu_luat trả về.")
LAW_TOOLS = [
    _tool("tra_cuu_luat",
          "Tra cứu quy định pháp luật Việt Nam HIỆN HÀNH (thuế, kế toán, hóa đơn, BHXH, lao động, doanh nghiệp) trên "
          "các trang thuế – kế toán lớn (MISA AMIS, MISA SME, meInvoice – bài mới nhất, đọc trực tuyến) và kho kiến "
          "thức đã kiểm chứng trên thuvienphapluat.vn. Trả về đoạn bài viết, ngày cập nhật, link. Gọi trước khi trả lời "
          "bất kỳ câu hỏi luật nào; câu hỏi nhiều ý thì gọi nhiều lần, mỗi lần một ý; kết quả chưa sát thì gọi lại với "
          "từ khóa khác.",
          {"cau_hoi": S, "tu_khoa": {**S, "description": "2-5 từ khóa tiếng Việt ngắn gọn, vd 'phạt chậm nộp tờ khai'"}}),
]


def _oa_tools() -> list[dict[str, Any]]:
    return [{"type": "function", "function": {"name": t["name"], "description": t["description"],
                                              "parameters": t["input_schema"]}} for t in APP_TOOLS + LAW_TOOLS]


KB_FILE = Path(__file__).resolve().parent / "ai_luat.md"
_GROUNDING = {"off_until": 0.0}  # gói free không có Google Search -> tạm bỏ qua 6 giờ cho đỡ tốn lượt


def _plain(s: str) -> str:
    import re
    import unicodedata
    s = unicodedata.normalize("NFD", s.lower().replace("đ", "d"))
    return re.sub(r"\s+", " ", "".join(c for c in s if unicodedata.category(c) != "Mn"))


def kb_search(question: str, top: int = 3) -> list[dict[str, str]]:
    """Tìm trong kho kiến thức thuế (app/ai_luat.md) các mục khớp câu hỏi nhất."""
    import re
    try:
        text = KB_FILE.read_text(encoding="utf-8")
    except OSError:
        return []
    q = _plain(question)
    qwords = {w for w in re.findall(r"[a-z0-9/%,.]+", q) if len(w) > 1}
    scored = []
    for sec in text.split("\n## ")[1:]:
        title, _, body = sec.partition("\n")
        m = re.search(r"Từ khóa:(.*)", body)
        kws = [_plain(k.strip()) for k in (m.group(1).split(",") if m else []) if k.strip()]
        score = sum(3 + len(k.split()) for k in kws if k and k in q)  # cụm từ khóa xuất hiện trong câu hỏi
        score += sum(1 for w in qwords if w in _plain(title + " " + body[:400]))
        if score:
            scored.append((score, title.strip(), body.strip()))
    scored.sort(key=lambda x: -x[0])
    return [{"muc": t, "noi_dung": b} for _, t, b in scored[:top]]


# Trang kiến thức thuế – kế toán lớn, cập nhật thường xuyên, luôn dẫn văn bản; có API tìm kiếm công khai (WordPress).
# (thuvienphapluat.vn, luatvietnam.vn bật kiểm tra chống robot Cloudflare -> phần mềm không tra tự động, chỉ đưa link.)
LAW_SITES = [("amis.misa.vn", "MISA AMIS"), ("sme.misa.vn", "MISA SME"), ("www.meinvoice.vn", "MISA meInvoice")]
_UA = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/130.0 "
                     "Safari/537.36 AWBook", "Accept-Language": "vi-VN,vi;q=0.9"}


_STOP = set("co duoc khong la gi the nao bao nhieu va cua cho khi thi neu nhu cac nhung mot de voi tu den trong tren "
             "duoi theo ve hay hoac phai can toi minh anh chi em cong ty doanh nghiep nam thang quy moi nhat hien nay "
             "ra sao lam sao quy dinh huong dan".split())
_GENERIC = {"thue", "gtgt", "hoa", "don", "2025", "2026", "2027", "ke", "toan"}


def _terms(s: str) -> list[str]:
    import re
    return [w for w in re.findall(r"[a-z0-9%/.,]+", _plain(s)) if len(w) > 1 and w not in _STOP]


def web_law_search(query: str, top: int = 3, keywords: str = "") -> list[dict[str, Any]]:
    """Tìm bài viết thuế mới nhất trên các trang lớn rồi đọc nội dung (đoạn liên quan) để AI trả lời có dẫn nguồn."""
    import html as _html
    import re
    from concurrent.futures import ThreadPoolExecutor

    import httpx
    from bs4 import BeautifulSoup

    words = list(dict.fromkeys(_terms(keywords + " " + query)))
    key_words = [w for w in words if w not in _GENERIC] or words
    # vài cách tìm: từ khóa AI rút ra; các từ quan trọng nhất của câu hỏi (API tìm kiếm WordPress cần câu ngắn)
    plain_q = [w for w in re.findall(r"\w+", query.lower()) if _plain(w) not in _STOP and len(w) > 1]
    queries = list(dict.fromkeys(q for q in (keywords.strip(), " ".join(plain_q[:4]), " ".join(plain_q[-4:])) if q))

    def search(arg: tuple[tuple[str, str], str]) -> list[dict[str, Any]]:
        (host, name), q = arg
        try:
            r = httpx.get(f"https://{host}/wp-json/wp/v2/search", params={"search": q, "per_page": 8,
                                                                           "subtype": "post"}, headers=_UA, timeout=15)
            items = r.json() if r.status_code == 200 else []
        except Exception:  # noqa: BLE001
            return []
        out = []
        for it in items if isinstance(items, list) else []:
            title = _html.unescape(it.get("title") or "")
            tp = _plain(title)
            hits = sum(1 for w in key_words if w in tp)
            score = hits * 2 + sum(1 for w in words if w in tp) + (1 if re.search(r"202[5-9]", title) else 0)
            if hits >= min(2, len(key_words)):  # phải khớp ít nhất 2 từ quan trọng -> tránh bài lạc đề
                out.append({"site": name, "host": host, "id": it.get("id"), "title": title, "url": it.get("url"),
                            "score": score})
        return out

    def read(it: dict[str, Any]) -> dict[str, Any]:
        try:
            r = httpx.get(f"https://{it['host']}/wp-json/wp/v2/posts/{it['id']}",
                          params={"_fields": "content,modified,date"}, headers=_UA, timeout=20)
            d = r.json()
            text = BeautifulSoup(d["content"]["rendered"], "html.parser").get_text("\n", strip=True)
        except Exception:  # noqa: BLE001
            return it | {"noi_dung": ""}
        paras = [p for p in text.split("\n") if len(p) > 30]
        ranked = sorted(range(len(paras)), key=lambda i: -sum(1 for w in key_words if w in _plain(paras[i])))
        keep = sorted(set(ranked[:14]) | set(range(min(3, len(paras)))))  # đoạn đầu + đoạn khớp nhất
        excerpt = "\n".join(paras[i] for i in keep)[:4500]
        return it | {"cap_nhat": (d.get("modified") or d.get("date") or "")[:10], "noi_dung": excerpt}

    with ThreadPoolExecutor(9) as ex:
        found = [x for lst in ex.map(search, [(s, q) for s in LAW_SITES for q in queries]) for x in lst]
        found.sort(key=lambda x: -x["score"])
        picked, seen = [], set()
        for x in found:
            if x["url"] and x["title"] not in seen and x["score"] > 0:
                seen.add(x["title"])
                picked.append(x)
            if len(picked) >= top:
                break
        docs = list(ex.map(read, picked))
    return [{k: d.get(k) for k in ("site", "title", "url", "cap_nhat", "noi_dung")} for d in docs if d.get("noi_dung")]


def law_lookup(question: str, keywords: str = "") -> dict[str, Any]:
    """Tra luật: bài mới nhất trên trang thuế lớn + kho kiến thức đã kiểm chứng (+ Google Search nếu khóa có quyền)."""
    import time

    kb = kb_search(question + " " + keywords)
    try:
        web = web_law_search(question, keywords=keywords)
    except Exception:  # noqa: BLE001 - mất mạng -> vẫn còn kho kiến thức
        web = []
    kb_out = {"nguon_tra": "Bài viết mới nhất trên MISA AMIS / MISA SME / meInvoice (tra trực tuyến lúc hỏi) + kho kiến "
                           "thức AWBook (đã kiểm chứng trên thuvienphapluat.vn 17/11/2026)",
              "bai_viet_truc_tuyen": web, "ket_qua": kb,
              "huong_dan": "Đối chiếu bài trực tuyến với kho kiến thức; ưu tiên thông tin MỚI NHẤT (xem ngày cap_nhat) "
                           "và văn bản còn hiệu lực; nêu số hiệu văn bản, điều khoản như trong nguồn; ghi rõ nguồn từng ý. "
                           "Nguồn mâu thuẫn hoặc mục 'CẦN ĐỐI CHIẾU' thì nói rõ chưa chắc. Không có thông tin thì nói "
                           "chưa tra được, KHÔNG bịa luật, gợi ý khách hỏi cơ quan thuế quản lý."}
    if time.time() < _GROUNDING["off_until"]:
        return kb_out
    out = _grounded(question)
    if out.get("loi"):
        if "Hết lượt" in out["loi"] or "quota" in out["loi"].lower():
            _GROUNDING["off_until"] = time.time() + 6 * 3600
        return kb_out | {"ghi_chu": "Không tra Google Search được (" + out["loi"][:120] + ")"}
    return out | {"kho_kien_thuc": kb}


def _grounded(question: str) -> dict[str, Any]:
    """Tra luật bằng Gemini + Google Search (grounding)."""
    import httpx
    key = key_for("gemini")
    if not key:
        return {"loi": "Chưa có khóa Gemini nên chưa tra Google Search được."}
    prompt = (f"Hôm nay là {date.today():%d/%m/%Y}. Tra cứu quy định pháp luật Việt Nam đang có hiệu lực cho câu hỏi sau và "
              f"trả lời bằng tiếng Việt: {question}\n\nYêu cầu: chỉ dựa vào nguồn chính thống/uy tín ("
              + ", ".join(LAW_DOMAINS) + "). Nêu rõ: tên văn bản số/năm, Điều/khoản/điểm, ngày hiệu lực, đã bị sửa đổi "
              "hoặc thay thế chưa, quy định chuyển tiếp; mức tiền/tỷ lệ cụ thể. Phân biệt quy định trực tiếp và suy luận. "
              "Không bịa số hiệu văn bản; nếu không tìm thấy thì nói rõ.")
    last = ""
    for model in (model_for("gemini") if provider() == "gemini" else PROVIDERS["gemini"]["model"],
                  PROVIDERS["gemini"]["fallback"]):
        r = httpx.post(f"https://generativelanguage.googleapis.com/v1beta/models/{model}:generateContent",
                       headers={"x-goog-api-key": key},
                       json={"contents": [{"role": "user", "parts": [{"text": prompt}]}], "tools": [{"google_search": {}}]},
                       timeout=120)
        if r.status_code == 429:
            last = "Hết lượt tra cứu miễn phí của Gemini (thử lại sau)."
            continue
        if r.status_code >= 400:
            last = f"Gemini báo lỗi {r.status_code}: {r.text[:300]}"
            continue
        cand = (r.json().get("candidates") or [{}])[0]
        text = "".join(p.get("text", "") for p in (cand.get("content") or {}).get("parts", []))
        srcs = [{"title": (c.get("web") or {}).get("title", ""), "url": (c.get("web") or {}).get("uri", "")}
                for c in (cand.get("groundingMetadata") or {}).get("groundingChunks", [])]
        return {"noi_dung": text, "nguon": [s for s in srcs if s["url"]][:10]}
    return {"loi": last}


def _post_oa(p: str, model: str, msgs: list[dict[str, Any]]) -> dict[str, Any]:
    import httpx
    cfg = PROVIDERS[p]
    k = key_for(p)
    headers = {"Authorization": f"Bearer {k}"} if k else {}
    import time
    for attempt in range(2):  # 500/503 = máy chủ AI quá tải tạm thời -> thử lại 1 lần rồi đổi model dự phòng
        t0 = time.time()
        try:
            r = httpx.post(cfg["base"] + "/chat/completions", headers=headers,
                           timeout=httpx.Timeout(40.0, connect=10.0),  # treo lâu -> đổi model, không bắt khách chờ
                           json={"model": model, "messages": msgs, "tools": _oa_tools(), "tool_choice": "auto"})
        except httpx.TimeoutException:
            _log(f"AI {model} TIMEOUT sau {time.time() - t0:.1f}s")
            raise _RateLimited("timeout") from None
        _log(f"AI {model} HTTP {r.status_code} {time.time() - t0:.1f}s, {len(json.dumps(msgs, ensure_ascii=False))} ký tự gửi")
        if r.status_code not in (500, 502, 503, 504):
            break
        time.sleep(2)
    if r.status_code in (429, 500, 502, 503, 504):  # hết lượt / vẫn quá tải -> đổi sang model dự phòng
        raise _RateLimited(r.text[:200])
    if r.status_code in (401, 403):
        raise RuntimeError(f"Khóa API {cfg['name']} không hợp lệ hoặc chưa được cấp quyền.")
    if r.status_code >= 400:
        raise RuntimeError(f"{cfg['name']} báo lỗi {r.status_code}: {r.text[:300]}")
    return r.json()


class _RateLimited(Exception):
    pass


def _log(msg: str) -> None:
    """Nhật ký thời gian từng bước của trợ lý (data/ai.log) – để tìm chỗ chậm."""
    try:
        f = ROOT / "data" / "ai.log"
        if f.exists() and f.stat().st_size > 2_000_000:
            f.unlink()
        with f.open("a", encoding="utf-8") as fh:
            fh.write(f"{datetime.now():%Y-%m-%d %H:%M:%S} {msg}\n")
    except OSError:
        pass


_BUSY: dict[tuple[str, str], float] = {}  # (nhà cung cấp, model) -> hết lượt/quá tải đến lúc nào


def _next_model(p: str, model: str) -> str | None:
    """Model kế tiếp còn lượt trong chuỗi dự phòng (None nếu hết)."""
    import time
    cfg = PROVIDERS[p]
    chain = list(dict.fromkeys([model_for(p)] + cfg.get("chain", []) + ([cfg["fallback"]] if cfg.get("fallback") else [])))
    rest = chain[chain.index(model) + 1:] if model in chain else chain
    return next((m for m in rest if time.time() >= _BUSY.get((p, m), 0)), None)


def _chat_free(p: str, chat_id: str | None, text: str) -> dict[str, Any]:
    import httpx
    import time
    cfg = PROVIDERS[p]
    if not cfg.get("no_key") and not key_for(p):
        from .config import FROZEN
        raise RuntimeError("Trợ lý AI chưa được kích hoạt – vui lòng liên hệ nhà cung cấp phần mềm." if FROZEN else
                           f"Chưa có khóa API {cfg['name']} – lấy tại {cfg['key_url']} rồi lưu làm khóa dựng sẵn (⚙).")
    with _lock:
        c = CHATS.get(chat_id or "") or {"id": uuid.uuid4().hex, "messages": [], "busy": False, "provider": p}
        if c["busy"]:
            raise RuntimeError("Trợ lý đang xử lý câu trước, chờ chút nhé.")
        c["busy"] = True
        CHATS[c["id"]] = c
    msgs = c["messages"]
    start = len(msgs)
    msgs.append({"role": "user", "content": f"[Hôm nay {datetime.now():%d/%m/%Y %H:%M}] {text}"})
    actions: list[str] = []
    sources: dict[str, str] = {}
    evidence: list[str] = []  # kết quả công cụ (tra luật, số liệu) – để các AI trong nhóm cùng soát
    model = model_for(p)
    reply = ""
    try:
        for _ in range(MAX_STEPS):
            body = [{"role": "system", "content": SYSTEM_FREE}] + msgs
            data = None
            while data is None:
                if time.time() < _BUSY.get((p, model), 0):  # model này vừa hết lượt -> dùng model kế
                    nxt = _next_model(p, model)
                    if not nxt:
                        raise RuntimeError(f"{cfg['name']} đang hết lượt miễn phí hoặc quá tải – anh/chị thử lại sau ít "
                                           "phút nhé.")
                    model = nxt
                try:
                    data = _post_oa(p, model, body)
                except _RateLimited as e:
                    m = str(e).lower()
                    _BUSY[(p, model)] = time.time() + (1800 if "quota" in m or "429" in m else 300 if "timeout" in m else 60)
                    nxt = _next_model(p, model)
                    if not nxt:
                        raise RuntimeError(f"{cfg['name']} đang hết lượt miễn phí hoặc quá tải – anh/chị thử lại sau ít "
                                           "phút nhé.") from None
                    actions.append(f"Chuyển sang model {nxt} ({model} hết lượt / quá tải)")
                    model = nxt
            msg = ((data.get("choices") or [{}])[0]).get("message") or {}
            msg = {k: v for k, v in msg.items() if v is not None}
            msg["role"] = "assistant"
            msgs.append(msg)
            calls = msg.get("tool_calls") or []
            if not calls:
                reply = (msg.get("content") or "").strip()
                break
            for tc in calls:
                fn = tc.get("function") or {}
                name = fn.get("name", "")
                try:
                    args = json.loads(fn.get("arguments") or "{}")
                except ValueError:
                    args = {}
                actions.append(("Tra cứu luật: " + str(args.get("cau_hoi", ""))) if name == "tra_cuu_luat"
                               else f"Thực hiện: {name} {json.dumps(args, ensure_ascii=False)}")
                try:
                    out = (law_lookup(args.get("cau_hoi", ""), args.get("tu_khoa", "")) if name == "tra_cuu_luat"
                           else run_tool(name, args))
                    if name == "tra_cuu_luat":
                        for s in out.get("nguon", []):
                            sources.setdefault(s["url"], s["title"] or s["url"])
                        for w in out.get("bai_viet_truc_tuyen", []):
                            sources.setdefault(w["url"], f"{w['site']} – {w['title']} ({w.get('cap_nhat', '')})")
                        for m in out.get("ket_qua", []) + out.get("kho_kien_thuc", []):
                            for u in __import__("re").findall(r"https?://\S+", m["noi_dung"]):
                                u = u.rstrip(" ;,.")
                                host = (__import__("urllib.parse").parse.urlparse(u).hostname or "").replace("www.", "")
                                sources.setdefault(u, f"{host} – {m['muc']}")
                    content = json.dumps(out, ensure_ascii=False, default=str)[:20000]
                    evidence.append(f"[{name}] " + content[:6000])
                except Exception as e:  # noqa: BLE001
                    content = f"Lỗi: {e}"
                msgs.append({"role": "tool", "tool_call_id": tc.get("id", ""), "name": name, "content": content})
        else:
            reply = reply or "Em đã làm nhiều bước nhưng chưa xong – anh/chị hỏi tiếp để em tiếp tục nhé."
    except httpx.ConnectError:
        del msgs[start:]
        raise RuntimeError(f"Không kết nối được {cfg['name']}"
                           + (" – hãy cài và mở Ollama trước." if p == "ollama" else " – kiểm tra mạng.")) from None
    except Exception:
        del msgs[start:]
        raise
    finally:
        c["busy"] = False
    return {"chat_id": c["id"], "reply": reply, "actions": actions, "provider": cfg["name"], "model": model,
            "sources": [{"url": u, "title": t} for u, t in list(sources.items())[:12]], "evidence": evidence}


# ------------------------------------------------------------------ nhóm chat nhiều AI
# Trưởng nhóm (AI chính) trả lời + thao tác app; các thành viên khác đọc câu hỏi, câu trả lời và cùng nguồn đã tra để
# soát lỗi, bổ sung (không gọi công cụ -> không chạy trùng thao tác). Thêm/bớt thành viên: data/ai.json "group".
GROUP_DEFAULT = [
    {"id": "thue", "name": "Chuyên viên thuế", "provider": "gemini", "model": "gemini-3.5-flash-lite", "color": "#1a7f37",
     "role": "chuyên viên thuế: soát căn cứ pháp lý (số hiệu văn bản, điều khoản, hiệu lực, ngưỡng tiền), chỉ ra chỗ "
             "sai hoặc thiếu so với nguồn đã tra"},
    {"id": "rui-ro", "name": "Kiểm soát rủi ro", "provider": "gemini", "model": "gemma-4-26b-a4b-it", "color": "#9a6700",
     "role": "kiểm soát rủi ro thuế: nêu rủi ro bị loại chi phí / không được khấu trừ / bị phạt, việc khách cần làm ngay "
             "và hồ sơ cần lưu"},
]
LEAD_NAME = "Kế toán trưởng"


def group_members() -> list[dict[str, Any]]:
    g = _load().get("group")
    return g if isinstance(g, list) and g else GROUP_DEFAULT


def _post_plain(p: str, model: str, msgs: list[dict[str, Any]]) -> str:
    import re

    import httpx
    k = key_for(p)
    if p == "claude":
        import anthropic
        r = anthropic.Anthropic(api_key=k).messages.create(
            model=model, max_tokens=2000, system=msgs[0]["content"], messages=msgs[1:])
        return "".join(b.text for b in r.content if b.type == "text")
    r = httpx.post(PROVIDERS[p]["base"] + "/chat/completions", headers={"Authorization": f"Bearer {k}"} if k else {},
                   json={"model": model, "messages": msgs}, timeout=60)
    r.raise_for_status()
    text = (r.json()["choices"][0]["message"].get("content") or "")
    return re.sub(r"<thought>.*?</thought>", "", text, flags=re.S).strip()  # Gemma in kèm phần suy nghĩ


def group_chat(chat_id: str | None, text: str, context: str = "") -> dict[str, Any]:
    from concurrent.futures import ThreadPoolExecutor

    r = chat(chat_id, text, context)
    if not r.get("reply"):
        return r | {"group": []}
    evidence = "\n\n".join(r.get("evidence", []))[:14000] or "(trưởng nhóm không tra nguồn nào)"

    def review(m: dict[str, Any]) -> dict[str, Any] | None:
        if m["provider"] not in PROVIDERS or not (key_for(m["provider"]) or PROVIDERS[m["provider"]].get("no_key")):
            return None
        sys_msg = (f"Bạn là {m['name']} trong nhóm kế toán dịch vụ của phần mềm AWBook – vai trò: {m['role']}. "
                   f"{LEAD_NAME} vừa trả lời khách. Đọc câu hỏi, câu trả lời và NGUỒN ĐÃ TRA, rồi góp ý NGẮN (tối đa "
                   "120 chữ, tiếng Việt, xưng 'em', gọi khách 'anh/chị'). Nếu nguồn đã tra ghi 'CẦN ĐỐI CHIẾU' hoặc các "
                   "nguồn mâu thuẫn mà câu trả lời lại khẳng định chắc chắn thì phải nêu ra. Mở đầu bằng '✅ Đồng ý' nếu câu trả lời đúng "
                   "và đủ (chỉ thêm 1-2 ý giá trị nếu có), hoặc '⚠️ Cần lưu ý' nếu có điểm sai/thiếu. Chỉ dựa vào nguồn "
                   "đã tra và hiểu biết chắc chắn; không bịa số hiệu văn bản; không lặp lại câu trả lời.")
        user = (f"Hôm nay {datetime.now():%d/%m/%Y}. CÂU HỎI CỦA KHÁCH:\n{text}\n\nCÂU TRẢ LỜI CỦA {LEAD_NAME.upper()}:\n"
                f"{r['reply'][:6000]}\n\nNGUỒN ĐÃ TRA / SỐ LIỆU APP:\n{evidence}")
        import time
        tries = [m["model"]] + [x for x in m.get("backup", ["gemini-3.1-flash-lite", "gemini-flash-lite-latest"])
                                if x != m["model"]]
        for i, model in enumerate(tries[:3]):
            try:
                out = _post_plain(m["provider"], model, [{"role": "system", "content": sys_msg},
                                                         {"role": "user", "content": user}])
                if out:
                    return {"name": m["name"], "model": model, "color": m.get("color", "#57606a"), "text": out}
            except Exception:  # noqa: BLE001 - bận / hết lượt / giới hạn mỗi phút -> chờ chút, thử model khác
                time.sleep(2 + 2 * i)
        return None  # cả nhóm dự phòng đều bận -> bỏ qua, không làm hỏng câu trả lời chính

    with ThreadPoolExecutor(4) as ex:
        group = [g for g in ex.map(review, group_members()) if g]
    return r | {"group": group, "lead": LEAD_NAME}


_UI = threading.local()  # thao tác màn hình (mở trang, mở file) mà trợ lý yêu cầu trong lượt chat hiện tại


def chat(chat_id: str | None, text: str, context: str = "") -> dict[str, Any]:
    p = provider()
    old = CHATS.get(chat_id or "")
    if old and old.get("provider", "claude") != p:
        chat_id = None  # đổi nhà cung cấp AI -> bắt đầu hội thoại mới (định dạng lịch sử khác nhau)
    if context:
        text = f"[Màn hình hiện tại: {context[:1500]}]\n{text}"
    _UI.items = []
    if p == "claude":
        r = _chat_claude(chat_id, text)
        CHATS[r["chat_id"]]["provider"] = "claude"
        r = r | {"provider": PROVIDERS["claude"]["name"], "model": MODEL}
    else:
        r = _chat_free(p, chat_id, text)
    return r | {"ui": list(getattr(_UI, "items", []))}
