"""Hệ connector lấy PDF gốc từ trang tra cứu của nhà cung cấp hóa đơn điện tử.

Thứ tự thử cho mỗi hóa đơn:
  1. Hàm Python riêng của NCC (plugins.CUSTOM_FETCHERS) nếu có.
  2. Các mẫu URL trong providers.json của NCC nhận diện được.
  3. Mọi link tra cứu in trong XML hóa đơn (dò link PDF trong trang HTML, 1 cấp).
Không tìm thấy -> trả về None kèm nhật ký các lần thử để người dùng tải tay.
"""
from __future__ import annotations

import json
import re
import string
from dataclasses import dataclass, field
from typing import Any, Callable
from urllib.parse import urljoin, urlparse

import httpx
from bs4 import BeautifulSoup

from ..config import settings
from ..gdt.client import UA


@dataclass
class InvoiceRef:
    seller_mst: str
    buyer_mst: str | None
    khmshdon: str
    khhdon: str
    shdon: str
    msttcgp: str | None
    mccqt: str | None
    lookup_code: str | None
    lookup_urls: list[str] = field(default_factory=list)

    def vars(self) -> dict[str, str | None]:
        return {
            "seller_mst": self.seller_mst, "buyer_mst": self.buyer_mst,
            "khmshdon": self.khmshdon, "khhdon": self.khhdon, "shdon": self.shdon,
            "mccqt": self.mccqt, "lookup_code": self.lookup_code,
            "lookup_url": self.lookup_urls[0] if self.lookup_urls else None,
        }


@dataclass
class Provider:
    code: str
    name: str
    mst_tcgp: list[str]
    domains: list[str]
    pdf_urls: list[str]
    verified: bool = False
    note: str | None = None

    def match_score(self, ref: InvoiceRef) -> int:
        score = 0
        if ref.msttcgp and ref.msttcgp in self.mst_tcgp:
            score += 10
        for u in ref.lookup_urls:
            host = urlparse(u).hostname or ""
            if any(host == d or host.endswith("." + d) for d in self.domains):
                score += 5
        return score


@dataclass
class PdfResult:
    content: bytes
    source_url: str
    provider_code: str | None


@dataclass
class FetchOutcome:
    result: PdfResult | None
    provider: Provider | None
    attempts: list[dict[str, Any]]
    needs_captcha: bool = False


class NeedsCaptcha(Exception):
    pass


def load_providers() -> list[Provider]:
    data = json.loads(settings.providers_file.read_text(encoding="utf-8"))
    keys = Provider.__dataclass_fields__.keys()
    return [Provider(**{k: v for k, v in p.items() if k in keys}) for p in data["providers"]]


def detect_provider(ref: InvoiceRef, providers: list[Provider] | None = None) -> Provider | None:
    providers = providers or load_providers()
    scored = sorted(((p.match_score(ref), p) for p in providers), key=lambda x: -x[0])
    return scored[0][1] if scored and scored[0][0] > 0 else None


def _fill(template: str, values: dict[str, str | None]) -> str | None:
    names = [f for _, f, _, _ in string.Formatter().parse(template) if f]
    if any(not values.get(n) for n in names):
        return None
    return template.format(**{n: values[n] for n in names})


def _is_pdf(r: httpx.Response) -> bool:
    return r.content[:5] == b"%PDF-" or "application/pdf" in r.headers.get("content-type", "")


PDF_HINT = re.compile(r"pdf|download|taive|tai-ve|xemhoadon|inhoadon|print", re.I)
JS_URL = re.compile(r"""["']([^"'\s]*(?:\.pdf|pdf[^"'\s]*|download[^"'\s]*))["']""", re.I)


def pdf_matches_invoice(content: bytes, ref: InvoiceRef) -> bool:
    """PDF phải chứa số hóa đơn + (ký hiệu hoặc MST người bán) – tránh nhận nhầm PDF hướng dẫn, mẫu..."""
    import io

    from pypdf import PdfReader
    try:
        reader = PdfReader(io.BytesIO(content))
        text = " ".join((p.extract_text() or "") for p in reader.pages[:3])
    except Exception:  # noqa: BLE001
        return False
    flat = re.sub(r"\s+", "", text).upper()
    if not flat:
        return False  # PDF ảnh, không kiểm chứng được -> không coi là bản gốc
    num = (ref.shdon or "").lstrip("0") or "0"
    has_num = re.search(rf"(?<!\d)0*{re.escape(num)}(?!\d)", re.sub(r"\s+", " ", text)) is not None
    has_sym = bool(ref.khhdon) and ref.khhdon.upper() in flat
    has_mst = bool(ref.seller_mst) and ref.seller_mst in flat
    return has_num and (has_sym or has_mst)


def fetch_pdf_from_url(http: httpx.Client, url: str, depth: int = 1,
                       accept: Callable[[bytes], bool] | None = None) -> tuple[bytes, str] | None:
    r = http.get(url)
    if r.status_code >= 400:
        raise httpx.HTTPStatusError(f"HTTP {r.status_code}", request=r.request, response=r)
    if _is_pdf(r):
        if r.content[:5] != b"%PDF-":
            return None  # content-type PDF nhưng nội dung không phải PDF (trang lỗi)
        if accept and not accept(r.content):
            return None  # PDF không phải của hóa đơn này
        return r.content, str(r.url)
    ctype = r.headers.get("content-type", "")
    if "html" not in ctype or depth <= 0:
        return None
    html = r.text
    soup = BeautifulSoup(html, "html.parser")
    cands: list[str] = []
    for tag, attr in (("iframe", "src"), ("embed", "src"), ("object", "data"), ("a", "href")):
        for el in soup.find_all(tag):
            v = el.get(attr)
            if v and not v.startswith(("javascript:", "#", "mailto:")):
                if tag != "a" or PDF_HINT.search(v) or PDF_HINT.search(el.get_text() or ""):
                    cands.append(v)
    for s in soup.find_all("script"):
        cands += JS_URL.findall(s.string or "")
    seen = set()
    for c in cands:
        abs_url = urljoin(str(r.url), c)
        if abs_url in seen or not abs_url.startswith("http"):
            continue
        seen.add(abs_url)
        if len(seen) > 6:
            break
        try:
            got = fetch_pdf_from_url(http, abs_url, depth - 1, accept)
        except httpx.HTTPError:
            continue
        if got:
            return got
    if "captcha" in html.lower():
        raise NeedsCaptcha("Trang tra cứu yêu cầu nhập captcha")
    return None


# Hàm tải PDF viết riêng cho từng NCC (khi mẫu URL không đủ, vd cần POST/API riêng).
# Khai báo trong app/providers/plugins.py:  CUSTOM_FETCHERS["misa"] = my_func
CustomFetcher = Callable[[httpx.Client, InvoiceRef], "PdfResult | None"]


def fetch_original_pdf(ref: InvoiceRef, providers: list[Provider] | None = None) -> FetchOutcome:
    from .plugins import BROWSER_HEADERS, CUSTOM_FETCHERS

    providers = providers or load_providers()
    provider = detect_provider(ref, providers)
    attempts: list[dict[str, Any]] = []
    needs_captcha = False
    values = ref.vars()

    urls: list[str] = []
    if provider:
        urls += [u for u in (_fill(t, values) for t in provider.pdf_urls) if u]
    elif ref.lookup_code:
        # Không nhận diện được NCC (vd dữ liệu từ TConnect chỉ có mã tra cứu):
        # thử các mẫu URL chỉ cần mã tra cứu của mọi NCC.
        for p in providers:
            urls += [u for u in (_fill(t, values) for t in p.pdf_urls if "{lookup_url}" not in t) if u]
    urls += ref.lookup_urls
    urls = list(dict.fromkeys(urls))

    with httpx.Client(
        timeout=httpx.Timeout(40.0, connect=15.0),
        follow_redirects=True,
        headers=BROWSER_HEADERS,
    ) as http:
        if provider and provider.code in CUSTOM_FETCHERS:
            try:
                res = CUSTOM_FETCHERS[provider.code](http, ref)
                if res and pdf_matches_invoice(res.content, ref):
                    return FetchOutcome(res, provider, attempts)
                attempts.append({"step": f"plugin:{provider.code}", "result": "không có PDF"})
            except Exception as e:  # noqa: BLE001 - ghi nhận mọi lỗi plugin
                attempts.append({"step": f"plugin:{provider.code}", "error": str(e)[:300]})

        for url in urls:
            try:
                got = fetch_pdf_from_url(http, url, accept=lambda c: pdf_matches_invoice(c, ref))
            except NeedsCaptcha as e:
                needs_captcha = True
                attempts.append({"url": url, "error": str(e)})
                continue
            except httpx.HTTPError as e:
                attempts.append({"url": url, "error": str(e)[:300]})
                continue
            if got:
                content, final_url = got
                return FetchOutcome(
                    PdfResult(content, final_url, provider.code if provider else None), provider, attempts
                )
            attempts.append({"url": url, "result": "không tìm thấy PDF khớp hóa đơn"})

    if not urls:
        attempts.append({"error": "XML không có link/mã tra cứu và NCC chưa có mẫu URL phù hợp"})
    return FetchOutcome(None, provider, attempts, needs_captcha)
