"""Lớp truy cập Supabase (Postgres + Storage)."""
from __future__ import annotations

from functools import lru_cache
from typing import Any

from supabase import Client, create_client

from .config import settings

INVOICE_KEY = "owner_mst,direction,nbmst,khmshdon,khhdon,shdon"


class DBNotConfigured(RuntimeError):
    pass


@lru_cache(maxsize=1)
def client() -> Client:
    if not settings.supabase_url or not settings.supabase_key:
        raise DBNotConfigured("Chưa cấu hình SUPABASE_URL / SUPABASE_SERVICE_KEY trong file .env")
    return create_client(settings.supabase_url, settings.supabase_key)


# ---------- companies ----------
def upsert_company(mst: str, name: str | None = None) -> None:
    row: dict[str, Any] = {"mst": mst}
    if name:
        row["name"] = name
    client().table("companies").upsert(row, on_conflict="mst").execute()


def get_company(mst: str) -> dict[str, Any] | None:
    res = client().table("companies").select("*").eq("mst", mst).limit(1).execute()
    return res.data[0] if res.data else None


def list_companies() -> list[dict[str, Any]]:
    return client().table("companies").select("mst,name,tconnect_account_id").order("mst").execute().data or []


def set_tconnect_account(mst: str, account_id: str) -> None:
    client().table("companies").upsert({"mst": mst, "tconnect_account_id": account_id},
                                       on_conflict="mst").execute()


# ---------- invoices ----------
def upsert_invoice(row: dict[str, Any]) -> dict[str, Any]:
    res = client().table("invoices").upsert(row, on_conflict=INVOICE_KEY).execute()
    return res.data[0]


def update_invoice(invoice_id: str, fields: dict[str, Any]) -> dict[str, Any]:
    res = client().table("invoices").update(fields).eq("id", invoice_id).execute()
    return res.data[0] if res.data else {}


def get_invoice(invoice_id: str) -> dict[str, Any] | None:
    res = client().table("invoices").select("*").eq("id", invoice_id).limit(1).execute()
    return res.data[0] if res.data else None


LIST_COLUMNS = (
    "id,owner_mst,direction,nbmst,nbten,nmmst,nmten,khmshdon,khhdon,shdon,tdlap,"
    "tgtcthue,tgtthue,tgtttbso,tthai,provider_code,lookup_code,lookup_url,"
    "xml_path,pdf_path,pdf_status,pdf_is_original,pdf_error"
)


def list_invoices(
    owner_mst: str | None = None,
    direction: str | None = None,
    date_from: str | None = None,
    date_to: str | None = None,
    pdf_status: str | None = None,
    q: str | None = None,
    limit: int = 200,
    offset: int = 0,
) -> list[dict[str, Any]]:
    query = client().table("invoices").select(LIST_COLUMNS)
    if owner_mst:
        query = query.eq("owner_mst", owner_mst)
    if direction:
        query = query.eq("direction", direction)
    if date_from:
        query = query.gte("tdlap", f"{date_from}T00:00:00+07:00")
    if date_to:
        query = query.lte("tdlap", f"{date_to}T23:59:59+07:00")
    if pdf_status:
        query = query.eq("pdf_status", pdf_status)
    if q:
        safe = q.replace(",", " ").replace("(", " ").replace(")", " ").strip()
        if safe:
            query = query.or_(f"nbten.ilike.*{safe}*,nmten.ilike.*{safe}*,nbmst.ilike.*{safe}*,shdon.eq.{safe}")
    res = query.order("tdlap", desc=True).range(offset, offset + limit - 1).execute()
    return res.data or []


def insert_job(row: dict[str, Any]) -> None:
    client().table("sync_jobs").insert(row).execute()


# ---------- storage ----------
def upload(path: str, data: bytes, content_type: str) -> str:
    client().storage.from_(settings.bucket).upload(
        path, data, {"content-type": content_type, "upsert": "true"}
    )
    return path


def download(path: str) -> bytes:
    return client().storage.from_(settings.bucket).download(path)


def signed_url(path: str, expires: int = 3600) -> str:
    res = client().storage.from_(settings.bucket).create_signed_url(path, expires)
    return res.get("signedURL") or res.get("signedUrl") or res.get("signed_url")
