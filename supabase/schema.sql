-- Chạy file này trong Supabase Dashboard > SQL Editor (một lần).

create extension if not exists pgcrypto;

-- Doanh nghiệp đã đồng bộ (không lưu mật khẩu cổng thuế).
create table if not exists public.companies (
  mst         text primary key,
  name        text,
  tconnect_account_id text,                           -- accountId bên TConnect (nếu dùng)
  created_at  timestamptz not null default now()
);
alter table public.companies add column if not exists tconnect_account_id text;

-- Hóa đơn đầu vào (purchase) / đầu ra (sold).
create table if not exists public.invoices (
  id               uuid primary key default gen_random_uuid(),
  owner_mst        text not null,                       -- MST doanh nghiệp đăng nhập
  direction        text not null check (direction in ('purchase', 'sold')),
  nbmst            text not null,                       -- MST người bán
  khmshdon         text not null,                       -- ký hiệu mẫu số
  khhdon           text not null,                       -- ký hiệu hóa đơn
  shdon            text not null,                       -- số hóa đơn
  tdlap            timestamptz,                         -- ngày lập
  nbten            text,
  nmmst            text,
  nmten            text,
  tgtcthue         numeric,                             -- tổng tiền chưa thuế
  tgtthue          numeric,                             -- tổng tiền thuế
  tgtttbso         numeric,                             -- tổng thanh toán
  tthai            int,                                 -- trạng thái hóa đơn
  ttxly            int,                                 -- trạng thái xử lý
  mhdon            text,                                -- mã CQT
  msttcgp          text,                                -- MST tổ chức cung cấp giải pháp HĐĐT
  source_query     text,                                -- endpoint đã lấy
  provider_code    text,                                -- nhà cung cấp nhận diện được
  lookup_code      text,                                -- mã tra cứu
  lookup_url       text,                                -- link tra cứu
  xml_path         text,                                -- đường dẫn trong Storage
  pdf_path         text,
  pdf_status       text not null default 'pending',     -- pending | original | rendered | manual | failed
  pdf_is_original  boolean,
  pdf_source_url   text,
  pdf_error        text,
  raw              jsonb,                               -- bản ghi gốc từ cổng thuế
  parsed           jsonb,                               -- dữ liệu đọc từ XML
  created_at       timestamptz not null default now(),
  updated_at       timestamptz not null default now(),
  unique (owner_mst, direction, nbmst, khmshdon, khhdon, shdon)
);

create index if not exists invoices_owner_dir_date on public.invoices (owner_mst, direction, tdlap desc);
create index if not exists invoices_pdf_status on public.invoices (pdf_status);

create or replace function public.touch_updated_at() returns trigger
language plpgsql set search_path = '' as $$
begin new.updated_at = now(); return new; end $$;

drop trigger if exists invoices_touch on public.invoices;
create trigger invoices_touch before update on public.invoices
  for each row execute function public.touch_updated_at();

-- Nhật ký các lần đồng bộ.
create table if not exists public.sync_jobs (
  id          uuid primary key default gen_random_uuid(),
  owner_mst   text not null,
  direction   text not null,
  date_from   date not null,
  date_to     date not null,
  status      text not null,
  total       int default 0,
  pdf_ok      int default 0,
  pdf_failed  int default 0,
  log         jsonb,
  started_at  timestamptz not null default now(),
  finished_at timestamptz
);

-- Bật RLS, không tạo policy: chỉ backend dùng service role key mới truy cập được.
alter table public.companies enable row level security;
alter table public.invoices  enable row level security;
alter table public.sync_jobs enable row level security;

-- Bucket riêng tư chứa XML/PDF.
insert into storage.buckets (id, name, public)
values ('invoices', 'invoices', false)
on conflict (id) do nothing;
