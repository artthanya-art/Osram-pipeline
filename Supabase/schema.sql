-- OSRAM Sales Pipeline — Supabase schema
-- Run this whole file once in your Supabase project's SQL editor (Database > SQL Editor > New query).

create extension if not exists pgcrypto;

-- ---------------- zones ----------------
create table if not exists zones (
  name text primary key
);

insert into zones (name) values
  ('Red Zone'), ('Blue Zone'), ('Green Zone'), ('Yellow Zone'), ('Orange Zone')
on conflict (name) do nothing;

-- ---------------- users ----------------
-- NOTE: this app implements its own login (Sales ID + password) instead of Supabase Auth,
-- to keep the existing "admin approves new accounts" workflow. See README.md for the
-- security trade-offs of that choice before using this with real/sensitive data.
create table if not exists users (
  sales_id   text primary key,
  password   text not null,
  name       text not null,
  zone       text not null,
  role       text not null default 'sales' check (role in ('admin','sales')),
  status     text not null default 'pending' check (status in ('pending','approved','rejected')),
  created_at timestamptz not null default now()
);

-- Demo accounts so the app is usable immediately (the app also seeds these client-side
-- on first load if the users table is empty, so this insert is just a convenience).
insert into users (sales_id, password, name, zone, role, status) values
  ('admin', 'admin123', 'ผู้ดูแลระบบ (Demo)', 'Red Zone', 'admin', 'approved'),
  ('demo',  'demo123',  'พนักงานขาย (Demo)',  'Red Zone', 'sales', 'approved')
on conflict (sales_id) do nothing;

-- ---------------- pipeline entries ----------------
create table if not exists pipeline_entries (
  id                   uuid primary key default gen_random_uuid(),
  sales_id             text not null references users(sales_id),
  sales_name           text,
  zone                 text,
  customer_id          text,
  customer_name        text not null,
  customer_type        text,
  customer_segment     text,
  product_type         text,
  item_code            text,
  item_description     text,
  qty                  numeric,
  uom                  text,
  price                numeric,
  competitor_name      text,
  competitor_price     numeric,
  project_close_year   text,
  project_close_month  text,
  quotation_number     text,
  delivery_method      text,
  kpi_register         text,
  action_plan_month    text,
  progress             text,
  deliver_in_months    text,
  delivery_start_year  text,
  delivery_start_month text,
  start_working_month  text,
  visit_date           text,
  grade                text,
  created_at           timestamptz not null default now(),
  updated_at           timestamptz not null default now()
);

create index if not exists pipeline_entries_sales_id_idx on pipeline_entries (sales_id);
create index if not exists pipeline_entries_created_at_idx on pipeline_entries (created_at desc);

-- ---------------- Row Level Security ----------------
-- The app authenticates its own users against the `users` table and calls Supabase with the
-- public anon key for everything (there is no Supabase Auth session). That means these
-- policies must allow the anon key to read/write, or the app simply won't work.
--
-- This is fine for an internal/demo tool, but it means ANYONE who has your anon key
-- (visible in the browser bundle) can read and write every row in every table below.
-- Before using this with real business data, restrict writes behind a server-side
-- component (e.g. a Supabase Edge Function that checks a session/secret) instead of
-- calling these tables directly from the browser.

alter table zones enable row level security;
alter table users enable row level security;
alter table pipeline_entries enable row level security;

drop policy if exists "public read zones" on zones;
drop policy if exists "public write zones" on zones;
create policy "public read zones" on zones for select using (true);
create policy "public write zones" on zones for insert with check (true);

drop policy if exists "public read users" on users;
drop policy if exists "public insert users" on users;
drop policy if exists "public update users" on users;
create policy "public read users" on users for select using (true);
create policy "public insert users" on users for insert with check (true);
create policy "public update users" on users for update using (true);

drop policy if exists "public read entries" on pipeline_entries;
drop policy if exists "public insert entries" on pipeline_entries;
drop policy if exists "public update entries" on pipeline_entries;
drop policy if exists "public delete entries" on pipeline_entries;
create policy "public read entries" on pipeline_entries for select using (true);
create policy "public insert entries" on pipeline_entries for insert with check (true);
create policy "public update entries" on pipeline_entries for update using (true);
create policy "public delete entries" on pipeline_entries for delete using (true);
