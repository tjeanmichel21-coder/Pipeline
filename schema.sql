-- ============================================================
-- Pipeline shared database v2 — WITH SEATS (run in Supabase SQL editor)
-- Safe to re-run on top of v1: it drops the old open policies.
-- ============================================================

create table if not exists itbs (id text primary key, data jsonb not null, updated_at timestamptz default now());
create table if not exists jobs (id text primary key, data jsonb not null, updated_at timestamptz default now());
create table if not exists accounts (id text primary key, data jsonb not null, updated_at timestamptz default now());
create table if not exists contacts (id text primary key, data jsonb not null, updated_at timestamptz default now());
create table if not exists tasks (id text primary key, data jsonb not null, updated_at timestamptz default now());
create table if not exists app_settings (id text primary key, data jsonb not null default '{}'::jsonb);

-- ClearBid integration mailbox
create table if not exists estimates (
  id uuid primary key default gen_random_uuid(),
  itb_id text references itbs(id) on delete cascade,
  source text default 'clearbid',
  status text default 'created',
  total numeric,
  items jsonb default '[]'::jsonb,
  project jsonb default '{}'::jsonb,   -- {projectName, client, contact, phone, address} for brand-new bids
  pdf_url text,
  created_at timestamptz default now(),
  unique (itb_id)
);
alter table estimates add column if not exists project jsonb default '{}'::jsonb;

-- ===================== SEATS =====================
create table if not exists seats (
  id uuid primary key default gen_random_uuid(),
  email text unique not null,
  role text not null default 'member',   -- 'owner' or 'member'
  created_at timestamptz default now()
);

-- >>> EDIT THIS LINE: put YOUR email here before running, then sign up with it in the app <<<
insert into seats (email, role) values ('YOUR_EMAIL_HERE@example.com', 'owner')
on conflict (email) do update set role = 'owner';

-- ===================== POLICIES =====================
alter table itbs enable row level security;
alter table jobs enable row level security;
alter table accounts enable row level security;
alter table contacts enable row level security;
alter table tasks enable row level security;
alter table app_settings enable row level security;
alter table estimates enable row level security;
alter table seats enable row level security;

-- remove v1 open policies if present
drop policy if exists "open itbs" on itbs;
drop policy if exists "open jobs" on jobs;
drop policy if exists "open settings" on app_settings;
drop policy if exists "open estimates" on estimates;
drop policy if exists "open accounts" on accounts;
drop policy if exists "open contacts" on contacts;
drop policy if exists "open tasks" on tasks;

-- helper: signed-in user holds a seat
create or replace function has_seat() returns boolean
language sql stable security definer set search_path = public as
$$ select exists (select 1 from seats where lower(email) = lower(auth.jwt()->>'email')) $$;

create or replace function is_owner() returns boolean
language sql stable security definer set search_path = public as
$$ select exists (select 1 from seats where lower(email) = lower(auth.jwt()->>'email') and role = 'owner') $$;

-- CRM data: seat holders only
do $$
declare t text;
begin
  foreach t in array array['itbs','jobs','accounts','contacts','tasks','app_settings'] loop
    execute format('drop policy if exists "seat access" on %I', t);
    execute format('create policy "seat access" on %I for all using (has_seat()) with check (has_seat())', t);
  end loop;
end $$;

-- seats: any seat holder can view the team; only owners can add/remove
drop policy if exists "seats read" on seats;
drop policy if exists "seats insert" on seats;
drop policy if exists "seats delete" on seats;
drop policy if exists "seats update" on seats;
create policy "seats read"   on seats for select using (has_seat());
create policy "seats insert" on seats for insert with check (is_owner());
create policy "seats delete" on seats for delete using (is_owner());
create policy "seats update" on seats for update using (is_owner());

-- estimates: open mailbox so ClearBid (anon key, no user session) can publish.
-- Lock this later by moving the bridge into a serverless function with a service-role key.
drop policy if exists "estimates insert" on estimates;
drop policy if exists "estimates update" on estimates;
drop policy if exists "estimates select" on estimates;
create policy "estimates insert" on estimates for insert with check (true);
create policy "estimates update" on estimates for update using (true);
create policy "estimates select" on estimates for select using (true);
