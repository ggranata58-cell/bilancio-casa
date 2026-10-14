-- Bilancio Casa: cloud persistence for the manually maintained cash-wallet balance.
-- Safe to run once in Supabase SQL Editor.
create table if not exists public.bilancio_cash_settings (
  user_id uuid primary key references auth.users(id) on delete cascade,
  cash_balance numeric(12,2) not null default 50.00 check (cash_balance >= 0),
  updated_at timestamptz not null default now()
);

alter table public.bilancio_cash_settings enable row level security;

drop policy if exists "Users can read own cash settings" on public.bilancio_cash_settings;
create policy "Users can read own cash settings"
  on public.bilancio_cash_settings for select
  to authenticated
  using (auth.uid() = user_id);

drop policy if exists "Users can insert own cash settings" on public.bilancio_cash_settings;
create policy "Users can insert own cash settings"
  on public.bilancio_cash_settings for insert
  to authenticated
  with check (auth.uid() = user_id);

drop policy if exists "Users can update own cash settings" on public.bilancio_cash_settings;
create policy "Users can update own cash settings"
  on public.bilancio_cash_settings for update
  to authenticated
  using (auth.uid() = user_id)
  with check (auth.uid() = user_id);

grant select, insert, update on public.bilancio_cash_settings to authenticated;
