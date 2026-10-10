-- Honey till — shared Supabase backend
-- Run this once in the Supabase SQL editor (Project > SQL Editor > New query).
--
-- Notes:
--   * ids are client-generated (crypto.randomUUID()) so sales/batches can be
--     recorded while offline, then synced later without a server roundtrip.
--   * columns that mirror JS field names (locId, accountId) are quoted to
--     keep the exact same shape on the client — no case-mapping layer needed.
--   * stock is NOT a column anywhere. It is always derived as
--     sum(batches.qty for product) - sum(sales.qty for product).

create table if not exists locations (
  id text primary key,
  name text not null
);

create table if not exists products (
  id text primary key,
  name text not null,
  type text not null default 'Other',
  price jsonb not null default '{}'::jsonb
);

create table if not exists accounts (
  id text primary key,
  name text not null,
  method text not null default 'TWINT',
  common boolean not null default false
);

-- Stock intake / correction events. A sale never touches this table;
-- stock-on-hand is sum(batches.qty) - sum(sales.qty) for a product.
create table if not exists batches (
  id text primary key,
  pid text not null references products(id) on delete cascade,
  qty numeric not null,
  ts bigint not null,
  note text
);

-- Sales are append-only from the till's point of view (undo deletes the
-- row rather than reversing a stock counter, so two phones selling at the
-- same time only ever add independent rows — never a shared number).
create table if not exists sales (
  id text primary key,
  ticket text not null,
  ts bigint not null,
  pid text not null,
  name text not null,
  type text,
  price numeric not null,
  list numeric not null,
  mode text not null default 'full',
  qty numeric not null,
  note text,
  "locId" text not null,
  location text,
  "accountId" text,
  account text,
  method text
);

-- Money paid from a person's account into the common honey account.
create table if not exists transfers (
  id text primary key,
  ts bigint not null,
  "accountId" text not null,
  name text,
  amount numeric not null,
  note text
);

-- A market day: a named event with its own product/price picks, pulled from
-- (but not overwriting) a location's normal price list. Sales made while a
-- market is active record its id/name alongside the usual location/account.
create table if not exists markets (
  id text primary key,
  name text not null,
  "locId" text not null,
  items jsonb not null default '[]'::jsonb,
  "startedAt" bigint not null,
  "endedAt" bigint
);

-- Additive: a market can run over several days (ended, then reopened).
-- closedAt marks one as finished for good so it stops being offered.
alter table markets add column if not exists "closedAt" bigint;

-- Additive: which market (if any) a sale belongs to. Nullable so existing
-- sales rows are unaffected.
alter table sales add column if not exists "marketId" text;
alter table sales add column if not exists market text;

-- Additive: cash tendered and change given, only set for cash sales.
alter table sales add column if not exists "cashReceived" numeric;
alter table sales add column if not exists "cashChange" numeric;

-- Every time a sale is moved to a different account (e.g. an unassigned
-- sale attached to a person, or a sale booked to the wrong TWINT), one row
-- lands here: who did it, when, and what the sale said before. Append-only:
-- there is deliberately no update or delete policy below, so history can't
-- be rewritten from the app. The item/amount summary is copied in so the
-- record still reads correctly after the sales themselves are cleared.
create table if not exists sale_changes (
  id text primary key,
  ts bigint not null,
  ticket text not null,
  "saleIds" jsonb not null default '[]'::jsonb,
  items text,
  amount numeric not null,
  "fromAccountId" text,
  "fromAccount" text,
  "fromMethod" text,
  "toAccountId" text not null,
  "toAccount" text not null,
  "toMethod" text,
  "changedBy" text not null,
  "authUserId" text,
  note text
);

alter table locations enable row level security;
alter table products  enable row level security;
alter table accounts  enable row level security;
alter table batches   enable row level security;
alter table sales     enable row level security;
alter table transfers enable row level security;
alter table markets   enable row level security;
alter table sale_changes enable row level security;

-- Single shared workspace: any signed-in user may read/write everything.
create policy "authenticated full access" on locations for all
  using (auth.role() = 'authenticated') with check (auth.role() = 'authenticated');
create policy "authenticated full access" on products for all
  using (auth.role() = 'authenticated') with check (auth.role() = 'authenticated');
create policy "authenticated full access" on accounts for all
  using (auth.role() = 'authenticated') with check (auth.role() = 'authenticated');
create policy "authenticated full access" on batches for all
  using (auth.role() = 'authenticated') with check (auth.role() = 'authenticated');
create policy "authenticated full access" on sales for all
  using (auth.role() = 'authenticated') with check (auth.role() = 'authenticated');
create policy "authenticated full access" on transfers for all
  using (auth.role() = 'authenticated') with check (auth.role() = 'authenticated');
create policy "authenticated full access" on markets for all
  using (auth.role() = 'authenticated') with check (auth.role() = 'authenticated');
-- sale_changes: read and append only. No update/delete policy on purpose.
create policy "authenticated read" on sale_changes for select
  using (auth.role() = 'authenticated');
create policy "authenticated append" on sale_changes for insert
  with check (auth.role() = 'authenticated');

-- Realtime: on most Supabase projects the "supabase_realtime" publication
-- is created FOR ALL TABLES, so every table above is already broadcasting
-- changes with no extra step. If your project isn't set up that way,
-- check Database > Replication in the dashboard and toggle these 6 tables
-- on there instead of via SQL (avoids fighting an existing publication).

-- Expenses: money someone paid out of their own pocket for the business
-- (a machine, jars, a stand fee). Beezness owes it back until "repaidAt"
-- is set. Additive and optional: the app runs without it, just without
-- the Expenses tab. Safe to run on its own in the SQL editor.
create table if not exists expenses (
  id text primary key,
  ts bigint not null,
  "paidBy" text not null,
  what text not null,
  amount numeric not null,
  "repaidAt" bigint
);
alter table expenses enable row level security;
drop policy if exists "authenticated full access" on expenses;
create policy "authenticated full access" on expenses for all
  using (auth.role() = 'authenticated') with check (auth.role() = 'authenticated');
