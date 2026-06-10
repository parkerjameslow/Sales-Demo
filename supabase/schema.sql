-- Demo Request Queue — database setup
-- Run this once in your Supabase project: SQL Editor → New query → paste → Run.

create table public.requests (
  id              bigint generated always as identity primary key,
  created_at      timestamptz not null default now(),
  request_name    text not null,
  requester       text not null,
  date_requested  date,
  applications    text not null,
  jira            text,
  sprint          text,
  priority        text not null check (priority in ('Critical', 'High', 'Medium', 'Low')),
  description     text not null,
  status          text not null default 'New'
                  check (status in ('New', 'In Review', 'Approved', 'Scheduled',
                                    'In Progress', 'Completed', 'Declined'))
);

-- Row Level Security: the site talks to this table with the public "anon" key,
-- so policies below define exactly what visitors can do.
alter table public.requests enable row level security;

-- Reps (anyone with the form link) can submit requests.
create policy "anyone can submit a request"
  on public.requests for insert to anon
  with check (true);

-- The dashboard (anyone with the link) can read the queue.
create policy "anyone can read the queue"
  on public.requests for select to anon
  using (true);

-- The dashboard can update triage fields (status, priority, jira, sprint).
-- Tighten this later if you move to passcode/login access.
create policy "anyone can triage requests"
  on public.requests for update to anon
  using (true)
  with check (true);
