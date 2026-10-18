-- Chay rieng file nay trong Supabase > SQL Editor > New query.
-- Bang quan he cap tren doc lap voi bang agent_stats.
create table if not exists public.agent_hierarchy (
  agent_id text primary key,
  parent_agent_id text,
  updated_at timestamptz not null default now(),
  constraint agent_hierarchy_not_self check (parent_agent_id is null or parent_agent_id <> agent_id)
);

alter table public.agent_hierarchy enable row level security;
grant select, insert, update, delete on public.agent_hierarchy to authenticated;

drop policy if exists "Authenticated users can read agent hierarchy" on public.agent_hierarchy;
create policy "Authenticated users can read agent hierarchy"
  on public.agent_hierarchy for select to authenticated using (true);

drop policy if exists "Authenticated users can insert agent hierarchy" on public.agent_hierarchy;
create policy "Authenticated users can insert agent hierarchy"
  on public.agent_hierarchy for insert to authenticated with check (true);

drop policy if exists "Authenticated users can update agent hierarchy" on public.agent_hierarchy;
create policy "Authenticated users can update agent hierarchy"
  on public.agent_hierarchy for update to authenticated using (true) with check (true);

drop policy if exists "Authenticated users can delete agent hierarchy" on public.agent_hierarchy;
create policy "Authenticated users can delete agent hierarchy"
  on public.agent_hierarchy for delete to authenticated using (true);

create index if not exists agent_hierarchy_parent_idx on public.agent_hierarchy(parent_agent_id);
