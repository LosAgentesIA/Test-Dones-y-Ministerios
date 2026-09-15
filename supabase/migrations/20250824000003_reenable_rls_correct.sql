-- Re-enable RLS with correct policies
alter table public.personas enable row level security;
alter table public.resultados enable row level security;

-- Drop old policies
drop policy if exists "allow_insert_personas" on public.personas;
drop policy if exists "allow_insert_resultados" on public.resultados;
drop policy if exists "allow_all_personas" on public.personas;
drop policy if exists "allow_all_resultados" on public.resultados;
drop policy if exists "anon_insert_personas" on public.personas;
drop policy if exists "anon_insert_resultados" on public.resultados;
drop policy if exists "auth_select_personas" on public.personas;
drop policy if exists "auth_select_resultados" on public.resultados;

-- Allow anyone (anon) to insert (public includes anon)
create policy "allow_insert_personas" on public.personas
  for insert to public with check (true);

create policy "allow_insert_resultados" on public.resultados
  for insert to public with check (true);

-- Only authenticated (admin) can select
create policy "allow_select_personas" on public.personas
  for select to authenticated using (true);

create policy "allow_select_resultados" on public.resultados
  for select to authenticated using (true);

-- Also allow authenticated to insert (for admin tests)
create policy "allow_authenticated_insert_personas" on public.personas
  for insert to authenticated with check (true);

create policy "allow_authenticated_insert_resultados" on public.resultados
  for insert to authenticated with check (true);
