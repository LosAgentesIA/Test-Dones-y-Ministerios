-- Fix RLS for anon insert (public includes anon)
drop policy if exists "anon_insert_personas" on public.personas;
drop policy if exists "anon_insert_resultados" on public.resultados;

create policy "allow_insert_personas" on public.personas
  for insert to public with check (true);

create policy "allow_insert_resultados" on public.resultados
  for insert to public with check (true);

-- Ensure select for authenticated remains
-- (already exists, but recreate to be safe)
drop policy if exists "auth_select_personas" on public.personas;
create policy "auth_select_personas" on public.personas
  for select to authenticated using (true);

drop policy if exists "auth_select_resultados" on public.resultados;
create policy "auth_select_resultados" on public.resultados
  for select to authenticated using (true);
