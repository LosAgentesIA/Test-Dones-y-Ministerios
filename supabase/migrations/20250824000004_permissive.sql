-- Permissive for anon insert (public includes anon)
drop policy if exists "allow_insert_personas" on public.personas;
drop policy if exists "allow_insert_resultados" on public.resultados;
drop policy if exists "allow_select_personas" on public.personas;
drop policy if exists "allow_select_resultados" on public.resultados;
drop policy if exists "allow_authenticated_insert_personas" on public.personas;
drop policy if exists "allow_authenticated_insert_resultados" on public.resultados;
drop policy if exists "allow_all_personas" on public.personas;
drop policy if exists "allow_all_resultados" on public.resultados;

create policy "allow_all_personas" on public.personas
  for all to public using (true) with check (true);

create policy "allow_all_resultados" on public.resultados
  for all to public using (true) with check (true);
