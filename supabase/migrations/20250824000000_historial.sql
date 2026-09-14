-- Historial de tests por persona (nombre+email) - IMES
-- Proyecto: xstrpgjvflpajgcbzsai

create table if not exists public.personas (
  id uuid primary key default gen_random_uuid(),
  nombre text not null,
  email_norm text not null unique,
  created_at timestamptz default now()
);

create table if not exists public.resultados (
  id uuid primary key default gen_random_uuid(),
  persona_id uuid references public.personas(id) on delete set null,
  nombre text not null,
  email text not null,
  email_norm text not null,
  test_tipo text not null,
  fecha timestamptz default now(),
  respuestas jsonb not null,
  puntajes jsonb not null
);

alter table public.personas enable row level security;
alter table public.resultados enable row level security;

drop policy if exists "anon_insert_personas" on public.personas;
create policy "anon_insert_personas" on public.personas
  for insert to anon with check (true);

drop policy if exists "anon_insert_resultados" on public.resultados;
create policy "anon_insert_resultados" on public.resultados
  for insert to anon with check (true);

drop policy if exists "auth_select_personas" on public.personas;
create policy "auth_select_personas" on public.personas
  for select to authenticated using (true);

drop policy if exists "auth_select_resultados" on public.resultados;
create policy "auth_select_resultados" on public.resultados
  for select to authenticated using (true);

create index if not exists idx_resultados_email_norm_fecha on public.resultados(email_norm, fecha desc);
create index if not exists idx_personas_email_norm on public.personas(email_norm);
