-- ─────────────────────────────────────────────────────────────
--  Eventrack · Esquema de Supabase
--  Pega TODO esto en Supabase → SQL Editor → New query → Run.
-- ─────────────────────────────────────────────────────────────

-- 1) Tabla que guarda el estado completo de la app en una sola fila.
create table if not exists public.eventrack_state (
  id text primary key,
  data jsonb,
  updated_at timestamptz default now()
);

-- 2) Activar Row Level Security.
alter table public.eventrack_state enable row level security;

-- 3) Políticas: cualquiera con la clave pública puede leer y escribir.
--    (Equipo de confianza con el enlace. Si más adelante quieres login,
--     se cambian estas políticas.)
drop policy if exists "eventrack lectura publica" on public.eventrack_state;
create policy "eventrack lectura publica"
  on public.eventrack_state for select using (true);

drop policy if exists "eventrack insertar publico" on public.eventrack_state;
create policy "eventrack insertar publico"
  on public.eventrack_state for insert with check (true);

drop policy if exists "eventrack actualizar publico" on public.eventrack_state;
create policy "eventrack actualizar publico"
  on public.eventrack_state for update using (true) with check (true);

-- 4) Habilitar la sincronización en tiempo real para esta tabla.
alter publication supabase_realtime add table public.eventrack_state;
