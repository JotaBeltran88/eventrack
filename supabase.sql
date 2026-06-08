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


-- ─────────────────────────────────────────────────────────────
--  ACCESO POR CÓDIGO (Admin / Contador)
--  Los códigos se guardan aquí pero NO son legibles desde la app:
--  la comprobación se hace con funciones del servidor (SECURITY DEFINER).
-- ─────────────────────────────────────────────────────────────

-- 5) Tabla de códigos (sin políticas de lectura => nadie puede leerla con la clave pública).
create table if not exists public.eventrack_config (
  id text primary key,
  admin_code text not null,
  contador_code text not null
);
alter table public.eventrack_config enable row level security;

-- 6) Códigos iniciales (CÁMBIALOS luego desde la app, en ⚙ Acceso).
insert into public.eventrack_config (id, admin_code, contador_code)
values ('codes', 'admin1234', 'contar1234')
on conflict (id) do nothing;

-- 6b) Tercer rol: código de Inventario (puede editar el stock inicial).
alter table public.eventrack_config add column if not exists inventario_code text;
update public.eventrack_config set inventario_code = 'inventario1234'
  where id = 'codes' and (inventario_code is null or inventario_code = '');

-- 7) Función de login: devuelve 'admin', 'inventario', 'contador' o '' según el código.
create or replace function public.eventrack_login(code text)
returns text
language sql
security definer
set search_path = public
as $$
  select case
    when code = (select admin_code from public.eventrack_config where id = 'codes') then 'admin'
    when code = (select inventario_code from public.eventrack_config where id = 'codes') then 'inventario'
    when code = (select contador_code from public.eventrack_config where id = 'codes') then 'contador'
    else ''
  end;
$$;

-- 8) Función para cambiar los códigos: exige el código de admin actual.
drop function if exists public.eventrack_set_codes(text, text, text);
create or replace function public.eventrack_set_codes(current_admin text, new_admin text, new_contador text, new_inventario text)
returns boolean
language plpgsql
security definer
set search_path = public
as $$
declare ok boolean;
begin
  select (current_admin = admin_code) into ok from public.eventrack_config where id = 'codes';
  if not ok then return false; end if;
  update public.eventrack_config
    set admin_code = new_admin, contador_code = new_contador, inventario_code = new_inventario
    where id = 'codes';
  return true;
end;
$$;

-- 9) Permitir que la app (clave pública) pueda EJECUTAR las funciones (no leer la tabla).
grant execute on function public.eventrack_login(text) to anon, authenticated;
grant execute on function public.eventrack_set_codes(text, text, text, text) to anon, authenticated;
