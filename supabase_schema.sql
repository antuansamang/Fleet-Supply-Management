-- Explorer Fleet Control · esquema de base de datos
-- Ejecutar en Supabase: Dashboard -> SQL Editor -> New query -> pegar todo -> Run

create table if not exists hx_requisiciones (
  id text primary key,
  barco_id text not null,
  data jsonb not null,
  updated_at timestamptz not null default now()
);

create table if not exists hx_facturas (
  id text primary key,
  barco_id text not null,
  data jsonb not null,
  updated_at timestamptz not null default now()
);

create table if not exists hx_config (
  id text primary key,
  data jsonb not null,
  updated_at timestamptz not null default now()
);

create index if not exists hx_requisiciones_barco_idx on hx_requisiciones (barco_id);
create index if not exists hx_facturas_barco_idx on hx_facturas (barco_id);

-- Seguridad a nivel de fila (RLS)
alter table hx_requisiciones enable row level security;
alter table hx_facturas enable row level security;
alter table hx_config enable row level security;

-- Nota de seguridad: estas políticas permiten lectura/escritura a cualquiera que tenga
-- la "anon key" pública (que va incrustada en el sitio, así que es efectivamente pública
-- si el sitio es público). Esto es aceptable como punto de partida para una herramienta
-- interna de la flota, pero si más adelante quieres restringir quién puede editar,
-- se puede añadir autenticación (login) y cambiar estas políticas para exigir un usuario
-- autenticado.

drop policy if exists "acceso total anon" on hx_requisiciones;
create policy "acceso total anon" on hx_requisiciones
  for all using (true) with check (true);

drop policy if exists "acceso total anon" on hx_facturas;
create policy "acceso total anon" on hx_facturas
  for all using (true) with check (true);

drop policy if exists "acceso total anon" on hx_config;
create policy "acceso total anon" on hx_config
  for all using (true) with check (true);
