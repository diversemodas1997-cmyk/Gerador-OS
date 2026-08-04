-- =====================================================================
-- TABELA compras_materiais — ponte com o programa de Contabilidade
-- =====================================================================
-- O sistema de Contabilidade (Flask/Render) importa as NF-e, traduz cada
-- item para (tecido, cor, kg) e GRAVA aqui usando a service_role key.
-- O Gerador-OS apenas LÊ esta tabela e soma os registros como ENTRADAS
-- no painel de estoque (origem "NF").
--
-- Como rodar (uma vez):
--   1. https://supabase.com/dashboard → projeto ckkqrjkhorvaahyazqsr
--   2. SQL Editor → New query → cole tudo → Run.
-- =====================================================================

create table if not exists compras_materiais (
  id            uuid primary key default gen_random_uuid(),
  -- chave única da origem (chave da NF + índice do item) — garante que
  -- reenviar a mesma NF NÃO duplique a entrada (upsert por id_origem).
  id_origem     text unique,
  nota_fiscal   text,            -- número da NF (exibição)
  chave_nfe     text,            -- chave de acesso (44 dígitos), opcional
  fornecedor    text,
  data          date,
  tecido_nome   text not null,   -- precisa bater com o cadastro de Tecidos do Gerador-OS
  cor_nome      text,            -- precisa bater com o cadastro de Cores
  quantidade_kg numeric not null,
  valor_total   numeric,         -- opcional (referência de custo)
  obs           text,
  empresa_cnpj  text,            -- CNPJ da empresa na contabilidade (rastreio)
  created_at    timestamptz default now(),
  updated_at    timestamptz default now()
);

create index if not exists idx_compras_tecido on compras_materiais (tecido_nome);
create index if not exists idx_compras_data   on compras_materiais (data);

-- RLS: qualquer usuário autenticado do Gerador-OS pode LER. A escrita é feita
-- pela contabilidade com service_role (que ignora RLS), então não há policy
-- de insert/update para 'authenticated' — leitura apenas.
alter table compras_materiais enable row level security;

drop policy if exists "compras: authenticated select" on compras_materiais;
create policy "compras: authenticated select"
  on compras_materiais for select
  to authenticated
  using (true);

-- Realtime: faz as compras aparecerem na hora no estoque do Gerador-OS.
do $$
begin
  if not exists (select 1 from pg_publication where pubname = 'supabase_realtime') then
    create publication supabase_realtime;
  end if;
end $$;

do $$
begin
  if not exists (
    select 1 from pg_publication_tables
     where pubname = 'supabase_realtime' and tablename = 'compras_materiais'
  ) then
    alter publication supabase_realtime add table compras_materiais;
  end if;
end $$;

-- Verificação:
-- select * from pg_policies where tablename='compras_materiais';
-- select * from pg_publication_tables where pubname='supabase_realtime' and tablename='compras_materiais';
