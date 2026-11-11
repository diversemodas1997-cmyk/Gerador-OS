-- =====================================================================
-- TABELA os_arquivadas — as OS antigas, fora do blob
-- =====================================================================
-- 11/11/2026, Junior: "insira a automação do programa em arquivar OS
-- antigas, deixando acesso imediato às 100 OS mais recentes".
--
-- POR QUE UMA TABELA, E NÃO O BLOB
-- Todo o programa vive numa linha só (shared_data), e `ordens` era quase
-- metade dela: 466 OS, ~4 MB de JSON, subindo e descendo a cada gravação de
-- qualquer computador. A maior parte era OS finalizada há meses. Aqui cada
-- OS arquivada é uma linha, gravada uma vez; cada computador a baixa uma vez
-- só e guarda a cópia no navegador (ver ARQUIVO DE OS no app.js).
--
-- O QUE VAI PARA CÁ: OS fora das 100 de número mais alto E finalizada ou
-- cancelada (com a família conjugada inteira). Quem decide é o programa, no
-- computador do admin. A OS continua na memória de todos — o estoque e os
-- relatórios seguem vendo tudo; só o blob emagrece.
--
-- `dados` é TEXTO, e não jsonb, de propósito: o jsonb reordena as chaves, e o
-- programa compara o texto da OS para saber se alguém mexeu nela depois de
-- arquivada (mexeu = ela volta para as ativas sozinha).
--
-- `seq` cresce a cada gravação ou alteração da linha: é por ele que cada
-- computador baixa só o que mudou desde a última vez.
--
-- `removida` = a OS voltou para as ativas (alguém mexeu nela) ou foi
-- excluída. A linha fica, marcada, para os outros computadores tirarem da
-- cópia deles.
--
-- Como rodar (uma vez, em cada servidor; tudo aqui é idempotente):
--   · servidor da fábrica:
--       docker exec -i supabase-db psql -U postgres -d postgres < sql\supabase-os-arquivadas.sql
--   · nuvem: https://supabase.com/dashboard → SQL Editor → cole tudo → Run.
-- Sem esta tabela o programa simplesmente não arquiva (e avisa no console).
-- =====================================================================

create sequence if not exists os_arquivadas_seq;

create table if not exists os_arquivadas (
  id            text primary key,          -- o id da OS (o mesmo de shared_data)
  os            text,                      -- o número, para consulta à mão
  dados         text not null,             -- a OS inteira, em JSON (texto: ver acima)
  removida      boolean not null default false,
  arquivada_em  timestamptz not null default now(),
  arquivada_por uuid,
  atualizado_em timestamptz not null default now(),
  seq           bigint not null default nextval('os_arquivadas_seq')
);

create index if not exists idx_os_arquivadas_seq on os_arquivadas (seq);
create index if not exists idx_os_arquivadas_os on os_arquivadas (os);

-- Toda alteração ganha um seq novo: é o que faz o outro computador perceber
-- que a OS voltou para as ativas (ou foi arquivada de novo).
create or replace function os_arquivadas_carimbar() returns trigger
language plpgsql as $$
begin
  new.seq := nextval('os_arquivadas_seq');
  new.atualizado_em := now();
  return new;
end $$;

drop trigger if exists trg_os_arquivadas_carimbar on os_arquivadas;
create trigger trg_os_arquivadas_carimbar
  before update on os_arquivadas
  for each row execute function os_arquivadas_carimbar();

-- RLS — as mesmas contas que gravam o blob.
alter table os_arquivadas enable row level security;

-- LER: todo mundo logado (a OS arquivada continua sendo consultada).
drop policy if exists "os_arquivadas: authenticated select" on os_arquivadas;
create policy "os_arquivadas: authenticated select"
  on os_arquivadas for select to authenticated using (true);

-- ARQUIVAR: só o admin (é a rotina dele que arquiva).
drop policy if exists "os_arquivadas: admin insert" on os_arquivadas;
create policy "os_arquivadas: admin insert"
  on os_arquivadas for insert to authenticated
  with check (exists (select 1 from public.user_roles where user_id = auth.uid() and role = 'admin'));

-- ALTERAR: qualquer conta logada. Quem marca o status de uma OS arquivada
-- devolve-a às ativas, e esse usuário pode não ser admin. É o mesmo poder
-- que ele já tem sobre a OS no blob.
drop policy if exists "os_arquivadas: authenticated update" on os_arquivadas;
create policy "os_arquivadas: authenticated update"
  on os_arquivadas for update to authenticated using (true) with check (true);

-- APAGAR linha: só o admin, à mão. O programa nunca apaga — marca `removida`.
drop policy if exists "os_arquivadas: admin delete" on os_arquivadas;
create policy "os_arquivadas: admin delete"
  on os_arquivadas for delete to authenticated
  using (exists (select 1 from public.user_roles where user_id = auth.uid() and role = 'admin'));

grant select, insert, update, delete on os_arquivadas to authenticated;
grant usage, select on sequence os_arquivadas_seq to authenticated;
