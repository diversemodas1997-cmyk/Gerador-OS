-- =====================================================================
-- TABELA mensagens — o recado de todo mundo, num canal só
-- =====================================================================
-- O programa ganhou um campo de mensagens onde toda a fábrica se fala: um
-- canal único, sem conversa privada e sem grupo. Quem manda escreve para
-- todos; quem abre lê tudo o que foi dito.
--
-- POR QUE UMA TABELA, E NÃO O BLOB DOS DADOS
-- Todo o resto do programa vive numa linha só (shared_data), que desce e sobe
-- INTEIRA — hoje quase 2 MB. Cada mensagem enviada custaria uma subida dessas,
-- e cada mensagem recebida, uma descida, para todos os computadores. Uma
-- tabela própria custa alguns bytes por mensagem e chega na hora pelo
-- Realtime — que é como a conversa tem de ser.
--
-- Como rodar (uma vez, em cada servidor):
--   · servidor da fábrica:
--       docker exec -i supabase-db psql -U postgres -d postgres < sql\supabase-mensagens.sql
--   · nuvem: https://supabase.com/dashboard → SQL Editor → cole tudo → Run.
--     (Na nuvem o programa abre em modo consulta, então lá isto só serve para
--      LER o que foi dito na fábrica.)
-- =====================================================================

create table if not exists mensagens (
  id        uuid primary key default gen_random_uuid(),
  criado_em timestamptz not null default now(),
  autor_id  uuid,            -- quem escreveu (auth.uid()), para a tranca do RLS
  autor     text,            -- o LOGIN no momento em que escreveu, congelado
  texto     text not null
);

create index if not exists idx_mensagens_criado_em on mensagens (criado_em);

-- RLS — a conversa é de quem tem conta no programa.
alter table mensagens enable row level security;

-- LER: todo mundo logado lê tudo. É um canal só, e é esse o ponto.
drop policy if exists "mensagens: authenticated select" on mensagens;
create policy "mensagens: authenticated select"
  on mensagens for select
  to authenticated
  using (true);

-- ESCREVER: cada um manda em NOME PRÓPRIO. O `autor_id = auth.uid()` é o que
-- impede alguém de gravar um recado assinado por outra pessoa.
drop policy if exists "mensagens: authenticated insert" on mensagens;
create policy "mensagens: authenticated insert"
  on mensagens for insert
  to authenticated
  with check (autor_id = auth.uid());

-- APAGAR: a própria mensagem, e o admin qualquer uma. Ninguém apaga o recado
-- do outro — o mesmo princípio da observação da folha de OS.
drop policy if exists "mensagens: apagar a propria" on mensagens;
create policy "mensagens: apagar a propria"
  on mensagens for delete
  to authenticated
  using (
    autor_id = auth.uid()
    or exists (
      select 1 from public.user_roles
       where user_id = auth.uid() and role = 'admin'
    )
  );

-- Realtime: a mensagem aparece na tela dos outros sem ninguém recarregar.
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
     where pubname = 'supabase_realtime' and tablename = 'mensagens'
  ) then
    alter publication supabase_realtime add table mensagens;
  end if;
end $$;

-- Verificação:
-- select * from pg_policies where tablename='mensagens';
-- select * from pg_publication_tables where pubname='supabase_realtime' and tablename='mensagens';
