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

-- Corrigido quando? Nulo = como saiu. Serve so para a tela escrever "(editado)"
-- ao lado da hora: quem le precisa saber que aquele texto nao e exatamente o
-- que estava ali antes. A HORA DO RECADO continua sendo `criado_em` — e por ela
-- que a conversa se ordena, e corrigir nao faz o recado pular para o fim.
-- (Coluna acrescentada em 27/08/2026, junto com a correcao de 5 minutos.
--  Rodar este arquivo de novo num servidor que ja tinha a tabela e seguro:
--  tudo aqui e idempotente.)
alter table mensagens add column if not exists editado_em timestamptz;

-- A QUEM ESTE RECADO RESPONDE. Nulo = recado solto, que e o caso comum. Com
-- valor, a tela desenha a mensagem recuada logo abaixo daquela a que ela
-- responde — um nivel so, de proposito: conversa de fabrica se le de cima para
-- baixo, e arvore de resposta dentro de resposta vira lista que ninguem segue.
--
-- `on delete set null`, e NAO cascade: apagar o proprio recado nao pode levar
-- junto as respostas dos outros. Sem a mae, a resposta volta a ser um recado
-- solto na conversa — continua la, que e o que importa.
-- (Coluna acrescentada em 27/08/2026.)
alter table mensagens add column if not exists responde_a uuid
  references mensagens(id) on delete set null;
create index if not exists idx_mensagens_responde_a on mensagens (responde_a);

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

-- CORRIGIR: so o AUTOR, e so nos 5 primeiros minutos.
--
-- Por que existe: quem manda recado erra o numero da OS, o tamanho, a palavra.
-- Reescrever logo e conserto. Reescrever depois e outra coisa — quem ja leu
-- "manda 200" e volta e encontra "manda 400" nao tem como saber que mudou, e o
-- canal deixa de valer como registro do que foi combinado.
--
-- Por que a TRANCA E AQUI: no programa o lapis some quando o prazo acaba, mas
-- isso e o relogio DAQUELE computador — atrasado, adiantado, ou a tela aberta
-- desde antes. `now()` aqui e o relogio do servidor, um so para todo mundo.
--
-- O ADMIN NAO TEM PRAZO (27/08/2026, decisao do Junior): ele corrige as
-- PROPRIAS mensagens a qualquer momento. O relogio existe para o recado da
-- producao nao ser reescrito depois de lido; a conta que administra o programa
-- e a que responde pelo canal, e corrige o que ficou errado ali no dia seguinte
-- se for o caso. A dispensa e SO do prazo.
--
-- NEM O ADMIN corrige recado ALHEIO (e ele pode APAGAR, isso segue valendo):
-- apagar deixa claro que sumiu; reescrever poe a palavra de um na boca do
-- outro. Mesmo principio da observacao da folha de OS. Por isso o
-- `autor_id = auth.uid()` fica FORA do parenteses: ele vale para todo mundo, e
-- o que o admin dispensa e so a segunda condicao.
--
-- O `with check` repete o dono e o prazo porque a linha e conferida DEPOIS da
-- alteracao: sem ele, um update poderia trocar `autor_id` ou empurrar
-- `criado_em` para a frente e renovar o proprio prazo.
drop policy if exists "mensagens: corrigir a propria em 5 min" on mensagens;
create policy "mensagens: corrigir a propria em 5 min"
  on mensagens for update
  to authenticated
  using (
    autor_id = auth.uid()
    and (
      criado_em > now() - interval '5 minutes'
      or exists (
        select 1 from public.user_roles
         where user_id = auth.uid() and role = 'admin'
      )
    )
  )
  with check (
    autor_id = auth.uid()
    and (
      criado_em > now() - interval '5 minutes'
      or exists (
        select 1 from public.user_roles
         where user_id = auth.uid() and role = 'admin'
      )
    )
  );

-- NUM UPDATE, SO O TEXTO MUDA.
--
-- O `with check` acima nao basta sozinho, e isto foi MEDIDO: ele confere a
-- linha como ela FICOU, e nao como ela era. Um update que faca
-- `criado_em = now()` passa na conferencia (a linha fica dentro do prazo) e
-- renova o prazo para sempre — alem de mudar a hora do recado na conversa dos
-- outros. RLS nao enxerga o valor antigo; trigger enxerga.
--
-- Entao a hora, o dono e o nome do autor voltam a ser o que eram, e o carimbo
-- de "editado" e do SERVIDOR, nao do que o programa mandou. Sobra o texto, que
-- e o que se quis deixar corrigir.
--
-- Vale para todo mundo, inclusive service_role: manutencao que precise mexer
-- de verdade numa linha desliga o gatilho na transacao
-- (`alter table mensagens disable trigger trg_mensagens_so_o_texto`).
create or replace function mensagens_so_o_texto() returns trigger
language plpgsql as $$
begin
  new.id        := old.id;
  new.criado_em := old.criado_em;
  new.autor_id  := old.autor_id;
  new.autor     := old.autor;
  -- Corrigir o texto nao remaneja a conversa: a resposta nao muda de mae.
  -- O NULO passa, e passa por um motivo MEDIDO: apagar a mae dispara o
  -- `on delete set null` da chave estrangeira, que chega aqui como um update
  -- pondo `responde_a = null`. Devolvendo o valor antigo, o gatilho recolocava
  -- o id da linha que estava sendo apagada — e o proprio banco recusava o
  -- DELETE por violacao da chave. Ou seja: sem esta excecao, nenhuma mensagem
  -- com resposta podia ser apagada.
  -- O que sobra de brecha e um cliente desamarrar a PROPRIA resposta (ela vira
  -- recado solto na conversa); trocar de mae, nao.
  if new.responde_a is not null then new.responde_a := old.responde_a; end if;
  if new.texto is distinct from old.texto then
    new.editado_em := now();
  else
    new.editado_em := old.editado_em;
  end if;
  return new;
end $$;

drop trigger if exists trg_mensagens_so_o_texto on mensagens;
create trigger trg_mensagens_so_o_texto
  before update on mensagens
  for each row execute function mensagens_so_o_texto();

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

-- =====================================================================
-- REACOES — o polegar de "vi e concordo"
-- =====================================================================
-- Metade dos recados da fabrica so pede um "ok". Escrever "ok" gasta uma linha
-- da conversa por pessoa; o polegar responde sem empurrar o resto para cima.
--
-- Tabela propria, e nao uma coluna em `mensagens`: reagir e mexer na linha de
-- OUTRA pessoa, e a politica de UPDATE de mensagens diz justamente que ninguem
-- escreve na linha alheia. Aqui cada reacao e uma linha de quem reagiu, com o
-- dono na chave — ninguem reage pelos outros e ninguem tira a reacao alheia.
--
-- A chave primaria (mensagem_id, user_id, reacao) e a trava do clique repetido:
-- reagir duas vezes na mesma mensagem nao cria duas linhas.
create table if not exists mensagem_reacoes (
  mensagem_id uuid not null references mensagens(id) on delete cascade,
  user_id     uuid not null,
  reacao      text not null default '+1',
  criado_em   timestamptz not null default now(),
  primary key (mensagem_id, user_id, reacao)
);

create index if not exists idx_mensagem_reacoes_msg on mensagem_reacoes (mensagem_id);

alter table mensagem_reacoes enable row level security;

-- LER: quem tem conta le todas. O numero ao lado do polegar e publico, como o
-- recado.
drop policy if exists "reacoes: authenticated select" on mensagem_reacoes;
create policy "reacoes: authenticated select"
  on mensagem_reacoes for select to authenticated using (true);

-- REAGIR: em nome proprio, e so.
drop policy if exists "reacoes: reagir em nome proprio" on mensagem_reacoes;
create policy "reacoes: reagir em nome proprio"
  on mensagem_reacoes for insert to authenticated
  with check (user_id = auth.uid());

-- TIRAR: a propria reacao. Nem o admin tira a dos outros — nao ha nada a
-- moderar num polegar, e a linha diz quem a pos.
drop policy if exists "reacoes: tirar a propria" on mensagem_reacoes;
create policy "reacoes: tirar a propria"
  on mensagem_reacoes for delete to authenticated
  using (user_id = auth.uid());

do $$
begin
  if not exists (
    select 1 from pg_publication_tables
     where pubname = 'supabase_realtime' and tablename = 'mensagem_reacoes'
  ) then
    alter publication supabase_realtime add table mensagem_reacoes;
  end if;
end $$;

-- =====================================================================
-- perfis() — a lista de nomes para a mencao com @
-- =====================================================================
-- Digitar "@" no campo de recado abre a lista de quem existe no programa. Ate
-- 27/08/2026 essa lista nao existia para quem nao e admin: `user_roles` so tem
-- o id (e usuario comum nem tem linha la), e a lista completa mora em
-- `auth.users`, fora do alcance da chave anonima. Quem quisesse mencionar
-- alguem tinha de acertar o nome de cabeca — e era o que estava acontecendo,
-- "@enfesto.corte:" digitado a mao nos recados.
--
-- Esta funcao devolve o MINIMO para a mencao funcionar: o id e o LOGIN (o
-- pedaco antes do @ do e-mail), que e o nome que a fabrica ja usa e ja aparece
-- em cada recado e em cada observacao de OS. Nao devolve e-mail completo, nem
-- papel, nem data, nem nada de senha.
--
-- `security definer` porque quem chama nao tem acesso ao schema auth; o
-- `search_path` fixo e a regra de ouro dessas funcoes (sem ele, um schema
-- plantado no caminho troca o que `users` significa). Conta apagada nao entra.
create or replace function public.perfis()
returns table (user_id uuid, login text)
language sql
security definer
set search_path = public, auth
as $$
  select u.id, split_part(u.email, '@', 1)
    from auth.users u
   where u.deleted_at is null
     and coalesce(u.email, '') <> ''
   order by 2
$$;

revoke all on function public.perfis() from public;
grant execute on function public.perfis() to authenticated;

-- Verificação:
-- select * from public.perfis();
-- select * from pg_policies where tablename='mensagens';
-- select * from pg_policies where tablename='mensagem_reacoes';
-- select tgname from pg_trigger where tgrelid='mensagens'::regclass and not tgisinternal;
-- select * from pg_publication_tables where pubname='supabase_realtime' and tablename='mensagens';
