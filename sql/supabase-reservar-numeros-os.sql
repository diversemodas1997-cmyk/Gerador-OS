-- =====================================================================
-- FUNÇÃO reservar_numeros_os — o número da OS sai do BANCO, um de cada vez
-- =====================================================================
-- (28/10/2026, Jones: "conceder capacidade de criação de OS para outro usuário
-- além do admin, de forma que não aconteça conflito de dois usuários criarem
-- ordens de serviço diferentes com número igual".)
--
-- O PROBLEMA
-- O número da OS era escolhido NO NAVEGADOR: o maior número da lista que aquela
-- tela tinha na memória, mais um. Duas pessoas aprovando planejamentos ao mesmo
-- tempo liam a mesma lista e escolhiam o mesmo número; o servidor junta as OS
-- por id (são OS diferentes), então as duas entravam — com número igual.
--
-- A SOLUÇÃO
-- Quem aprova relê as OS do servidor e calcula o próximo número a partir delas
-- (`minimo`, a mesma regra de sempre, que ignora o número solto digitado por
-- engano). Depois pede ao banco um BLOCO de `qtd` números. O banco guarda o
-- último número já entregue (os_numeracao) e, com a linha TRAVADA (for update),
-- devolve o maior entre `minimo` e "último entregue + 1". Duas aprovações ao
-- mesmo tempo esperam uma pela outra na trava e saem com blocos diferentes —
-- mesmo antes de qualquer uma das OS chegar ao servidor.
--
-- O QUE MUDA NA PRÁTICA: o número de uma OS excluída não volta a ser usado
-- (o banco lembra que já o entregou). É de propósito: a folha e a etiqueta
-- daquele número podem já ter sido impressas.
--
-- Como rodar (uma vez, em cada servidor):
--   · fábrica: docker exec -i supabase-db psql -U postgres -d postgres < sql\supabase-reservar-numeros-os.sql
--   · nuvem:   SQL Editor → cole → Run.
-- Enquanto ela não existir, o app avisa no console e escolhe o número como
-- antes (relendo o servidor), sem a trava do banco.
-- =====================================================================

create table if not exists public.os_numeracao (
  id              text primary key,
  ultimo          integer not null default 0,
  atualizado_em   timestamptz not null default now(),
  atualizado_por  uuid
);

-- Ninguém lê nem escreve a tabela direto: só a função abaixo (security definer).
alter table public.os_numeracao enable row level security;
revoke all on table public.os_numeracao from anon, authenticated;

create or replace function public.reservar_numeros_os(
  qtd     integer,
  minimo  integer
)
returns integer
language plpgsql
security definer
set search_path = public
as $$
declare
  atual  integer;
  inicio integer;
begin
  if auth.uid() is null then
    raise exception 'reservar_numeros_os: é preciso estar logado';
  end if;
  if qtd is null or qtd < 1 or qtd > 60 then
    raise exception 'reservar_numeros_os: qtd fora do intervalo (1 a 60): %', qtd;
  end if;

  insert into public.os_numeracao (id, ultimo) values ('main', 0)
  on conflict (id) do nothing;

  -- A TRAVA: a segunda aprovação espera aqui até a primeira terminar.
  select ultimo into atual from public.os_numeracao where id = 'main' for update;

  inicio := greatest(atual + 1, coalesce(minimo, 1), 1);
  if inicio + qtd - 1 > 9999 then
    raise exception 'reservar_numeros_os: a numeração passaria de 9999';
  end if;

  update public.os_numeracao
     set ultimo = inicio + qtd - 1, atualizado_em = now(), atualizado_por = auth.uid()
   where id = 'main';

  return inicio;
end;
$$;

revoke all on function public.reservar_numeros_os(integer, integer) from public;
grant execute on function public.reservar_numeros_os(integer, integer) to authenticated;

-- Verificação:
-- select * from public.os_numeracao;


-- =====================================================================
-- A TRAVA DO PLANEJAMENTO (30/10/2026) — o mesmo planejamento não é aprovado
-- duas vezes ao mesmo tempo
-- =====================================================================
-- O QUE ACONTECEU: às 08:56:31 de 30/10/2026 o Admin e a Nathaly clicaram em
-- "Aprovar" no MESMO planejamento, cada um no seu computador. A trava acima
-- protegia o NÚMERO (cada um recebeu o seu: 700 e 701), mas não o
-- PLANEJAMENTO: os dois releram o servidor antes de qualquer um gravar, os dois
-- acharam o planejamento ainda na lista, e os dois geraram a OS. A gravação que
-- chegou por último apagou a outra, e o número 700 ficou pulado.
--
-- A CORREÇÃO: a reserva do número passa a dizer QUAL planejamento está sendo
-- aprovado (`plano`). O banco anota o planejamento em os_aprovacoes na mesma
-- transação da reserva; quem chega em segundo recebe o erro
-- "plano_ja_aprovado|<quem>|<quando>" e não ganha número nenhum.
--
-- ANOTAÇÃO VELHA NÃO TRAVA PARA SEMPRE: aprovação que reservou e não terminou
-- (internet caiu no meio, número já tomado) deixaria o planejamento preso. Depois
-- de 2 minutos a anotação pode ser retomada — aprovar leva segundos, e o app
-- confere antes, relendo o servidor, se o planejamento ainda existe (se a outra
-- aprovação terminou, ele já virou OS e saiu da lista).
--
-- DEVOLVER O NÚMERO: quando a aprovação desiste depois de reservar, o app chama
-- liberar_aprovacao_os. Ela apaga a anotação e, se ninguém reservou nada depois,
-- devolve o bloco (o contador volta para antes dele) — assim a desistência não
-- pula número.
--
-- O app continua funcionando sem estas funções: avisa no console e reserva sem
-- a trava do planejamento, como antes.
-- =====================================================================

create table if not exists public.os_aprovacoes (
  plano_id        text primary key,
  inicio          integer,
  qtd             integer,
  aprovado_por    uuid,
  aprovado_em     timestamptz not null default now()
);

alter table public.os_aprovacoes enable row level security;
revoke all on table public.os_aprovacoes from anon, authenticated;

create or replace function public.reservar_numeros_os(
  qtd     integer,
  minimo  integer,
  plano   text
)
returns integer
language plpgsql
security definer
set search_path = public
as $$
#variable_conflict use_variable
declare
  atual   integer;
  inicio  integer;
  outro   record;
begin
  if auth.uid() is null then
    raise exception 'reservar_numeros_os: é preciso estar logado';
  end if;
  if qtd is null or qtd < 1 or qtd > 60 then
    raise exception 'reservar_numeros_os: qtd fora do intervalo (1 a 60): %', qtd;
  end if;

  if plano is not null and plano <> '' then
    -- Anota o planejamento. INSERIR PRIMEIRO é o que trava de verdade: com
    -- duas aprovações ao mesmo tempo, a segunda espera a primeira terminar
    -- dentro deste insert e só então vê o conflito. (Ler antes, com "for
    -- update", não trava nada quando a linha ainda não existe — as duas
    -- passariam.)
    insert into public.os_aprovacoes (plano_id, aprovado_por, aprovado_em)
    values (plano, auth.uid(), now())
    on conflict (plano_id) do nothing;
    if not found then
      -- Já anotado. Há menos de 2 minutos: recusa, dizendo quem e quando.
      select a.aprovado_em, u.email into outro
        from public.os_aprovacoes a left join auth.users u on u.id = a.aprovado_por
       where a.plano_id = plano for update of a;
      if outro.aprovado_em > now() - interval '2 minutes' then
        raise exception 'plano_ja_aprovado|%|%', coalesce(outro.email, ''),
          to_char(outro.aprovado_em at time zone 'UTC', 'YYYY-MM-DD"T"HH24:MI:SS"Z"');
      end if;
      -- Anotação velha: a aprovação de antes não terminou. Retoma.
      update public.os_aprovacoes
         set aprovado_por = auth.uid(), aprovado_em = now(), inicio = null, qtd = null
       where plano_id = plano;
    end if;
  end if;

  insert into public.os_numeracao (id, ultimo) values ('main', 0)
  on conflict (id) do nothing;

  select ultimo into atual from public.os_numeracao where id = 'main' for update;
  inicio := greatest(atual + 1, coalesce(minimo, 1), 1);
  if inicio + qtd - 1 > 9999 then
    raise exception 'reservar_numeros_os: a numeração passaria de 9999';
  end if;

  update public.os_numeracao
     set ultimo = inicio + qtd - 1, atualizado_em = now(), atualizado_por = auth.uid()
   where id = 'main';

  if plano is not null and plano <> '' then
    -- (A tabela tem colunas com os nomes das variáveis; o #variable_conflict
    -- do começo faz o lado direito ser a variável.)
    update public.os_aprovacoes set inicio = inicio, qtd = qtd
     where plano_id = plano;
  end if;
  return inicio;
end;
$$;

revoke all on function public.reservar_numeros_os(integer, integer, text) from public;
grant execute on function public.reservar_numeros_os(integer, integer, text) to authenticated;

-- A aprovação desistiu depois de reservar: solta o planejamento e, se o bloco
-- dele ainda é o último entregue, devolve os números. Só quem anotou solta.
-- `devolver` = false quando o motivo foi um número do bloco já ser de outra OS:
-- devolvido, o bloco voltaria igual na próxima tentativa e falharia de novo.
create or replace function public.liberar_aprovacao_os(plano text, devolver boolean default true)
returns void
language plpgsql
security definer
set search_path = public
as $$
declare
  a record;
begin
  if auth.uid() is null or plano is null then return; end if;
  delete from public.os_aprovacoes
   where plano_id = plano and aprovado_por = auth.uid()
  returning inicio, qtd into a;
  if devolver and a.inicio is not null and a.qtd is not null then
    update public.os_numeracao
       set ultimo = a.inicio - 1, atualizado_em = now(), atualizado_por = auth.uid()
     where id = 'main' and ultimo = a.inicio + a.qtd - 1;
  end if;
end;
$$;

revoke all on function public.liberar_aprovacao_os(text, boolean) from public;
grant execute on function public.liberar_aprovacao_os(text, boolean) to authenticated;

-- O PostgREST guarda a lista de funções em cache: sem isto, as novas só
-- apareceriam para o app depois de reiniciar o container.
notify pgrst, 'reload schema';

-- Verificação:
-- select * from public.os_aprovacoes order by aprovado_em desc limit 10;
