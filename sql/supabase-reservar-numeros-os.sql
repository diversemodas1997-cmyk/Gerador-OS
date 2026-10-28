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
