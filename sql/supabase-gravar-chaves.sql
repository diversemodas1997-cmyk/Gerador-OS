-- =====================================================================
-- FUNÇÃO gravar_chaves — subir só a parte que mudou
-- =====================================================================
-- O PROBLEMA QUE ELA RESOLVE
-- Todos os dados do programa vivem numa linha só (shared_data.id='main'), numa
-- coluna jsonb. Para trocar uma chave — marcar uma etapa mexe em `ordens` —, o
-- app mandava a linha INTEIRA de volta: em 26/08/2026 eram 2,5 MB por gravação,
-- e cada OS nova acrescenta ~5 KB. O ícone ficava em "Salvando" por segundos, as
-- gravações entravam em fila e quem fechasse o programa nesse meio-tempo
-- arriscava perder a última alteração.
--
-- A LEITURA já tinha sido resolvida (mapa de versões por chave na tabela-sinal:
-- cada máquina baixa só a chave que mudou). Faltava a ESCRITA, que é o que esta
-- função faz: o app manda apenas as chaves sujas, e o PRÓPRIO BANCO as costura
-- dentro do jsonb que já está lá.
--
--   data || pares   →  troca as chaves recebidas, mantém todas as outras.
--
-- O QUE ISSO GANHA, com os números de hoje (blob de 2,5 MB):
--   · marcar status, mexer em expedição, plano, compra, configuração… passam a
--     mandar KILOBYTES em vez de 2,5 MB;
--   · salvar uma OS ainda manda a chave `ordens` inteira (1,35 MB), porque as
--     251 OS são UM valor dentro do jsonb — mas para de arrastar junto
--     `operacoes` (500 KB), `grades`, `desenhos` e o resto.
-- Descer abaixo desse piso exige quebrar `ordens` em um registro por linha, que
-- é outra obra — e esta função continua valendo quando ela for feita.
--
-- O QUE ELA NÃO FAZ: apagar chave. `||` só acrescenta e substitui — e o app
-- nunca apaga chave do blob. Restaurar backup e importar dados marcam TODAS as
-- chaves como sujas, então continuam reescrevendo tudo, como antes.
--
-- SEGURANÇA: `security invoker` de propósito. A função roda com as permissões
-- de quem chamou, então continua valendo o RLS da tabela (só usuário logado
-- escreve). Ela não é uma porta nova — é o mesmo update de sempre, feito do
-- lado do servidor para não ter de trafegar o resto.
--
-- Como rodar (uma vez, em cada servidor):
--   · fábrica: docker exec -i supabase-db psql -U postgres -d postgres < sql\supabase-gravar-chaves.sql
--   · nuvem:   SQL Editor → cole → Run.
--
-- Enquanto ela não existir, o app percebe e volta sozinho a mandar o blob
-- inteiro — mais lento, porém correto. Nada quebra por falta dela.
-- =====================================================================

create or replace function public.gravar_chaves(
  pares   jsonb,
  quem    uuid,
  carimbo timestamptz
)
returns void
language plpgsql
security invoker
set search_path = public
as $$
begin
  if pares is null or jsonb_typeof(pares) <> 'object' then
    raise exception 'gravar_chaves: pares tem de ser um objeto jsonb';
  end if;

  update public.shared_data
     set data       = coalesce(data, '{}'::jsonb) || pares,
         updated_at = coalesce(carimbo, now()),
         updated_by = quem
   where id = 'main';

  -- Primeira gravação da vida (banco recém-criado): a linha ainda não existe.
  if not found then
    insert into public.shared_data (id, data, updated_at, updated_by)
    values ('main', pares, coalesce(carimbo, now()), quem);
  end if;
end;
$$;

-- Quem pode chamar: usuário logado. A tranca de verdade continua sendo o RLS da
-- tabela — esta linha só evita que a função apareça para quem nem login tem.
revoke all on function public.gravar_chaves(jsonb, uuid, timestamptz) from public;
grant execute on function public.gravar_chaves(jsonb, uuid, timestamptz) to authenticated;

-- Verificação:
-- select proname, prosecdef from pg_proc where proname = 'gravar_chaves';


-- =====================================================================
-- FUNÇÃO gravar_chaves_se — grava só se ninguém gravou desde a leitura
-- =====================================================================
-- (30/10/2026) O QUE ACONTECEU: o Admin e a Nathaly aprovaram ao mesmo tempo e
-- os dois computadores gravaram a lista de OS no mesmo segundo. Cada um tinha
-- relido o servidor e juntado as suas mudanças (merge por registro, no app),
-- mas entre a leitura e a gravação o outro gravou — e `data || pares` troca a
-- chave inteira: a lista de quem chegou por último apagou a OS do outro.
--
-- A CORREÇÃO é a concorrência otimista que faltava: o app manda junto o
-- `updated_at` da versão que ele leu (`esperado`). Com a linha travada, o banco
-- confere:
--   · igual       → ninguém gravou no meio: grava e devolve true;
--   · diferente   → alguém gravou: NÃO grava e devolve false. O app relê,
--                   junta de novo (agora com a mudança do outro) e tenta outra
--                   vez.
-- Assim a junção sempre parte da versão que está de fato no banco, e nenhuma
-- gravação simultânea apaga a outra.
--
-- Linha sem carimbo (updated_at nulo) e linha que ainda não existe gravam
-- direto: não há versão com que comparar.
--
-- Como rodar: igual à gravar_chaves (acima). Enquanto ela não existir, o app
-- avisa no console e grava pela gravar_chaves, sem a conferência — como antes.
-- =====================================================================
create or replace function public.gravar_chaves_se(
  pares    jsonb,
  quem     uuid,
  carimbo  timestamptz,
  esperado timestamptz
)
returns boolean
language plpgsql
security invoker
set search_path = public
as $$
declare
  atual timestamptz;
begin
  if pares is null or jsonb_typeof(pares) <> 'object' then
    raise exception 'gravar_chaves_se: pares tem de ser um objeto jsonb';
  end if;
  select updated_at into atual from public.shared_data where id = 'main' for update;
  if not found then
    insert into public.shared_data (id, data, updated_at, updated_by)
    values ('main', pares, coalesce(carimbo, now()), quem);
    return true;
  end if;
  if atual is not null and (esperado is null or atual <> esperado) then
    return false;
  end if;
  update public.shared_data
     set data       = coalesce(data, '{}'::jsonb) || pares,
         updated_at = coalesce(carimbo, now()),
         updated_by = quem
   where id = 'main';
  return true;
end;
$$;

revoke all on function public.gravar_chaves_se(jsonb, uuid, timestamptz, timestamptz) from public;
grant execute on function public.gravar_chaves_se(jsonb, uuid, timestamptz, timestamptz) to authenticated;

notify pgrst, 'reload schema';

-- Verificação:
-- select proname from pg_proc where proname in ('gravar_chaves', 'gravar_chaves_se');
