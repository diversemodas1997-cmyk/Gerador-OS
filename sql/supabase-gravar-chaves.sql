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
