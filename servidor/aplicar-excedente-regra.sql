-- =====================================================================
-- APLICA A REGRA DO EXCEDENTE DE ENFESTO EM TODAS AS GRADES
-- =====================================================================
-- Este arquivo é a SEGUNDA metade da operação. A primeira é:
--
--   node servidor\aplicar-excedente-regra.js ^
--        --entrada <backup do grades> --saida %TEMP%\grades-novo.json --aplicar
--   docker cp %TEMP%\grades-novo.json supabase-db:/tmp/grades-novo.json
--
-- O script Node é quem aplica a regra — e ele NÃO tem uma cópia dela: recorta
-- `excedenteRegraDaFase` do app.js de verdade. Este SQL só troca a chave
-- `grades` do blob pelo arquivo que ele produziu.
--
-- Rodar (no servidor):
--   docker cp servidor\aplicar-excedente-regra.sql supabase-db:/tmp/aplicar.sql
--   docker exec supabase-db psql -U postgres -d postgres -v ON_ERROR_STOP=1 -f /tmp/aplicar.sql
--
-- SEGURANÇA: tudo numa transação, com duas travas antes de gravar. Qualquer
-- uma que dispare aborta e NADA é escrito.
-- =====================================================================
BEGIN;

-- TRAVA 1: o arquivo novo tem mesmo um array de grades com conteúdo?
-- Grava-se por cima do cadastro inteiro; um arquivo truncado ou vazio aqui
-- apagaria as 128 grades de uma vez.
DO $$
DECLARE n int;
BEGIN
  SELECT jsonb_array_length(pg_read_file('/tmp/grades-novo.json')::jsonb) INTO n;
  IF n IS NULL OR n < 100 THEN
    RAISE EXCEPTION 'o arquivo novo tem % grades (esperado 100+) — abortando', n;
  END IF;
  RAISE NOTICE 'arquivo novo conferido: % grades', n;
END $$;

-- TRAVA 2: o servidor não pode ter MENOS grades do que o arquivo — isso seria
-- sinal de que o arquivo é de outro banco, ou de outro momento.
DO $$
DECLARE atual int; novo int;
BEGIN
  SELECT jsonb_array_length((data->>'grades')::jsonb) INTO atual
    FROM shared_data WHERE id = 'main';
  SELECT jsonb_array_length(pg_read_file('/tmp/grades-novo.json')::jsonb) INTO novo;
  IF novo < atual THEN
    RAISE EXCEPTION 'o arquivo tem % grades e o servidor tem % — abortando', novo, atual;
  END IF;
  RAISE NOTICE 'servidor: % grades  ->  novo: %', atual, novo;
END $$;

-- A troca. `_device` vira 'manutencao' de propósito: é por ele que as abas
-- abertas reconhecem que a gravação veio de FORA e recarregam a tela. Deixando
-- o device de quem gravou por último, a máquina dele acharia que a mudança é
-- dela mesma e não se atualizaria.
UPDATE shared_data
   SET data = jsonb_set(
                jsonb_set(data, '{grades}', to_jsonb(pg_read_file('/tmp/grades-novo.json'))),
                '{_device}', to_jsonb('manutencao'::text)),
       updated_at = now()
 WHERE id = 'main';

-- O aviso para as outras máquinas, dizendo que SÓ `grades` mudou: assim elas
-- baixam essa chave e não o blob inteiro (ver a leitura parcial no app.js).
UPDATE sync_signal
   SET updated_at = now(),
       device_id = 'manutencao',
       key_versions = coalesce(key_versions, '{}'::jsonb)
                      || jsonb_build_object('grades',
                           to_char(now() AT TIME ZONE 'UTC', 'YYYY-MM-DD"T"HH24:MI:SS.MS"Z"'))
 WHERE id = 'main';

COMMIT;

-- Conferência: deve sobrar pouca coisa sem excedente, e as fases de viés
-- devem estar todas em 0.
SELECT coalesce(nullif(fa->>'excedente',''), '(vazio)') AS excedente, count(*)
  FROM shared_data, jsonb_array_elements((data->>'grades')::jsonb) gr,
       jsonb_array_elements(coalesce(gr->'fases','[]'::jsonb)) fa
 WHERE id = 'main'
 GROUP BY 1 ORDER BY 2 DESC;
